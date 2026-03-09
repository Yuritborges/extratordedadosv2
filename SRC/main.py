"""
main.py  —  COFRE BRASUL v9.3  (VERSÃO DEFINITIVA)
===================================================
Projeto : Extrator de Insumos FDE
Autor   : Iury  |  Brasul Construtora

O QUE ESSE CÓDIGO FAZ:
  Lê PDFs de atestados FDE, detecta páginas de medição e quantitativo,
  faz OCR nos códigos de serviço, cruza com a Base Mestra e salva tudo
  num Excel organizado e colorido: Cofre_Brasul.xlsx.

COMO INSTALAMOS AS DEPENDÊNCIAS:
  pip install pytesseract pillow pdfplumber python-Levenshtein openpyxl
  Tesseract: https://github.com/UB-Mannheim/tesseract/wiki

ESTRATÉGIA DE CAPTURA (o que aprendemos testando os PDFs reais):
  - Acumulado de Medição     → coluna de código em x = 3~22% da largura
  - Continuação do Acumulado → mesmo layout, cabeçalho diferente
  - Quantitativa padrão      → coluna de código em x = 11~23%
  - Quantitativa contrato    → código em pipes, OCR na metade esquerda

CORREÇÕES ACUMULADAS ATÉ ESSA VERSÃO:
  v7 → coluna correta por tipo de página + regex tolerante
  v8 → detecção de páginas de continuação + grupos 15-18 + formato contrato
  v9.0 → nome da obra corrigido (ESCOLA: antes de tudo), MODO_PASTA=True, rglob
  v9.1 → candidatos de 4 dígitos: retorna 1 candidato (não 2), consultando a Base
          isso eliminou falsos positivos como 02.02.090, 02.02.907, 02.03.041
  v9.2 → Base Mestra expandida de 244 → 3160 itens (tabela FDE completa Abr/2022)
  v9.3 → filtro terceiro grupo > 510 elimina lixo OCR (datas, valores monetários)
          dicionário _GRUPOS_GLOBAIS para itens .000 (grupos/subgrupos de totalização)
          filtro de grupo ajustado para 01-16 (FDE só vai até grupo 16)
          preservação de preenchimentos manuais entre rodadas
"""

import re
import sys
import unicodedata
from pathlib import Path

# ─────────────────────────────────────────────────────────
#  CONFIGURAÇÃO  —  só mexa aqui
# ─────────────────────────────────────────────────────────

# pasta raiz do projeto
BASE_DIR = Path(r"C:\Users\Iury\Documents\PROJETO EXTRATOR DE DADOS VERSÃO 2")

# True  → processa TODOS os PDFs da pasta input (e subpastas) de uma vez
# False → processa só o CAMINHO_PDF abaixo (bom pra testes)
MODO_PASTA = True

CAMINHO_PDF  = BASE_DIR / "DATA" / "input" / "Ana Luiza Florence Borges I.pdf"
PASTA_INPUT  = BASE_DIR / "DATA" / "input"
PASTA_OUTPUT = BASE_DIR / "DATA" / "output"
CAMINHO_BASE = BASE_DIR / "Base_Mestra_FDE.xlsx"
NOME_SAIDA   = "Cofre_Brasul.xlsx"

DPI_DETECT    = 200    # DPI rápido pra detectar tipo de página
DPI_OCR       = 400    # DPI alto pra ler os códigos com mais precisão
CONTRASTE     = 2.5    # contraste aplicado na imagem antes do OCR

TESSERACT_CMD = r"C:\Program Files\Tesseract-OCR\tesseract.exe"

# ─────────────────────────────────────────────────────────

try:
    import pytesseract
    pytesseract.pytesseract.tesseract_cmd = TESSERACT_CMD
    from PIL import Image, ImageEnhance
    import pdfplumber
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
except ImportError as e:
    sys.exit(
        f"\nFaltando dependência: {e}\n"
        f"Rode: pip install pytesseract pillow pdfplumber openpyxl\n"
    )

# Levenshtein mede similaridade entre strings — usamos no fuzzy de descrição
try:
    import Levenshtein as _lev
    def _ratio(a, b): return _lev.ratio(a, b)
except ImportError:
    # versão manual caso não esteja instalado
    def _ratio(s1, s2):
        if s1 == s2: return 1.0
        l1, l2 = len(s1), len(s2)
        if not l1 or not l2: return 0.0
        dp = list(range(l2 + 1))
        for i in range(1, l1 + 1):
            prev, dp[0] = dp[0], i
            for j in range(1, l2 + 1):
                tmp   = dp[j]
                dp[j] = prev if s1[i-1] == s2[j-1] else 1 + min(prev, dp[j], dp[j-1])
                prev  = tmp
        return 1.0 - dp[l2] / max(l1, l2)


# ══════════════════════════════════════════════════════════
#  1.  BASE MESTRA
#  Carrega o Excel com os itens FDE pra cruzar com os PDFs
# ══════════════════════════════════════════════════════════

def carregar_base(path: Path):
    """
    Lê o Base_Mestra_FDE_EXPANDIDA.xlsx e monta dois índices:
      por_cod  → cod7 (7 dígitos sem pontos)  →  item completo
      por_desc → [(desc_normalizada, item)]   →  pra busca fuzzy por descrição
    """
    if not path.exists():
        sys.exit(f"\nBase Mestra não encontrada:\n  {path}\n")

    wb = openpyxl.load_workbook(str(path), data_only=True)
    ws = wb.active
    por_cod  = {}
    por_desc = []
    pula_header = True

    for row in ws.iter_rows(values_only=True):
        if pula_header:
            pula_header = False
            continue
        cod  = str(row[0]).strip() if row[0] else ''
        desc = str(row[1]).strip() if len(row) > 1 and row[1] else ''
        un   = str(row[2]).strip() if len(row) > 2 and row[2] else ''
        cod7 = re.sub(r'\D', '', cod)
        if len(cod7) != 7:
            continue
        item = {'codigo': cod, 'descricao': desc, 'unidade': un}
        por_cod[cod7] = item
        por_desc.append((_norm(desc), item))

    print(f"  Base Mestra: {len(por_cod)} itens carregados")
    return por_cod, por_desc


# Dicionário de grupos/subgrupos globais (.000)
# Esses itens aparecem nas planilhas como linhas de totalização de grupo.
# A tabela FDE não os lista individualmente — são descrições genéricas.
_GRUPOS_GLOBAIS = {
    '0100000': ('SERVIÇOS GERAIS', '%'),
    '0200000': ('FUNDAÇÕES', '%'),
    '0300000': ('ESTRUTURA', '%'),
    '0400000': ('DEMOLIÇÕES', '%'),
    '0500000': ('ESQUADRIAS DE MADEIRA', '%'),
    '0600000': ('ESQUADRIAS METÁLICAS', '%'),
    '0700000': ('COBERTURA', '%'),
    '0800000': ('INST. HIDRÁULICAS E SANITÁRIAS', '%'),
    '0900000': ('INST. ELÉTRICAS E LÓGICA', '%'),
    '1000000': ('FORROS', '%'),
    '1100000': ('IMPERMEABILIZAÇÕES', '%'),
    '1200000': ('REVESTIMENTOS', '%'),
    '1300000': ('PISOS', '%'),
    '1400000': ('VIDROS', '%'),
    '1500000': ('PINTURAS', '%'),
    '1600000': ('SERVIÇOS COMPLEMENTARES', '%'),
    # subgrupos frequentes
    '0102000': ('SERVIÇOS GERAIS - SERVIÇOS INICIAIS', '%'),
    '0202000': ('FUNDAÇÕES - ESTACAS', '%'),
    '0203000': ('FUNDAÇÕES - FORMAS', '%'),
    '0204000': ('FUNDAÇÕES - ARMADURAS', '%'),
    '0205000': ('FUNDAÇÕES - CONCRETO', '%'),
    '0302000': ('ESTRUTURA - SUPER ESTRUTURA', '%'),
    '0304000': ('ESTRUTURA - ESTRUTURA METÁLICA', '%'),
    '0307000': ('ESTRUTURA - MADEIRAMENTO', '%'),
    '0350000': ('ESTRUTURA - DEMOLIÇÕES', '%'),
    '0502000': ('ESQUADRIAS - PORTAS INTERNAS', '%'),
    '0560000': ('ESQUADRIAS - RETIRADAS DE MADEIRA', '%'),
    '0580000': ('ESQUADRIAS - FERRAGENS', '%'),
    '0601000': ('ESQ. METÁLICAS - PORTÕES E GRADES', '%'),
    '0603000': ('ESQ. METÁLICAS - GUARDA-CORPOS E CORRIMÃOS', '%'),
    '0660000': ('ESQ. METÁLICAS - RETIRADAS', '%'),
    '0702000': ('COBERTURA - ESTRUTURA METÁLICA', '%'),
    '0703000': ('COBERTURA - TELHAS', '%'),
    '0704000': ('COBERTURA - CUMEEIRAS E RUFOS', '%'),
    '0705000': ('COBERTURA - FECHAMENTOS E VEDAÇÕES', '%'),
    '0760000': ('COBERTURA - RETIRADAS DE ESTRUTURA', '%'),
    '0770000': ('COBERTURA - RETIRADAS DE TELHAS', '%'),
    '0780000': ('COBERTURA - MADEIRAMENTO', '%'),
    '0800000': ('INST. HIDRÁULICAS E SANITÁRIAS', '%'),
    '0807000': ('INST. HIDRÁULICAS - ÁGUA FRIA', '%'),
    '0808000': ('INST. HIDRÁULICAS - ESGOTO', '%'),
    '0811000': ('INST. HIDRÁULICAS - ÁGUAS PLUVIAIS TUBULAÇÕES', '%'),
    '0812000': ('INST. HIDRÁULICAS - ÁGUAS PLUVIAIS CALHAS', '%'),
    '0813000': ('INST. HIDRÁULICAS - TUBULAÇÕES GERAIS', '%'),
    '0814000': ('INST. HIDRÁULICAS - REGISTROS E VÁLVULAS', '%'),
    '0816000': ('INST. HIDRÁULICAS - APARELHOS SANITÁRIOS', '%'),
    '0819000': ('INST. HIDRÁULICAS - RETIRADAS', '%'),
    '0850000': ('INST. HIDRÁULICAS - DEMOLIÇÕES', '%'),
    '0860000': ('INST. HIDRÁULICAS - RETIRADAS DIVERSAS', '%'),
    '0900000': ('INST. ELÉTRICAS E LÓGICA', '%'),
    '0902000': ('INST. ELÉTRICAS - ENTRADA DE ENERGIA', '%'),
    '0905000': ('INST. ELÉTRICAS - ELETRODUTOS', '%'),
    '0907000': ('INST. ELÉTRICAS - FIOS E CABOS', '%'),
    '0908000': ('INST. ELÉTRICAS - PONTOS ELÉTRICOS', '%'),
    '0909000': ('INST. ELÉTRICAS - ILUMINAÇÃO', '%'),
    '0912000': ('INST. ELÉTRICAS - EXAUSTÃO', '%'),
    '0913000': ('INST. ELÉTRICAS - SPDA/ATERRAMENTO', '%'),
    '0919000': ('INST. ELÉTRICAS - RETIRADAS', '%'),
    '0960000': ('INST. ELÉTRICAS - RETIRADAS DIVERSAS', '%'),
    '0964000': ('INST. ELÉTRICAS - RETIRADAS DE APARELHOS', '%'),
    '0974000': ('INST. ELÉTRICAS - RECOLOCAÇÕES', '%'),
    '0982000': ('INST. ELÉTRICAS - POSTES E SUPORTES', '%'),
    '0984000': ('INST. ELÉTRICAS - INTERRUPTORES E TOMADAS', '%'),
    '0985000': ('INST. ELÉTRICAS - INFRAESTRUTURA LÓGICA', '%'),
    '1050000': ('FORROS - DEMOLIÇÕES', '%'),
    '1150000': ('IMPERMEABILIZAÇÕES - DEMOLIÇÕES', '%'),
    '1202000': ('REVESTIMENTOS - ARGAMASSAS', '%'),
    '1207000': ('REVESTIMENTOS - PASTILHAS', '%'),
    '1302000': ('PISOS - CIMENTADOS', '%'),
    '1305000': ('PISOS - CERÂMICOS', '%'),
    '1306000': ('PISOS - SOLEIRAS', '%'),
    '1380000': ('PISOS - DEMOLIÇÕES', '%'),
    '1501000': ('PINTURAS - ESTRUTURA METÁLICA', '%'),
    '1502000': ('PINTURAS - PAREDES E TETOS', '%'),
    '1503000': ('PINTURAS - ESQUADRIAS METÁLICAS', '%'),
    '1550000': ('PINTURAS - REMOÇÕES', '%'),
    '1603000': ('SERV. COMPL. - JARDINAGEM', '%'),
    '1606000': ('SERV. COMPL. - INSTALAÇÕES PROVISÓRIAS', '%'),
    '1608000': ('SERV. COMPL. - SINALIZAÇÃO', '%'),
    '1650000': ('SERV. COMPL. - DEMOLIÇÕES', '%'),
    '1680000': ('SERV. COMPL. - SERVIÇOS FINAIS', '%'),
}


def _norm(txt: str) -> str:
    """Remove acentos, sobe pra maiúsculo, tira especiais — pra comparação."""
    txt = unicodedata.normalize('NFKD', txt)
    txt = ''.join(c for c in txt if not unicodedata.combining(c))
    return re.sub(r'\s+', ' ', re.sub(r'[^A-Z0-9 ]', ' ', txt.upper())).strip()


# ══════════════════════════════════════════════════════════
#  2.  MATCH COM A BASE MESTRA
#  Cruza o código OCR com a Base pra buscar descrição e UN
# ══════════════════════════════════════════════════════════

def match_item(cod7: str, por_cod: dict, por_desc: list) -> dict:
    """
    Busca o código na Base. Sempre retorna um item:
      → achou na base  : retorna com código, descrição e unidade corretos
      → não achou      : retorna stub com o código OCR preservado (desc vazia)

    IMPORTANTE: fuzzy de código foi removido porque substituía silenciosamente
    códigos corretos. Ex: OCR capturava 02.02.097 → fuzzy trocava por 02.02.098.
    Agora o código capturado é sempre preservado.
    """
    if cod7 in por_cod:
        return por_cod[cod7]
    # não achou na base principal — tenta nos grupos globais .000
    if cod7 in _GRUPOS_GLOBAIS:
        desc, un = _GRUPOS_GLOBAIS[cod7]
        cod_fmt = f"{cod7[:2]}.{cod7[2:4]}.{cod7[4:]}"
        return {'codigo': cod_fmt, 'descricao': desc, 'unidade': un}
    # genuinamente ausente: preserva o código OCR para preenchimento manual posterior
    cod_fmt = f"{cod7[:2]}.{cod7[2:4]}.{cod7[4:]}"
    return {'codigo': cod_fmt, 'descricao': '', 'unidade': ''}


# ══════════════════════════════════════════════════════════
#  3.  REGEX E EXTRAÇÃO DE CÓDIGOS
#  O coração do extrator — captura os códigos XX.XX.XXX
# ══════════════════════════════════════════════════════════

# Regex ultra-tolerante: aceita erros comuns de OCR
#   O/o/Q/q/A/a → 0   (zero confundido com letra arredondada)
#   I/l         → 1   (um confundido com maiúscula/minúscula)
#   separadores: ponto, vírgula, pipe, espaço, traço — até 2 chars entre grupos
#   terceiro grupo: 2, 3 ou 4 dígitos (OCR às vezes insere dígito extra)
_RE_TOL = re.compile(
    r'(?<![0-9A-Za-z])'
    r'([0-2OoAaQq9][0-9OoQq])'    # grupo 1: dois chars (ex: 02, 09, 16)
    r'[.,_|\-\s\\]{0,2}'           # separador flexível
    r'([0-9OoQq]{2})'              # grupo 2: dois chars
    r'[.,_|\-\s\\]{0,2}'
    r'([0-9OoQq]{2,4})'            # grupo 3: 2, 3 ou 4 chars
    r'(?![0-9A-Za-z])'
)

# corrige letras que o OCR confunde com dígitos
_OCR_CORR = str.maketrans('OoQqIlAa', '00000100')


def _norm_grupo(a: str, b: str, c: str) -> str:
    """
    Converte os 3 grupos do regex pra formato XX.XX.XXX.
    Letras → dígitos, grupo curto → completa com zero, grupo longo → corta.
    """
    a2 = a.translate(_OCR_CORR)
    b2 = b.translate(_OCR_CORR)
    c2 = c.translate(_OCR_CORR)
    if   len(c2) == 2: c2 += '0'    # 07.02.00  → 07.02.000
    elif len(c2) == 4: c2  = c2[:3] # fallback antes de chamar _melhor_candidato
    return f"{a2}.{b2}.{c2}"


def _melhor_candidato_4dig(a2: str, b2: str, c4: str, por_cod: dict):
    """
    Quando o OCR produz 4 dígitos no terceiro grupo, um dígito foi inserido errado.
    Testamos remover 1 dígito de cada uma das 4 posições:
      '0907' → pos0='907'  pos1='007'  pos2='097'  pos3='090'

    Testamos empiricamente com os PDFs reais:
      0907 → pos2 = '097' ✓   (02.02.097)
      0041 → pos2 = '001' ✓   (02.03.001  Forma de Madeira Maciça)
      0278 → pos2 = '028' ✓   (02.05.028)
      0090 → pos2 = '000' ✓   (07.05.000)
      0002 → pos0 = '002' ✓   (02.04.002)

    Retorna APENAS UM código — o melhor candidato.
    Isso elimina os falsos positivos 02.02.090, 02.02.907, 02.03.041 que apareciam
    quando incluíamos todos os candidatos de uma vez.

    Decisão:
      1. Se algum candidato está na Base → usa esse (dado concreto)
      2. Senão → usa posição 2 (c[:2]+c[3:]) que acertou em 4/5 casos testados
    """
    cands = [
        a2 + b2 + c4[1:],            # pos 0
        a2 + b2 + c4[0] + c4[2:],    # pos 1
        a2 + b2 + c4[:2] + c4[3:],   # pos 2  ← melhor empiricamente
        a2 + b2 + c4[:3],            # pos 3
    ]
    validos = [c for c in cands
               if len(c) == 7 and c.isdigit() and 1 <= int(c[:2]) <= 16]
    if not validos:
        return None
    # prefere o que está na base
    for c in validos:
        if c in por_cod:
            return c
    # padrão: posição 2
    pos2 = a2 + b2 + c4[:2] + c4[3:]
    return pos2 if pos2 in validos else validos[0]


def _extrair_codigos(texto: str, por_cod: dict) -> list:
    """
    Aplica o regex no texto OCR e retorna lista de cod7 únicos e válidos.
    Grupos FDE válidos: 01 a 20.

    Para 4 dígitos no terceiro grupo: usa _melhor_candidato_4dig que consulta
    a Base Mestra e retorna exatamente 1 código (sem falsos positivos duplos).
    """
    resultado = []
    vistos    = set()

    for a, b, c in _RE_TOL.findall(texto):
        a2 = a.translate(_OCR_CORR)
        b2 = b.translate(_OCR_CORR)
        c2 = c.translate(_OCR_CORR)

        if len(c2) == 4:
            # 4 dígitos: escolhe o melhor candidato consultando a base
            cod7 = _melhor_candidato_4dig(a2, b2, c2, por_cod)
        elif len(c2) == 2:
            # 2 dígitos: completa com zero
            raw = a2 + b2 + c2 + '0'
            cod7 = raw if (len(raw) == 7 and raw.isdigit()) else None
        else:
            # 3 dígitos: caminho normal
            raw = a2 + b2 + c2
            cod7 = raw if (len(raw) == 7 and raw.isdigit()) else None

        if not cod7 or len(cod7) != 7:
            continue
        if not (1 <= int(cod7[:2]) <= 16):
            continue
        # terceiro grupo > 510 é lixo OCR (datas, valores, numeração de contrato)
        # maior código legítimo na tabela FDE é 16.03.510
        if int(cod7[4:]) > 510:
            continue
        if cod7 not in vistos:
            vistos.add(cod7)
            resultado.append(cod7)

    return resultado


# ══════════════════════════════════════════════════════════
#  4.  UTILITÁRIOS DE IMAGEM E OCR
# ══════════════════════════════════════════════════════════

def _prep_img(page, dpi: int):
    """Converte página em imagem cinza com contraste alto. Melhora muito o OCR."""
    img = page.to_image(resolution=dpi).original.convert('L')
    return ImageEnhance.Contrast(img).enhance(CONTRASTE)


def _ocr(img_pil, config='--psm 6 -l por+eng') -> str:
    """Roda o Tesseract na imagem e devolve o texto."""
    return pytesseract.image_to_string(img_pil, config=config)


# ══════════════════════════════════════════════════════════
#  5.  DETECÇÃO DO TIPO DE PÁGINA
#  Lê o cabeçalho pra saber que tipo de planilha é
# ══════════════════════════════════════════════════════════

def detectar_tipo(page) -> str:
    """
    OCR rápido no cabeçalho (top 32%) e classifica:
      'ACUMULADO'      → Planilha de Acumulado de Medição (pág inicial)
      'ACUMULADO'      → Continuação do Acumulado (sem a palavra no topo)
      'QUANTITATIVA'   → Planilha Quantitativa padrão
      'QUANT_CONTRATO' → Quantitativa formato contrato (códigos em pipes)
      ''               → Ignora (capa, assinaturas, texto corrido)
    """
    img_top = _prep_img(page, DPI_DETECT)
    W, H    = img_top.size
    img_cab = img_top.crop((0, int(H * 0.03), W, int(H * 0.32)))
    txt     = _ocr(img_cab, '--psm 3').upper()

    # página inicial do Acumulado de Medição
    if re.search(r'ACUMULADO', txt) and re.search(r'MEDI|CRITER|CONTRAT', txt):
        return 'ACUMULADO'

    # páginas de continuação do Acumulado
    # (repetem o cabeçalho de colunas: CRITERIO + UNITARIA + CONTRATO)
    if (re.search(r'CRIT', txt) and
            re.search(r'UNITARI|UNITATIA|GLOBAL', txt) and
            re.search(r'CONTRAT|DATA|NRO|MEDI|DT', txt)):
        return 'ACUMULADO'

    # Planilha Quantitativa padrão
    if re.search(r'QUANTITATIV', txt):
        return 'QUANTITATIVA'
    if re.search(r'DESCRI.{1,5}O\s+ATIVIDADE|CODIGO\s+ATIVIDADE', txt):
        return 'QUANTITATIVA'

    # Quantitativa no formato CONTRATO (ex: Ana Luiza pág 6)
    if (re.search(r'CONTRATO\s*:', txt) and
            re.search(r'COD\.?\s*OBRA|PREDIO|\bPI\b', txt)):
        return 'QUANT_CONTRATO'

    return ''


# ══════════════════════════════════════════════════════════
#  6.  EXTRAÇÃO DE CÓDIGOS POR TIPO DE PÁGINA
# ══════════════════════════════════════════════════════════

def _ocr_coluna(page, x0: float, x1: float, y0: float, y1: float,
                por_cod: dict) -> list:
    """
    OCR numa faixa vertical da página (onde ficam os códigos).
    Upscale 2x + sharpen antes do Tesseract — fez muita diferença
    em colunas estreitas de ~5% da largura.
    """
    img  = _prep_img(page, DPI_OCR)
    W, H = img.size
    faixa = img.crop((int(W * x0), int(H * y0), int(W * x1), int(H * y1)))
    faixa = faixa.resize((faixa.width * 2, faixa.height * 2), Image.LANCZOS)
    faixa = ImageEnhance.Sharpness(faixa).enhance(2.0)
    return _extrair_codigos(_ocr(faixa), por_cod)


def _ocr_meia_pagina(page, por_cod: dict) -> list:
    """
    OCR na metade esquerda da página inteira.
    Usado no QUANT_CONTRATO onde os códigos ficam em | pipes | sem coluna fixa.
    """
    img  = _prep_img(page, 300)
    W, H = img.size
    metade = img.crop((0, int(H * 0.12), int(W * 0.55), int(H * 0.97)))
    return _extrair_codigos(_ocr(metade), por_cod)


def processar_pagina(page, tipo: str, por_cod: dict, por_desc: list) -> list:
    """
    Extrai itens de uma página usando o método correto pra cada tipo:
      ACUMULADO      → coluna x=3~22%
      QUANTITATIVA   → coluna x=11~23%
      QUANT_CONTRATO → metade esquerda da página
    """
    if tipo == 'ACUMULADO':
        codigos = _ocr_coluna(page, 0.03, 0.22, 0.22, 0.95, por_cod)
    elif tipo == 'QUANTITATIVA':
        codigos = _ocr_coluna(page, 0.11, 0.23, 0.20, 0.95, por_cod)
    else:  # QUANT_CONTRATO
        codigos = _ocr_meia_pagina(page, por_cod)

    itens = []
    for cod7 in codigos:
        item = match_item(cod7, por_cod, por_desc)
        if item:
            itens.append(item)
    return itens


# ══════════════════════════════════════════════════════════
#  7.  EXTRAÇÃO DO NOME DA OBRA
#  Procura o nome da escola/obra no cabeçalho do PDF
# ══════════════════════════════════════════════════════════

def extrair_nome(page) -> str:
    """
    Lê o cabeçalho e procura o nome da obra/escola.
    Padrões em ordem de prioridade (mais confiável primeiro):
      ESCOLA: 62110 - EE PROFA ANA LUIZA...
      62110 - EE PROFA ANA LUIZA...
      NOME INTERV. EE PROFA ANA LUIZA...
      PRÉDIO: 62110 - NOME...

    Retorna title case: "62110 - Ee Profa Ana Luiza Florence Borges"
    """
    img = _prep_img(page, DPI_DETECT)
    txt = _ocr(img, '--psm 3 -l por+eng').upper()

    padroes = [
        # ESCOLA: CÓDIGO - NOME (mais confiável — é um campo rotulado)
        r'ESCOLA\s*[:\|]?\s*(\d{5,6}\s*[-–]\s*[A-Z][\w\s\.\-]+)',
        # CÓDIGO - tipo de escola + NOME
        r'(\d{5,6}\s*[-–]\s*(?:EE|EM|EMEF|EMEFM|ETEC|CEI|CIEJA)\s+[A-Z][\w\s\.\-]{5,60})',
        # campo "Nome Interv." das planilhas de medição
        r'NOME\s+INTERV\.?\s*[:\|]?\s*([A-Z][\w\s\.\-]{5,70})',
        # campo "Prédio: CÓDIGO - NOME"
        r'PR[EÉ]DIO\s*[:\|]?\s*\d{5,6}\s*[-–]\s*([A-Z][\w\s\.\-]{5,60})',
    ]

    for pat in padroes:
        m = re.search(pat, txt)
        if m:
            nome = m.group(1).strip()
            # corta no primeiro separador forte
            nome = re.split(r'\s{3,}|\||\n|CONTRATO|FISCAL|PI\s*:|DIRETORIA', nome)[0].strip()
            nome = re.sub(r'[\s\.\,\-]+$', '', nome)
            if 6 < len(nome) < 90:
                return nome.title()

    return 'OBRA_DESCONHECIDA'


# ══════════════════════════════════════════════════════════
#  8.  PROCESSA UM PDF COMPLETO
#  Passa por todas as páginas e junta tudo
# ══════════════════════════════════════════════════════════

def processar_pdf(caminho: Path, por_cod: dict, por_desc: list,
                  ja_processados=None) -> dict:
    """
    Abre o PDF, processa cada página e devolve:
      { 'nome': nome da obra, 'arq': nome do arquivo, 'itens': lista de itens }

    Se um item aparece nos dois tipos de planilha → tipo vira 'AMBOS'.
    """
    print(f"\n  >> {caminho.name}")

    if not caminho.exists():
        print(f"     arquivo não encontrado, pulando...")
        return None

    if ja_processados is not None and caminho.name in ja_processados:
        print(f"     esse PDF já foi processado antes, pulando...")
        return None

    itens_totais = {}        # cod7 → item (deduplicado automaticamente)
    nome_obra    = 'OBRA_DESCONHECIDA'

    with pdfplumber.open(str(caminho)) as pdf:
        n = len(pdf.pages)
        print(f"     {n} páginas no total")

        for num, page in enumerate(pdf.pages, 1):
            try:
                tipo = detectar_tipo(page)

                if not tipo:
                    print(f"     pág {num}/{n}: ignorando")
                    continue

                labels = {
                    'ACUMULADO':      'Acumulado de Medição',
                    'QUANTITATIVA':   'Quantitativa',
                    'QUANT_CONTRATO': 'Quantitativa (formato contrato)',
                }
                print(f"     pág {num}/{n}: {labels[tipo]} ...", end=' ', flush=True)

                itens = processar_pagina(page, tipo, por_cod, por_desc)
                print(f"capturei {len(itens)} itens")

                # extrai o nome da obra na primeira página relevante
                if nome_obra == 'OBRA_DESCONHECIDA':
                    nome_obra = extrair_nome(page)

                # acumula os itens, marca tipo e deduplicação
                tipo_saida = 'QUANTITATIVA' if tipo == 'QUANT_CONTRATO' else tipo
                for it in itens:
                    cod7 = re.sub(r'\D', '', it['codigo'])
                    if not cod7:
                        continue
                    if cod7 not in itens_totais:
                        itens_totais[cod7] = {**it, 'tipo': tipo_saida}
                    else:
                        # aparece nos dois tipos → AMBOS
                        t_ant = itens_totais[cod7]['tipo']
                        if t_ant != tipo_saida and 'AMBOS' not in t_ant:
                            itens_totais[cod7]['tipo'] = 'AMBOS'

            except Exception as e:
                print(f"\n     pág {num}: deu ruim → {e}")
                import traceback
                traceback.print_exc()

    total = len(itens_totais)
    print(f"\n     resultado: {nome_obra}  |  {total} itens únicos")
    return {
        'nome':  nome_obra,
        'arq':   caminho.name,
        'itens': list(itens_totais.values()),
    }


# ══════════════════════════════════════════════════════════
#  9.  GERA O COFRE_BRASUL.XLSX
# ══════════════════════════════════════════════════════════

_COR_CAB  = PatternFill("solid", fgColor="1F4E79")
_COR_ACUM = PatternFill("solid", fgColor="DEEAF1")
_COR_QUAN = PatternFill("solid", fgColor="E2EFDA")
_COR_BOTH = PatternFill("solid", fgColor="FFF2CC")
_COR_PAR  = PatternFill("solid", fgColor="F5F5F5")
_FONT_CAB  = Font(bold=True, color="FFFFFF", size=10)
_FONT_NORM = Font(size=9)
_AL_C = Alignment(horizontal="center", vertical="center")
_AL_E = Alignment(horizontal="left",   vertical="center", wrap_text=True)
_BRD  = Border(
    left   = Side(style="thin", color="CCCCCC"),
    right  = Side(style="thin", color="CCCCCC"),
    top    = Side(style="thin", color="CCCCCC"),
    bottom = Side(style="thin", color="CCCCCC"),
)


def _carregar_descricoes_manuais(pasta: Path) -> dict:
    """
    Lê o Cofre anterior (se existir) e salva todas as descrições que
    foram preenchidas manualmente pelo usuário.
    Retorna dicionário: cod7 → {'descricao': ..., 'unidade': ...}

    Isso garante que NENHUM preenchimento manual seja perdido entre rodadas.
    O programa sempre preserva o que o usuário escreveu.
    """
    caminho = pasta / NOME_SAIDA
    if not caminho.exists():
        return {}

    manuais = {}
    try:
        wb = openpyxl.load_workbook(str(caminho), data_only=True)
        ws = wb.active
        for row in ws.iter_rows(min_row=2, values_only=True):
            if not row or len(row) < 6:
                continue
            cod_raw = str(row[3]).strip() if row[3] else ''
            desc    = str(row[4]).strip() if row[4] else ''
            un      = str(row[5]).strip() if row[5] else ''
            cod7    = re.sub(r'\D', '', cod_raw)
            if len(cod7) == 7 and desc:
                manuais[cod7] = {'descricao': desc, 'unidade': un}
        print(f"  preservando {len(manuais)} descrições do Cofre anterior")
    except Exception as e:
        print(f"  aviso: não consegui ler Cofre anterior ({e})")
    return manuais


def gerar_excel(resultados: list, pasta: Path) -> Path:
    """
    Salva o Cofre_Brasul.xlsx.

    PROTEÇÃO DE PREENCHIMENTO MANUAL:
    Antes de gerar, lê o Cofre anterior e salva todas as descrições que
    o usuário preencheu manualmente. Se um item sair sem descrição da Base
    mas tiver descrição no Cofre anterior → usa a descrição manual.
    Assim o trabalho manual NUNCA é perdido entre rodadas.
    """
    pasta.mkdir(parents=True, exist_ok=True)
    caminho = pasta / NOME_SAIDA

    # salva preenchimentos manuais do Cofre anterior ANTES de sobrescrever
    manuais = _carregar_descricoes_manuais(pasta)

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Cofre_Brasul"

    # cabeçalho
    for cell, txt in zip(ws[1], ["Obra", "Obra_Arq", "Tipo", "Cod", "Desc", "UN"]):
        cell.value     = txt
        cell.fill      = _COR_CAB
        cell.font      = _FONT_CAB
        cell.alignment = _AL_C
    ws.row_dimensions[1].height = 22

    # dados
    linha_n = 2
    for r in resultados:
        if not r:
            continue
        for it in sorted(r['itens'], key=lambda x: x.get('codigo', '')):
            tipo = it.get('tipo', '')
            desc = it.get('descricao', '')
            un   = it.get('unidade', '')
            cod7 = re.sub(r'\D', '', it.get('codigo', ''))

            # --- PRESERVAÇÃO MANUAL ---
            # Se a Base não tem descrição mas o usuário preencheu antes → usa a manual
            if not desc and cod7 in manuais:
                desc = manuais[cod7]['descricao']
                un   = un or manuais[cod7]['unidade']
            # --------------------------

            if 'AMBOS' in tipo:           fill = _COR_BOTH
            elif 'QUANT' in tipo.upper(): fill = _COR_QUAN
            elif 'ACUM'  in tipo.upper(): fill = _COR_ACUM
            else:                         fill = _COR_PAR

            ws.append([r['nome'], r['arq'], tipo,
                       it.get('codigo',''), desc, un])
            for col in range(1, 7):
                c = ws.cell(linha_n, col)
                c.fill = fill; c.font = _FONT_NORM; c.border = _BRD; c.alignment = _AL_E
            ws.cell(linha_n, 3).alignment = _AL_C
            ws.cell(linha_n, 4).alignment = _AL_C
            ws.cell(linha_n, 6).alignment = _AL_C
            linha_n += 1

    ws.column_dimensions['A'].width = 42
    ws.column_dimensions['B'].width = 36
    ws.column_dimensions['C'].width = 28
    ws.column_dimensions['D'].width = 14
    ws.column_dimensions['E'].width = 72
    ws.column_dimensions['F'].width =  8
    ws.freeze_panes = "A2"

    # legenda
    ws.append([])
    ws.append(["LEGENDA DE CORES:"])
    ws.cell(ws.max_row, 1).font = Font(bold=True, size=9)
    for fill, txt in [
        (_COR_ACUM, "Azul    — ACUMULADO DE MEDIÇÃO"),
        (_COR_QUAN, "Verde   — PLANILHA DE QUANTITATIVOS"),
        (_COR_BOTH, "Amarelo — Aparece nas DUAS tabelas"),
    ]:
        ws.append(["", txt])
        ws.cell(ws.max_row, 1).fill   = fill
        ws.cell(ws.max_row, 1).border = _BRD
        ws.cell(ws.max_row, 2).font   = Font(italic=True, size=9)

    wb.save(str(caminho))
    return caminho

def main():
    print()
    print("=" * 60)
    print("  COFRE BRASUL — Extrator FDE v9.1")
    print("=" * 60)

    por_cod, por_desc = carregar_base(CAMINHO_BASE)

    if not MODO_PASTA:
        # modo de teste: processa só um PDF
        print(f"\nmodo de teste — processando só {CAMINHO_PDF.name}\n")
        resultado = processar_pdf(CAMINHO_PDF, por_cod, por_desc)
        if resultado:
            caminho_xls = gerar_excel([resultado], PASTA_OUTPUT)
            print(f"\narquivo salvo: {caminho_xls}")
            print(f"itens capturados: {len(resultado['itens'])}")
    else:
        # modo pasta: processa todos os PDFs de uma vez (incluindo subpastas)
        pdfs = sorted(PASTA_INPUT.rglob("*.pdf"))
        if not pdfs:
            sys.exit(f"\nnão encontrei nenhum PDF em:\n  {PASTA_INPUT}\n")

        print(f"\nmodo pasta — vou processar {len(pdfs)} PDF(s)")
        print(f"pasta: {PASTA_INPUT}\n")

        resultados, erros = [], []
        for i, pdf in enumerate(pdfs, 1):
            print(f"[{i}/{len(pdfs)}]", end='')
            try:
                r = processar_pdf(pdf, por_cod, por_desc)
                if r:
                    resultados.append(r)
            except Exception as e:
                print(f"  deu problema: {e}")
                erros.append(pdf.name)

        caminho_xls = gerar_excel(resultados, PASTA_OUTPUT)
        total_itens = sum(len(r['itens']) for r in resultados)

        print()
        print("=" * 60)
        print(f"  tudo pronto!")
        print(f"  arquivo: {caminho_xls}")
        print(f"  obras processadas: {len(resultados)}")
        print(f"  itens no total: {total_itens}")
        if erros:
            print(f"  PDFs com erro: {erros}")
        print("=" * 60)

    print("\nConcluído! ✓\n")


if __name__ == "__main__":
    main()
