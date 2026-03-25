"""
Configurações centralizadas do projeto.
Todos os caminhos e constantes do sistema.
"""

from pathlib import Path

# ═══════════════════════════════════════════════════════════════════════════
# 1. CAMINHOS DO PROJETO
# ═══════════════════════════════════════════════════════════════════════════

# Detecta a raiz do projeto automaticamente
BASE_DIR = Path(__file__).resolve().parent.parent

# Pastas principais
DATA_DIR = BASE_DIR / "DATA"
INPUT_DIR = DATA_DIR / "input"
OUTPUT_DIR = DATA_DIR / "output"
LOGS_DIR = BASE_DIR / "logs"
ASSETS_DIR = BASE_DIR / "assets"
ICONS_DIR = ASSETS_DIR / "icons"
IMAGES_DIR = ASSETS_DIR / "images"

# ═══════════════════════════════════════════════════════════════════════════
# 2. ARQUIVOS ESPECÍFICOS
# ═══════════════════════════════════════════════════════════════════════════

CAMINHO_BASE = INPUT_DIR / "Base_Mestra_FDE.xlsx"
NOME_SAIDA = "Cofre_Brasul.xlsx"
CAMINHO_COFRE = OUTPUT_DIR / NOME_SAIDA

# ═══════════════════════════════════════════════════════════════════════════
# 3. CONFIGURAÇÕES DO SISTEMA
# ═══════════════════════════════════════════════════════════════════════════

MODO_PASTA = True
PASTA_INPUT = INPUT_DIR
PASTA_OUTPUT = OUTPUT_DIR
CAMINHO_PDF = PASTA_INPUT / "exemplo.pdf"

# ═══════════════════════════════════════════════════════════════════════════
# 4. CONFIGURAÇÕES DE OCR
# ═══════════════════════════════════════════════════════════════════════════

DPI_DETECT = 200
DPI_OCR = 400
CONTRASTE = 2.5

# ═══════════════════════════════════════════════════════════════════════════
# 5. TESSERACT
# ═══════════════════════════════════════════════════════════════════════════

_TESS_DEFAULT = r"C:\Program Files\Tesseract-OCR\tesseract.exe"
TESSERACT_CMD = _TESS_DEFAULT if Path(_TESS_DEFAULT).exists() else "tesseract"

# ═══════════════════════════════════════════════════════════════════════════
# 6. LOGS
# ═══════════════════════════════════════════════════════════════════════════

LOG_LEVEL = "INFO"
LOG_FORMAT = "%(asctime)s - %(name)s - %(levelname)s - %(message)s"
LOG_FILE = LOGS_DIR / "aplicacao.log"

# Garantir que as pastas existem
for folder in [INPUT_DIR, OUTPUT_DIR, LOGS_DIR]:
    folder.mkdir(parents=True, exist_ok=True)