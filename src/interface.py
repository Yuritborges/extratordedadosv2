"""
Sistema de Gestão de Insumos - Brasul Construtora

Interface gráfica para visualização e gerenciamento dos códigos FDE extraídos
dos atestados de obras. Permite filtrar por tipo, buscar por código/descrição,
importar novos PDFs e exportar relatórios.
"""

import os, sys, threading, shutil, unicodedata
import pandas as pd
import customtkinter as ctk
from tkinter import ttk, filedialog, messagebox
from PIL import Image, ImageTk
import ctypes
import sys
from pathlib import Path

# ═══════════════════════════════════════════════════════════════════════════════
# CONFIGURAÇÕES DO PROJETO
# ═══════════════════════════════════════════════════════════════════════════════
# Puxamos as configurações centralizadas para encontrar os assets (ícone e logo)
# de forma consistente, independente de onde o programa está sendo executado.
# ═══════════════════════════════════════════════════════════════════════════════

BASE_DIR = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(BASE_DIR))

from config.settings import ICONS_DIR, IMAGES_DIR

# Caminhos para os arquivos de imagem
icone_path = ICONS_DIR / "iconebrasul2.ico"
logo_path = IMAGES_DIR / "LOGOTIPOBRASUL.png"


# ═══════════════════════════════════════════════════════════════════════════════
# FUNÇÃO AUXILIAR PARA ARQUIVOS EMPACOTADOS
# ═══════════════════════════════════════════════════════════════════════════════
# Quando geramos um .exe com PyInstaller, os arquivos são extraídos para uma
# pasta temporária. Essa função ajuda a localizar os arquivos corretamente
# tanto no modo desenvolvimento quanto no modo executável.
# ═══════════════════════════════════════════════════════════════════════════════

def resource_path(rel):
    try:
        base = sys._MEIPASS
    except Exception:
        base = os.path.abspath(".")
    return os.path.join(base, rel)


# ═══════════════════════════════════════════════════════════════════════════════
# IMPORTAÇÃO DO MÓDULO DE EXTRAÇÃO
# ═══════════════════════════════════════════════════════════════════════════════
# Tenta importar as funções principais do main.py. Se falhar, desabilita a
# funcionalidade de importar PDFs (útil quando a interface está sendo executada
# sozinha ou em modo de teste).
# ═══════════════════════════════════════════════════════════════════════════════

try:
    from main import processar_pdf, carregar_base, gerar_excel
    MAIN_OK = True
except ImportError:
    MAIN_OK = False

# ═══════════════════════════════════════════════════════════════════════════════
# CONFIGURAÇÃO VISUAL (CORES E FONTES)
# ═══════════════════════════════════════════════════════════════════════════════
# Aqui definimos o tema claro e a paleta de cores que vai ser usada em toda
# interface. Escolhi tons modernos e suaves para deixar a visualização mais
# agradável. As cores foram pensadas para destacar os diferentes tipos de itens.
# ═══════════════════════════════════════════════════════════════════════════════

ctk.set_appearance_mode("light")

# Cores principais
AMARELO = "#FFCC00"          # Cor da Brasul
AMARELO_ESC = "#E6B800"
AMARELO_LITE = "#FFF8D6"
PRETO = "#1E293B"            # Azul escuro moderno
BRANCO = "#FFFFFF"
FUNDO = "#F8FAFC"            # Fundo suave da tela
CARD_BG = "#FFFFFF"          # Fundo dos cards
BORDA = "#E2E8F0"            # Cor das bordas
AZUL = "#3B82F6"             # Para itens do tipo ACUMULADO
AZUL_LITE = "#EFF6FF"
VERDE = "#22C55E"            # Para itens do tipo QUANTITATIVA
VERDE_LITE = "#F0FDF4"
LILAS = "#8B5CF6"            # Para itens AMBOS (quando aparecem nos dois tipos)
LILAS_LITE = "#F5F3FF"
CINZA_TXT = "#64748B"        # Textos secundários
CINZA_LINHA = "#F8FAFC"

# Fontes usadas na interface
F_TITULO = ("Segoe UI", 22, "bold")
F_SECAO = ("Segoe UI", 14, "bold")
F_LABEL = ("Segoe UI", 11)
F_BTN = ("Segoe UI", 12, "bold")
F_BUSCA = ("Segoe UI", 14)
F_KPI_N = ("Segoe UI", 28, "bold")   # Números grandes dos KPIs
F_KPI_L = ("Segoe UI", 10, "bold")   # Labels dos KPIs
F_COD = ("Consolas", 11, "bold")     # Fonte monoespaçada para códigos


# ═══════════════════════════════════════════════════════════════════════════════
# CLASSE PRINCIPAL - CofreBrasul
# ═══════════════════════════════════════════════════════════════════════════════
# Esta é a janela principal do sistema. Ela gerencia:
#   - A barra lateral com os KPIs e botões de ação
#   - A área principal com a busca e a tabela de resultados
#   - O carregamento e filtro dos dados do Excel
#   - A importação de novos PDFs (em thread separada)
# ═══════════════════════════════════════════════════════════════════════════════

class CofreBrasul(ctk.CTk):

    # ──────────────────────────────────────────────────────────────────────────
    # CONSTRUTOR - Inicializa a janela e carrega os dados
    # ──────────────────────────────────────────────────────────────────────────
    def __init__(self):
        super().__init__()

        # Define os caminhos das pastas baseado em onde o programa está rodando
        if getattr(sys, 'frozen', False):
            # Modo executável (.exe)
            self.dir_base = os.path.dirname(sys.executable)
        else:
            # Modo desenvolvimento (código fonte)
            src = os.path.dirname(os.path.abspath(__file__))
            self.dir_base = os.path.dirname(src)

        self.pasta_input = os.path.join(self.dir_base, "DATA", "input")
        self.pasta_output = os.path.join(self.dir_base, "DATA", "output")
        self.caminho_xls = os.path.join(self.pasta_output, "Cofre_Brasul.xlsx")
        self.caminho_base = os.path.join(self.dir_base, "DATA", "input", "Base_Mestra_FDE.xlsx")

        # Configurações da janela
        self.title("BRASUL — Sistema de Gestão de Insumos")
        self.geometry("1680x920")
        self.minsize(1280, 780)
        self.configure(fg_color=FUNDO)

        # Ícone da janela
        if icone_path.exists():
            try:
                ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID("brasul.cofre.v2")
                self.iconbitmap(str(icone_path))
            except Exception:
                pass

        # Estado da aplicação
        self.df_completo: pd.DataFrame = pd.DataFrame()   # Todos os dados
        self.df_filtro: pd.DataFrame = pd.DataFrame()     # Dados após filtro
        self._tipo_ativo = "TODOS"                        # Filtro ativo
        self._sort_col = None                             # Coluna de ordenação
        self._sort_asc = True                             # Ordem ascendente?
        self._timer_busca = None                          # Timer para busca com debounce

        # Carrega os dados do Excel
        self._carregar_dados()

        # Monta a interface
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)

        self._build_sidebar()   # Barra lateral com KPIs e botões
        self._build_main()      # Área principal com busca e tabela
        self._atualizar_kpis()  # Atualiza os números dos cards

        self.entry_busca.focus()

    # ──────────────────────────────────────────────────────────────────────────
    # CARREGAMENTO DOS DADOS
    # ──────────────────────────────────────────────────────────────────────────
    def _carregar_dados(self):
        """Carrega o arquivo Excel Cofre_Brasul.xlsx e prepara o DataFrame."""
        self.df_completo = pd.DataFrame()
        if os.path.exists(self.caminho_xls):
            try:
                self.df_completo = pd.read_excel(self.caminho_xls).fillna('')
                # Garante que todas as colunas necessárias existem
                for col in ['Obra', 'Obra_Arq', 'Tipo', 'Cod', 'Desc', 'UN']:
                    if col not in self.df_completo.columns:
                        self.df_completo[col] = ''
            except Exception as e:
                print(f"Erro ao carregar dados: {e}")

    def _norm(self, txt: str) -> str:
        """Normaliza texto removendo acentos e convertendo para maiúsculo.
        Usado para buscas case-insensitive e sem acentos."""
        t = unicodedata.normalize('NFD', str(txt))
        return ''.join(c for c in t if unicodedata.category(c) != 'Mn').upper()

    # ══════════════════════════════════════════════════════════════════════════
    # BARRA LATERAL (SIDEBAR)
    # ══════════════════════════════════════════════════════════════════════════
    # Contém o logotipo, os KPIs (indicadores) e os botões de ação.
    # ══════════════════════════════════════════════════════════════════════════

    def _build_sidebar(self):
        """Monta a barra lateral esquerda com todos os elementos."""
        self.sidebar = ctk.CTkFrame(self, width=300, fg_color=BRANCO, corner_radius=0,
                                    border_width=0)
        self.sidebar.grid(row=0, column=0, sticky="nsew")
        self.sidebar.grid_propagate(False)
        self.sidebar.grid_columnconfigure(0, weight=1)

        # Barra amarela no topo (detalhe visual)
        topo = ctk.CTkFrame(self.sidebar, height=6, fg_color=AMARELO, corner_radius=0)
        topo.pack(fill="x")

        # Logotipo da Brasul
        logo_frame = ctk.CTkFrame(self.sidebar, fg_color="transparent")
        logo_frame.pack(padx=24, pady=(28, 16), fill="x")

        if logo_path.exists():
            try:
                img_raw = Image.open(logo_path)
                self.logo_img = ctk.CTkImage(img_raw, size=(250, 200))
                ctk.CTkLabel(logo_frame, image=self.logo_img, text="").pack(anchor="center")
            except Exception as e:
                print(f"Erro ao carregar logo: {e}")
                ctk.CTkLabel(logo_frame, text="BRASUL", font=("Segoe UI", 28, "bold"),
                             text_color=PRETO).pack(anchor="center")
        else:
            ctk.CTkLabel(logo_frame, text="BRASUL", font=("Segoe UI", 28, "bold"),
                         text_color=PRETO).pack(anchor="center")

        # Subtítulo (deixei vazio para dar espaço)
        ctk.CTkLabel(self.sidebar, text="",
                     font=ctk.CTkFont(family="Segoe UI", size=12),
                     text_color=CINZA_TXT).pack(anchor="center", pady=(0, 20))

        # Separador
        ctk.CTkFrame(self.sidebar, height=1, fg_color=BORDA).pack(fill="x", padx=0)

        # ── KPIs (Cards com números) ──
        kpi_container = ctk.CTkFrame(self.sidebar, fg_color="transparent")
        kpi_container.pack(fill="x", padx=18, pady=20)
        kpi_container.grid_columnconfigure((0, 1), weight=1)

        self.kpi_obras = self._kpi(kpi_container, "OBRAS", 0, 0, AZUL, "🏢")
        self.kpi_itens = self._kpi(kpi_container, "ITENS", 1, 0, VERDE, "📦")
        self.kpi_acum = self._kpi(kpi_container, "ACUMULADO", 0, 1, AZUL, "📊")
        self.kpi_quant = self._kpi(kpi_container, "QUANTITATIVA", 1, 1, VERDE, "📋")

        # Separador
        ctk.CTkFrame(self.sidebar, height=1, fg_color=BORDA).pack(fill="x", padx=0)

        # ── Botões de ação ──
        btn_area = ctk.CTkFrame(self.sidebar, fg_color="transparent")
        btn_area.pack(fill="x", padx=18, pady=20)

        self.btn_import = ctk.CTkButton(
            btn_area, text="⬇  IMPORTAR PDF",
            fg_color=AMARELO, text_color=PRETO,
            hover_color=AMARELO_ESC, height=52, corner_radius=12,
            font=F_BTN, command=self._importar_thread)
        self.btn_import.pack(fill="x", pady=(0, 10))

        ctk.CTkButton(
            btn_area, text="⬆  EXPORTAR BUSCA",
            fg_color=VERDE, text_color=BRANCO,
            hover_color="#16A34A", height=52, corner_radius=12,
            font=F_BTN, command=self._exportar).pack(fill="x", pady=(0, 10))

        ctk.CTkButton(
            btn_area, text="↺  RECARREGAR DADOS",
            fg_color=FUNDO, text_color=PRETO,
            hover_color=BORDA, border_width=1, border_color=BORDA,
            height=48, corner_radius=12,
            font=ctk.CTkFont(family="Segoe UI", size=11, weight="bold"),
            command=self._recarregar).pack(fill="x")

        # ── Barra de progresso (fica escondida até ser usada) ──
        self.progresso = ctk.CTkProgressBar(self.sidebar, mode="indeterminate",
                                            progress_color=AMARELO, fg_color=BORDA,
                                            height=6, corner_radius=3)
        self.lbl_status = ctk.CTkLabel(self.sidebar, text="⏳  Processando PDF...",
                                       font=ctk.CTkFont(family="Segoe UI", size=10,
                                                        slant="italic"),
                                       text_color=CINZA_TXT)

        # ── Rodapé ──
        ctk.CTkFrame(self.sidebar, height=1, fg_color=BORDA).pack(side="bottom",
                                                                  fill="x", pady=(0, 0))
        ctk.CTkLabel(self.sidebar, text="Brasul Construtora  ·  v2.0",
                     font=ctk.CTkFont(family="Segoe UI", size=10),
                     text_color="#AAAAAA").pack(side="bottom", pady=12)

    def _kpi(self, parent, label, col, row, cor, icone=""):
        """Cria um card de KPI com ícone e número."""
        frame = ctk.CTkFrame(parent, fg_color=FUNDO, corner_radius=12, height=95)
        frame.grid(row=row, column=col, padx=6, pady=6, sticky="ew")
        frame.grid_propagate(False)

        inner = ctk.CTkFrame(frame, fg_color="transparent")
        inner.pack(expand=True, fill="both", padx=12, pady=10)

        # Linha superior com ícone e número
        top_row = ctk.CTkFrame(inner, fg_color="transparent")
        top_row.pack(fill="x", expand=True)

        if icone:
            ctk.CTkLabel(top_row, text=icone, font=ctk.CTkFont(size=20),
                         text_color=cor).pack(side="left")

        lbl_n = ctk.CTkLabel(top_row, text="0", font=F_KPI_N, text_color=cor)
        lbl_n.pack(side="right")

        ctk.CTkLabel(inner, text=label, font=F_KPI_L,
                     text_color=CINZA_TXT).pack(pady=(5, 0))

        setattr(self, f"_kn_{label.lower().split()[0]}", lbl_n)
        return lbl_n

    # ══════════════════════════════════════════════════════════════════════════
    # ÁREA PRINCIPAL
    # ══════════════════════════════════════════════════════════════════════════
    # Contém o cabeçalho, o campo de busca, os botões de filtro e a tabela
    # com os resultados.
    # ══════════════════════════════════════════════════════════════════════════

    def _build_main(self):
        """Monta a área principal da interface (busca e tabela)."""
        self.main = ctk.CTkFrame(self, fg_color="transparent")
        self.main.grid(row=0, column=1, padx=28, pady=24, sticky="nsew")
        self.main.grid_rowconfigure(2, weight=1)
        self.main.grid_columnconfigure(0, weight=1)

        # ── Cabeçalho com título e data ──
        topo = ctk.CTkFrame(self.main, fg_color="transparent")
        topo.grid(row=0, column=0, sticky="ew", pady=(0, 16))
        topo.grid_columnconfigure(0, weight=1)

        left = ctk.CTkFrame(topo, fg_color="transparent")
        left.pack(side="left")

        ctk.CTkLabel(left, text="Painel de Insumos",
                     font=F_TITULO, text_color=PRETO).pack(anchor="w")
        ctk.CTkLabel(left, text="Consulte e gerencie todos os serviços extraídos dos PDFs de Atestado",
                     font=F_LABEL, text_color=CINZA_TXT).pack(anchor="w")

        # Data atual formatada
        from datetime import date
        meses = ["", "Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho",
                 "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"]
        d = date.today()
        ctk.CTkLabel(topo,
                     text=f"{d.day} de {meses[d.month]} de {d.year}",
                     font=F_LABEL, text_color=CINZA_TXT).pack(side="right")

        # ── Card de busca ──
        busca_card = ctk.CTkFrame(self.main, fg_color=CARD_BG, corner_radius=16,
                                  border_width=1, border_color=BORDA)
        busca_card.grid(row=1, column=0, sticky="ew", pady=(0, 16))

        # Linha do campo de busca
        linha1 = ctk.CTkFrame(busca_card, fg_color="transparent")
        linha1.pack(fill="x", padx=20, pady=(18, 12))

        ctk.CTkLabel(linha1, text="🔍",
                     font=ctk.CTkFont(size=22)).pack(side="left", padx=(0, 12))

        self.entry_busca = ctk.CTkEntry(
            linha1,
            placeholder_text="Busque por código (ex: 02.01.001) ou descrição (ex: ESCAVAÇÃO, AÇO, CONCRETO, PINTURA)...",
            height=56, border_width=0,
            fg_color="#F8F9FB",
            font=F_BUSCA,
            text_color=PRETO,
            corner_radius=12)
        self.entry_busca.pack(side="left", fill="x", expand=True, padx=(0, 12))
        self.entry_busca.bind("<Return>", lambda e: self._pesquisar())
        self.entry_busca.bind("<KeyRelease>", self._debounce_busca)

        ctk.CTkButton(
            linha1, text="BUSCAR",
            fg_color=AMARELO, text_color=PRETO,
            hover_color=AMARELO_ESC,
            width=130, height=56, corner_radius=12, font=F_BTN,
            command=self._pesquisar).pack(side="left", padx=(0, 8))

        ctk.CTkButton(
            linha1, text="✕",
            fg_color="#F0F0F0", text_color=PRETO,
            hover_color=BORDA,
            width=54, height=56, corner_radius=12, font=F_BTN,
            command=self._limpar).pack(side="left")

        # ── Filtros por tipo (apenas 3: Todos, Acumulado, Quantitativa) ──
        linha2 = ctk.CTkFrame(busca_card, fg_color="transparent")
        linha2.pack(fill="x", padx=20, pady=(0, 16))

        ctk.CTkLabel(linha2, text="Filtrar por tipo:",
                     font=F_LABEL, text_color=CINZA_TXT).pack(side="left", padx=(0, 12))

        self._chips = {}
        cores_botoes = {
            "TODOS": (PRETO, BRANCO),
            "ACUMULADO": (AZUL, BRANCO),
            "QUANTITATIVA": (VERDE, BRANCO),
        }

        for nome, val in [("Todos", "TODOS"),
                          ("Acumulado", "ACUMULADO"),
                          ("Quantitativa", "QUANTITATIVA")]:
            cor_fundo, cor_texto = cores_botoes[val]
            btn = ctk.CTkButton(
                linha2, text=nome, width=130, height=36,
                corner_radius=18,
                fg_color=cor_fundo if val == "TODOS" else "#EFEFEF",
                text_color=cor_texto if val == "TODOS" else CINZA_TXT,
                hover_color=cor_fundo,
                font=ctk.CTkFont(family="Segoe UI", size=11, weight="bold"),
                command=lambda v=val, ca=cor_fundo, ct=cor_texto: self._set_tipo(v, ca, ct))
            btn.pack(side="left", padx=5)
            self._chips[val] = (btn, cor_fundo, cor_texto)

        # ── Tabela de resultados ──
        tab_card = ctk.CTkFrame(self.main, fg_color=CARD_BG, corner_radius=16,
                                border_width=1, border_color=BORDA)
        tab_card.grid(row=2, column=0, sticky="nsew")
        tab_card.grid_rowconfigure(1, weight=1)
        tab_card.grid_columnconfigure(0, weight=1)

        # Cabeçalho da tabela (título e contagem)
        cab = ctk.CTkFrame(tab_card, fg_color="transparent")
        cab.grid(row=0, column=0, columnspan=2, sticky="ew", padx=20, pady=(16, 10))

        self.lbl_titulo_res = ctk.CTkLabel(
            cab, text="Aguardando pesquisa…",
            font=F_SECAO, text_color=PRETO)
        self.lbl_titulo_res.pack(side="left")

        self.lbl_contagem = ctk.CTkLabel(cab, text="",
                                         font=F_LABEL, text_color=CINZA_TXT)
        self.lbl_contagem.pack(side="right")

        # Configuração da Treeview (tabela)
        style = ttk.Style()
        style.theme_use("default")
        style.configure("Brasul.Treeview",
                        background=BRANCO,
                        fieldbackground=BRANCO,
                        rowheight=46,
                        font=('Segoe UI', 10),
                        borderwidth=0, relief="flat")
        style.configure("Brasul.Treeview.Heading",
                        font=('Segoe UI', 10, 'bold'),
                        background="#F8FAFC",
                        foreground=PRETO,
                        relief="flat",
                        borderwidth=0, padding=12)
        style.map("Brasul.Treeview",
                  background=[('selected', AMARELO_LITE)],
                  foreground=[('selected', PRETO)])

        colunas = ("tipo", "obra", "cod", "desc", "un")
        self.tabela = ttk.Treeview(tab_card, columns=colunas,
                                   show="headings",
                                   style="Brasul.Treeview",
                                   selectmode="extended")

        headers_cfg = {
            "tipo": ("TIPO", 110, "center"),
            "obra": ("ESCOLA / OBRA", 320, "w"),
            "cod": ("CÓDIGO", 140, "center"),
            "desc": ("DESCRIÇÃO DO SERVIÇO", 580, "w"),
            "un": ("UN", 80, "center"),
        }
        for col, (h, w, anc) in headers_cfg.items():
            self.tabela.heading(col, text=h,
                                command=lambda c=col: self._ordenar(c))
            self.tabela.column(col, width=w, anchor=anc, minwidth=60)

        # Cores para diferentes tipos de itens
        self.tabela.tag_configure("ACUMULADO", background=AZUL_LITE, foreground=AZUL)
        self.tabela.tag_configure("QUANTITATIVA", background=VERDE_LITE, foreground=VERDE)
        self.tabela.tag_configure("AMBOS", background=LILAS_LITE, foreground=LILAS)
        self.tabela.tag_configure("PAR", background=BRANCO, foreground=PRETO)
        self.tabela.tag_configure("IMPAR", background=CINZA_LINHA, foreground=PRETO)

        # Barras de rolagem
        vsb = ttk.Scrollbar(tab_card, orient="vertical", command=self.tabela.yview)
        hsb = ttk.Scrollbar(tab_card, orient="horizontal", command=self.tabela.xview)
        self.tabela.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        self.tabela.grid(row=1, column=0, sticky="nsew", padx=(16, 0), pady=(0, 0))
        vsb.grid(row=1, column=1, sticky="ns")
        hsb.grid(row=2, column=0, sticky="ew", padx=(16, 0))

    # ══════════════════════════════════════════════════════════════════════════
    # ATUALIZAÇÃO DOS KPIs
    # ══════════════════════════════════════════════════════════════════════════
    def _atualizar_kpis(self):
        """Atualiza os números dos cards (OBRAS, ITENS, ACUMULADO, QUANTITATIVA)."""
        if self.df_completo.empty:
            for attr in ['kpi_obras', 'kpi_itens', 'kpi_acum', 'kpi_quant']:
                getattr(self, attr).configure(text="0")
            return

        df = self.df_completo
        obras = df['Obra'].nunique()
        # Considera apenas códigos válidos (formato XX.XX.XXX)
        validos = df[df['Cod'].str.match(r'\d{2}\.\d{2}', na=False)]
        itens = len(validos)
        acum = validos[validos['Tipo'].str.contains('ACUMULADO', na=False)].shape[0]
        quant = validos[validos['Tipo'].str.contains('QUANTITATIVA', na=False)].shape[0]

        self.kpi_obras.configure(text=str(obras))
        self.kpi_itens.configure(text=str(itens))
        self.kpi_acum.configure(text=str(acum))
        self.kpi_quant.configure(text=str(quant))

    # ══════════════════════════════════════════════════════════════════════════
    # BUSCA E FILTROS
    # ══════════════════════════════════════════════════════════════════════════
    def _debounce_busca(self, _=None):
        """Aguarda um pequeno intervalo antes de executar a busca.
        Isso evita processar a cada tecla digitada."""
        if self._timer_busca:
            self.after_cancel(self._timer_busca)
        self._timer_busca = self.after(380, self._pesquisar)

    def _pesquisar(self, _=None):
        """Filtra os dados com base no termo de busca e no tipo selecionado."""
        termo = self.entry_busca.get().strip()
        tipo = self._tipo_ativo

        if self.df_completo.empty:
            self.lbl_titulo_res.configure(text="Nenhum dado carregado.")
            return

        df = self.df_completo.copy()

        # Filtro pelo termo de busca (código, descrição ou obra)
        if termo:
            t = self._norm(termo)
            mask = (
                    df['Desc'].apply(lambda x: t in self._norm(x)) |
                    df['Cod'].apply(lambda x: t in self._norm(x)) |
                    df['Obra'].apply(lambda x: t in self._norm(x))
            )
            df = df[mask]

        # Filtro pelo tipo (apenas 3 opções)
        if tipo != "TODOS":
            if tipo == "ACUMULADO":
                df = df[df['Tipo'].str.contains('ACUMULADO', na=False)]
            elif tipo == "QUANTITATIVA":
                df = df[df['Tipo'].str.contains('QUANTITATIVA', na=False)]

        self.df_filtro = df
        self._popular_tabela(df)

    def _set_tipo(self, val, cor_ativa, cor_txt):
        """Altera o filtro ativo e atualiza a aparência dos botões."""
        self._tipo_ativo = val
        for v, (btn, ca, ct) in self._chips.items():
            if v == val:
                btn.configure(fg_color=ca, text_color=ct)
            else:
                btn.configure(fg_color="#EFEFEF", text_color=CINZA_TXT)
        self._pesquisar()

    # ══════════════════════════════════════════════════════════════════════════
    # POPULAÇÃO DA TABELA
    # ══════════════════════════════════════════════════════════════════════════
    def _popular_tabela(self, df: pd.DataFrame):
        """Preenche a tabela com os dados filtrados."""
        self.tabela.delete(*self.tabela.get_children())

        icone_tipo = {
            'ACUMULADO': '📊  Acumulado',
            'QUANTITATIVA': '📋  Quantitativa',
            'ACUMULADO e QUANTITATIVA': '🔄  Ambos',
        }

        MAX = 2000
        for i, (_, row) in enumerate(df.head(MAX).iterrows()):
            tipo_raw = str(row.get('Tipo', ''))

            # Define ícone e tag baseado no tipo
            if 'ACUMULADO' in tipo_raw and 'QUANTITATIVA' in tipo_raw:
                label = icone_tipo['ACUMULADO e QUANTITATIVA']
                tag = "AMBOS"
            elif 'ACUMULADO' in tipo_raw:
                label = icone_tipo['ACUMULADO']
                tag = "ACUMULADO"
            elif 'QUANTITATIVA' in tipo_raw:
                label = icone_tipo['QUANTITATIVA']
                tag = "QUANTITATIVA"
            else:
                label = tipo_raw
                tag = "PAR" if i % 2 == 0 else "IMPAR"

            self.tabela.insert("", "end", tags=(tag,), values=(
                label,
                str(row.get('Obra', '')),
                str(row.get('Cod', '')),
                str(row.get('Desc', '')),
                str(row.get('UN', '')),
            ))

        total = len(df)
        exib = min(total, MAX)
        suf = f"  (exibindo {exib} de {total})" if total > MAX else ""

        self.lbl_titulo_res.configure(text="Resultados")
        self.lbl_contagem.configure(
            text=f"🔎  {total} item{'ns' if total != 1 else ''} encontrado{'s' if total != 1 else ''}{suf}")

    def _limpar_tabela(self):
        """Limpa a tabela e reseta o estado do filtro."""
        self.tabela.delete(*self.tabela.get_children())
        self.df_filtro = pd.DataFrame()
        self.lbl_titulo_res.configure(text="Aguardando pesquisa…")
        self.lbl_contagem.configure(text="")

    def _limpar(self):
        """Limpa o campo de busca e reseta a tabela."""
        self.entry_busca.delete(0, 'end')
        self._limpar_tabela()
        self.entry_busca.focus()

    # ══════════════════════════════════════════════════════════════════════════
    # ORDENAÇÃO
    # ══════════════════════════════════════════════════════════════════════════
    def _ordenar(self, col):
        """Ordena a tabela pela coluna clicada (alterna ascendente/descendente)."""
        if self._sort_col == col:
            self._sort_asc = not self._sort_asc
        else:
            self._sort_col = col
            self._sort_asc = True

        mapa = {"tipo": "Tipo", "obra": "Obra", "cod": "Cod", "desc": "Desc", "un": "UN"}
        col_df = mapa.get(col, col)

        if not self.df_filtro.empty and col_df in self.df_filtro.columns:
            self._popular_tabela(
                self.df_filtro.sort_values(col_df, ascending=self._sort_asc))

    # ══════════════════════════════════════════════════════════════════════════
    # IMPORTAÇÃO DE PDF (OCR)
    # ══════════════════════════════════════════════════════════════════════════
    def _importar_thread(self):
        """Inicia a importação de um PDF em uma thread separada."""
        if not MAIN_OK:
            messagebox.showwarning(
                "Módulo não encontrado",
                "O arquivo 'main.py' não foi localizado.\n"
                "Verifique se ele está na mesma pasta que interface.py.")
            return

        caminho = filedialog.askopenfilename(
            title="Selecionar PDF para importar",
            filetypes=[("Arquivo PDF", "*.pdf")])
        if not caminho:
            return

        # Copia o PDF para a pasta de entrada
        os.makedirs(self.pasta_input, exist_ok=True)
        dest = os.path.join(self.pasta_input, os.path.basename(caminho))
        shutil.copy(caminho, dest)

        # Atualiza a interface para mostrar o progresso
        self.btn_import.configure(state="disabled", text="⏳  Lendo PDF...")
        self.progresso.pack(padx=18, pady=6, fill="x")
        self.progresso.start()
        self.lbl_status.pack(padx=18, pady=(0, 6))

        threading.Thread(target=self._run_ocr, args=(dest,), daemon=True).start()

    def _arquivos_ja_processados(self) -> set:
        """Retorna o conjunto de nomes de arquivos já processados."""
        if os.path.exists(self.caminho_xls):
            try:
                df = pd.read_excel(self.caminho_xls, usecols=['Obra_Arq']).fillna('')
                return set(df['Obra_Arq'].astype(str).tolist())
            except Exception:
                pass
        return set()

    def _run_ocr(self, caminho_pdf: str):
        """Executa o OCR em um PDF em background."""
        erro_msg = None
        n_itens = 0
        duplicado = False
        try:
            from pathlib import Path
            import time

            nome_arq = os.path.basename(caminho_pdf)
            ja_proc = self._arquivos_ja_processados()

            if nome_arq in ja_proc:
                duplicado = True
            else:
                por_cod, por_desc = carregar_base(Path(self.caminho_base))
                resultado = processar_pdf(Path(caminho_pdf), por_cod, por_desc,
                                          ja_processados=ja_proc)

                if resultado is None:
                    duplicado = True
                else:
                    rows = []
                    for it in resultado.get('itens', []):
                        rows.append({
                            'Obra': resultado.get('nome', 'DESCONHECIDA'),
                            'Obra_Arq': resultado.get('arq', ''),
                            'Tipo': it.get('tipo', ''),
                            'Cod': it.get('codigo', ''),
                            'Desc': it.get('descricao', ''),
                            'UN': it.get('unidade', ''),
                        })
                    novo = pd.DataFrame(rows)
                    n_itens = len(rows)

                    if not novo.empty:
                        os.makedirs(self.pasta_output, exist_ok=True)
                        if os.path.exists(self.caminho_xls):
                            existente = pd.read_excel(self.caminho_xls).fillna('')
                            combinado = pd.concat([existente, novo], ignore_index=True)
                            combinado.drop_duplicates(
                                subset=['Obra_Arq', 'Cod', 'Tipo'], inplace=True)
                        else:
                            combinado = novo

                        # Tenta salvar, com retry em caso de arquivo bloqueado
                        for tentativa in range(5):
                            try:
                                combinado.to_excel(self.caminho_xls, index=False)
                                break
                            except PermissionError:
                                if tentativa == 4:
                                    raise PermissionError(
                                        "Arquivo em uso! Feche o Cofre_Brasul.xlsx "
                                        "no Excel e tente importar novamente.")
                                time.sleep(1.5)

        except Exception as e:
            import traceback
            traceback.print_exc()
            erro_msg = str(e)

        # Chama as funções de callback na thread principal
        if duplicado:
            self.after(0, self._ocr_duplicado)
        elif erro_msg:
            self.after(0, lambda msg=erro_msg: self._ocr_erro(msg))
        else:
            self.after(0, lambda n=n_itens: self._ocr_ok(n))

    def _ocr_ok(self, n_itens: int):
        """Callback quando o OCR termina com sucesso."""
        self._fim_ocr()
        self._carregar_dados()
        self._atualizar_kpis()
        messagebox.showinfo(
            "Importação concluída",
            f"PDF processado com sucesso!\n\n"
            f"{n_itens} {'item importado' if n_itens == 1 else 'itens importados'} para o Cofre.")
        self._limpar()

    def _ocr_duplicado(self):
        """Callback quando o PDF já foi processado anteriormente."""
        self._fim_ocr()
        messagebox.showwarning(
            "PDF já importado",
            "Este arquivo já foi processado e está no Cofre.\n\n"
            "Para reprocessar, remova o arquivo do Cofre_Brasul.xlsx primeiro.")

    def _ocr_erro(self, msg: str):
        """Callback quando ocorre um erro no OCR."""
        self._fim_ocr()
        messagebox.showerror("❌  Erro no processamento", msg)

    def _fim_ocr(self):
        """Limpa os elementos de progresso e reativa o botão de importar."""
        self.progresso.stop()
        self.progresso.pack_forget()
        self.lbl_status.pack_forget()
        self.btn_import.configure(state="normal", text="⬇  IMPORTAR PDF")

    # ══════════════════════════════════════════════════════════════════════════
    # EXPORTAÇÃO E RECARREGAMENTO
    # ══════════════════════════════════════════════════════════════════════════
    def _exportar(self):
        """Exporta os dados filtrados para um arquivo Excel."""
        if self.df_filtro.empty:
            messagebox.showwarning(
                "Sem dados para exportar",
                "Realize uma busca primeiro para exportar os resultados.")
            return

        path = filedialog.asksaveasfilename(
            title="Salvar resultado",
            defaultextension=".xlsx",
            initialfile="Relatorio_Busca_Atestados_Brasul.xlsx",
            filetypes=[("Planilha Excel", "*.xlsx")])
        if not path:
            return

        try:
            self.df_filtro.to_excel(path, index=False)
            messagebox.showinfo(
                "✅  Exportado",
                f"Planilha salva com {len(self.df_filtro)} registros!")
        except Exception as e:
            messagebox.showerror("Erro ao exportar", str(e))

    def _recarregar(self):
        """Recarrega os dados do Excel e atualiza a interface."""
        self._carregar_dados()
        self._atualizar_kpis()
        self._limpar()
        total = len(self.df_completo)
        messagebox.showinfo(
            "✅  Dados atualizados",
            f"Banco recarregado com {total} registros.")


# ═══════════════════════════════════════════════════════════════════════════════
# PONTO DE ENTRADA
# ═══════════════════════════════════════════════════════════════════════════════
if __name__ == "__main__":
    app = CofreBrasul()
    app.mainloop()