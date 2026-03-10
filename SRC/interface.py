
import os, sys, threading, shutil, unicodedata
import pandas as pd
import customtkinter as ctk
from tkinter import ttk, filedialog, messagebox
from PIL import Image, ImageTk
import ctypes

# ─────────────────────────────────────────────────────────────────────────────
#  RECURSOS (compatível com PyInstaller .exe)
# ─────────────────────────────────────────────────────────────────────────────
def resource_path(rel):
    try:
        base = sys._MEIPASS
    except Exception:
        base = os.path.abspath(".")
    return os.path.join(base, rel)

# ─────────────────────────────────────────────────────────────────────────────
#  IMPORTA O EXTRATOR (main.py v7.0)
# ─────────────────────────────────────────────────────────────────────────────
try:
    from main import processar_pdf, carregar_base, gerar_excel
    MAIN_OK = True
except ImportError:
    MAIN_OK = False

# ─────────────────────────────────────────────────────────────────────────────
#  TEMA
# ─────────────────────────────────────────────────────────────────────────────
ctk.set_appearance_mode("light")

# Paleta
AMARELO      = "#FFCC00"
AMARELO_ESC  = "#E6B800"
AMARELO_LITE = "#FFF8D6"
PRETO        = "#1C1C1E"
BRANCO       = "#FFFFFF"
FUNDO        = "#F0F2F5"
CARD_BG      = "#FFFFFF"
BORDA        = "#E2E5EA"
AZUL         = "#2563EB"   # ACUMULADO
AZUL_LITE    = "#EFF6FF"
VERDE        = "#16A34A"   # QUANTITATIVA
VERDE_LITE   = "#F0FDF4"
LILAS        = "#7C3AED"   # AMBOS
LILAS_LITE   = "#F5F3FF"
CINZA_TXT    = "#6B7280"
CINZA_LINHA  = "#F8FAFC"

# Fontes — definidas como tuplas simples (compatível antes de criar a janela)
F_TITULO  = ("Segoe UI", 20, "bold")
F_SECAO   = ("Segoe UI", 13, "bold")
F_LABEL   = ("Segoe UI", 11)
F_BTN     = ("Segoe UI", 12, "bold")
F_BUSCA   = ("Segoe UI", 14)
F_KPI_N   = ("Segoe UI", 26, "bold")
F_KPI_L   = ("Segoe UI",  9, "bold")
F_COD     = ("Consolas",  11, "bold")


# ═════════════════════════════════════════════════════════════════════════════
#  APLICAÇÃO
# ═════════════════════════════════════════════════════════════════════════════
class CofreBrasul(ctk.CTk):

    # ── INIT ──────────────────────────────────────────────────────────────────
    def __init__(self):
        super().__init__()

        # Caminhos
        if getattr(sys, 'frozen', False):
            self.dir_base = os.path.dirname(sys.executable)
        else:
            src = os.path.dirname(os.path.abspath(__file__))
            self.dir_base = os.path.dirname(src)

        self.pasta_input  = os.path.join(self.dir_base, "DATA", "input")
        self.pasta_output = os.path.join(self.dir_base, "DATA", "output")
        self.caminho_xls  = os.path.join(self.pasta_output, "Cofre_Brasul.xlsx")
        self.caminho_base = os.path.join(self.dir_base, "Base_Mestra_FDE.xlsx")

        # Janela
        self.title("BRASUL — Sistema de Gestão de Insumos Brasul Construtora")
        self.geometry("1680x920")
        self.minsize(1280, 780)
        self.configure(fg_color=FUNDO)

        # Ícone
        ico = resource_path("iconebrasul2.ico")
        if os.path.exists(ico):
            try:
                ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID("brasul.cofre.v2")
                img_ico = Image.open(ico)
                self.icon_photo = ImageTk.PhotoImage(img_ico)
                self.after(200, lambda: self.wm_iconphoto(False, self.icon_photo))
            except Exception:
                pass

        # Estado
        self.df_completo: pd.DataFrame = pd.DataFrame()
        self.df_filtro:   pd.DataFrame = pd.DataFrame()
        self._tipo_ativo  = "TODOS"
        self._sort_col    = None
        self._sort_asc    = True
        self._timer_busca = None

        self._carregar_dados()

        # Grid raiz
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)

        self._build_sidebar()
        self._build_main()
        self._atualizar_kpis()
        self.entry_busca.focus()

    # ── DADOS ─────────────────────────────────────────────────────────────────
    def _carregar_dados(self):
        self.df_completo = pd.DataFrame()
        if os.path.exists(self.caminho_xls):
            try:
                self.df_completo = pd.read_excel(self.caminho_xls).fillna('')
                # Garante colunas mínimas
                for col in ['Obra', 'Obra_Arq', 'Tipo', 'Cod', 'Desc', 'UN']:
                    if col not in self.df_completo.columns:
                        self.df_completo[col] = ''
            except Exception as e:
                print(f"Erro ao carregar dados: {e}")

    def _norm(self, txt: str) -> str:
        t = unicodedata.normalize('NFD', str(txt))
        return ''.join(c for c in t if unicodedata.category(c) != 'Mn').upper()

    # ══════════════════════════════════════════════════════════════════════════
    #  SIDEBAR
    # ══════════════════════════════════════════════════════════════════════════
    def _build_sidebar(self):
        self.sidebar = ctk.CTkFrame(self, width=285, fg_color=BRANCO, corner_radius=0,
                                    border_width=0)
        self.sidebar.grid(row=0, column=0, sticky="nsew")
        self.sidebar.grid_propagate(False)
        self.sidebar.grid_columnconfigure(0, weight=1)

        # Faixa amarela no topo da sidebar
        topo = ctk.CTkFrame(self.sidebar, height=6, fg_color=AMARELO, corner_radius=0)
        topo.pack(fill="x")

        # ── Logo ──
        logo_frame = ctk.CTkFrame(self.sidebar, fg_color="transparent")
        logo_frame.pack(padx=24, pady=(28, 6), fill="x")

        caminho_logo = os.path.join(os.path.dirname(os.path.abspath(__file__)), "LOGOTIPOBRASUL.png")
        if os.path.exists(caminho_logo):
            img_raw = Image.open(caminho_logo)
            self.logo_img = ctk.CTkImage(img_raw, size=(210, 98))
            ctk.CTkLabel(logo_frame, image=self.logo_img, text="").pack(anchor="w")
        else:
            ctk.CTkLabel(logo_frame, text="BRASUL", font=("Segoe UI", 24, "bold"),
                         text_color=PRETO).pack(anchor="w")

        ctk.CTkLabel(self.sidebar, text="",
                     font=ctk.CTkFont(family="Segoe UI", size=9, weight="bold"),
                     text_color=CINZA_TXT).pack(anchor="w", padx=26, pady=(0, 20))

        # ── Separador ──
        ctk.CTkFrame(self.sidebar, height=1, fg_color=BORDA).pack(fill="x", padx=0)

        # ── KPIs ──
        kpi_container = ctk.CTkFrame(self.sidebar, fg_color="transparent")
        kpi_container.pack(fill="x", padx=18, pady=18)
        kpi_container.grid_columnconfigure((0, 1), weight=1)

        self.kpi_obras  = self._kpi(kpi_container, "OBRAS",       0, 0)
        self.kpi_itens  = self._kpi(kpi_container, "ITENS",        1, 0)
        self.kpi_acum   = self._kpi(kpi_container, "ACUMULADO",   0, 1, AZUL)
        self.kpi_quant  = self._kpi(kpi_container, "QUANTITATIVA",1, 1, VERDE)

        # ── Separador ──
        ctk.CTkFrame(self.sidebar, height=1, fg_color=BORDA).pack(fill="x", padx=0)

        # ── Botões de ação ──
        btn_area = ctk.CTkFrame(self.sidebar, fg_color="transparent")
        btn_area.pack(fill="x", padx=18, pady=20)

        self.btn_import = ctk.CTkButton(
            btn_area, text="⬇   IMPORTAR PDF",
            fg_color=AMARELO, text_color=PRETO,
            hover_color=AMARELO_ESC, height=50, corner_radius=10,
            font=F_BTN, command=self._importar_thread)
        self.btn_import.pack(fill="x", pady=(0, 10))

        ctk.CTkButton(
            btn_area, text="⬆   EXPORTAR BUSCA",
            fg_color="#22C55E", text_color=BRANCO,
            hover_color="#16A34A", height=50, corner_radius=10,
            font=F_BTN, command=self._exportar).pack(fill="x", pady=(0, 10))

        ctk.CTkButton(
            btn_area, text="↺   RECARREGAR DADOS",
            fg_color=FUNDO, text_color=PRETO,
            hover_color=BORDA, border_width=1, border_color=BORDA,
            height=44, corner_radius=10,
            font=ctk.CTkFont(family="Segoe UI", size=11, weight="bold"),
            command=self._recarregar).pack(fill="x")

        # ── Barra de progresso (escondida) ──
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
                     font=ctk.CTkFont(family="Segoe UI", size=9),
                     text_color="#AAAAAA").pack(side="bottom", pady=10)

    def _kpi(self, parent, label, col, row, cor=PRETO):
        frame = ctk.CTkFrame(parent, fg_color=FUNDO, corner_radius=10)
        frame.grid(row=row, column=col, padx=5, pady=5, sticky="ew")

        lbl_n = ctk.CTkLabel(frame, text="–", font=F_KPI_N, text_color=cor)
        lbl_n.pack(pady=(12, 0))

        ctk.CTkLabel(frame, text=label, font=F_KPI_L,
                     text_color=CINZA_TXT).pack(pady=(0, 10))

        setattr(self, f"_kn_{label.lower().split()[0]}", lbl_n)
        return lbl_n

    # ══════════════════════════════════════════════════════════════════════════
    #  ÁREA PRINCIPAL
    # ══════════════════════════════════════════════════════════════════════════
    def _build_main(self):
        self.main = ctk.CTkFrame(self, fg_color="transparent")
        self.main.grid(row=0, column=1, padx=28, pady=24, sticky="nsew")
        self.main.grid_rowconfigure(2, weight=1)
        self.main.grid_columnconfigure(0, weight=1)

        # ── Cabeçalho ──────────────────────────────────────────────────────
        topo = ctk.CTkFrame(self.main, fg_color="transparent")
        topo.grid(row=0, column=0, sticky="ew", pady=(0, 16))
        topo.grid_columnconfigure(0, weight=1)

        left = ctk.CTkFrame(topo, fg_color="transparent")
        left.pack(side="left")

        ctk.CTkLabel(left, text="Painel de Insumos",
                     font=F_TITULO, text_color=PRETO).pack(anchor="w")
        ctk.CTkLabel(left, text="Consulte e gerencie todos os serviços extraídos dos PDFs FDE",
                     font=F_LABEL, text_color=CINZA_TXT).pack(anchor="w")

        # Data
        from datetime import date
        meses = ["","Janeiro","Fevereiro","Março","Abril","Maio","Junho",
                 "Julho","Agosto","Setembro","Outubro","Novembro","Dezembro"]
        d = date.today()
        ctk.CTkLabel(topo,
                     text=f"{d.day} de {meses[d.month]} de {d.year}",
                     font=F_LABEL, text_color=CINZA_TXT).pack(side="right")

        # ── Card de busca ───────────────────────────────────────────────────
        busca_card = ctk.CTkFrame(self.main, fg_color=CARD_BG, corner_radius=14,
                                  border_width=1, border_color=BORDA)
        busca_card.grid(row=1, column=0, sticky="ew", pady=(0, 16))

        # Linha 1: entrada + botão
        linha1 = ctk.CTkFrame(busca_card, fg_color="transparent")
        linha1.pack(fill="x", padx=18, pady=(16, 8))

        # Ícone lupa
        ctk.CTkLabel(linha1, text="🔍",
                     font=ctk.CTkFont(size=20)).pack(side="left", padx=(0, 10))

        self.entry_busca = ctk.CTkEntry(
            linha1,
            placeholder_text="Busque por código  (ex: 02.01.001)  ou descrição  (ex: ESCAVAÇÃO, AÇO, CONCRETO, PINTURA)...",
            height=54, border_width=0,
            fg_color="#F8F9FB",
            font=F_BUSCA,
            text_color=PRETO,
            corner_radius=10)
        self.entry_busca.pack(side="left", fill="x", expand=True, padx=(0, 12))
        self.entry_busca.bind("<Return>",    lambda e: self._pesquisar())
        self.entry_busca.bind("<KeyRelease>", self._debounce_busca)

        ctk.CTkButton(
            linha1, text="BUSCAR",
            fg_color=AMARELO, text_color=PRETO,
            hover_color=AMARELO_ESC,
            width=130, height=54, corner_radius=10, font=F_BTN,
            command=self._pesquisar).pack(side="left", padx=(0, 8))

        ctk.CTkButton(
            linha1, text="✕",
            fg_color="#F0F0F0", text_color=PRETO,
            hover_color=BORDA,
            width=54, height=54, corner_radius=10, font=F_BTN,
            command=self._limpar).pack(side="left")

        # Linha 2: filtros de tipo (chips)
        linha2 = ctk.CTkFrame(busca_card, fg_color="transparent")
        linha2.pack(fill="x", padx=18, pady=(0, 14))

        ctk.CTkLabel(linha2, text="Filtrar por tipo:",
                     font=F_LABEL, text_color=CINZA_TXT).pack(side="left", padx=(0, 12))

        self._chips = {}
        for nome, val, cor_a, cor_t in [
            ("Todos",         "TODOS",        PRETO,   BRANCO),
            ("Acumulado",     "ACUMULADO",    AZUL,    BRANCO),
            ("Quantitativa",  "QUANTITATIVA", VERDE,   BRANCO),
            ("Ambos",         "AMBOS",        LILAS,   BRANCO),
        ]:
            btn = ctk.CTkButton(
                linha2, text=nome, width=130, height=34,
                corner_radius=17,
                fg_color=PRETO if val == "TODOS" else "#EFEFEF",
                text_color=BRANCO if val == "TODOS" else CINZA_TXT,
                hover_color=cor_a,
                font=ctk.CTkFont(family="Segoe UI", size=11, weight="bold"),
                command=lambda v=val, ca=cor_a, ct=cor_t: self._set_tipo(v, ca, ct))
            btn.pack(side="left", padx=4)
            self._chips[val] = (btn, cor_a, cor_t)

        # ── Card da tabela ──────────────────────────────────────────────────
        tab_card = ctk.CTkFrame(self.main, fg_color=CARD_BG, corner_radius=14,
                                border_width=1, border_color=BORDA)
        tab_card.grid(row=2, column=0, sticky="nsew")
        tab_card.grid_rowconfigure(1, weight=1)
        tab_card.grid_columnconfigure(0, weight=1)

        # Sub-cabeçalho da tabela
        cab = ctk.CTkFrame(tab_card, fg_color="transparent")
        cab.grid(row=0, column=0, columnspan=2, sticky="ew", padx=18, pady=(14, 8))

        self.lbl_titulo_res = ctk.CTkLabel(
            cab, text="Aguardando pesquisa…",
            font=F_SECAO, text_color=PRETO)
        self.lbl_titulo_res.pack(side="left")

        self.lbl_contagem = ctk.CTkLabel(cab, text="",
                                         font=F_LABEL, text_color=CINZA_TXT)
        self.lbl_contagem.pack(side="right")

        # Treeview — estilo refinado
        style = ttk.Style()
        style.theme_use("default")
        style.configure("Brasul.Treeview",
                        background=BRANCO,
                        fieldbackground=BRANCO,
                        rowheight=44,
                        font=('Segoe UI', 10),
                        borderwidth=0, relief="flat")
        style.configure("Brasul.Treeview.Heading",
                        font=('Segoe UI', 10, 'bold'),
                        background="#F8FAFC",
                        foreground=PRETO,
                        relief="flat",
                        borderwidth=0, padding=10)
        style.map("Brasul.Treeview",
                  background=[('selected', AMARELO_LITE)],
                  foreground=[('selected', PRETO)])

        colunas = ("tipo", "obra", "cod", "desc", "un")
        self.tabela = ttk.Treeview(tab_card, columns=colunas,
                                   show="headings",
                                   style="Brasul.Treeview",
                                   selectmode="extended")

        headers_cfg = {
            "tipo": ("TIPO",              110, "center"),
            "obra": ("ESCOLA / OBRA",     310, "w"),
            "cod":  ("CÓDIGO",            130, "center"),
            "desc": ("DESCRIÇÃO DO SERVIÇO", 560, "w"),
            "un":   ("UN",                70,  "center"),
        }
        for col, (h, w, anc) in headers_cfg.items():
            self.tabela.heading(col, text=h,
                                command=lambda c=col: self._ordenar(c))
            self.tabela.column(col, width=w, anchor=anc, minwidth=60)

        # Tags de cor
        self.tabela.tag_configure("ACUMULADO",   background=AZUL_LITE,  foreground=AZUL)
        self.tabela.tag_configure("QUANTITATIVA",background=VERDE_LITE, foreground=VERDE)
        self.tabela.tag_configure("AMBOS",       background=LILAS_LITE, foreground=LILAS)
        self.tabela.tag_configure("PAR",         background=BRANCO,     foreground=PRETO)
        self.tabela.tag_configure("IMPAR",       background=CINZA_LINHA,foreground=PRETO)

        # Scrollbars
        vsb = ttk.Scrollbar(tab_card, orient="vertical",   command=self.tabela.yview)
        hsb = ttk.Scrollbar(tab_card, orient="horizontal", command=self.tabela.xview)
        self.tabela.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        self.tabela.grid(row=1, column=0, sticky="nsew", padx=(14, 0), pady=(0, 0))
        vsb.grid(row=1, column=1, sticky="ns")
        hsb.grid(row=2, column=0, sticky="ew", padx=(14, 0))

    # ══════════════════════════════════════════════════════════════════════════
    #  KPIs
    # ══════════════════════════════════════════════════════════════════════════
    def _atualizar_kpis(self):
        if self.df_completo.empty:
            for attr in ['kpi_obras','kpi_itens','kpi_acum','kpi_quant']:
                getattr(self, attr).configure(text="0")
            return

        df = self.df_completo
        obras = df['Obra'].nunique()
        # só linhas com código FDE válido
        validos = df[df['Cod'].str.match(r'\d{2}\.\d{2}', na=False)]
        itens  = len(validos)
        acum   = validos[validos['Tipo'].str.contains('ACUMULADO',    na=False)].shape[0]
        quant  = validos[validos['Tipo'].str.contains('QUANTITATIVA', na=False)].shape[0]

        self.kpi_obras.configure(text=str(obras))
        self.kpi_itens.configure(text=str(itens))
        self.kpi_acum.configure(text=str(acum))
        self.kpi_quant.configure(text=str(quant))

    # ══════════════════════════════════════════════════════════════════════════
    #  BUSCA
    # ══════════════════════════════════════════════════════════════════════════
    def _debounce_busca(self, _=None):
        if self._timer_busca:
            self.after_cancel(self._timer_busca)
        self._timer_busca = self.after(380, self._pesquisar)

    def _pesquisar(self, _=None):
        termo = self.entry_busca.get().strip()
        tipo  = self._tipo_ativo

        if not termo and tipo == "TODOS":
            self._limpar_tabela()
            return

        if self.df_completo.empty:
            self.lbl_titulo_res.configure(text="Nenhum dado carregado.")
            return

        df = self.df_completo.copy()

        # ── Filtro texto ──
        if termo:
            t = self._norm(termo)
            mask = (
                df['Desc'].apply(lambda x: t in self._norm(x)) |
                df['Cod'].apply(lambda x: t in self._norm(x))  |
                df['Obra'].apply(lambda x: t in self._norm(x))
            )
            df = df[mask]

        # ── Filtro tipo ──
        if tipo != "TODOS":
            if tipo == "AMBOS":
                df = df[df['Tipo'].str.upper().str.strip() == 'AMBOS']
            else:
                df = df[df['Tipo'].str.upper().str.strip() == tipo]

        self.df_filtro = df
        self._popular_tabela(df)

    def _set_tipo(self, val, cor_ativa, cor_txt):
        self._tipo_ativo = val
        # Atualiza visual dos chips
        for v, (btn, ca, ct) in self._chips.items():
            if v == val:
                btn.configure(fg_color=ca if ca != PRETO else PRETO,
                               text_color=ct)
            else:
                btn.configure(fg_color="#EFEFEF", text_color=CINZA_TXT)
        # Chip "Todos" especial
        if val == "TODOS":
            self._chips["TODOS"][0].configure(fg_color=PRETO, text_color=BRANCO)
        self._pesquisar()

    # ══════════════════════════════════════════════════════════════════════════
    #  TABELA
    # ══════════════════════════════════════════════════════════════════════════
    def _popular_tabela(self, df: pd.DataFrame):
        self.tabela.delete(*self.tabela.get_children())

        icone_tipo = {
            'ACUMULADO':                 '📊  Acumulado',
            'QUANTITATIVA':              '📋  Quantitativa',
            'ACUMULADO e QUANTITATIVA':  '🔄  Ambos',
        }

        MAX = 2000
        for i, (_, row) in enumerate(df.head(MAX).iterrows()):
            tipo_raw = str(row.get('Tipo', ''))
            label    = icone_tipo.get(tipo_raw, tipo_raw)

            if   'ACUMULADO' in tipo_raw and 'QUANTITATIVA' in tipo_raw: tag = "AMBOS"
            elif 'ACUMULADO'    in tipo_raw: tag = "ACUMULADO"
            elif 'QUANTITATIVA' in tipo_raw: tag = "QUANTITATIVA"
            elif i % 2 == 0:                 tag = "PAR"
            else:                            tag = "IMPAR"

            self.tabela.insert("", "end", tags=(tag,), values=(
                label,
                str(row.get('Obra', '')),
                str(row.get('Cod',  '')),
                str(row.get('Desc', '')),
                str(row.get('UN',   '')),
            ))

        total = len(df)
        exib  = min(total, MAX)
        suf   = f"  (exibindo {exib} de {total})" if total > MAX else ""

        self.lbl_titulo_res.configure(text="Resultados")
        self.lbl_contagem.configure(
            text=f"🔎  {total} item{'ns' if total!=1 else ''} encontrado{'s' if total!=1 else ''}{suf}")

    def _limpar_tabela(self):
        self.tabela.delete(*self.tabela.get_children())
        self.df_filtro = pd.DataFrame()
        self.lbl_titulo_res.configure(text="Aguardando pesquisa…")
        self.lbl_contagem.configure(text="")

    def _limpar(self):
        self.entry_busca.delete(0, 'end')
        self._limpar_tabela()
        self.entry_busca.focus()

    # ══════════════════════════════════════════════════════════════════════════
    #  ORDENAÇÃO
    # ══════════════════════════════════════════════════════════════════════════
    def _ordenar(self, col):
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
    #  IMPORTAR PDF
    # ══════════════════════════════════════════════════════════════════════════
    def _importar_thread(self):
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

        os.makedirs(self.pasta_input, exist_ok=True)
        dest = os.path.join(self.pasta_input, os.path.basename(caminho))
        shutil.copy(caminho, dest)

        # Feedback visual
        self.btn_import.configure(state="disabled", text="⏳  Lendo PDF...")
        self.progresso.pack(padx=18, pady=6, fill="x")
        self.progresso.start()
        self.lbl_status.pack(padx=18, pady=(0, 6))

        threading.Thread(target=self._run_ocr, args=(dest,), daemon=True).start()

    def _arquivos_ja_processados(self) -> set:
        """Retorna conjunto dos nomes de arquivo ja no Cofre_Brasul.xlsx."""
        if os.path.exists(self.caminho_xls):
            try:
                df = pd.read_excel(self.caminho_xls, usecols=['Obra_Arq']).fillna('')
                return set(df['Obra_Arq'].astype(str).tolist())
            except Exception:
                pass
        return set()

    def _run_ocr(self, caminho_pdf: str):
        erro_msg    = None
        n_itens     = 0
        duplicado   = False
        try:
            from pathlib import Path
            import time

            nome_arq = os.path.basename(caminho_pdf)

            # Verifica duplicidade ANTES de rodar OCR
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
                    # Monta DataFrame com os itens extraidos
                    rows = []
                    for it in resultado.get('itens', []):
                        rows.append({
                            'Obra':     resultado.get('nome', 'DESCONHECIDA'),
                            'Obra_Arq': resultado.get('arq',  ''),
                            'Tipo':     it.get('tipo',      ''),
                            'Cod':      it.get('codigo',    ''),
                            'Desc':     it.get('descricao', ''),
                            'UN':       it.get('unidade',   ''),
                        })
                    novo = pd.DataFrame(rows)
                    n_itens = len(rows)

                    # Integra ao banco existente
                    if not novo.empty:
                        os.makedirs(self.pasta_output, exist_ok=True)
                        if os.path.exists(self.caminho_xls):
                            existente = pd.read_excel(self.caminho_xls).fillna('')
                            combinado = pd.concat([existente, novo], ignore_index=True)
                            combinado.drop_duplicates(
                                subset=['Obra_Arq', 'Cod', 'Tipo'], inplace=True)
                        else:
                            combinado = novo

                        # Salva com retry - caso o arquivo esteja aberto no Excel
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

        # Volta para a thread principal
        if duplicado:
            self.after(0, self._ocr_duplicado)
        elif erro_msg:
            self.after(0, lambda msg=erro_msg: self._ocr_erro(msg))
        else:
            self.after(0, lambda n=n_itens: self._ocr_ok(n))

    def _ocr_ok(self, n_itens: int):
        self._fim_ocr()
        self._carregar_dados()
        self._atualizar_kpis()
        messagebox.showinfo(
            "Importacao concluida",
            f"PDF processado com sucesso!\n\n"
            f"{n_itens} {'item importado' if n_itens==1 else 'itens importados'} para o Cofre.")
        self._limpar()

    def _ocr_duplicado(self):
        self._fim_ocr()
        messagebox.showwarning(
            "PDF ja importado",
            "Este arquivo ja foi processado e esta no Cofre.\n\n"
            "Para reprocessar, remova o arquivo do Cofre_Brasul.xlsx primeiro.")

    def _ocr_erro(self, msg: str):
        self._fim_ocr()
        messagebox.showerror("❌  Erro no processamento", msg)

    def _fim_ocr(self):
        self.progresso.stop()
        self.progresso.pack_forget()
        self.lbl_status.pack_forget()
        self.btn_import.configure(state="normal", text="⬇   IMPORTAR PDF")

    # ══════════════════════════════════════════════════════════════════════════
    #  EXPORTAR
    # ══════════════════════════════════════════════════════════════════════════
    def _exportar(self):
        if self.df_filtro.empty:
            messagebox.showwarning(
                "Sem dados para exportar",
                "Realize uma busca primeiro para exportar os resultados.")
            return
        path = filedialog.asksaveasfilename(
            title="Salvar resultado",
            defaultextension=".xlsx",
            initialfile="Busca_Cofre_Brasul.xlsx",
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

    # ══════════════════════════════════════════════════════════════════════════
    #  RECARREGAR
    # ══════════════════════════════════════════════════════════════════════════
    def _recarregar(self):
        self._carregar_dados()
        self._atualizar_kpis()
        self._limpar()
        total = len(self.df_completo)
        messagebox.showinfo(
            "✅  Dados atualizados",
            f"Banco recarregado com {total} registros.")


# ─────────────────────────────────────────────────────────────────────────────
#  ENTRY POINT
# ─────────────────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    app = CofreBrasul()
    app.mainloop()
