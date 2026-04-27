"""
tela_tabulacao.py  —  Pipeline completo: Codificação → Tabulação → PPT
=======================================================================
Integra ao app.py existente:
    from tela_tabulacao import TelaTabulacao
    TelaTabulacao(self.root, self.codificador, self.banco)

Fluxo em 6 etapas visuais:
  1. Upload da base crua SurveyMonkey
  2. IA sugere colunas abertas -> usuário confirma
  3. IA codifica as colunas selecionadas
  4. Revisão humana das codificações (TelaRevisao existente)
  5. Detecção do tipo de todas as perguntas -> usuário revisa
  6. Gerar Excel + PowerPoint com 1 clique
"""

import re
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from pathlib import Path

import pandas as pd

# ─── Módulos do sistema ───────────────────────────────────────────────────────
try:
    from tabulador   import carregar_base, detectar_perguntas, exportar_excel, TIPOS_LABEL
    from gerador_ppt import gerar_ppt
    from tela_revisao import TelaRevisao
    from codificador  import CodificadorIA, TIPOS_PERGUNTA
    _OK = True
except ImportError as _e:
    _OK = False
    _ERR = str(_e)
    TIPOS_PERGUNTA = {}
    TIPOS_LABEL    = {}

# ─── Paleta (idêntica ao app.py) ─────────────────────────────────────────────
NAV_ACTIVE = "#1d4ed8"
BG         = "#f9fafb"
CARD       = "#ffffff"
BORDER     = "#e5e7eb"
BORDER2    = "#d1d5db"
AZUL       = "#1d4ed8"
AZUL_DARK  = "#1e3a8a"
AZUL_LIGHT = "#eff6ff"
AZUL_MID   = "#bfdbfe"
VERDE      = "#059669"
VERDE_LIGHT= "#ecfdf5"
ROXO       = "#7c3aed"
OURO       = "#d97706"
OURO_LIGHT = "#fffbeb"
TXT1="#111827"; TXT2="#374151"; TXT3="#6b7280"; TXT4="#9ca3af"; TXT5="#d1d5db"
F_SEC  = ("Segoe UI", 10, "bold")
F_BODY = ("Segoe UI", 9)
F_SMALL= ("Segoe UI", 8)
LOG_BG ="#0d1117"; LOG_GRN="#3fb950"; LOG_BLU="#79c0ff"; LOG_RED="#f85149"

_TIPOS_COD = {k: v["label"] for k, v in TIPOS_PERGUNTA.items()}

# ─── Regex para detecção de tipo de coluna aberta ────────────────────────────
_RE_MARC = re.compile(r"marca|instituição|lembra.*visto|quais.*marcas", re.I)
_RE_PAL  = re.compile(r"uma palavra|defina.*palavra|palavra.*associa", re.I)
_RE_SAT  = re.compile(r"por que avalia|motivo|justif|como você avalia", re.I)
_RE_NPS  = re.compile(r"de 0 a 10|probabilidade.*recomendar", re.I)
_IGNORAR = {
    "respondent_id","collector_id","date_created","date_modified",
    "ip_address","email_address","first_name","last_name","custom_1",
    "Pesquisador:","Entrevistador:","Comentário:","Comentário","Observações:",
    "Observaçõees","Setor:","Setor","ip_address",
}

# Valores inválidos que não devem ser codificados
_INVALIDOS_COD = {
    "", "-", "nan", "none", "r:", "n/a", "na",
    "não soube responder", "nao soube responder",
    "não respondeu", "nao respondeu",
    "não se aplica", "nao se aplica",
    "sem resposta", "não informado", "não informado",
}


# â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•
class TelaTabulacao(tk.Frame):
    """
    Abre o pipeline completo como janela filha do app.py.

    Uso no app.py:
        from tela_tabulacao import TelaTabulacao
        TelaTabulacao(self.root, self.codificador, self.banco)
    """

    def __init__(self, parent, codificador=None, banco=None, inline=False):
        super().__init__(parent, bg=BG)

        self.inline = inline
        self.parent = parent

        if not inline:
            self.master.title("Pipeline de Pesquisa  |  IFec RJ")
            sw, sh = self.winfo_screenwidth(), self.winfo_screenheight()
            w = min(int(sw * 0.92), 1380)
            h = min(int(sh * 0.90), 900)
            self.master.geometry(f"{w}x{h}+{(sw-w)//2}+{(sh-h)//2}")
            self.master.configure(bg=BG)
            self.master.resizable(True, True)

        self._codificador = codificador or (CodificadorIA() if _OK else None)
        self._banco = banco
        self._df: pd.DataFrame | None = None
        self._arquivo = ""
        self._cols_cod: list[dict] = []
        self._perguntas: list[dict] = []
        self._etapa = 0

        self._build()

        if not _OK:
            messagebox.showerror(
                "Módulo ausente",
                f"Não foi possível importar módulos:\n{_ERR}\n\n"
                "Verifique se tabulador.py, gerador_ppt.py e tela_revisao.py "
                "estão na mesma pasta que app.py."
            )

    # --- Layout principal ─────────────────────────────────────────────────────
    def _build(self):
        # Topbar
        top = tk.Frame(self, bg=CARD, height=58)
        top.pack(fill=tk.X)
        top.pack_propagate(False)
        tk.Frame(top, bg=AZUL, width=5).pack(side=tk.LEFT, fill=tk.Y)
        tf = tk.Frame(top, bg=CARD, padx=16)
        tf.pack(side=tk.LEFT, fill=tk.Y)
        tk.Label(tf, text="Pipeline de Pesquisa",
                 bg=CARD, fg=TXT1,
                 font=("Segoe UI", 15, "bold")).pack(anchor="w", pady=(10, 0))
        tk.Label(tf,
                 text="Codificação com IA  ->  Tabulação  ->  Excel + PowerPoint  |  IFec RJ",
                 bg=CARD, fg=TXT4, font=F_SMALL).pack(anchor="w")
        tk.Frame(top, bg=BORDER, height=1).pack(side=tk.BOTTOM, fill=tk.X)

        # Stepper
        self._build_stepper()

        # Corpo com notebook
        body = tk.Frame(self, bg=BG)
        body.pack(fill=tk.BOTH, expand=True, padx=14, pady=(8, 4))
        body.grid_columnconfigure(0, weight=1)
        body.grid_rowconfigure(0, weight=1)

        self._nb = ttk.Notebook(body)
        self._nb.grid(row=0, column=0, sticky="nsew")

        self._tab1 = self._nova_aba("1  Upload")
        self._tab2 = self._nova_aba("2  Colunas abertas")
        self._tab3 = self._nova_aba("3  Tabulacao")

        self._build_tab1()
        self._build_tab2()
        self._build_tab3()

        # Log
        lf = tk.Frame(body, bg=LOG_BG)
        lf.grid(row=1, column=0, sticky="ew", pady=(4, 0))
        self._log_w = tk.Text(lf, height=4, bg=LOG_BG, fg="#e6edf3",
                               font=("Consolas", 8), relief="flat",
                               state="disabled", wrap=tk.WORD)
        self._log_w.pack(fill=tk.BOTH, padx=8, pady=4)
        for tag, fg in [("ok", LOG_GRN), ("err", LOG_RED), ("inf", LOG_BLU)]:
            self._log_w.tag_config(tag, foreground=fg)

    def _nova_aba(self, label):
        f = tk.Frame(self._nb, bg=BG)
        self._nb.add(f, text=f"  {label}  ")
        return f

    # ─── Stepper ─────────────────────────────────────────────────────────────
    def _build_stepper(self):
        sf = tk.Frame(self, bg=CARD, pady=8)
        sf.pack(fill=tk.X)
        tk.Frame(sf, bg=BORDER, height=1).pack(fill=tk.X, side=tk.BOTTOM)
        row = tk.Frame(sf, bg=CARD)
        row.pack()
        steps = ["1 Upload", "2 Detectar colunas", "3 Codificar",
                 "4 Revisar", "5 Conferir perguntas", "6 Gerar"]
        self._steps = []
        for i, s in enumerate(steps):
            lbl = tk.Label(row, text=s, bg=CARD, fg=TXT4,
                           font=("Segoe UI", 8, "bold"), padx=10)
            lbl.pack(side=tk.LEFT)
            self._steps.append(lbl)
            if i < len(steps) - 1:
                tk.Label(row, text="->", bg=CARD, fg=TXT5,
                         font=F_SMALL).pack(side=tk.LEFT)
        self._go(0)

    def _go(self, n):
        self._etapa = n
        for i, lbl in enumerate(self._steps):
            if i < n:   lbl.config(fg=VERDE,  font=("Segoe UI", 8, "bold"))
            elif i == n: lbl.config(fg=AZUL,   font=("Segoe UI", 8, "bold"))
            else:        lbl.config(fg=TXT4,   font=("Segoe UI", 8))

    # ─── Tab 1: Upload ────────────────────────────────────────────────────────
    def _build_tab1(self):
        f = self._tab1
        f.grid_columnconfigure(0, weight=1)

        # Card arquivo + metadados
        card = tk.Frame(f, bg=CARD, padx=18, pady=16)
        card.grid(row=0, column=0, sticky="ew", padx=16, pady=(16, 8))
        card.grid_columnconfigure(1, weight=1)

        tk.Label(card, text="Base SurveyMonkey (.xlsx):",
                 bg=CARD, fg=TXT2, font=F_SEC).grid(row=0, column=0, sticky="w")
        self._lbl_arq = tk.Label(card,
                                  text="  Nenhum arquivo selecionado",
                                  bg=AZUL_LIGHT, fg=TXT3,
                                  font=F_BODY, padx=10, pady=5, anchor="w")
        self._lbl_arq.grid(row=0, column=1, sticky="ew", padx=(10, 10))
        tk.Button(card, text="Selecionar",
                  bg=AZUL, fg="white", font=F_BODY,
                  relief="flat", padx=12, pady=5, cursor="hand2",
                  activebackground=AZUL_DARK, activeforeground="white",
                  command=self._selecionar
                  ).grid(row=0, column=2)

        for row_i, (lbl, var_name, default) in enumerate([
            ("Titulo da pesquisa:", "_v_titulo", "Pesquisa IFec RJ"),
            ("Subtitulo / periodo:", "_v_sub",   ""),
        ], start=1):
            tk.Label(card, text=lbl, bg=CARD, fg=TXT2,
                     font=F_BODY).grid(row=row_i, column=0, sticky="w",
                                        pady=(10 if row_i==1 else 6, 0))
            v = tk.StringVar(value=default)
            setattr(self, var_name, v)
            tk.Entry(card, textvariable=v, font=F_BODY, bg=BG,
                     relief="flat", highlightthickness=1,
                     highlightbackground=BORDER2
                     ).grid(row=row_i, column=1, sticky="ew",
                             padx=(10, 10),
                             pady=(10 if row_i==1 else 6, 0))

        # Métricas
        mf = tk.Frame(f, bg=BG)
        mf.grid(row=1, column=0, sticky="ew", padx=16, pady=(0, 8))
        mf.grid_columnconfigure((0, 1, 2), weight=1)
        self._mw = {}
        for i, (key, val, sub, bg, fg) in enumerate([
            ("respondentes", "-", "respondentes",   AZUL_LIGHT,  AZUL),
            ("colunas",      "-", "colunas totais", VERDE_LIGHT, VERDE),
            ("status",       "-", "status",         OURO_LIGHT,  OURO),
        ]):
            c = tk.Frame(mf, bg=CARD, padx=12, pady=8)
            c.grid(row=0, column=i, sticky="nsew",
                   padx=(0, 10 if i < 2 else 0))
            tk.Frame(c, bg=fg, height=3).pack(fill=tk.X)
            lv = tk.Label(c, text=val, bg=CARD, fg=TXT1,
                          font=("Segoe UI", 20, "bold"))
            lv.pack(anchor="w")
            ls = tk.Label(c, text=sub, bg=CARD, fg=TXT4, font=F_SMALL)
            ls.pack(anchor="w")
            self._mw[key] = (lv, ls)

        # Botão avançar
        self._btn_av1 = tk.Button(
            f, text="Avançar -> Detectar colunas abertas",
            bg=AZUL, fg="white", font=("Segoe UI", 10, "bold"),
            relief="flat", padx=16, pady=8, cursor="hand2",
            state="disabled",
            activebackground=AZUL_DARK, activeforeground="white",
            command=self._ir_colunas)
        self._btn_av1.grid(row=2, column=0, sticky="w", padx=16, pady=8)

    # ─── Tab 2: Colunas abertas ───────────────────────────────────────────────
    def _build_tab2(self):
        f = self._tab2
        f.grid_columnconfigure(0, weight=1)
        f.grid_rowconfigure(1, weight=1)

        # Cabeçalho
        hf = tk.Frame(f, bg=CARD, padx=16, pady=10)
        hf.grid(row=0, column=0, sticky="ew")
        tk.Label(hf, text="Colunas com texto aberto detectadas",
                 bg=CARD, fg=TXT1, font=F_SEC).pack(side=tk.LEFT)
        tk.Label(hf,
                 text="  Duplo clique ou Espaço para marcar/desmarcar.",
                 bg=CARD, fg=TXT4, font=F_SMALL).pack(side=tk.LEFT)
        for txt, cmd in [
            ("+ Adicionar coluna",  self._adicionar_coluna),
            ("− Remover selecionada", self._remover_coluna),
            ("Marcar tudo",         self._marcar_tudo),
            ("Desmarcar tudo",      self._desmarcar_tudo),
        ]:
            bg = VERDE_LIGHT if "Adicionar" in txt else (
                 "#fef2f2"   if "Remover"   in txt else AZUL_LIGHT)
            fg = VERDE if "Adicionar" in txt else (
                 "#dc2626" if "Remover" in txt else AZUL)
            tk.Button(hf, text=txt, bg=bg, fg=fg,
                      font=F_SMALL, relief="flat", padx=6, pady=2,
                      cursor="hand2", command=cmd).pack(side=tk.RIGHT, padx=2)

        # Treeview
        tvf = tk.Frame(f, bg=CARD, padx=12, pady=8)
        tvf.grid(row=1, column=0, sticky="nsew", padx=14, pady=(0, 6))
        tvf.grid_columnconfigure(0, weight=1)
        tvf.grid_rowconfigure(0, weight=1)

        cols = ("cod", "coluna", "tipo", "n", "amostra")
        self._tv2 = ttk.Treeview(tvf, columns=cols,
                                  show="headings", selectmode="browse")
        for col, hdr, w in [
            ("cod",    "Codificar?", 80),
            ("coluna", "Coluna",     300),
            ("tipo",   "Tipo IA",    170),
            ("n",      "Respostas",  80),
            ("amostra","Amostra",    340),
        ]:
            self._tv2.heading(col, text=hdr)
            self._tv2.column(col, width=w,
                              stretch=(col in ("coluna", "amostra")))
        vsb = ttk.Scrollbar(tvf, orient="vertical", command=self._tv2.yview)
        self._tv2.configure(yscrollcommand=vsb.set)
        self._tv2.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        self._tv2.tag_configure("sim", background="#f0fdf4")
        self._tv2.tag_configure("nao", background="#f9fafb", foreground=TXT4)
        self._tv2.bind("<Double-1>", self._toggle_col)
        self._tv2.bind("<space>",    self._toggle_col)

        # Botão codificar
        bf = tk.Frame(f, bg=BG, padx=14, pady=8)
        bf.grid(row=2, column=0, sticky="ew")
        self._btn_cod = tk.Button(
            bf, text="⚡ Codificar com IA →",
            bg=ROXO, fg="white", font=("Segoe UI", 10, "bold"),
            relief="flat", padx=16, pady=8, cursor="hand2",
            state="disabled",
            activebackground="#6d28d9", activeforeground="white",
            command=self._codificar)
        self._btn_cod.pack(side=tk.LEFT)
        self._lbl_cod = tk.Label(bf, text="", bg=BG, fg=TXT3, font=F_SMALL)
        self._lbl_cod.pack(side=tk.LEFT, padx=10)
        self._prog2 = ttk.Progressbar(bf, mode="indeterminate", length=180)
        self._prog2.pack(side=tk.LEFT)

    # ///////// Tab 3: Tabulação \\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    def _build_tab3(self):
        f = self._tab3
        f.grid_columnconfigure(0, weight=1)
        f.grid_rowconfigure(1, weight=1)

        hf = tk.Frame(f, bg=CARD, padx=16, pady=10)
        hf.grid(row=0, column=0, sticky="ew")
        tk.Label(hf, text="Perguntas para tabulação",
                 bg=CARD, fg=TXT1, font=F_SEC).pack(side=tk.LEFT)
        tk.Label(hf,
                 text="  Duplo clique para editar tipo/nota. "
                      "Espaçoo para ativar/desativar.",
                 bg=CARD, fg=TXT4, font=F_SMALL).pack(side=tk.LEFT)
        for txt, cmd in [("Ativar tudo", self._ativar_tudo),
                          ("Desativar tudo", self._desativar_tudo)]:
            tk.Button(hf, text=txt, bg=AZUL_LIGHT, fg=AZUL,
                      font=F_SMALL, relief="flat", padx=6, pady=2,
                      cursor="hand2", command=cmd).pack(side=tk.RIGHT, padx=2)

        tvf = tk.Frame(f, bg=CARD, padx=12, pady=8)
        tvf.grid(row=1, column=0, sticky="nsew", padx=14, pady=(0, 6))
        tvf.grid_columnconfigure(0, weight=1)
        tvf.grid_rowconfigure(0, weight=1)

        cols3 = ("ativo", "num", "tipo", "pergunta", "nota")
        self._tv3 = ttk.Treeview(tvf, columns=cols3,
                                  show="headings", selectmode="browse")
        for col, hdr, w in [
            ("ativo",   "✓",        42),
            ("num",     "N°",       52),
            ("tipo",    "Tipo",    115),
            ("pergunta","Pergunta", 460),
            ("nota",    "Nota",    200),
        ]:
            self._tv3.heading(col, text=hdr)
            self._tv3.column(col, width=w,
                              stretch=(col == "pergunta"))
        vsb3 = ttk.Scrollbar(tvf, orient="vertical", command=self._tv3.yview)
        self._tv3.configure(yscrollcommand=vsb3.set)
        self._tv3.grid(row=0, column=0, sticky="nsew")
        vsb3.grid(row=0, column=1, sticky="ns")
        self._tv3.tag_configure("ativa",  background="#f0fdf4")
        self._tv3.tag_configure("inativa",background="#f9fafb", foreground=TXT4)
        self._tv3.tag_configure("aberta", background="#fefce8")
        self._tv3.bind("<Double-1>", self._editar_perg)
        self._tv3.bind("<space>",    self._toggle_perg)

        # Botões gerar
        bf = tk.Frame(f, bg=BG, padx=14, pady=10)
        bf.grid(row=2, column=0, sticky="ew")
        self._btn_gerar = tk.Button(
            bf, text="Gerar Excel + PowerPoint",
            bg=VERDE, fg="white", font=("Segoe UI", 10, "bold"),
            relief="flat", padx=16, pady=8, cursor="hand2",
            state="disabled",
            activebackground="#047857", activeforeground="white",
            command=self._gerar)
        self._btn_gerar.pack(side=tk.LEFT)
        self._btn_xls = tk.Button(
            bf, text="Só Excel",
            bg=AZUL_LIGHT, fg=AZUL, font=F_BODY,
            relief="flat", padx=10, pady=8, cursor="hand2",
            state="disabled",
            command=lambda: self._gerar(ppt=False))
        self._btn_xls.pack(side=tk.LEFT, padx=(8, 0))
        self._btn_ppt = tk.Button(
            bf, text="Só PowerPoint",
            bg=AZUL_LIGHT, fg=AZUL, font=F_BODY,
            relief="flat", padx=10, pady=8, cursor="hand2",
            state="disabled",
            command=lambda: self._gerar(excel=False))
        self._btn_ppt.pack(side=tk.LEFT, padx=(8, 0))
        self._prog3 = ttk.Progressbar(bf, mode="indeterminate", length=180)
        self._prog3.pack(side=tk.RIGHT)
        self._lbl_g = tk.Label(bf, text="", bg=BG, fg=TXT3, font=F_SMALL)
        self._lbl_g.pack(side=tk.RIGHT, padx=6)

    # ////////////// Helpers \\\\\\\\\\\\\\\\\\\\\\\\\\\
    def _lg(self, msg, tag=""):
        self._log_w.configure(state="normal")
        self._log_w.insert(tk.END, msg + "\n", tag)
        self._log_w.see(tk.END)
        self._log_w.configure(state="disabled")

    def _upd(self, key, val, sub=None):
        if key in self._mw:
            self._mw[key][0].config(text=str(val))
            if sub is not None:
                self._mw[key][1].config(text=sub)

    def _set_btns_gerar(self, state):
        for b in (self._btn_gerar, self._btn_xls, self._btn_ppt):
            b.config(state=state)

    # ///////////////////////////////////////////////////////////
    # ETAPA 1 | ” Selecionar e carregar arquivo
    # //////////////////////////////////////////////////////////
    def _selecionar(self):
        path = filedialog.askopenfilename(
            title="Selecionar base SurveyMonkey",
            filetypes=[("Excel", "*.xlsx"), ("Todos", "*.*")])
        if not path:
            return
        self._arquivo = path
        self._lbl_arq.config(text=f"  {Path(path).name}", fg=TXT1)
        self._v_titulo.set(Path(path).stem.replace("_", " "))
        self._lg(f"Carregando: {Path(path).name}", "inf")
        threading.Thread(target=self._carregar_t, daemon=True).start()

    def _carregar_t(self):
        try:
            df = carregar_base(self._arquivo)
            self._df = df
            self.after(0, lambda: self._carregar_ok(df))
        except Exception as e:
            self.after(0, lambda: self._lg(f"ERRO ao carregar: {e}", "err"))

    def _carregar_ok(self, df):
        self._upd("respondentes", len(df), "respondentes")
        self._upd("colunas", len(df.columns), "colunas totais")
        self._upd("status", "Pronto", "arquivo carregado")
        self._btn_av1.config(state="normal")
        self._go(1)
        self._lg(f"✓ {len(df)} respondentes  |  {len(df.columns)} colunas", "ok")

    # \\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    # ETAPA 2 | Detectar colunas abertas
    # \\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    def _ir_colunas(self):
        self._nb.select(self._tab2)
        self._go(1)
        self._btn_cod.config(state="disabled")
        threading.Thread(target=self._detectar_cols_t, daemon=True).start()

    def _detectar_cols_t(self):
        try:
            # Passo 1: heurística rápida
            self.after(0, lambda: self._lbl_cod.config(
                text="Identificando colunas abertas..."))
            cols = _detectar_colunas_abertas(self._df)

            # Passo 2: IA classifica o tipo de cada coluna
            if cols and self._codificador:
                self.after(0, lambda: self._lbl_cod.config(
                    text=f"IA classificando {len(cols)} coluna(s)..."))
                cols = _classificar_tipos_com_ia(cols, self._codificador)

            self.after(0, lambda: self._lbl_cod.config(text=""))
            self.after(0, lambda: self._popular_tv2(cols))
        except Exception as e:
            self.after(0, lambda: self._lg(f"ERRO detecção colunas: {e}", "err"))
            self.after(0, lambda: self._lbl_cod.config(text=""))

    def _popular_tv2(self, cols):
        self._cols_cod = cols
        self._tv2.delete(*self._tv2.get_children())
        for c in cols:
            tag = "sim" if c["codificar"] else "nao"
            self._tv2.insert("", tk.END, values=(
                "✓" if c["codificar"] else "✗",
                c["coluna"][:80],
                _TIPOS_COD.get(c["tipo"], c["tipo"]),
                c["n"],
                c["amostra"][:70],
            ), tags=(tag,))
        self._btn_cod.config(state="normal")
        n_sug = sum(1 for c in cols if c["codificar"])
        self._lg(f"✓ {len(cols)} colunas abertas  |  {n_sug} sugeridas", "ok")

    def _toggle_col(self, event=None):
        item = (self._tv2.identify_row(event.y)
                if event and hasattr(event, "y") else self._tv2.focus())
        if not item:
            return
        idx = self._tv2.index(item)
        c   = self._cols_cod[idx]
        c["codificar"] = not c["codificar"]
        tag  = "sim" if c["codificar"] else "nao"
        vals = self._tv2.item(item, "values")
        self._tv2.item(item,
                        values=("✓" if c["codificar"] else "✗",) + vals[1:],
                        tags=(tag,))

    def _marcar_tudo(self):
        for c in self._cols_cod:
            c["codificar"] = True
        self._popular_tv2(self._cols_cod)

    def _desmarcar_tudo(self):
        for c in self._cols_cod:
            c["codificar"] = False
        self._popular_tv2(self._cols_cod)

    def _adicionar_coluna(self):
        """Permite ao usuário adicionar uma coluna que a IA não detectou."""
        if self._df is None:
            return
        todas = list(self._df.columns)
        ja_tem = {c["coluna"] for c in self._cols_cod}
        disponiveis = [c for c in todas if c not in ja_tem
                       and c not in _IGNORAR
                       and not c.startswith("Original_")
                       and not c.endswith("_cod")]
        if not disponiveis:
            messagebox.showinfo("Adicionar coluna",
                                "Todas as colunas disponíveis já estão na lista.")
            return
        _SeletorColuna(self, disponiveis, self._df, self._cols_cod,
                       callback=lambda: self._popular_tv2(self._cols_cod))

    def _remover_coluna(self):
        """Remove a coluna selecionada da lista."""
        item = self._tv2.focus()
        if not item:
            messagebox.showwarning("Remover", "Selecione uma coluna na lista.")
            return
        idx = self._tv2.index(item)
        col_nome = self._cols_cod[idx]["coluna"]
        if messagebox.askyesno("Remover coluna",
                               f"Remover '{col_nome[:60]}' da lista?"):
            self._cols_cod.pop(idx)
            self._popular_tv2(self._cols_cod)

    # /////////////////////////////////////
    # ETAPA 3 | Codificaçã com IA
    # /////////////////////////////////////
    def _codificar(self):
        sel = [c for c in self._cols_cod if c["codificar"]]
        if not sel:
            messagebox.showwarning("Nenhuma coluna",
                                   "Marque ao menos uma coluna para codificar.")
            return
        self._btn_cod.config(state="disabled")
        self._prog2.start(12)
        self._go(2)
        self._lg(f"Codificando {len(sel)} coluna(s) com IA...", "inf")
        threading.Thread(target=self._codificar_t, args=(sel,), daemon=True).start()

    def _codificar_t(self, sel):
        fila = []
        try:
            for col_info in sel:
                col  = col_info["coluna"]
                tipo = col_info["tipo"]

                # ///// Filtrar respostas válidas \\\\\\\\\\\\\
                _INVALIDOS = _INVALIDOS_COD
                try:
                    dado = self._df[col]
                    # Se retornar DataFrame (coluna duplicada), pegar a primeira
                    if isinstance(dado, pd.DataFrame):
                        dado = dado.iloc[:, 0]
                    serie = (dado
                             .dropna()
                             .astype(str)
                             .pipe(lambda s: s[~s.str.strip().str.lower().isin(_INVALIDOS)])
                             .pipe(lambda s: s[~s.str.strip().str.upper().str.startswith("NÃƒO SE APLICA")])
                             .pipe(lambda s: s[s.str.strip().str.len() > 1]))
                    respostas = serie.tolist()
                except Exception as e_read:
                    self.after(0, lambda c=col, e=e_read:
                        self._lg(f"  ERRO lendo '{c[:40]}': {e}", "err"))
                    continue

                if not respostas:
                    self.after(0, lambda c=col:
                        self._lg(f"  AVISO: '{c[:40]}' â€” sem respostas válidas para codificar", "inf"))
                    continue

                n_resp = len(respostas)
                self.after(0, lambda c=col, n=n_resp, tp=tipo:
                    self._lg(f"  Codificando: '{c[:50]}' | {n} respostas | tipo: {tp}", "inf"))
                self.after(0, lambda c=col, n=n_resp:
                    self._lbl_cod.config(text=f"'{c[:28]}...' ({n} respostas)"))

                # ///////////// Codificar ////////////////
                try:
                    def _cb(i, t, r, cat, _col=col):
                        self.after(0, lambda: self._lbl_cod.config(
                            text=f"'{_col[:22]}' {i+1}/{t}"))

                    res = self._codificador.codificar_lote_modo(
                        respostas, tipo=tipo, modo="simples",
                        callback_progresso=_cb)

                except Exception as e_api:
                    self.after(0, lambda c=col, e=e_api:
                        self._lg(f"  ERRO API '{c[:40]}': {e}", "err"))
                    continue

                # //////////// Aplicar resultado à base ////////////////////
                codificados = res.get("resultado", [])

                if not codificados:
                    self.after(0, lambda c=col:
                        self._lg(f"  AVISO: '{c[:40]}' — nenhum resultado retornado pela IA", "err"))
                    continue

                # Alinhar tamanhos se necessário
                if len(codificados) != n_resp:
                    self.after(0, lambda c=col, nr=n_resp, nc=len(codificados):
                        self._lg(f"  AVISO '{c[:30]}': {nr} respostas vs {nc} codificados", "inf"))
                    codificados = (codificados + [""] * n_resp)[:n_resp]

                col_nova = col + "_cod"
                mapa     = dict(zip(respostas, codificados))
                self._df[col_nova] = self._df[col].apply(
                    lambda v: mapa.get(str(v).strip(), "") if pd.notna(v) else "")

                n_cod = sum(1 for c_ in codificados if c_ and c_ != "SEM_RESPOSTA")
                self.after(0, lambda c=col, cn=col_nova, n=n_cod:
                    self._lg(f"  ✓ '{c[:40]}' → '{cn[:40]}' ({n} codificados)", "ok"))

                # ////////////// Preparar revisão //////////////////
                itens = [{"resposta": r, "categoria": c_}
                         for r, c_ in zip(respostas, codificados)
                         if c_ and c_ not in ("SEM_RESPOSTA", "ERRO")]
                if self._banco and itens:
                    s2 = self._banco.selecionar_para_revisao(itens, n=5)
                elif itens:
                    import random
                    s2 = random.sample(itens, min(5, len(itens)))
                else:
                    s2 = []
                if s2:
                    fila.append({"aba": col[:40], "tipo": tipo, "exemplos": s2})

            self.after(0, lambda: self._cod_ok(fila))

        except Exception as e:
            self.after(0, lambda: self._lg(f"ERRO codificação: {e}", "err"))
            self.after(0, self._prog2.stop)
            self.after(0, lambda: self._btn_cod.config(state="normal"))


    def _cod_ok(self, fila):
        self._prog2.stop()
        self._lbl_cod.config(text="Concluído ✓")
        self._go(3)
        self._lg("✓ Codificação concluída. Abrindo revisão...", "ok")

        if fila and self._banco:
            sheets_fake = {item["aba"]: self._df for item in fila}
            TelaRevisao(self, self._banco, fila,
                        sheets=sheets_fake,
                        on_close=self._apos_revisao)
        else:
            self._apos_revisao()

    # /////////////////////////////////////////////////
    # ETAPA 4 | ApÃ³s revisão â†’ detectar perguntas
    # /////////////////////////////////////////////////
    def _apos_revisao(self):
        self._go(4)
        self._lg("Detectando perguntas para tabulação...", "inf")
        threading.Thread(target=self._detectar_pergs_t, daemon=True).start()

    def _detectar_pergs_t(self):
        try:
            pergs = detectar_perguntas(self._df)
            self.after(0, lambda: self._popular_tv3(pergs))
        except Exception as e:
            self.after(0, lambda: self._lg(f"ERRO perguntas: {e}", "err"))

    def _popular_tv3(self, pergs):
        self._perguntas = pergs
        self._tv3.delete(*self._tv3.get_children())
        for p in pergs:
            ativo = p.get("ativo", True)
            tag   = ("aberta"  if p["tipo"] == "ABERTA" else
                     "ativa"   if ativo else "inativa")
            self._tv3.insert("", tk.END, values=(
                "✓" if ativo else "✗",
                p["num"],
                TIPOS_LABEL.get(p["tipo"], p["tipo"]),
                p["pergunta"][:90],
                p.get("nota", ""),
            ), tags=(tag,))
        n_ativas = sum(1 for p in pergs
                       if p.get("ativo") and p["tipo"] != "IGNORAR")
        self._nb.select(self._tab3)
        self._go(4)
        self._set_btns_gerar("normal")
        self._lg(f"✓ {len(pergs)} perguntas  |  {n_ativas} ativas", "ok")

    # ─── Editar / toggle perguntas ────────────────────────────────────────────
    def _editar_perg(self, event=None):
        item = (self._tv3.identify_row(event.y)
                if event else self._tv3.focus())
        if not item:
            return
        idx = self._tv3.index(item)
        p   = self._perguntas[idx]
        _EditorPergunta(self, p,
                        callback=lambda: self._atualizar_p(idx, item))

    def _atualizar_p(self, idx, item):
        p     = self._perguntas[idx]
        ativo = p.get("ativo", True)
        tag   = ("aberta"  if p["tipo"] == "ABERTA" else
                 "ativa"   if ativo else "inativa")
        self._tv3.item(item, values=(
            "✓" if ativo else "✗", p["num"],
            TIPOS_LABEL.get(p["tipo"], p["tipo"]),
            p["pergunta"][:90], p.get("nota", ""),
        ), tags=(tag,))

    def _toggle_perg(self, event=None):
        item = (self._tv3.identify_row(event.y)
                if event and hasattr(event, "y") else self._tv3.focus())
        if not item:
            return
        idx = self._tv3.index(item)
        p   = self._perguntas[idx]
        p["ativo"] = not p.get("ativo", True)
        self._atualizar_p(idx, item)

    def _ativar_tudo(self):
        for p in self._perguntas:
            p["ativo"] = True
        self._popular_tv3(self._perguntas)

    def _desativar_tudo(self):
        for p in self._perguntas:
            p["ativo"] = False
        self._popular_tv3(self._perguntas)

    # ///////////////////////////////
    # ETAPA 5 | Gerar Excel + PPT
    # ///////////////////////////////
    def _gerar(self, excel=True, ppt=True):
        ativas = [p for p in self._perguntas
                  if p.get("ativo") and p["tipo"] != "IGNORAR"]
        if not ativas:
            messagebox.showwarning("Sem perguntas", "Ative ao menos uma pergunta.")
            return

        saida_dir = filedialog.askdirectory(title="Pasta de saída")
        if not saida_dir:
            return

        titulo    = self._v_titulo.get().strip() or "Pesquisa IFec RJ"
        subtitulo = self._v_sub.get().strip()
        stem      = Path(self._arquivo).stem
        saida_xlsx = str(Path(saida_dir) / f"Tabulação_{stem}.xlsx")
        saida_pptx = str(Path(saida_dir) / f"Apresentação_{stem}.pptx")

        self._set_btns_gerar("disabled")
        self._prog3.start(12)
        self._go(5)

        def _run():
            try:
                # ── Sanitiza o DataFrame antes de exportar ────────────────────
                # O SurveyMonkey gera colunas duplicadas que viram MultiIndex
                # ou colunas onde o valor é uma Series — isso quebra o openpyxl.
                df_exp = self._df.copy()

                # Achatar MultiIndex de colunas
                if isinstance(df_exp.columns, pd.MultiIndex):
                    df_exp.columns = [
                        "_".join(str(c) for c in col).strip("_")
                        for col in df_exp.columns
                    ]

                # Garantir nomes de colunas únicos
                cols_seen: dict = {}
                new_cols = []
                for c in df_exp.columns:
                    c_str = str(c)
                    if c_str in cols_seen:
                        cols_seen[c_str] += 1
                        new_cols.append(f"{c_str}_{cols_seen[c_str]}")
                    else:
                        cols_seen[c_str] = 0
                        new_cols.append(c_str)
                df_exp.columns = new_cols

                # Colapsar qualquer coluna que ainda seja DataFrame/Series aninhada
                for col in df_exp.columns:
                    if isinstance(df_exp[col], pd.DataFrame):
                        df_exp[col] = df_exp[col].iloc[:, 0]

                # Converter tudo para tipos simples (evita int() on Series)
                for col in df_exp.columns:
                    try:
                        df_exp[col] = df_exp[col].astype(object)
                    except Exception:
                        pass

                if excel:
                    self.after(0, lambda: self._lbl_g.config(text="Gerando Excel..."))
                    exportar_excel(df_exp, ativas, saida=saida_xlsx,
                                   titulo=titulo, total_respostas=len(df_exp))
                    self.after(0, lambda: self._lg(
                        f"✓ Excel: {Path(saida_xlsx).name}", "ok"))
                if ppt:
                    self.after(0, lambda: self._lbl_g.config(text="Gerando PPT..."))
                    gerar_ppt(df_exp, ativas, saida=saida_pptx,
                              titulo=titulo, subtitulo=subtitulo)
                    self.after(0, lambda: self._lg(
                        f"✓ PPT: {Path(saida_pptx).name}", "ok"))
                self.after(0, _done)
            except Exception as e:
                self.after(0, lambda: self._lg(f"ERRO: {e}", "err"))
                self.after(0, lambda: messagebox.showerror("Erro", str(e)))
                self.after(0, _done)

        def _done():
            self._prog3.stop()
            self._lbl_g.config(text="")
            self._set_btns_gerar("normal")
            messagebox.showinfo("Concluído âœ“",
                                f"Arquivos salvos em:\n{saida_dir}")

        threading.Thread(target=_run, daemon=True).start()


# â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•
# DETECÇÃO DE COLUNAS ABERTAS — heurística rápida (pré-filtro)
# â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•
def _detectar_colunas_abertas(df: pd.DataFrame) -> list[dict]:
    """
    Passo 1: heurística rápida para pré-filtrar colunas com texto livre.
    O tipo final é definido pela IA em _classificar_tipos_com_ia().
    """
    resultado = []
    for col in df.columns:
        if col in _IGNORAR or col.startswith("Original_") or col.endswith("_cod"):
            continue
        dado = df[col]
        if isinstance(dado, pd.DataFrame):
            dado = dado.iloc[:, 0]
        serie = (dado.dropna().astype(str)
                 .pipe(lambda s: s[~s.str.strip().str.lower().isin(_INVALIDOS_COD)])
                 .pipe(lambda s: s[~s.str.strip().str.upper().str.startswith("NÃƒO SE APLICA")])
                 .pipe(lambda s: s[s.str.strip().str.len() > 1]))
        n = len(serie)
        if n < 3:
            continue
        if pd.to_numeric(serie, errors="coerce").notna().mean() > 0.7:
            continue
        if serie.nunique() / n <= 0.30:
            continue
        if _RE_NPS.search(col):
            continue

        amostra_vals = serie.sample(min(3, n), random_state=42).tolist()
        amostra = "  |  ".join(amostra_vals)

        resultado.append({
            "coluna":    col,
            "tipo":      "livre",          # tipo provisório — IA refina depois
            "n":         n,
            "amostra":   amostra[:120],
            "codificar": True,
        })
    return resultado


def _classificar_tipos_com_ia(colunas: list[dict], codificador) -> list[dict]:
    """
    Passo 2: usa a IA para classificar o tipo correto de cada coluna aberta.
    Chama a OpenAI uma vez com todas as colunas para minimizar latência.
    Tipos disponíveis: reconhecimento_marca, satisfacao, definicao_palavra,
                       local_moradia, livre
    """
    if not colunas or codificador is None:
        return colunas

    tipos_disponiveis = {
        "reconhecimento_marca": "Pergunta sobre marcas/instituições que o participante lembrou de ver",
        "satisfacao":           "Pergunta sobre motivo de avaliação, satisfação, o que faltou, por que gostou/não gostou, sugestões",
        "definicao_palavra":    "Pergunta pedindo uma palavra que define a experiência",
        "local_moradia":        "Pergunta sobre cidade, bairro ou estado onde mora",
        "livre":                "Outro tipo de pergunta aberta que não se encaixa acima",
    }

    linhas = []
    for i, c in enumerate(colunas):
        amostra = c["amostra"][:80].replace("\n", " ")
        linhas.append(f'{i+1}. Coluna: "{c["coluna"][:80]}"\n   Amostra: {amostra}')

    prompt = f"""Você é especialista em análise de pesquisas quantitativas.
Classifique cada coluna abaixo em exatamente um dos tipos:

{chr(10).join(f'- {k}: {v}' for k, v in tipos_disponiveis.items())}

Colunas para classificar:
{chr(10).join(linhas)}

Responda SOMENTE em JSON válido, sem markdown:
{{"classificacoes": [{{"indice": 1, "tipo": "..."}}, ...]}}

Regras:
- "satisfacao" cobre: motivo, por que avalia, o que faltou, sugestão, melhoria, ponto positivo/negativo
- "reconhecimento_marca" só se a pergunta pede marcas/empresas vistas
- "definicao_palavra" só se pede uma única palavra
- Se tiver dúvida, use "livre"
"""

    try:
        import json, re as _re
        resp = codificador._chamar_gpt(
            "Você classifica tipos de perguntas abertas de pesquisas. Responda em JSON.",
            prompt,
            max_tokens=800,
            modelo="gpt-4o"
        )
        texto = _re.sub(r"```[a-z]*", "", resp).strip("`").strip()
        dados = json.loads(texto)
        mapa  = {item["indice"]-1: item["tipo"]
                 for item in dados.get("classificacoes", [])}
        for i, c in enumerate(colunas):
            if i in mapa and mapa[i] in tipos_disponiveis:
                c["tipo"] = mapa[i]
    except Exception as e:
        print(f"[IA tipos] Erro ao classificar: {e}")

    return colunas


# â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•
# POPUP: selecionar coluna para adicionar manualmente
# â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•
class _SeletorColuna(tk.Toplevel):
    def __init__(self, parent, disponiveis: list, df, cols_cod: list, callback=None):
        super().__init__(parent)
        self._disponiveis = disponiveis
        self._df          = df
        self._cols_cod    = cols_cod
        self._cb          = callback
        self.title("Adicionar coluna")
        self.configure(bg=CARD)
        self.resizable(False, False)
        self.grab_set()
        sw, sh = self.winfo_screenwidth(), self.winfo_screenheight()
        self.geometry(f"620x400+{(sw-620)//2}+{(sh-400)//2}")
        self._build()

    def _build(self):
        f = tk.Frame(self, bg=CARD, padx=18, pady=14)
        f.pack(fill=tk.BOTH, expand=True)
        f.grid_columnconfigure(0, weight=1)
        f.grid_rowconfigure(1, weight=1)

        tk.Label(f, text="Selecione a coluna a adicionar:",
                 bg=CARD, fg=TXT2, font=F_SEC).grid(row=0, column=0, sticky="w", pady=(0,6))

        # Listbox com scroll
        lf = tk.Frame(f, bg=BORDER2, padx=1, pady=1)
        lf.grid(row=1, column=0, sticky="nsew", pady=(0, 10))
        self._lb = tk.Listbox(lf, font=F_BODY, relief="flat",
                              selectbackground=AZUL_LIGHT,
                              selectforeground=AZUL,
                              activestyle="none", height=14)
        vsb = ttk.Scrollbar(lf, orient="vertical", command=self._lb.yview)
        self._lb.configure(yscrollcommand=vsb.set)
        self._lb.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        vsb.pack(side=tk.RIGHT, fill=tk.Y)
        for c in self._disponiveis:
            self._lb.insert(tk.END, c)
        self._lb.bind("<Double-1>", lambda e: self._confirmar())

        # Tipo
        tf = tk.Frame(f, bg=CARD)
        tf.grid(row=2, column=0, sticky="ew", pady=(0, 10))
        tk.Label(tf, text="Tipo IA:", bg=CARD, fg=TXT2,
                 font=F_BODY).pack(side=tk.LEFT, padx=(0, 8))
        self._v_tipo = tk.StringVar(value="satisfacao")
        tipos_opts = list(_TIPOS_COD.keys()) if _TIPOS_COD else ["livre"]
        ttk.Combobox(tf, textvariable=self._v_tipo, values=tipos_opts,
                     state="readonly", font=F_BODY, width=26).pack(side=tk.LEFT)

        bf = tk.Frame(f, bg=CARD)
        bf.grid(row=3, column=0, sticky="e")
        tk.Button(bf, text="Cancelar", bg=BG, fg=TXT2, font=F_BODY,
                  relief="flat", padx=10, pady=5, cursor="hand2",
                  command=self.destroy).pack(side=tk.LEFT, padx=(0, 8))
        tk.Button(bf, text="Adicionar", bg=AZUL, fg="white", font=F_BODY,
                  relief="flat", padx=10, pady=5, cursor="hand2",
                  command=self._confirmar).pack(side=tk.LEFT)

    def _confirmar(self):
        sel = self._lb.curselection()
        if not sel:
            messagebox.showwarning("Selecione", "Escolha uma coluna da lista.",
                                   parent=self)
            return
        col = self._disponiveis[sel[0]]
        serie = (self._df[col].dropna().astype(str)
                 .pipe(lambda s: s[~s.str.strip().isin(["", "-", "nan"])]))
        n = len(serie)
        amostra = "  |  ".join(serie.sample(min(3, n), random_state=42).tolist()) if n else ""
        self._cols_cod.append({
            "coluna":    col,
            "tipo":      self._v_tipo.get(),
            "n":         n,
            "amostra":   amostra[:120],
            "codificar": True,
        })
        if self._cb:
            self._cb()
        self.destroy()


# â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•
# POPUP: editar pergunta
# â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•
class _EditorPergunta(tk.Toplevel):
    TIPOS = list(TIPOS_LABEL.keys())

    def __init__(self, parent, pergunta, callback=None):
        super().__init__(parent)
        self._p  = pergunta
        self._cb = callback
        self.title(f"Editar — {pergunta['num']}")
        self.configure(bg=CARD)
        self.resizable(False, False)
        self.grab_set()
        sw, sh = self.winfo_screenwidth(), self.winfo_screenheight()
        self.geometry(f"580x310+{(sw-580)//2}+{(sh-310)//2}")
        self._build()

    def _build(self):
        f = tk.Frame(self, bg=CARD, padx=22, pady=18)
        f.pack(fill=tk.BOTH, expand=True)
        f.grid_columnconfigure(1, weight=1)

        tk.Label(f, text="Pergunta:", bg=CARD, fg=TXT2,
                 font=F_SEC).grid(row=0, column=0, sticky="nw")
        txt = tk.Text(f, height=3, font=F_BODY, bg=BG, fg=TXT1,
                      relief="flat", wrap=tk.WORD)
        txt.insert("1.0", self._p["pergunta"])
        txt.config(state="disabled")
        txt.grid(row=0, column=1, sticky="ew", padx=(10, 0), pady=(0, 10))

        tk.Label(f, text="Tipo:", bg=CARD, fg=TXT2,
                 font=F_SEC).grid(row=1, column=0, sticky="w", pady=(0, 8))
        self._v_tipo = tk.StringVar(value=self._p["tipo"])
        ttk.Combobox(f, textvariable=self._v_tipo, values=self.TIPOS,
                     state="readonly", font=F_BODY, width=24
                     ).grid(row=1, column=1, sticky="w",
                             padx=(10, 0), pady=(0, 8))

        tk.Label(f, text="Nota rodapé:", bg=CARD, fg=TXT2,
                 font=F_SEC).grid(row=2, column=0, sticky="w", pady=(0, 8))
        self._v_nota = tk.StringVar(value=self._p.get("nota", ""))
        tk.Entry(f, textvariable=self._v_nota, font=F_BODY, bg=BG,
                 relief="flat", highlightthickness=1,
                 highlightbackground=BORDER2
                 ).grid(row=2, column=1, sticky="ew",
                         padx=(10, 0), pady=(0, 8))

        self._v_ativo = tk.BooleanVar(value=self._p.get("ativo", True))
        tk.Checkbutton(f, text="Incluir na tabulação",
                       variable=self._v_ativo, bg=CARD, fg=TXT1,
                       font=F_BODY, activebackground=CARD
                       ).grid(row=3, column=1, sticky="w",
                               padx=(10, 0), pady=(0, 14))

        bf = tk.Frame(f, bg=CARD)
        bf.grid(row=4, column=0, columnspan=2, sticky="e")
        tk.Button(bf, text="Cancelar", bg=BG, fg=TXT2, font=F_BODY,
                  relief="flat", padx=10, pady=5, cursor="hand2",
                  command=self.destroy).pack(side=tk.LEFT, padx=(0, 8))
        tk.Button(bf, text="Salvar", bg=AZUL, fg="white", font=F_BODY,
                  relief="flat", padx=10, pady=5, cursor="hand2",
                  command=self._salvar).pack(side=tk.LEFT)

    def _salvar(self):
        self._p["tipo"]  = self._v_tipo.get()
        self._p["nota"]  = self._v_nota.get().strip()
        self._p["ativo"] = self._v_ativo.get()
        if self._cb:
            self._cb()
        self.destroy()


# â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•
# PONTO DE ENTRADA STANDALONE
# â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•
if __name__ == "__main__":
    root = tk.Tk()
    root.withdraw()
    TelaTabulacao(root)
    root.mainloop()

