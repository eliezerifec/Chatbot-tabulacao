"""
Codificador de Pesquisas — IFec RJ
Interface responsiva — adapta a qualquer tamanho de tela
"""
import json
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import tkinter.simpledialog
import threading
import pandas as pd
from pathlib import Path
from codificador import CodificadorIA
from tela_revisao import TelaRevisao
from tela_refino_codebook import TelaRefinoCodebook
try:
    from tela_tabulacao import TelaTabulacao
    _PIPELINE_OK = True
except ImportError:
    _PIPELINE_OK = False

try:
    from PIL import Image, ImageTk
    PIL_OK = True
except ImportError:
    PIL_OK = False

# ── Paleta ────────────────────────────────────────────────────────────────────
NAV_BG     = "#111827"
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
VERDE_BTN  = "#10b981"
VERDE_LIGHT= "#ecfdf5"
ROXO       = "#7c3aed"
ROXO_LIGHT = "#f5f3ff"
ROXO_MID   = "#ddd6fe"
OURO       = "#d97706"
OURO_LIGHT = "#fffbeb"
TXT1  = "#111827"
TXT2  = "#374151"
TXT3  = "#6b7280"
TXT4  = "#9ca3af"
TXT5  = "#d1d5db"

F_SEC   = ("Segoe UI", 10, "bold")
F_BODY  = ("Segoe UI", 9)
F_SMALL = ("Segoe UI", 8)
F_MICRO = ("Segoe UI", 7, "bold")
F_MONO  = ("Consolas", 9)
F_NAV   = ("Segoe UI", 15)

LOG_BG  = "#0d1117"
LOG_FG  = "#e6edf3"
LOG_GRN = "#3fb950"
LOG_BLU = "#79c0ff"
LOG_RED = "#f85149"


def _hr(p, c=BORDER, pady=0):
    tk.Frame(p, bg=c, height=1).pack(fill=tk.X, pady=pady)


def card_frame(parent, px=16, py=14, expand=False, mb=10):
    outer = tk.Frame(parent, bg=BORDER2)
    outer.pack(fill=tk.BOTH if expand else tk.X, expand=expand, pady=(0, mb))
    inner = tk.Frame(outer, bg=CARD, padx=px, pady=py)
    inner.pack(fill=tk.BOTH, expand=expand, pady=(0, 1))
    return inner


def sec_header(parent, icon, title, sub="", right_fn=None):
    f = tk.Frame(parent, bg=CARD)
    f.pack(fill=tk.X, pady=(0, 8))
    lf = tk.Frame(f, bg=CARD)
    lf.pack(side=tk.LEFT)
    pill = tk.Frame(lf, bg=AZUL_LIGHT, padx=7, pady=3)
    pill.pack(side=tk.LEFT, padx=(0, 10))
    tk.Label(pill, text=icon, bg=AZUL_LIGHT, fg=AZUL,
             font=("Segoe UI", 11)).pack()
    tf = tk.Frame(lf, bg=CARD)
    tf.pack(side=tk.LEFT)
    tk.Label(tf, text=title, bg=CARD, fg=TXT1, font=F_SEC).pack(anchor="w")
    if sub:
        tk.Label(tf, text=sub, bg=CARD, fg=TXT4, font=F_SMALL).pack(anchor="w")
    if right_fn:
        right_fn(f)
    _hr(parent, pady=(2, 8))


class _WrapFrame(tk.Frame):
    """Frame que distribui filhos em linhas — quebra automaticamente."""
    def __init__(self, parent, **kw):
        super().__init__(parent, **kw)
        self.bind("<Configure>", self._reflow)

    def _reflow(self, event=None):
        max_w = self.winfo_width()
        if max_w < 2:
            return
        x, y, linha_h = 0, 0, 0
        for w in self.winfo_children():
            w.update_idletasks()
            ww = w.winfo_reqwidth()
            wh = w.winfo_reqheight()
            if x + ww > max_w and x > 0:
                x = 0
                y += linha_h + 4
                linha_h = 0
            w.place(x=x, y=y)
            x += ww + 6
            linha_h = max(linha_h, wh)
        total_h = y + linha_h + 4
        self.config(height=max(total_h, 24))


# ══════════════════════════════════════════════════════════════════════════════
class ModoDropdown:
    ITEMS = [
        ("Simples",               "simples",      False, CARD,       TXT1),
        ("Multipla",              "multipla",     False, CARD,       TXT1),
        (None,                    None,           True,  None,       None),
        ("Semiaberta - Simples",  "semi_simples", False, ROXO_LIGHT, ROXO),
        ("Semiaberta - Multipla", "semi_multipla",False, ROXO_LIGHT, ROXO),
    ]
    LABELS = {
        "simples":       "Simples",
        "multipla":      "Multipla",
        "semi_simples":  "Semi-Simples",
        "semi_multipla": "Semi-Multipla",
    }

    def __init__(self, parent, on_change=None):
        self.on_change = on_change
        self._val = "simples"
        self._win = None
        self.btn = tk.Button(
            parent,
            text=f"  {self.LABELS['simples']}  v",
            bg=AZUL_LIGHT, fg=AZUL,
            font=("Segoe UI", 8, "bold"),
            relief="flat", cursor="hand2",
            padx=6, pady=3, bd=0,
            activebackground=AZUL_MID, activeforeground=AZUL,
            command=self._toggle)

    def pack(self, **kw):
        self.btn.pack(**kw)

    def grid(self, **kw):
        self.btn.grid(**kw)

    def get(self):
        return self._val

    def set(self, val):
        self._val = val
        self.btn.config(text=f"  {self.LABELS.get(val, val)}  v")

    def _toggle(self):
        if self._win and self._win.winfo_exists():
            self._win.destroy()
            self._win = None
        else:
            self._open()

    def _open(self):
        popup = tk.Toplevel()
        popup.overrideredirect(True)
        popup.configure(bg=BORDER2)
        popup.attributes("-topmost", True)
        self._win = popup
        wrap = tk.Frame(popup, bg=CARD)
        wrap.pack(fill=tk.BOTH, padx=1, pady=1)
        for label, valor, is_sep, ibg, ifg in self.ITEMS:
            if is_sep:
                sep = tk.Frame(wrap, bg="#f3f4f6", pady=3)
                sep.pack(fill=tk.X)
                tk.Label(sep, text="  SEMIABERTA", bg="#f3f4f6", fg=TXT4,
                         font=F_MICRO, anchor="w", padx=10).pack(fill=tk.X)
                tk.Frame(sep, bg=BORDER, height=1).pack(fill=tk.X)
                continue
            is_sel = (self._val == valor)
            row_bg = AZUL_LIGHT if is_sel else ibg
            row = tk.Frame(wrap, bg=row_bg, cursor="hand2")
            row.pack(fill=tk.X)
            ind = tk.Frame(row, bg=AZUL if is_sel else row_bg, width=3)
            ind.pack(side=tk.LEFT, fill=tk.Y)
            inner = tk.Frame(row, bg=row_bg, padx=12, pady=8)
            inner.pack(side=tk.LEFT, fill=tk.X, expand=True)
            lbl = tk.Label(inner, text=label, bg=row_bg,
                           fg=AZUL if is_sel else ifg,
                           font=("Segoe UI", 9, "bold") if is_sel else F_BODY,
                           anchor="w")
            lbl.pack(side=tk.LEFT, fill=tk.X, expand=True)
            if is_sel:
                tk.Label(inner, text="v", bg=row_bg, fg=AZUL,
                         font=("Segoe UI", 9, "bold")).pack(side=tk.RIGHT)

            def _sel(v=valor, p=popup):
                self.set(v)
                p.destroy()
                self._win = None
                if self.on_change:
                    self.on_change(v)

            for widget in (row, inner, lbl):
                widget.bind("<Button-1>", lambda e, fn=_sel: fn())
                widget.bind("<Enter>",
                            lambda e, r=row, i=inner, l=lbl:
                            [x.config(bg=AZUL_LIGHT) for x in (r, i, l)])
                widget.bind("<Leave>",
                            lambda e, r=row, i=inner, l=lbl, bg=ibg, v=valor:
                            [x.config(bg=bg) for x in (r, i, l)]
                            if self._val != v else None)
        _hr(wrap)
        popup.update_idletasks()
        bx = self.btn.winfo_rootx()
        by = self.btn.winfo_rooty() + self.btn.winfo_height()
        popup.geometry(f"220x{popup.winfo_reqheight()}+{bx}+{by}")
        popup.bind("<FocusOut>", lambda e: self._fechar(popup))
        popup.focus_set()

    def _fechar(self, p):
        try:
            p.destroy()
        except Exception:
            pass
        self._win = None


# ══════════════════════════════════════════════════════════════════════════════
class App:
    def __init__(self, root):
        self.root = root
        self.root.title("Codificador de Pesquisas  -  IFec RJ")
        # Tamanho inicial = 90% da tela disponível
        sw = root.winfo_screenwidth()
        sh = root.winfo_screenheight()
        w  = min(int(sw * 0.90), 1400)
        h  = min(int(sh * 0.88), 960)
        x  = (sw - w) // 2
        y  = (sh - h) // 2
        self.root.geometry(f"{w}x{h}+{x}+{y}")
        self.root.configure(bg=NAV_BG)
        self.root.resizable(True, True)
        self.root.minsize(800, 600)

        self.arquivo_dados          = None
        self.sheets: dict           = {}
        self.sheet_names: list      = []
        self.codificador            = CodificadorIA()
        self.banco                  = self.codificador.banco
        self.abas_config: dict      = {}
        self._fila_revisao          = []
        self._sessao_codebook: dict = {"items": []}
        # Pesquisa anterior — categorias por aba (nome_aba → lista de categorias)
        self.pesquisa_anterior_por_aba: dict = {}
        self.pesquisa_anterior_ativa: bool = False

        self._setup_styles()
        self._build()

    def _setup_styles(self):
        s = ttk.Style()
        s.theme_use("clam")
        s.configure("Blue.Horizontal.TProgressbar",
                    troughcolor=BORDER, background=AZUL,
                    lightcolor=AZUL, darkcolor=AZUL_DARK, thickness=6)
        s.configure("TScrollbar", background=BORDER,
                    troughcolor=BG, arrowcolor=TXT4, width=6)
        s.configure("TCombobox", fieldbackground=CARD, background=CARD,
                    foreground=TXT1, selectbackground=AZUL_LIGHT,
                    selectforeground=AZUL, padding=4, relief="flat", borderwidth=0)
        s.map("TCombobox", fieldbackground=[("readonly", CARD)])

    # ── Build ─────────────────────────────────────────────────────────────────
    def _build(self):
        # root usa grid para sidebar fixa + main expansível
        self.root.grid_rowconfigure(0, weight=1)
        self.root.grid_columnconfigure(0, weight=0)  # sidebar fixa
        self.root.grid_columnconfigure(1, weight=1)  # main expansível

        self._sidebar()
        self._main()

    # ── Sidebar ───────────────────────────────────────────────────────────────
    def _sidebar(self):
        sb = tk.Frame(self.root, bg=NAV_BG, width=62)
        sb.grid(row=0, column=0, sticky="ns")
        sb.pack_propagate(False)
        sb.grid_propagate(False)

        # Logo
        logo_f = tk.Frame(sb, bg=AZUL_DARK, height=62)
        logo_f.pack(fill=tk.X)
        logo_f.pack_propagate(False)
        if PIL_OK:
            try:
                img = Image.open(Path(__file__).parent / "logo_ifec.png")
                img = img.resize((46, 46), Image.LANCZOS)
                self._logo_img = ImageTk.PhotoImage(img)
                tk.Label(logo_f, image=self._logo_img,
                         bg=AZUL_DARK).pack(expand=True)
            except Exception:
                self._logo_txt(logo_f)
        else:
            self._logo_txt(logo_f)

        tk.Frame(sb, bg="#1f2937", height=1).pack(fill=tk.X)

        # ── Botão Codificação ────────────────────────────────────────────────
        self._nav_cod = tk.Frame(sb, bg=NAV_ACTIVE, height=62, cursor="hand2")
        self._nav_cod.pack(fill=tk.X)
        self._nav_cod.pack_propagate(False)
        tk.Frame(self._nav_cod, bg="#60a5fa", width=3).pack(side=tk.LEFT, fill=tk.Y)
        cod_inner = tk.Frame(self._nav_cod, bg=NAV_ACTIVE)
        cod_inner.pack(expand=True)
        tk.Label(cod_inner, text="⌨", bg=NAV_ACTIVE, fg="white",
                 font=("Segoe UI", 16)).pack()
        tk.Label(cod_inner, text="COD", bg=NAV_ACTIVE, fg="#bfdbfe",
                 font=("Segoe UI", 6, "bold")).pack()
        for w in [self._nav_cod, cod_inner] + list(cod_inner.winfo_children()):
            w.bind("<Button-1>", lambda e: self._mostrar_aba("cod"))

        tk.Frame(sb, bg="#1f2937", height=1).pack(fill=tk.X)

        # ── Botão Tabulação ──────────────────────────────────────────────────
        self._nav_tab = tk.Frame(sb, bg=NAV_BG, height=62, cursor="hand2")
        self._nav_tab.pack(fill=tk.X)
        self._nav_tab.pack_propagate(False)
        tab_inner = tk.Frame(self._nav_tab, bg=NAV_BG)
        tab_inner.pack(expand=True)
        tk.Label(tab_inner, text="📊", bg=NAV_BG, fg="#6b7280",
                 font=("Segoe UI", 14)).pack()
        tk.Label(tab_inner, text="TAB", bg=NAV_BG, fg="#6b7280",
                 font=("Segoe UI", 6, "bold")).pack()
        for w in [self._nav_tab, tab_inner] + list(tab_inner.winfo_children()):
            w.bind("<Button-1>", lambda e: self._mostrar_aba("tab"))
        self._nav_tab.bind("<Enter>", lambda e: self._nav_tab.config(bg="#1f2937"))
        self._nav_tab.bind("<Leave>", lambda e: self._nav_tab.config(bg=NAV_BG))

        # Rodapé
        tk.Frame(sb, bg="#1f2937", height=1).pack(side=tk.BOTTOM, fill=tk.X)
        bot = tk.Frame(sb, bg=NAV_BG, height=50)
        bot.pack(side=tk.BOTTOM, fill=tk.X)
        bot.pack_propagate(False)
        av = tk.Frame(bot, bg=NAV_ACTIVE, width=34, height=34)
        av.pack(expand=True)
        av.pack_propagate(False)
        tk.Label(av, text="JD", bg=NAV_ACTIVE, fg="white",
                 font=("Segoe UI", 8, "bold")).pack(expand=True)

    def _mostrar_aba(self, aba: str):
        """Troca entre o painel de codificação e o de tabulação na mesma janela."""
        if aba == "cod":
            # Esconde tabulação se estiver no grid
            try:
                self._frame_tab.grid_remove()
            except Exception:
                pass
            self._frame_cod.grid(row=1, column=0, sticky="nsew")
            # Sidebar: ativa COD
            self._nav_cod.config(bg=NAV_ACTIVE)
            for w in self._nav_cod.winfo_children():
                w.config(bg=NAV_ACTIVE)
                for ww in w.winfo_children():
                    try: ww.config(bg=NAV_ACTIVE, fg="white")
                    except Exception: pass
            # Sidebar: desativa TAB
            self._nav_tab.config(bg=NAV_BG)
            for w in self._nav_tab.winfo_children():
                w.config(bg=NAV_BG)
                for ww in w.winfo_children():
                    try: ww.config(bg=NAV_BG, fg="#6b7280")
                    except Exception: pass
        else:
            # Esconde codificação
            self._frame_cod.grid_remove()
            # Coloca tabulação no grid (primeira vez ou volta)
            self._frame_tab.grid(row=1, column=0, sticky="nsew")
            # Sidebar: ativa TAB
            self._nav_tab.config(bg="#4c1d95")
            for w in self._nav_tab.winfo_children():
                w.config(bg="#4c1d95")
                for ww in w.winfo_children():
                    try: ww.config(bg="#4c1d95", fg="white")
                    except Exception: pass
            # Sidebar: desativa COD
            self._nav_cod.config(bg=NAV_BG)
            for w in self._nav_cod.winfo_children():
                w.config(bg=NAV_BG)
                for ww in w.winfo_children():
                    try: ww.config(bg=NAV_BG, fg="#6b7280")
                    except Exception: pass
            # Inicializa tabulação na primeira visita
            if not self._tab_iniciada:
                self._iniciar_tabulacao()

    def _abrir_pipeline(self):
        """Mantido por compatibilidade — agora só troca de aba."""
        self._mostrar_aba("tab")

    def _iniciar_tabulacao(self):
        self._tab_iniciada = True

        if not _PIPELINE_OK:
            tk.Label(
                self._frame_tab,
                text="tela_tabulacao.py não encontrado.\n"
                     "Coloque o arquivo na mesma pasta que app.py.",
                bg=BG, fg=TXT3, font=F_BODY,
                justify="center"
            ).pack(expand=True)
            return

        for child in self._frame_tab.winfo_children():
            child.destroy()

        self._tab_view = TelaTabulacao(
            self._frame_tab,
            self.codificador,
            self.banco,
            inline=True,
        )

        self._tab_view.grid(row=0, column=0, sticky="nsew")


    def _logo_txt(self, p):
        tk.Label(p, text="IF", bg=AZUL_DARK, fg="#fbbf24",
                 font=("Segoe UI", 14, "bold")).pack(expand=True)

    # ── Área principal ────────────────────────────────────────────────────────
    def _main(self):
        main = tk.Frame(self.root, bg=BG)
        main.grid(row=0, column=1, sticky="nsew")
        main.grid_rowconfigure(0, weight=0)  # topbar
        main.grid_rowconfigure(1, weight=1)  # conteúdo
        main.grid_columnconfigure(0, weight=1)

        self._topbar(main)

        # ── Painel Codificação ────────────────────────────────────────────────
        self._frame_cod = tk.Frame(main, bg=BG)
        self._frame_cod.grid(row=1, column=0, sticky="nsew")
        self._frame_cod.grid_rowconfigure(0, weight=1)
        self._frame_cod.grid_columnconfigure(0, weight=1)

        wrap = tk.Frame(self._frame_cod, bg=BG)
        wrap.grid(row=0, column=0, sticky="nsew")
        wrap.grid_rowconfigure(0, weight=1)
        wrap.grid_columnconfigure(0, weight=1)

        canvas = tk.Canvas(wrap, bg=BG, highlightthickness=0)
        vsb    = ttk.Scrollbar(wrap, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=vsb.set)
        canvas.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")

        self.sf = tk.Frame(canvas, bg=BG)
        self._sf_id = canvas.create_window((0, 0), window=self.sf, anchor="nw")

        def _on_canvas_resize(e):
            canvas.itemconfig(self._sf_id, width=e.width)
        canvas.bind("<Configure>", _on_canvas_resize)
        self.sf.bind("<Configure>",
                     lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.bind_all("<MouseWheel>",
                        lambda e: canvas.yview_scroll(-1*(e.delta//120), "units"))

        self._build_content()

        # ── Painel Tabulação ──────────────────────────────────────────────────
        # Criado mas NUNCA colocado no grid ainda — _mostrar_aba faz isso
        self._frame_tab = tk.Frame(main, bg=BG)
        self._frame_tab.grid_rowconfigure(0, weight=1)
        self._frame_tab.grid_columnconfigure(0, weight=1)
        self._tab_iniciada = False
        # frame_cod já está visível por padrão (foi o último a fazer .grid)

    def _topbar(self, parent):
        top = tk.Frame(parent, bg=CARD, height=62)
        top.grid(row=0, column=0, sticky="ew")
        top.grid_columnconfigure(0, weight=1)
        top.grid_propagate(False)

        lf = tk.Frame(top, bg=CARD)
        lf.pack(side=tk.LEFT, fill=tk.Y, padx=20)
        tk.Label(lf, text="Codificador de Pesquisas",
                 bg=CARD, fg=TXT1,
                 font=("Segoe UI", 15, "bold")).pack(anchor="w", pady=(14, 0))
        tk.Label(lf, text="Analise automatica de respostas abertas com IA  -  IFec RJ",
                 bg=CARD, fg=TXT4, font=F_SMALL).pack(anchor="w")

        rf = tk.Frame(top, bg=CARD)
        rf.pack(side=tk.RIGHT, fill=tk.Y, padx=20)
        av = tk.Frame(rf, bg=AZUL, width=34, height=34)
        av.pack(side=tk.RIGHT, pady=14)
        av.pack_propagate(False)
        tk.Label(av, text="IFEC", bg=AZUL, fg="white",
                 font=("Segoe UI", 8, "bold")).pack(expand=True)

        tk.Frame(top, bg=BORDER, height=1).pack(side=tk.BOTTOM, fill=tk.X)

    # ── Conteúdo principal (dentro do scroll) ─────────────────────────────────
    def _build_content(self):
        # uniform="cols" faz as colunas respeitarem weight sem deixar
        # o conteúdo interno empurrar a proporção
        self.sf.grid_columnconfigure(0, weight=6, uniform="cols")
        self.sf.grid_columnconfigure(1, weight=4, uniform="cols")

        pad = {"padx": 20, "pady": 20}

        # ── Métricas (linha 0, span 2 colunas) ───────────────────────────────
        self._sec_metrics()

        # ── Coluna esquerda ───────────────────────────────────────────────────
        left = tk.Frame(self.sf, bg=BG)
        left.grid(row=1, column=0, sticky="nsew", padx=(20, 10), pady=(0, 16))
        left.grid_columnconfigure(0, weight=1)


        self._card_upload(left)
        self._card_abas(left)

        # ── Coluna direita ────────────────────────────────────────────────────
        right = tk.Frame(self.sf, bg=BG)
        right.grid(row=1, column=1, sticky="nsew", padx=(10, 20), pady=(0, 16))
        right.grid_columnconfigure(0, weight=1)


        self._card_pesquisa_anterior(right)
        self._card_categorias(right)
        self._card_contexto(right)

        # ── Ações (linha 2, span 2) ───────────────────────────────────────────
        self._sec_actions()

        # ── Log (linha 3, span 2) ─────────────────────────────────────────────
        self._sec_log()

    # ── Métricas ──────────────────────────────────────────────────────────────
    def _sec_metrics(self):
        mf = tk.Frame(self.sf, bg=BG)
        mf.grid(row=0, column=0, columnspan=2, sticky="ew",
                padx=20, pady=(20, 0))
        mf.grid_columnconfigure((0, 1, 2, 3), weight=1)

        defs = [
            ("D", AZUL_LIGHT,  AZUL,  "ARQUIVO",  "-",      "nenhum carregado"),
            ("A", AZUL_LIGHT,  AZUL,  "ABAS",     "0",      "aguardando"),
            ("M", OURO_LIGHT,  OURO,  "MODELO",   "gpt-5.4","Ag1: gpt-5.4 / Ag2: gpt-4o"),
            ("B", VERDE_LIGHT, VERDE, "BANCO IA", "-",      "exemplos salvos"),
        ]
        self._mw = {}
        for i, (ico, ibg, ifg, title, val, sub) in enumerate(defs):
            c = tk.Frame(mf, bg=CARD)
            c.grid(row=0, column=i, sticky="nsew",
                   padx=(0, 12 if i < 3 else 0))

            tk.Frame(c, bg=ifg, height=3).pack(fill=tk.X)
            body = tk.Frame(c, bg=CARD, padx=14, pady=12)
            body.pack(fill=tk.BOTH)
            body.grid_columnconfigure(0, weight=1)

            top_r = tk.Frame(body, bg=CARD)
            top_r.pack(fill=tk.X)

            tf = tk.Frame(top_r, bg=CARD)
            tf.pack(side=tk.LEFT, fill=tk.X, expand=True)
            lv = tk.Label(tf, text=val, bg=CARD, fg=TXT1,
                          font=("Segoe UI", 15, "bold"))
            lv.pack(anchor="w")
            ls = tk.Label(tf, text=sub, bg=CARD, fg=TXT3, font=F_SMALL)
            ls.pack(anchor="w")

            ico_f = tk.Frame(top_r, bg=ibg, width=36, height=36)
            ico_f.pack(side=tk.RIGHT, anchor="ne")
            ico_f.pack_propagate(False)
            tk.Label(ico_f, text=ico, bg=ibg, fg=ifg,
                     font=("Segoe UI", 12, "bold")).pack(expand=True)

            tk.Label(body, text=title, bg=CARD, fg=TXT5,
                     font=F_MICRO).pack(anchor="w", pady=(6, 0))
            self._mw[title] = (lv, ls)

        try:
            st = self.banco.stats()
            self._mw["BANCO IA"][0].config(text=str(st["total"]))
            self._mw["BANCO IA"][1].config(text=f"{st['taxa_acerto']}% precisao")
        except Exception:
            pass

    def _upd_m(self, key, val, sub=None):
        if key in self._mw:
            self._mw[key][0].config(text=str(val))
            if sub:
                self._mw[key][1].config(text=sub)

    # ── Card Upload ───────────────────────────────────────────────────────────
    def _card_upload(self, parent):
        c = card_frame(parent)
        sec_header(c, "^", "Upload da Base de Dados",
                   "XLSX ou CSV - todas as abas serao detectadas")

        dz = tk.Frame(c, bg="#f0f7ff", cursor="hand2")
        dz.pack(fill=tk.X)
        tk.Frame(c, bg=AZUL_MID, height=1).pack(fill=tk.X)
        inner = tk.Frame(dz, bg="#f0f7ff", pady=22)
        inner.pack(fill=tk.X)

        tk.Label(inner, text="^", bg="#f0f7ff", fg=AZUL,
                 font=("Segoe UI", 26)).pack()
        row_msg = tk.Frame(inner, bg="#f0f7ff")
        row_msg.pack(pady=(4, 0))
        tk.Label(row_msg, text="Arraste a planilha aqui",
                 bg="#f0f7ff", fg=TXT1,
                 font=("Segoe UI", 10, "bold")).pack(side=tk.LEFT)
        tk.Label(row_msg, text="  ou  ", bg="#f0f7ff", fg=TXT4,
                 font=F_BODY).pack(side=tk.LEFT)
        lnk = tk.Label(row_msg, text="clique para selecionar",
                       bg="#f0f7ff", fg=AZUL,
                       font=("Segoe UI", 9, "bold"), cursor="hand2")
        lnk.pack(side=tk.LEFT)
        lnk.bind("<Button-1>", lambda e: self._abrir_arquivo())
        inner.bind("<Button-1>", lambda e: self._abrir_arquivo())
        dz.bind("<Button-1>",   lambda e: self._abrir_arquivo())

        tk.Label(inner, text="Formatos: .xlsx   .csv",
                 bg="#f0f7ff", fg=TXT4, font=F_SMALL).pack(pady=(4, 0))

        # Linha inferior: nome do arquivo + botão Trocar
        self._upload_info_row = tk.Frame(c, bg=CARD)
        self._upload_info_row.pack(fill=tk.X, pady=(8, 0))

        self.lbl_arquivo = tk.Label(self._upload_info_row, text="", bg=CARD,
                                    fg=VERDE, font=("Segoe UI", 9, "bold"))
        self.lbl_arquivo.pack(side=tk.LEFT)

        self._btn_trocar = self._mkbtn(
            self._upload_info_row, "Trocar arquivo",
            self._abrir_arquivo,
            bg=AZUL_LIGHT, fg=AZUL, small=True)
        # Começa oculto — aparece só quando houver arquivo carregado
        self._btn_trocar.pack(side=tk.RIGHT)
        self._btn_trocar.pack_forget()

        self.lbl_status = tk.Label(c, text="", bg=CARD, fg=TXT4, font=F_SMALL)
        self.lbl_status.pack(anchor="w")

    # ── Card Abas ─────────────────────────────────────────────────────────────
    def _card_abas(self, parent):
        c = card_frame(parent, mb=0)

        def _btns_dir(f):
            bf = tk.Frame(f, bg=CARD)
            bf.pack(side=tk.RIGHT)
            self._mkbtn(bf, "Todas", self._sel_todas,
                        bg=AZUL_LIGHT, fg=AZUL, small=True).pack(
                            side=tk.LEFT, padx=(0, 4))
            self._mkbtn(bf, "Nenhuma", self._des_todas,
                        bg=BG, fg=TXT3, small=True).pack(side=tk.LEFT)

        sec_header(c, "=", "Configuracao das Abas",
                   "Tipo semantico e modo de resposta por aba",
                   right_fn=_btns_dir)

        # Cabeçalho da tabela — usa grid para alinhar com as linhas
        thead = tk.Frame(c, bg="#f9fafb")
        thead.pack(fill=tk.X)
        thead.grid_columnconfigure(0, minsize=26)   # checkbox
        thead.grid_columnconfigure(1, weight=1)     # aba
        thead.grid_columnconfigure(2, weight=2)     # entrada
        thead.grid_columnconfigure(3, weight=2)     # saida
        thead.grid_columnconfigure(4, weight=3)     # tipo
        thead.grid_columnconfigure(5, weight=2)     # modo
        thead.grid_columnconfigure(6, minsize=32)   # btn

        for col, txt in enumerate(["", "ABA", "ENTRADA", "SAIDA",
                                    "TIPO", "MODO", ""]):
            tk.Label(thead, text=txt, bg="#f9fafb", fg=TXT4,
                     font=F_MICRO, anchor="w",
                     padx=4, pady=5).grid(row=0, column=col, sticky="ew")
        _hr(c, pady=(2, 2))

        self.frame_abas = tk.Frame(c, bg=CARD)
        self.frame_abas.pack(fill=tk.X)
        tk.Label(self.frame_abas,
                 text="  <- Carregue um arquivo para ver as abas",
                 bg=CARD, fg=TXT4, font=F_BODY).pack(anchor="w", pady=14)

    # ── Card Pesquisa Anterior ────────────────────────────────────────────────
    def _card_pesquisa_anterior(self, parent):
        OURO2      = "#92400e"
        OURO_MID   = "#fde68a"
        OURO_DARK  = "#b45309"

        c = card_frame(parent)

        def _btn_dir(f):
            self._btn_limpar_ant = self._mkbtn(
                f, "Limpar", self._limpar_pesquisa_anterior,
                bg=BG, fg=TXT3, small=True)
            self._btn_limpar_ant.pack(side=tk.RIGHT)
            self._btn_limpar_ant.pack_forget()   # oculto até ter dados

        sec_header(c, "H", "Pesquisa Anterior",
                   "Reutiliza categorias ja validadas — IA só cria nova em ultimo caso",
                   right_fn=_btn_dir)

        # ── Zona de drop / clique ──────────────────────────────────────────
        self._ant_dropzone = tk.Frame(c, bg="#fffbeb", cursor="hand2")
        self._ant_dropzone.pack(fill=tk.X)
        tk.Frame(c, bg=OURO_MID, height=1).pack(fill=tk.X)

        ant_inner = tk.Frame(self._ant_dropzone, bg="#fffbeb", pady=16)
        ant_inner.pack(fill=tk.X)
        tk.Label(ant_inner, text="H", bg="#fffbeb", fg=OURO,
                 font=("Segoe UI", 22)).pack()
        row_msg2 = tk.Frame(ant_inner, bg="#fffbeb")
        row_msg2.pack(pady=(4, 0))
        tk.Label(row_msg2, text="Carregar planilha anterior",
                 bg="#fffbeb", fg=TXT1,
                 font=("Segoe UI", 10, "bold")).pack(side=tk.LEFT)
        tk.Label(row_msg2, text="  ou  ", bg="#fffbeb", fg=TXT4,
                 font=F_BODY).pack(side=tk.LEFT)
        lnk2 = tk.Label(row_msg2, text="selecionar arquivo",
                        bg="#fffbeb", fg=OURO,
                        font=("Segoe UI", 9, "bold"), cursor="hand2")
        lnk2.pack(side=tk.LEFT)
        lnk2.bind("<Button-1>", lambda e: self._carregar_pesquisa_anterior())
        ant_inner.bind("<Button-1>", lambda e: self._carregar_pesquisa_anterior())
        self._ant_dropzone.bind("<Button-1>",
                                lambda e: self._carregar_pesquisa_anterior())
        tk.Label(ant_inner, text="Formatos: .xlsx   .csv",
                 bg="#fffbeb", fg=TXT4, font=F_SMALL).pack(pady=(4, 0))

        # ── Painel de configuração (oculto até carregar arquivo) ───────────
        self._ant_config_frame = tk.Frame(c, bg=CARD)
        # (não empacotado ainda)

        cf = self._ant_config_frame

        # Cabeçalho da tabela aba→coluna
        thead = tk.Frame(cf, bg=CARD)
        thead.pack(fill=tk.X, pady=(8, 2))
        thead.grid_columnconfigure(0, weight=2)
        thead.grid_columnconfigure(1, weight=3)
        tk.Label(thead, text="ABA", bg=CARD, fg=TXT4,
                 font=F_MICRO).grid(row=0, column=0, sticky="w")
        tk.Label(thead, text="COLUNA DE CATEGORIAS", bg=CARD, fg=TXT4,
                 font=F_MICRO).grid(row=0, column=1, sticky="w", padx=(8, 0))

        # Frame que vai receber as linhas dinâmicas (uma por aba)
        self._ant_linhas_frame = tk.Frame(cf, bg=CARD)
        self._ant_linhas_frame.pack(fill=tk.X)
        # dict: nome_aba → ttk.Combobox
        self._ant_combos_aba: dict = {}

        _hr(cf, pady=(10, 0))

        # Status / pills
        self._ant_status = tk.Label(c, text="", bg=CARD, fg=TXT4, font=F_SMALL)
        self._ant_status.pack(anchor="w", pady=(6, 0))

        self._ant_tags_frame = _WrapFrame(c, bg=CARD)
        self._ant_tags_frame.pack(fill=tk.X, pady=(4, 0))

        # Badge "ATIVO" — aparece quando há categorias carregadas
        self._ant_badge = tk.Label(c, text="", bg=CARD, fg=CARD, font=F_SMALL)
        self._ant_badge.pack(anchor="w", pady=(2, 0))

        # Guarda referência ao frame do card para highlight
        self._ant_card = c

    def _carregar_pesquisa_anterior(self):
        path = filedialog.askopenfilename(
            title="Selecionar planilha anterior",
            filetypes=[("Planilhas", "*.xlsx *.csv"), ("Todos", "*.*")])
        if not path:
            return
        try:
            if path.endswith(".csv"):
                self._ant_sheets      = {"Planilha": pd.read_csv(path)}
                self._ant_sheet_names = ["Planilha"]
            else:
                xl                    = pd.ExcelFile(path)
                self._ant_sheet_names = xl.sheet_names
                self._ant_sheets      = {n: xl.parse(n) for n in xl.sheet_names}

            # Reconstrói as linhas dinâmicas (uma por aba)
            self._ant_combos_aba.clear()
            for w in self._ant_linhas_frame.winfo_children():
                w.destroy()

            self._ant_linhas_frame.grid_columnconfigure(0, weight=2)
            self._ant_linhas_frame.grid_columnconfigure(1, weight=3)

            PALAVRAS_CHAVE = ("codigo", "categ", "classe", "classif", "cod_ia", "ia")

            for i, nome_aba in enumerate(self._ant_sheet_names):
                cols = list(self._ant_sheets[nome_aba].columns)
                row_bg = CARD if i % 2 == 0 else "#fafafa"

                tk.Label(self._ant_linhas_frame,
                         text=f" {nome_aba[:18]}",
                         bg=row_bg, fg=TXT2, font=F_BODY,
                         anchor="w").grid(row=i, column=0, sticky="ew",
                                          pady=2)

                cb = ttk.Combobox(self._ant_linhas_frame,
                                  values=cols, state="readonly", font=F_BODY)
                cb.grid(row=i, column=1, sticky="ew", padx=(8, 0), pady=2)

                # Auto-detecta coluna mais provável
                palpite = next(
                    (c for c in cols if any(k in str(c).lower()
                     for k in PALAVRAS_CHAVE)),
                    cols[0] if cols else "")
                cb.set(palpite)

                # Re-extrai ao trocar coluna
                cb.bind("<<ComboboxSelected>>",
                        lambda e: self._extrair_categorias_anteriores())

                self._ant_combos_aba[nome_aba] = cb

            # Mostra painel, esconde dropzone
            self._ant_dropzone.pack_forget()
            self._ant_config_frame.pack(fill=tk.X)
            self._btn_limpar_ant.pack(side=tk.RIGHT)

            nome = Path(path).name
            self._ant_status.config(
                text=f"  {nome}  —  extraindo categorias...", fg=TXT3)
            self._log(f"Pesquisa anterior: {nome} — {len(self._ant_sheet_names)} aba(s)", "INFO")

            # Extrai automaticamente com as colunas detectadas
            self._extrair_categorias_anteriores()
        except Exception as e:
            messagebox.showerror("Erro ao abrir pesquisa anterior", str(e))

    def _extrair_categorias_anteriores(self):
        if not hasattr(self, "_ant_combos_aba") or not self._ant_combos_aba:
            return
        try:
            por_aba: dict = {}   # nome_aba → lista de categorias únicas

            for nome_aba, cb in self._ant_combos_aba.items():
                col = cb.get()
                if not col:
                    continue
                df = self._ant_sheets.get(nome_aba, pd.DataFrame())
                if col not in df.columns:
                    continue
                vals = (df[col].dropna()
                               .astype(str)
                               .str.strip()
                               .unique()
                               .tolist())
                cats = sorted(
                    v for v in vals
                    if v and v.lower() not in
                    ("nan", "none", "", "sem_resposta", "erro", "nao_classificado"))
                if cats:
                    por_aba[nome_aba] = cats

            if not por_aba:
                self._ant_status.config(
                    text="  Nenhuma categoria válida encontrada nas colunas selecionadas",
                    fg=OURO)
                return

            self.pesquisa_anterior_por_aba = por_aba
            self.pesquisa_anterior_ativa   = True

            total_cats = sum(len(v) for v in por_aba.values())
            info = f"{total_cats} categorias em {len(por_aba)} aba(s)"
            self._ant_status.config(
                text=f"  {info} — cada aba usará suas próprias categorias",
                fg=VERDE)
            self._ant_badge.config(
                text=f"  ✓  MODO ATIVO — categorias separadas por aba",
                bg=VERDE_LIGHT, fg=VERDE,
                font=("Segoe UI", 8, "bold"), padx=8, pady=4)
            self._ant_card.config(highlightbackground=VERDE,
                                  highlightthickness=2, highlightcolor=VERDE)

            # Renderiza pills agrupadas por aba
            for w in self._ant_tags_frame.winfo_children():
                w.destroy()
            for nome_aba, cats in por_aba.items():
                # Label da aba
                lbl_aba = tk.Label(self._ant_tags_frame,
                                   text=f" {nome_aba}: ",
                                   bg=AZUL_LIGHT, fg=AZUL,
                                   font=("Segoe UI", 7, "bold"),
                                   padx=3, pady=2)
                lbl_aba.place(x=0, y=0)
                for cat in cats:
                    tag = tk.Label(self._ant_tags_frame, text=f"  {cat}  ",
                                   bg=OURO_LIGHT, fg=OURO,
                                   font=("Segoe UI", 8, "bold"),
                                   padx=4, pady=2)
                    tag.place(x=0, y=0)
            self._ant_tags_frame._reflow()

            self._log(f"Pesquisa anterior: {info} ativas", "OK", "g")
        except Exception as e:
            messagebox.showerror("Erro ao extrair categorias", str(e))

    def _limpar_pesquisa_anterior(self):
        self.pesquisa_anterior_por_aba = {}
        self.pesquisa_anterior_ativa   = False
        for w in self._ant_tags_frame.winfo_children():
            w.destroy()
        for w in self._ant_linhas_frame.winfo_children():
            w.destroy()
        self._ant_combos_aba.clear()
        self._ant_status.config(text="", fg=TXT4)
        self._ant_badge.config(text="", bg=CARD)
        self._ant_config_frame.pack_forget()
        self._ant_dropzone.pack(fill=tk.X)
        self._btn_limpar_ant.pack_forget()
        self._ant_card.config(highlightthickness=0)
        self._log("Pesquisa anterior removida", "INFO")

    # ── Card Categorias ───────────────────────────────────────────────────────
    def _card_categorias(self, parent):
        c = card_frame(parent)

        def _btn_imp(f):
            self._mkbtn(f, "Importar", self._importar_codigos,
                        bg=AZUL_LIGHT, fg=AZUL, small=True).pack(side=tk.RIGHT)

        sec_header(c, "#", "Categorias", right_fn=_btn_imp)

        # WrapFrame: distribui pills em linhas automaticamente ao redimensionar
        self.frame_tags = _WrapFrame(c, bg=CARD)
        self.frame_tags.pack(fill=tk.X, pady=(0, 10))
        self._tags_placeholder = tk.Label(
            self.frame_tags, text="Nenhuma categoria adicionada",
            bg=CARD, fg=TXT4, font=F_SMALL)
        self._tags_placeholder.pack(anchor="w")

        _hr(c, pady=(0, 10))

        add = tk.Frame(c, bg=CARD)
        add.pack(fill=tk.X)
        add.grid_columnconfigure(0, weight=1)

        ef = tk.Frame(add, bg=BORDER2, padx=1, pady=1)
        ef.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 8))
        ei = tk.Frame(ef, bg=CARD)
        ei.pack(fill=tk.X)
        self.entry_cat = tk.Entry(ei, bg=CARD, fg=TXT4,
                                  font=F_BODY, relief="flat", bd=0)
        self.entry_cat.pack(fill=tk.X, ipady=6, padx=10)
        self.entry_cat.insert(0, "Adicionar categoria...")
        self.entry_cat.bind("<FocusIn>",
                            lambda e: self.entry_cat.delete(0, tk.END)
                            if "Adicionar" in self.entry_cat.get() else None)
        self.entry_cat.bind("<Return>", lambda e: self._add_cat())

        self._mkbtn(add, "Adicionar", self._add_cat,
                    bg=AZUL, small=True).pack(side=tk.LEFT)
        self.lbl_cats_info = tk.Label(c, text="", bg=CARD, fg=TXT4, font=F_SMALL)
        self.lbl_cats_info.pack(anchor="w", pady=(6, 0))

    # ── Card Contexto ─────────────────────────────────────────────────────────
    def _card_contexto(self, parent):
        c = card_frame(parent, mb=0)
        sec_header(c, "i", "Contexto Global",
                   "Instrucao extra para o tipo Personalizado")
        ef = tk.Frame(c, bg=BORDER2, padx=1, pady=1)
        ef.pack(fill=tk.X)
        ei = tk.Frame(ef, bg=CARD)
        ei.pack(fill=tk.X)
        self.text_ctx = tk.Text(ei, height=4, bg=CARD, fg=TXT2,
                                insertbackground=AZUL, font=F_BODY,
                                relief="flat", bd=0, padx=10, pady=8)
        self.text_ctx.pack(fill=tk.X)
        self.text_ctx.insert("1.0",
            "Pesquisa de satisfacao de evento.\nRespostas curtas de participantes.")

    # ── Ações ─────────────────────────────────────────────────────────────────
    def _sec_actions(self):
        bar = tk.Frame(self.sf, bg=BG)
        bar.grid(row=2, column=0, columnspan=2, sticky="ew",
                 padx=20, pady=(0, 12))
        bar.grid_columnconfigure(0, weight=3)
        bar.grid_columnconfigure(1, weight=3)
        bar.grid_columnconfigure(2, weight=3)
        bar.grid_columnconfigure(3, weight=3)
        bar.grid_columnconfigure(4, weight=2)

        self.btn_run = tk.Button(
            bar, text="  Iniciar Codificacao",
            bg=VERDE_BTN, fg="white",
            font=("Segoe UI", 11, "bold"), relief="flat",
            cursor="hand2", pady=12, bd=0,
            activebackground=VERDE, activeforeground="white",
            command=self._executar)
        self.btn_run.grid(row=0, column=0, sticky="ew", padx=(0, 8))

        self.btn_train = tk.Button(
            bar, text="  Treinar IA",
            bg=ROXO, fg="white",
            font=("Segoe UI", 11, "bold"), relief="flat",
            cursor="hand2", pady=12, bd=0, state="disabled",
            activebackground="#6d28d9", activeforeground="white",
            command=self._abrir_revisao)
        self.btn_train.grid(row=0, column=1, sticky="ew", padx=(0, 8))

        self.btn_refine = tk.Button(
            bar, text="  Refinar Codebook",
            bg=OURO, fg="white",
            font=("Segoe UI", 11, "bold"), relief="flat",
            cursor="hand2", pady=12, bd=0, state="disabled",
            activebackground="#b45309", activeforeground="white",
            command=self._abrir_refino_codebook)
        self.btn_refine.grid(row=0, column=2, sticky="ew", padx=(0, 8))

        tk.Button(
            bar, text="  Exportar Resultado",
            bg=AZUL, fg="white",
            font=("Segoe UI", 11, "bold"), relief="flat",
            cursor="hand2", pady=12, bd=0,
            activebackground=AZUL_DARK, activeforeground="white",
            command=self._exportar
        ).grid(row=0, column=3, sticky="ew", padx=(0, 8))

        # Card progresso
        pc = tk.Frame(bar, bg=CARD)
        pc.grid(row=0, column=4, sticky="nsew")
        pi = tk.Frame(pc, bg=CARD, padx=12, pady=9)
        pi.pack(fill=tk.BOTH)
        tr = tk.Frame(pi, bg=CARD)
        tr.pack(fill=tk.X)
        tk.Label(tr, text="Progresso", bg=CARD, fg=TXT2,
                 font=("Segoe UI", 9, "bold")).pack(side=tk.LEFT)
        self.lbl_pct = tk.Label(tr, text="0%", bg=CARD, fg=AZUL,
                                font=("Segoe UI", 9, "bold"))
        self.lbl_pct.pack(side=tk.RIGHT)
        self.progress = ttk.Progressbar(pi, mode="determinate",
                                        style="Blue.Horizontal.TProgressbar")
        self.progress.pack(fill=tk.X, pady=(4, 2))
        self.lbl_prog_info = tk.Label(pi, text="-", bg=CARD,
                                      fg=TXT4, font=F_SMALL)
        self.lbl_prog_info.pack(anchor="w")

    # ── Log ───────────────────────────────────────────────────────────────────
    def _sec_log(self):
        lf = tk.Frame(self.sf, bg=BG)
        lf.grid(row=3, column=0, columnspan=2, sticky="ew",
                padx=20, pady=(0, 20))
        lf.grid_columnconfigure(0, weight=1)

        c = card_frame(lf, mb=0)
        sec_header(c, ">", "Log do Sistema", "Saida em tempo real")
        self.log = tk.Text(c, height=8, bg=LOG_BG, fg=LOG_FG,
                           font=F_MONO, relief="flat", bd=0,
                           padx=14, pady=12, state="disabled")
        self.log.pack(fill=tk.BOTH)
        self.log.tag_configure("g", foreground=LOG_GRN)
        self.log.tag_configure("b", foreground=LOG_BLU)
        self.log.tag_configure("r", foreground=LOG_RED)
        self._log("Sistema iniciado - Aguardando arquivo", "INFO")
        self._log("Ag1: gpt-5.4  |  Ag2: gpt-4o  |  Banco ativo", "INFO", "b")

    # ── Helpers UI ────────────────────────────────────────────────────────────
    def _mkbtn(self, parent, text, cmd, bg=AZUL, fg="white",
               small=False, width=None):
        f  = ("Segoe UI", 7, "bold") if small else ("Segoe UI", 9, "bold")
        px, py = (7, 3) if small else (14, 7)
        kw = dict(text=text, bg=bg, fg=fg, font=f, relief="flat",
                  cursor="hand2", padx=px, pady=py, bd=0,
                  activebackground=AZUL_DARK, activeforeground="white",
                  command=cmd)
        if width:
            kw["width"] = width
        return tk.Button(parent, **kw)

    def _log(self, msg, prefix="INFO", tag="g"):
        self.log.config(state="normal")
        self.log.insert(tk.END, f" {prefix} ", "g")
        self.log.insert(tk.END, f" {msg}\n", tag)
        self.log.see(tk.END)
        self.log.config(state="disabled")

    def _set_progress(self, pct, info=""):
        self.progress.config(value=pct)
        self.lbl_pct.config(text=f"{int(pct)}%")
        if info:
            self.lbl_prog_info.config(text=info)

    def _normalizar_lista_categorias(self, valor):
        if valor is None:
            return []
        txt = str(valor).strip()
        if not txt or txt in ("SEM_RESPOSTA", "ERRO", "nan", "None"):
            return []
        return [parte.strip() for parte in txt.split(",") if parte.strip()]

    def _juntar_categorias(self, categorias):
        unicas = []
        for cat in categorias or []:
            cat_limpa = str(cat).strip()
            if cat_limpa and cat_limpa not in unicas:
                unicas.append(cat_limpa)
        return ", ".join(unicas)

    def _nova_sessao_codebook(self):
        self._sessao_codebook = {
            "items": [],
            "total_respostas": 0,
            "respostas_codificadas": 0,
        }

    def _registrar_item_codebook(self, nome_aba, row_idx, resposta, categorias,
                                 origem_por_categoria, cfg_cols):
        categorias_validas = self._normalizar_lista_categorias(
            self._juntar_categorias(categorias)
        )
        if not categorias_validas:
            return

        self._sessao_codebook["items"].append({
            "id": f"{nome_aba}::{row_idx}",
            "aba": nome_aba,
            "linha": int(row_idx),
            "resposta": str(resposta),
            "categorias": categorias_validas,
            "origens": {
                cat: origem_por_categoria.get(cat, "resultado")
                for cat in categorias_validas
            },
            **cfg_cols,
        })
        self._sessao_codebook["respostas_codificadas"] += 1

    def _atualizar_estado_refino(self):
        tem_itens = bool(self._sessao_codebook.get("items"))
        estado = "normal" if tem_itens else "disabled"
        try:
            self.btn_refine.config(state=estado)
        except Exception:
            pass

    def _abrir_refino_codebook(self):
        if not self._sessao_codebook.get("items"):
            messagebox.showinfo(
                "Refinar Codebook",
                "Execute uma codificacao primeiro para gerar o codebook da rodada."
            )
            return
        TelaRefinoCodebook(
            self.root,
            self._sessao_codebook,
            on_apply=self._aplicar_refino_codebook
        )

    def _aplicar_refino_codebook(self, itens_atualizados):
        atualizados = 0
        categorias_globais = []

        for item in itens_atualizados:
            aba = item["aba"]
            if aba not in self.sheets:
                continue

            df = self.sheets[aba]
            linha = int(item["linha"])
            categorias = self._normalizar_lista_categorias(
                self._juntar_categorias(item.get("categorias", []))
            )
            origens = item.get("origens", {})

            for cat in categorias:
                if cat not in categorias_globais:
                    categorias_globais.append(cat)

            if str(item.get("modo", "")).startswith("semi"):
                col_imp = item.get("col_imp")
                col_novo = item.get("col_novo")
                imputadas = [c for c in categorias if origens.get(c) == "imputado"]
                novas = [c for c in categorias if origens.get(c) != "imputado"]
                if col_imp:
                    df.at[linha, col_imp] = self._juntar_categorias(imputadas)
                if col_novo:
                    df.at[linha, col_novo] = self._juntar_categorias(novas)
            else:
                col_out = item.get("col_out")
                if col_out:
                    df.at[linha, col_out] = self._juntar_categorias(categorias)

            self.sheets[aba] = df
            atualizados += 1

        self._sessao_codebook["items"] = itens_atualizados
        self.codificador.categorias = list(dict.fromkeys(
            self.codificador.categorias + categorias_globais
        ))
        self._log(f"Codebook refinado aplicado em {atualizados} resposta(s)", "OK", "g")
        self._atualizar_estado_refino()

    def _sel_todas(self):
        for cfg in self.abas_config.values():
            cfg["var"].set(True)

    def _des_todas(self):
        for cfg in self.abas_config.values():
            cfg["var"].set(False)

    # ── Rebuild tabela de abas — layout em GRID para responsividade ───────────
    def _rebuild_abas(self):
        for w in self.frame_abas.winfo_children():
            w.destroy()
        self.abas_config.clear()

        if not self.sheet_names:
            tk.Label(self.frame_abas, text="  Nenhuma aba encontrada",
                     bg=CARD, fg=TXT4, font=F_BODY).pack(anchor="w", pady=10)
            return

        from codificador import TIPOS_PERGUNTA
        labels_tipo = [v["label"] for v in TIPOS_PERGUNTA.values()]
        keys_tipo   = list(TIPOS_PERGUNTA.keys())
        paleta = [(AZUL_LIGHT, AZUL), (OURO_LIGHT, OURO),
                  (VERDE_LIGHT, VERDE), (ROXO_LIGHT, ROXO)]

        # Colunas da tabela expandem com a janela
        self.frame_abas.grid_columnconfigure(0, minsize=26)   # checkbox
        self.frame_abas.grid_columnconfigure(1, weight=1)     # aba
        self.frame_abas.grid_columnconfigure(2, weight=2)     # entrada
        self.frame_abas.grid_columnconfigure(3, weight=2)     # saida
        self.frame_abas.grid_columnconfigure(4, weight=3)     # tipo
        self.frame_abas.grid_columnconfigure(5, weight=2)     # modo
        self.frame_abas.grid_columnconfigure(6, minsize=32)   # btn

        for i, nome in enumerate(self.sheet_names):
            df      = self.sheets[nome]
            cols    = list(df.columns)
            row_bg  = CARD if i % 2 == 0 else "#fafafa"
            ibg, ifg = paleta[i % len(paleta)]
            row = i

            # Linha separadora
            sep = tk.Frame(self.frame_abas, bg=BORDER, height=1)
            sep.grid(row=row*2, column=0, columnspan=7, sticky="ew")

            # Célula container para cada coluna
            def cell(col, bg=row_bg, px=4, py=6):
                f = tk.Frame(self.frame_abas, bg=bg, padx=px, pady=py)
                f.grid(row=row*2+1, column=col, sticky="nsew")
                return f

            # Checkbox
            var = tk.BooleanVar(value=True)
            cf = cell(0, px=2)
            tk.Checkbutton(cf, variable=var, bg=row_bg,
                           activebackground=row_bg,
                           selectcolor=CARD, relief="flat",
                           cursor="hand2").pack(expand=True)

            # Pill aba
            af = cell(1)
            pill = tk.Frame(af, bg=ibg)
            pill.pack(fill=tk.X)
            tk.Label(pill, text=f" {nome[:10]} ", bg=ibg, fg=ifg,
                     font=("Segoe UI", 8, "bold"),
                     padx=4, pady=2, anchor="w").pack(fill=tk.X)

            # Entrada
            ef2 = cell(2)
            cb_in = ttk.Combobox(ef2, values=cols, state="readonly", font=F_BODY)
            cb_in.pack(fill=tk.X)
            cb_in.set(cols[0] if cols else "")

            # Saída
            sf2 = cell(3)
            cb_out = ttk.Combobox(sf2, values=cols + ["codigo_ia"],
                                  state="readonly", font=F_BODY)
            cb_out.pack(fill=tk.X)
            cb_out.set("codigo_ia")

            # Tipo
            tf2 = cell(4)
            cb_tipo = ttk.Combobox(tf2, values=labels_tipo,
                                   state="readonly", font=F_BODY)
            cb_tipo.pack(fill=tk.X)
            cb_tipo.set(labels_tipo[0])

            # Modo
            mf2 = cell(5)
            modo_dd = ModoDropdown(mf2)
            modo_dd.btn.pack(fill=tk.X)

            # Botão nova coluna
            bf2 = cell(6, px=2)
            self._mkbtn(bf2, "...", lambda n=nome: self._nova_col(n),
                        bg=BORDER, fg=TXT3, small=True).pack(expand=True)

            # Linha extra semiaberta (oculta por padrão, span completo)
            semi_row = row*2 + 2
            semi_container = tk.Frame(self.frame_abas, bg=row_bg)

            semi_inner = tk.Frame(semi_container, bg=row_bg, padx=8, pady=6)
            semi_inner.pack(fill=tk.X)
            semi_inner.grid_columnconfigure(0, weight=1)

            # ── Linha 1: colunas de saída ─────────────────────────────────────
            cols_row = tk.Frame(semi_inner, bg=row_bg)
            cols_row.pack(fill=tk.X, pady=(0, 8))

            tk.Label(cols_row, text="Col. Imputacao:", bg=row_bg,
                     fg=TXT4, font=F_MICRO).pack(side=tk.LEFT, padx=(0, 4))
            cb_imp = ttk.Combobox(cols_row, values=cols + ["col_imputado"],
                                  state="readonly", width=16, font=F_BODY)
            cb_imp.pack(side=tk.LEFT, padx=(0, 20))
            cb_imp.set("col_imputado")

            tk.Label(cols_row, text="Col. Nova:", bg=row_bg,
                     fg=TXT4, font=F_MICRO).pack(side=tk.LEFT, padx=(0, 4))
            cb_novo = ttk.Combobox(cols_row, values=cols + ["col_nova"],
                                   state="readonly", width=16, font=F_BODY)
            cb_novo.pack(side=tk.LEFT)
            cb_novo.set("col_nova")

            # ── Separador ─────────────────────────────────────────────────────
            tk.Frame(semi_inner, bg=BORDER, height=1).pack(fill=tk.X, pady=(0, 6))

            # ── Linha 2: categorias pré-definidas desta aba ───────────────────
            cat_hdr = tk.Frame(semi_inner, bg=row_bg)
            cat_hdr.pack(fill=tk.X, pady=(0, 4))
            tk.Label(cat_hdr, text="Categorias pré-definidas:",
                     bg=row_bg, fg=TXT2, font=F_MICRO).pack(side=tk.LEFT)
            tk.Label(cat_hdr,
                     text="(a IA vai tentar encaixar aqui antes de criar novas)",
                     bg=row_bg, fg=TXT4, font=F_MICRO).pack(side=tk.LEFT, padx=(6, 0))

            # Pills das categorias desta aba
            cats_aba = []   # lista de strings — categorias desta aba
            tags_frame = _WrapFrame(semi_inner, bg=row_bg)
            tags_frame.pack(fill=tk.X, pady=(0, 6), ipady=2)
            lbl_sem_cats = tk.Label(tags_frame,
                                    text="Nenhuma — IA criará todas as categorias",
                                    bg=row_bg, fg=TXT4, font=F_SMALL)
            lbl_sem_cats.place(x=0, y=0)
            tags_frame.config(height=24)

            def _add_cat_aba(entry_widget, tags_f, cats_list,
                             lbl_vazio, rbg=row_bg):
                txt = entry_widget.get().strip()
                if not txt or "Digite" in txt:
                    return
                # Remove placeholder se existe
                try:
                    lbl_vazio.place_forget()
                except Exception:
                    pass
                for cat in [c.strip() for c in txt.split(",") if c.strip()]:
                    if cat not in cats_list:
                        cats_list.append(cat)
                        tag = tk.Label(tags_f, text=f"  {cat}  ",
                                       bg=ROXO_LIGHT, fg=ROXO,
                                       font=("Segoe UI", 8, "bold"),
                                       padx=4, pady=2, cursor="hand2")
                        tag.place(x=0, y=0)
                        # Clique remove a categoria
                        def _remove(ev, t=tag, c=cat, cl=cats_list, tf=tags_f,
                                    lv=lbl_vazio, rbg2=rbg):
                            cl.remove(c)
                            t.destroy()
                            tf._reflow()
                            if not cl:
                                lv.place(x=0, y=0)
                                tf.config(height=24)
                        tag.bind("<Button-1>", _remove)
                        tag.bind("<Enter>", lambda e, t=tag: t.config(bg="#f3e8ff"))
                        tag.bind("<Leave>", lambda e, t=tag: t.config(bg=ROXO_LIGHT))
                tags_f._reflow()
                entry_widget.delete(0, tk.END)

            # Entry + botão adicionar categoria
            add_row = tk.Frame(semi_inner, bg=row_bg)
            add_row.pack(fill=tk.X)
            add_row.grid_columnconfigure(0, weight=1)

            ef_wrap = tk.Frame(add_row, bg=BORDER2, padx=1, pady=1)
            ef_wrap.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 6))
            ef_inner = tk.Frame(ef_wrap, bg=CARD)
            ef_inner.pack(fill=tk.X)
            entry_cat_aba = tk.Entry(ef_inner, bg=CARD, fg=TXT4,
                                     font=F_BODY, relief="flat", bd=0)
            entry_cat_aba.pack(fill=tk.X, ipady=5, padx=8)
            entry_cat_aba.insert(0, "Digite uma categoria...")
            entry_cat_aba.bind("<FocusIn>",
                lambda ev, e=entry_cat_aba:
                e.delete(0, tk.END) if "Digite" in e.get() else None)
            entry_cat_aba.bind("<Return>",
                lambda ev, e=entry_cat_aba, tf=tags_frame,
                       cl=cats_aba, lv=lbl_sem_cats:
                _add_cat_aba(e, tf, cl, lv))

            self._mkbtn(add_row, "+ Adicionar",
                        lambda e=entry_cat_aba, tf=tags_frame,
                               cl=cats_aba, lv=lbl_sem_cats:
                        _add_cat_aba(e, tf, cl, lv),
                        bg=ROXO, small=True).pack(side=tk.LEFT)

            # Campo contexto livre (oculto por padrão)
            ctx_container = tk.Frame(self.frame_abas, bg=row_bg)
            ctx_inner = tk.Frame(ctx_container, bg=row_bg, padx=8, pady=4)
            ctx_inner.pack(fill=tk.X)
            tk.Label(ctx_inner, text="Contexto:", bg=row_bg,
                     fg=TXT4, font=F_MICRO).pack(side=tk.LEFT, padx=(0, 4))
            e_ctx = tk.Entry(ctx_inner, bg=BG, fg=TXT2, font=F_BODY,
                             relief="flat", bd=1,
                             highlightbackground=BORDER2,
                             highlightthickness=1)
            e_ctx.pack(side=tk.LEFT, fill=tk.X, expand=True)
            e_ctx.insert(0, "Contexto personalizado...")
            e_ctx.bind("<FocusIn>",
                       lambda ev, e=e_ctx:
                       e.delete(0, tk.END) if "Contexto" in e.get() else None)

            # ── Callbacks de visibilidade ─────────────────────────────────────
            def _on_tipo(event=None, ct=cb_tipo, cc=ctx_container,
                         kts=keys_tipo, lts=labels_tipo, r=semi_row, n_cols=7):
                idx = lts.index(ct.get()) if ct.get() in lts else 0
                if kts[idx] == "livre":
                    cc.grid(row=r+1, column=0, columnspan=n_cols, sticky="ew")
                else:
                    cc.grid_remove()

            def _on_modo(modo, sc=semi_container, r=semi_row, n_cols=7):
                if "semi" in modo:
                    sc.grid(row=r, column=0, columnspan=n_cols, sticky="ew")
                else:
                    sc.grid_remove()

            cb_tipo.bind("<<ComboboxSelected>>", _on_tipo)
            _on_tipo()
            modo_dd.on_change = _on_modo

            self.abas_config[nome] = {
                "var": var, "cb_in": cb_in, "cb_out": cb_out,
                "cb_tipo": cb_tipo, "modo_dd": modo_dd,
                "cb_imp": cb_imp, "cb_novo": cb_novo, "e_ctx": e_ctx,
                "cats_aba": cats_aba,   # categorias pré-definidas desta aba (semiaberta)
                "keys_tipo": keys_tipo, "labels_tipo": labels_tipo,
            }

        # Linha final separadora
        tk.Frame(self.frame_abas, bg=BORDER, height=1).grid(
            row=len(self.sheet_names)*2, column=0, columnspan=7, sticky="ew")

    def _nova_col(self, nome_aba):
        novo = tk.simpledialog.askstring(
            "Nova Coluna", f"Nome da nova coluna para '{nome_aba}':",
            parent=self.root)
        if novo:
            cfg = self.abas_config[nome_aba]
            for key in ("cb_out", "cb_imp", "cb_novo"):
                cb   = cfg[key]
                vals = list(cb["values"])
                if novo not in vals:
                    vals.append(novo)
                    cb["values"] = vals
            cfg["cb_out"].set(novo)

    # ── Ações funcionais ──────────────────────────────────────────────────────
    def _abrir_arquivo(self):
        path = filedialog.askopenfilename(
            title="Selecionar planilha",
            filetypes=[("Planilhas", "*.xlsx *.csv"), ("Todos", "*.*")])
        if not path:
            return
        self.arquivo_dados = path
        nome = Path(path).name
        try:
            self._nova_sessao_codebook()
            if path.endswith(".csv"):
                self.sheets      = {"Planilha": pd.read_csv(path)}
                self.sheet_names = ["Planilha"]
            else:
                xl               = pd.ExcelFile(path)
                self.sheet_names = xl.sheet_names
                self.sheets      = {n: xl.parse(n) for n in self.sheet_names}
            total = sum(len(d) for d in self.sheets.values())
            self.lbl_arquivo.config(text=f"  {nome}")
            self.lbl_status.config(
                text=f"{len(self.sheet_names)} aba(s)  {total} linhas")
            self._upd_m("ARQUIVO", Path(nome).stem[:12],
                        f"{len(self.sheet_names)} aba(s)")
            self._upd_m("ABAS", len(self.sheet_names), "selecionadas")
            self._rebuild_abas()
            self._btn_trocar.pack(side=tk.RIGHT)   # exibe botão Trocar
            self._atualizar_estado_refino()
            self._log(f"Arquivo: {nome} - {len(self.sheet_names)} aba(s), {total} linhas")
        except Exception as e:
            messagebox.showerror("Erro ao abrir", str(e))

    def _importar_codigos(self):
        path = filedialog.askopenfilename(
            filetypes=[("JSON", "*.json"), ("Excel", "*.xlsx"), ("Todos", "*.*")])
        if not path:
            return
        try:
            if path.endswith(".json"):
                with open(path, encoding="utf-8") as f:
                    dados = json.load(f)
            else:
                df2   = pd.read_excel(path)
                dados = dict(zip(df2.iloc[:, 0].astype(str),
                                 df2.iloc[:, 1].astype(str)))
            self.codificador.carregar_codigos(dados)
            self.lbl_cats_info.config(
                text=f"  {len(self.codificador.codigos_base)} mapeamentos",
                fg=VERDE)
        except Exception as e:
            messagebox.showerror("Erro", str(e))

    def _add_cat(self):
        txt = self.entry_cat.get().strip()
        if not txt or "Adicionar" in txt:
            return
        for w in self.frame_tags.winfo_children():
            if isinstance(w, tk.Label) and "Nenhuma" in str(w.cget("text")):
                w.destroy()
        for cat in [c.strip() for c in txt.split(",")]:
            if cat:
                self.codificador.adicionar_categoria(cat)
                tag = tk.Label(self.frame_tags, text=f"  {cat}  ",
                               bg=AZUL_LIGHT, fg=AZUL,
                               font=("Segoe UI", 8, "bold"),
                               padx=5, pady=3, cursor="hand2")
                tag.place(x=0, y=0)  # posição inicial; _reflow vai reposicionar
        self.frame_tags._reflow()
        self.entry_cat.delete(0, tk.END)

    def _executar(self):
        if not self.sheets:
            messagebox.showwarning("Atencao", "Carregue uma planilha primeiro!")
            return
        sel = [(n, c) for n, c in self.abas_config.items() if c["var"].get()]
        if not sel:
            messagebox.showwarning("Atencao", "Selecione ao menos uma aba!")
            return
        self.btn_run.config(state="disabled")
        self.btn_refine.config(state="disabled")
        ctx = self.text_ctx.get("1.0", tk.END).strip()
        threading.Thread(target=self._run_all, args=(sel, ctx), daemon=True).start()

    def _run_all(self, selecionadas, ctx_global):
        total_abas         = len(selecionadas)
        self._fila_revisao = []
        self._nova_sessao_codebook()
        resultado          = {}

        for idx_aba, (nome, cfg) in enumerate(selecionadas):
            col_in  = cfg["cb_in"].get()
            col_out = cfg["cb_out"].get() or "codigo_ia"
            modo    = cfg["modo_dd"].get()
            labels_tipo = cfg["labels_tipo"]
            keys_tipo   = cfg["keys_tipo"]
            label_sel   = cfg["cb_tipo"].get()
            idx_tipo    = labels_tipo.index(label_sel) if label_sel in labels_tipo else 0
            tipo        = keys_tipo[idx_tipo]
            ctx_custom  = ""
            if tipo == "livre":
                ctx_raw    = cfg["e_ctx"].get().strip()
                ctx_custom = ctx_raw if "Contexto" not in ctx_raw else ctx_global

            df        = self.sheets[nome]
            respostas = df[col_in].astype(str).tolist()
            total     = len(respostas)
            self._sessao_codebook["total_respostas"] += total
            self.root.after(0, lambda n=nome, t=total, tp=label_sel, m=modo:
                self._log(f"'{n}' | {tp} | Modo: {m} | {t} respostas"))

            def _make_cb(n, ia, ta):
                def cb(i_local, t_local, resp, cat):
                    pct  = ((ia + (i_local+1)/max(t_local,1))/ta)*100
                    info = f"'{n}'  {i_local+1}/{t_local}"
                    self.root.after(0, lambda p=pct, nf=info:
                        self._set_progress(p, nf))
                    self.root.after(0, lambda r=resp, c=cat, l=i_local+1, tt=t_local:
                        self._log(f"  {l}/{tt}  {r[:30]} -> {c}", "...", "b"))
                return cb

            try:
                # Categorias da pesquisa anterior: busca pela aba de mesmo nome.
                # Se não achar pelo nome exato, tenta correspondência parcial
                # (ex: "Q1 - Motivação" bate com "Motivação").
                # Só cai para lista vazia se realmente não encontrar nada.
                cats_anteriores = []
                if self.pesquisa_anterior_ativa and self.pesquisa_anterior_por_aba:
                    por_aba = self.pesquisa_anterior_por_aba
                    # 1. Nome exato
                    if nome in por_aba:
                        cats_anteriores = por_aba[nome]
                    else:
                        # 2. Correspondência parcial (contém ou está contido)
                        nome_lower = nome.lower()
                        for chave, cats in por_aba.items():
                            if nome_lower in chave.lower() or chave.lower() in nome_lower:
                                cats_anteriores = cats
                                break
                    if cats_anteriores:
                        self.root.after(0, lambda n=nome, nc=len(cats_anteriores):
                            self._log(f"  '{n}' → {nc} categorias da pesquisa anterior", "INFO", "b"))
                    else:
                        self.root.after(0, lambda n=nome:
                            self._log(f"  '{n}' → sem aba correspondente na pesquisa anterior, IA criará categorias", "AVISO", "b"))

                cats_imp = cats_anteriores or cfg.get("cats_aba") or self.codificador.categorias[:]

                resultado = self.codificador.codificar_lote_modo(
                    respostas, tipo=tipo, modo=modo,
                    contexto_custom=ctx_custom,
                    categorias_imputacao=cats_imp,
                    categorias_anteriores=cats_anteriores,
                    callback_progresso=_make_cb(nome, idx_aba, total_abas))

                if "imputado" in resultado:
                    col_imp  = cfg["cb_imp"].get() or "col_imputado"
                    col_novo = cfg["cb_novo"].get() or "col_nova"
                    df[col_imp]  = resultado["imputado"]
                    df[col_novo] = resultado["novo"]
                    for idx_linha, resposta_txt in enumerate(respostas):
                        cats_imp = self._normalizar_lista_categorias(resultado["imputado"][idx_linha])
                        cats_novo = self._normalizar_lista_categorias(resultado["novo"][idx_linha])
                        categorias_item = cats_imp + [c for c in cats_novo if c not in cats_imp]
                        origens_item = {cat: "imputado" for cat in cats_imp}
                        for cat in cats_novo:
                            origens_item[cat] = "novo"
                        self._registrar_item_codebook(
                            nome, idx_linha, resposta_txt, categorias_item, origens_item,
                            {
                                "modo": modo,
                                "tipo": tipo,
                                "col_in": col_in,
                                "col_imp": col_imp,
                                "col_novo": col_novo,
                            }
                        )
                    self.root.after(0, lambda n=nome, ci=col_imp, cn=col_novo:
                        self._log(f"  '{n}' - imp:'{ci}'  novo:'{cn}'", "OK", "g"))
                else:
                    df[col_out] = resultado["resultado"]
                    for idx_linha, resposta_txt in enumerate(respostas):
                        categorias_item = self._normalizar_lista_categorias(
                            resultado["resultado"][idx_linha]
                        )
                        self._registrar_item_codebook(
                            nome, idx_linha, resposta_txt, categorias_item,
                            {cat: "resultado" for cat in categorias_item},
                            {
                                "modo": modo,
                                "tipo": tipo,
                                "col_in": col_in,
                                "col_out": col_out,
                            }
                        )
                    self.root.after(0, lambda n=nome:
                        self._log(f"  '{n}' concluido", "OK", "g"))
            except Exception as e:
                df[col_out] = ["ERRO"] * total
                self.root.after(0, lambda ex=e:
                    self._log(f"ERRO: {ex}", "ERR", "r"))

            self.sheets[nome] = df
            codificadas = resultado.get("resultado",
                          resultado.get("imputado", [""] * total))
            itens = [{"resposta": r, "categoria": c}
                     for r, c in zip(respostas, codificadas)
                     if c not in ("SEM_RESPOSTA", "ERRO", "")]
            sel2 = self.banco.selecionar_para_revisao(itens, n=5)
            if sel2:
                self._fila_revisao.append(
                    {"aba": nome, "tipo": tipo, "exemplos": sel2})

        self.root.after(0, lambda: self._set_progress(
            100, f"{total_abas} aba(s) concluida(s)"))
        self.root.after(0, lambda: self.btn_run.config(state="normal"))
        self.root.after(0, lambda: self.btn_train.config(state="normal"))
        self.root.after(0, self._atualizar_estado_refino)
        self.root.after(0, self._revisao_auto)

    def _revisao_auto(self):
        if self._fila_revisao:
            TelaRevisao(self.root, self.banco, self._fila_revisao,
                        sheets=self.sheets)
            self.btn_train.config(state="disabled")
            self._fila_revisao = []

    def _abrir_revisao(self):
        if not self._fila_revisao:
            messagebox.showinfo("Treinar IA", "Execute uma codificacao primeiro.")
            return
        TelaRevisao(self.root, self.banco, self._fila_revisao, sheets=self.sheets)
        self.btn_train.config(state="disabled")
        self._fila_revisao = []

    def _exportar(self):
        if not self.sheets:
            messagebox.showwarning("Atencao", "Nenhum dado para exportar!")
            return
        path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel", "*.xlsx"), ("CSV", "*.csv")])
        if not path:
            return
        try:
            if path.endswith(".csv"):
                list(self.sheets.values())[0].to_csv(path, index=False)
            else:
                with pd.ExcelWriter(path, engine="openpyxl") as writer:
                    for n, d in self.sheets.items():
                        d.to_excel(writer, sheet_name=n, index=False)
            self._log(f"Exportado: {Path(path).name}", "OK", "g")
            messagebox.showinfo("Sucesso", f"Arquivo salvo:\n{path}")
        except Exception as e:
            messagebox.showerror("Erro ao exportar", str(e))


if __name__ == "__main__":
    root = tk.Tk()
    App(root)
    root.mainloop()
