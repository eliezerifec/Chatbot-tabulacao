"""
Codificador de Pesquisas — IFec RJ
Interface principal com modo dropdown hierárquico por aba
"""
# ── Fix OpenBLAS memory error — deve vir ANTES de qualquer import numérico ──
import os
os.environ.setdefault("OPENBLAS_NUM_THREADS", "1")
os.environ.setdefault("OMP_NUM_THREADS",      "1")
os.environ.setdefault("MKL_NUM_THREADS",      "1")
os.environ.setdefault("NUMEXPR_NUM_THREADS",  "1")
# ─────────────────────────────────────────────────────────────────────────────

import json
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import tkinter.simpledialog
import threading
import pandas as pd
from pathlib import Path
from codificador import CodificadorIA
from tela_revisao import TelaRevisao
from tela_tabulacao import TelaTabulacao   # ← v2.0

try:
    from PIL import Image, ImageTk
    PIL_OK = True
except ImportError:
    PIL_OK = False

NAV_BG     = "#111827"
NAV_ACTIVE = "#1d4ed8"
TOPBAR     = "#ffffff"
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
F_SEC    = ("Segoe UI", 10, "bold")
F_BODY   = ("Segoe UI", 9)
F_SMALL  = ("Segoe UI", 8)
F_MICRO  = ("Segoe UI", 7, "bold")
F_MONO   = ("Consolas", 9)
F_NAV    = ("Segoe UI", 15)
LOG_BG   = "#0d1117"
LOG_FG   = "#e6edf3"
LOG_GRN  = "#3fb950"
LOG_BLU  = "#79c0ff"
LOG_RED  = "#f85149"


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


class ModoDropdown:
    ITEMS = [
        ("Simples",                "simples",      False, CARD,       TXT1),
        ("Multipla",               "multipla",     False, CARD,       TXT1),
        (None,                     None,           True,  None,       None),
        ("Semiaberta - Simples",   "semi_simples", False, ROXO_LIGHT, ROXO),
        ("Semiaberta - Multipla",  "semi_multipla",False, ROXO_LIGHT, ROXO),
    ]
    LABELS = {
        "simples":       "Simples",
        "multipla":      "Multipla",
        "semi_simples":  "Semi Simples",
        "semi_multipla": "Semi Multipla",
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
            padx=8, pady=4, bd=0,
            activebackground=AZUL_MID, activeforeground=AZUL,
            command=self._toggle)

    def pack(self, **kw):
        self.btn.pack(**kw)

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
            inner = tk.Frame(row, bg=row_bg, padx=12, pady=9)
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
        popup.geometry(f"230x{popup.winfo_reqheight()}+{bx}+{by}")
        popup.bind("<FocusOut>", lambda e: self._close(popup))
        popup.focus_set()

    def _close(self, p):
        try:
            p.destroy()
        except Exception:
            pass
        self._win = None


class App:
    def __init__(self, root):
        self.root = root
        self.root.title("Codificador de Pesquisas  -  IFec RJ  v2.0")
        self.root.geometry("1160x940")
        self.root.configure(bg=NAV_BG)
        self.root.resizable(True, True)
        self.root.minsize(920, 680)
        self.arquivo_dados     = None
        self.sheets: dict      = {}
        self.sheet_names: list = []
        self.codificador       = CodificadorIA()
        self.banco             = self.codificador.banco
        self.abas_config: dict = {}
        self._fila_revisao     = []
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

    def _build(self):
        body = tk.Frame(self.root, bg=NAV_BG)
        body.pack(fill=tk.BOTH, expand=True)
        self._sidebar(body)
        self._main(body)

    def _sidebar(self, parent):
        sb = tk.Frame(parent, bg=NAV_BG, width=62)
        sb.pack(side=tk.LEFT, fill=tk.Y)
        sb.pack_propagate(False)
        logo_f = tk.Frame(sb, bg=AZUL_DARK, height=62)
        logo_f.pack(fill=tk.X)
        logo_f.pack_propagate(False)
        if PIL_OK:
            try:
                img = Image.open(Path(__file__).parent / "logo_ifec.png")
                img = img.resize((46, 46), Image.LANCZOS)
                self._logo_img = ImageTk.PhotoImage(img)
                tk.Label(logo_f, image=self._logo_img, bg=AZUL_DARK).pack(expand=True)
            except Exception:
                tk.Label(logo_f, text="IF", bg=AZUL_DARK, fg="#fbbf24",
                         font=("Segoe UI", 14, "bold")).pack(expand=True)
        else:
            tk.Label(logo_f, text="IF", bg=AZUL_DARK, fg="#fbbf24",
                     font=("Segoe UI", 14, "bold")).pack(expand=True)

        # (ícone, ativo, comando)  ← "=" = Codificador (ativo), "%" = Tabulação (novo v2.0)
        nav = [
            ("o", False, None),
            ("=", True,  None),
            ("%", False, self._abrir_tabulacao),   # ← v2.0
            ("*", False, None),
            ("+", False, None),
            ("~", False, None),
        ]
        for icon, active, cmd in nav:
            bg = NAV_ACTIVE if active else NAV_BG
            f  = tk.Frame(sb, bg=bg, height=52,
                          cursor="hand2" if cmd else "arrow")
            f.pack(fill=tk.X)
            f.pack_propagate(False)
            if active:
                tk.Frame(f, bg="#60a5fa", width=3).pack(side=tk.LEFT, fill=tk.Y)
            lbl = tk.Label(f, text=icon, bg=bg,
                           fg="#f9fafb" if active else "#6b7280",
                           font=F_NAV)
            lbl.pack(expand=True)
            if cmd:
                f.bind("<Button-1>",   lambda e, c=cmd: c())
                lbl.bind("<Button-1>", lambda e, c=cmd: c())
                f.bind("<Enter>",  lambda e, fr=f,  lb=lbl: [fr.config(bg="#1f2937"), lb.config(bg="#1f2937")])
                f.bind("<Leave>",  lambda e, fr=f,  lb=lbl: [fr.config(bg=NAV_BG),   lb.config(bg=NAV_BG)])
            elif not active:
                f.bind("<Enter>", lambda e, fr=f: fr.config(bg="#1f2937"))
                f.bind("<Leave>", lambda e, fr=f: fr.config(bg=NAV_BG))

        tk.Frame(sb, bg="#1f2937", height=1).pack(side=tk.BOTTOM, fill=tk.X)
        bot = tk.Frame(sb, bg=NAV_BG, height=60)
        bot.pack(side=tk.BOTTOM, fill=tk.X)
        bot.pack_propagate(False)
        av = tk.Frame(bot, bg=NAV_ACTIVE, width=34, height=34)
        av.pack(expand=True)
        av.pack_propagate(False)
        tk.Label(av, text="JD", bg=NAV_ACTIVE, fg="white",
                 font=("Segoe UI", 8, "bold")).pack(expand=True)

    def _main(self, parent):
        main = tk.Frame(parent, bg=BG)
        main.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self._topbar(main)
        wrap = tk.Frame(main, bg=BG)
        wrap.pack(fill=tk.BOTH, expand=True)
        canvas = tk.Canvas(wrap, bg=BG, highlightthickness=0)
        vsb    = ttk.Scrollbar(wrap, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=vsb.set)
        vsb.pack(side=tk.RIGHT, fill=tk.Y)
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.sf = tk.Frame(canvas, bg=BG, padx=24, pady=20)
        wid = canvas.create_window((0, 0), window=self.sf, anchor="nw")
        canvas.bind("<Configure>",
                    lambda e: canvas.itemconfig(wid, width=e.width))
        self.sf.bind("<Configure>",
                     lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.bind_all("<MouseWheel>",
                        lambda e: canvas.yview_scroll(-1*(e.delta//120), "units"))
        self._sec_metrics()
        self._sec_body()
        self._sec_actions()
        self._sec_log()

    def _topbar(self, parent):
        top = tk.Frame(parent, bg=TOPBAR, height=62)
        top.pack(fill=tk.X)
        top.pack_propagate(False)
        lf = tk.Frame(top, bg=TOPBAR)
        lf.pack(side=tk.LEFT, fill=tk.Y, padx=20)
        tk.Label(lf, text="Codificador de Pesquisas",
                 bg=TOPBAR, fg=TXT1,
                 font=("Segoe UI", 15, "bold")).pack(anchor="w", pady=(14, 0))
        tk.Label(lf, text="Analise automatica de respostas abertas com IA  -  IFec RJ  v2.0",
                 bg=TOPBAR, fg=TXT4, font=F_SMALL).pack(anchor="w")
        rf = tk.Frame(top, bg=TOPBAR)
        rf.pack(side=tk.RIGHT, fill=tk.Y, padx=20)
        # Botão Tabulação  ← v2.0
        tk.Button(
            rf, text="📊  Tabulação",
            bg=AZUL_LIGHT, fg=AZUL,
            font=("Segoe UI", 8, "bold"),
            relief="flat", bd=0, padx=10, pady=6,
            cursor="hand2",
            activebackground=AZUL_MID, activeforeground=AZUL,
            command=self._abrir_tabulacao,
        ).pack(side=tk.RIGHT, padx=(0, 12), pady=14)
        av = tk.Frame(rf, bg=AZUL, width=34, height=34)
        av.pack(side=tk.RIGHT, pady=14)
        av.pack_propagate(False)
        tk.Label(av, text="JD", bg=AZUL, fg="white",
                 font=("Segoe UI", 8, "bold")).pack(expand=True)
        tk.Frame(top, bg=BORDER, height=1).pack(side=tk.BOTTOM, fill=tk.X)

    def _sec_metrics(self):
        row = tk.Frame(self.sf, bg=BG)
        row.pack(fill=tk.X, pady=(0, 18))
        defs = [
            ("D", AZUL_LIGHT,  AZUL,  "ARQUIVO",  "-",      "nenhum carregado"),
            ("A", AZUL_LIGHT,  AZUL,  "ABAS",     "0",      "aguardando"),
            ("M", OURO_LIGHT,  OURO,  "MODELO",   "gpt-5.4", "Agente 1  ·  gpt-4o Ag2"),
            ("B", VERDE_LIGHT, VERDE, "BANCO IA", "-",      "exemplos salvos"),
        ]
        self._mw = {}
        for i, (ico, ibg, ifg, title, val, sub) in enumerate(defs):
            c = tk.Frame(row, bg=CARD)
            c.grid(row=0, column=i, sticky="nsew",
                   padx=(0, 10 if i < 3 else 0))
            row.grid_columnconfigure(i, weight=1)
            tk.Frame(c, bg=ifg, height=3).pack(fill=tk.X)
            body = tk.Frame(c, bg=CARD, padx=16, pady=14)
            body.pack(fill=tk.BOTH)
            top_r = tk.Frame(body, bg=CARD)
            top_r.pack(fill=tk.X)
            tf = tk.Frame(top_r, bg=CARD)
            tf.pack(side=tk.LEFT, fill=tk.X, expand=True)
            lv = tk.Label(tf, text=val, bg=CARD, fg=TXT1,
                          font=("Segoe UI", 16, "bold"))
            lv.pack(anchor="w")
            ls = tk.Label(tf, text=sub, bg=CARD, fg=TXT3, font=F_SMALL)
            ls.pack(anchor="w")
            ico_f = tk.Frame(top_r, bg=ibg, width=38, height=38)
            ico_f.pack(side=tk.RIGHT, anchor="ne")
            ico_f.pack_propagate(False)
            tk.Label(ico_f, text=ico, bg=ibg, fg=ifg,
                     font=("Segoe UI", 13, "bold")).pack(expand=True)
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

    def _sec_body(self):
        cols = tk.Frame(self.sf, bg=BG)
        cols.pack(fill=tk.BOTH, expand=True)
        cols.grid_columnconfigure(0, weight=60)
        cols.grid_columnconfigure(1, weight=40)
        left = tk.Frame(cols, bg=BG)
        left.grid(row=0, column=0, sticky="nsew", padx=(0, 12))
        self._card_upload(left)
        self._card_abas(left)
        right = tk.Frame(cols, bg=BG)
        right.grid(row=0, column=1, sticky="nsew")
        self._card_categorias(right)
        self._card_contexto(right)

    def _card_upload(self, parent):
        c = card_frame(parent)
        sec_header(c, "^", "Upload da Base de Dados",
                   "XLSX ou CSV - todas as abas serao detectadas")
        dz = tk.Frame(c, bg="#f0f7ff", cursor="hand2")
        dz.pack(fill=tk.X)
        tk.Frame(c, bg=AZUL_MID, height=1).pack(fill=tk.X)
        inner = tk.Frame(dz, bg="#f0f7ff", pady=26)
        inner.pack(fill=tk.X)
        tk.Label(inner, text="^", bg="#f0f7ff", fg=AZUL,
                 font=("Segoe UI", 28)).pack()
        row_msg = tk.Frame(inner, bg="#f0f7ff")
        row_msg.pack(pady=(4, 0))
        tk.Label(row_msg, text="Arraste a planilha aqui",
                 bg="#f0f7ff", fg=TXT1,
                 font=("Segoe UI", 11, "bold")).pack(side=tk.LEFT)
        tk.Label(row_msg, text="  ou  ", bg="#f0f7ff", fg=TXT4,
                 font=F_BODY).pack(side=tk.LEFT)
        lnk = tk.Label(row_msg, text="clique para selecionar",
                       bg="#f0f7ff", fg=AZUL,
                       font=("Segoe UI", 10, "bold"), cursor="hand2")
        lnk.pack(side=tk.LEFT)
        lnk.bind("<Button-1>", lambda e: self._abrir_arquivo())
        inner.bind("<Button-1>", lambda e: self._abrir_arquivo())
        dz.bind("<Button-1>", lambda e: self._abrir_arquivo())
        tk.Label(inner, text="Formatos aceitos  .xlsx   .csv",
                 bg="#f0f7ff", fg=TXT4, font=F_SMALL).pack(pady=(6, 0))
        self.lbl_arquivo = tk.Label(c, text="", bg=CARD,
                                    fg=VERDE, font=("Segoe UI", 9, "bold"))
        self.lbl_arquivo.pack(anchor="w", pady=(10, 0))
        self.lbl_status = tk.Label(c, text="", bg=CARD, fg=TXT4, font=F_SMALL)
        self.lbl_status.pack(anchor="w")

    def _card_abas(self, parent):
        c = card_frame(parent, mb=0)
        def _right_btns(f):
            bf = tk.Frame(f, bg=CARD)
            bf.pack(side=tk.RIGHT)
            self._mkbtn(bf, "Todas", self._sel_todas,
                        bg=AZUL_LIGHT, fg=AZUL, small=True).pack(
                            side=tk.LEFT, padx=(0, 4))
            self._mkbtn(bf, "Nenhuma", self._des_todas,
                        bg=BG, fg=TXT3, small=True).pack(side=tk.LEFT)
        sec_header(c, "=", "Configuracao das Abas",
                   "Tipo semantico e modo de resposta por aba",
                   right_fn=_right_btns)
        thead = tk.Frame(c, bg="#f9fafb")
        thead.pack(fill=tk.X, pady=(0, 2))
        COLS = [("", 2), ("ABA", 10), ("ENTRADA", 13), ("SAIDA", 11),
                ("TIPO", 20), ("MODO", 16), ("CONTEXTO", 15), ("", 4)]
        for txt, w in COLS:
            tk.Label(thead, text=txt, bg="#f9fafb", fg=TXT4,
                     font=F_MICRO, width=w, anchor="w",
                     padx=4, pady=5).pack(side=tk.LEFT)
        _hr(c, pady=(0, 2))
        self.frame_abas = tk.Frame(c, bg=CARD)
        self.frame_abas.pack(fill=tk.X)
        tk.Label(self.frame_abas,
                 text="  <- Carregue um arquivo para ver as abas",
                 bg=CARD, fg=TXT4, font=F_BODY).pack(anchor="w", pady=14)

    def _card_categorias(self, parent):
        c = card_frame(parent)
        def _right(f):
            self._mkbtn(f, "Importar", self._importar_codigos,
                        bg=AZUL_LIGHT, fg=AZUL, small=True).pack(side=tk.RIGHT)
        sec_header(c, "#", "Categorias de Codificacao",
                   "Usadas como guia para o modelo IA", right_fn=_right)
        self.frame_tags = tk.Frame(c, bg=CARD)
        self.frame_tags.pack(fill=tk.X, pady=(0, 10))
        tk.Label(self.frame_tags, text="Nenhuma categoria adicionada",
                 bg=CARD, fg=TXT4, font=F_SMALL).pack(anchor="w")
        _hr(c, pady=(0, 10))
        add = tk.Frame(c, bg=CARD)
        add.pack(fill=tk.X)
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
        self._mkbtn(add, "Adicionar", self._add_cat, bg=AZUL, small=True).pack(side=tk.LEFT)
        self.lbl_cats_info = tk.Label(c, text="", bg=CARD, fg=TXT4, font=F_SMALL)
        self.lbl_cats_info.pack(anchor="w", pady=(6, 0))

    def _card_contexto(self, parent):
        c = card_frame(parent, mb=0)
        sec_header(c, "i", "Contexto Global",
                   "Instrucao extra para o tipo 'Personalizado'")
        ef = tk.Frame(c, bg=BORDER2, padx=1, pady=1)
        ef.pack(fill=tk.X)
        ei = tk.Frame(ef, bg=CARD)
        ei.pack(fill=tk.X)
        self.text_ctx = tk.Text(ei, height=4, bg=CARD, fg=TXT2,
                                insertbackground=AZUL, font=F_BODY,
                                relief="flat", bd=0, padx=12, pady=10)
        self.text_ctx.pack(fill=tk.X)
        self.text_ctx.insert("1.0",
            "Pesquisa de satisfacao de evento.\nRespostas curtas de participantes.")

    def _sec_actions(self):
        bar = tk.Frame(self.sf, bg=BG)
        bar.pack(fill=tk.X, pady=(16, 14))
        bar.grid_columnconfigure(0, weight=3)
        bar.grid_columnconfigure(1, weight=3)
        bar.grid_columnconfigure(2, weight=3)
        bar.grid_columnconfigure(3, weight=2)
        self.btn_run = tk.Button(
            bar, text="  Iniciar Codificacao",
            bg=VERDE_BTN, fg="white",
            font=("Segoe UI", 11, "bold"), relief="flat",
            cursor="hand2", pady=13, bd=0,
            activebackground=VERDE, activeforeground="white",
            command=self._executar)
        self.btn_run.grid(row=0, column=0, sticky="ew", padx=(0, 8))
        self.btn_train = tk.Button(
            bar, text="  Treinar IA",
            bg=ROXO, fg="white",
            font=("Segoe UI", 11, "bold"), relief="flat",
            cursor="hand2", pady=13, bd=0, state="disabled",
            activebackground="#6d28d9", activeforeground="white",
            command=self._abrir_revisao)
        self.btn_train.grid(row=0, column=1, sticky="ew", padx=(0, 8))
        tk.Button(
            bar, text="  Exportar Resultado",
            bg=AZUL, fg="white",
            font=("Segoe UI", 11, "bold"), relief="flat",
            cursor="hand2", pady=13, bd=0,
            activebackground=AZUL_DARK, activeforeground="white",
            command=self._exportar
        ).grid(row=0, column=2, sticky="ew", padx=(0, 8))
        pc = tk.Frame(bar, bg=CARD)
        pc.grid(row=0, column=3, sticky="nsew")
        pi = tk.Frame(pc, bg=CARD, padx=14, pady=10)
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

    def _sec_log(self):
        c = card_frame(self.sf)
        sec_header(c, ">", "Log do Sistema", "Saida em tempo real")
        self.log = tk.Text(c, height=8, bg=LOG_BG, fg=LOG_FG,
                           font=F_MONO, relief="flat", bd=0,
                           padx=14, pady=12, state="disabled")
        self.log.pack(fill=tk.BOTH)
        self.log.tag_configure("g", foreground=LOG_GRN)
        self.log.tag_configure("b", foreground=LOG_BLU)
        self.log.tag_configure("r", foreground=LOG_RED)
        self._log("Sistema iniciado - Aguardando arquivo", "INFO")
        self._log("Agente 1: gpt-5.4  ·  Agente 2: gpt-4o  ·  Banco ativo", "INFO", "b")

    def _mkbtn(self, parent, text, cmd, bg=AZUL, fg="white",
               small=False, width=None):
        f  = ("Segoe UI", 7, "bold") if small else ("Segoe UI", 9, "bold")
        px, py = (8, 3) if small else (14, 7)
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

    def _sel_todas(self):
        for cfg in self.abas_config.values():
            cfg["var"].set(True)

    def _des_todas(self):
        for cfg in self.abas_config.values():
            cfg["var"].set(False)

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
        paleta = [
            (AZUL_LIGHT, AZUL), (OURO_LIGHT, OURO),
            (VERDE_LIGHT, VERDE), (ROXO_LIGHT, ROXO),
        ]
        for i, nome in enumerate(self.sheet_names):
            df      = self.sheets[nome]
            cols    = list(df.columns)
            row_bg  = CARD if i % 2 == 0 else "#fafafa"
            ibg, ifg = paleta[i % len(paleta)]

            row = tk.Frame(self.frame_abas, bg=row_bg, pady=7)
            row.pack(fill=tk.X)
            _hr(self.frame_abas, c="#f0f0f0")

            var = tk.BooleanVar(value=True)
            tk.Checkbutton(row, variable=var, bg=row_bg,
                           activebackground=row_bg,
                           selectcolor=CARD, relief="flat",
                           cursor="hand2").pack(side=tk.LEFT, padx=(4, 2))

            pill = tk.Frame(row, bg=ibg)
            pill.pack(side=tk.LEFT, padx=(2, 6))
            tk.Label(pill, text=f" {nome[:8]} ", bg=ibg, fg=ifg,
                     font=("Segoe UI", 8, "bold"),
                     padx=4, pady=2).pack()

            cb_in = ttk.Combobox(row, values=cols, state="readonly",
                                 width=12, font=F_BODY)
            cb_in.pack(side=tk.LEFT, padx=2)
            cb_in.set(cols[0] if cols else "")

            cb_out = ttk.Combobox(row, values=cols + ["codigo_ia"],
                                  state="readonly", width=11, font=F_BODY)
            cb_out.pack(side=tk.LEFT, padx=2)
            cb_out.set("codigo_ia")

            cb_tipo = ttk.Combobox(row, values=labels_tipo,
                                   state="readonly", width=19, font=F_BODY)
            cb_tipo.pack(side=tk.LEFT, padx=2)
            cb_tipo.set(labels_tipo[0])

            modo_dd = ModoDropdown(row)
            modo_dd.pack(side=tk.LEFT, padx=(6, 4))

            # Frame semiaberta - oculto por padrao
            semi_f = tk.Frame(row, bg=row_bg)

            tk.Label(semi_f, text="imp:", bg=row_bg, fg=TXT4,
                     font=F_MICRO).pack(side=tk.LEFT, padx=(4, 1))
            cb_imp = ttk.Combobox(semi_f, values=cols + ["col_imputado"],
                                  state="readonly", width=12, font=F_BODY)
            cb_imp.pack(side=tk.LEFT, padx=(0, 6))
            cb_imp.set("col_imputado")

            tk.Label(semi_f, text="novo:", bg=row_bg, fg=TXT4,
                     font=F_MICRO).pack(side=tk.LEFT, padx=(0, 1))
            cb_novo = ttk.Combobox(semi_f, values=cols + ["col_nova"],
                                   state="readonly", width=12, font=F_BODY)
            cb_novo.pack(side=tk.LEFT)
            cb_novo.set("col_nova")

            e_ctx = tk.Entry(row, bg=BG, fg=TXT2, font=F_BODY,
                             relief="flat", bd=1,
                             highlightbackground=BORDER2,
                             highlightthickness=1, width=14)
            e_ctx.insert(0, "Contexto...")
            e_ctx.bind("<FocusIn>",
                       lambda ev, e=e_ctx:
                       e.delete(0, tk.END) if "Contexto" in e.get() else None)

            self._mkbtn(row, "...", lambda n=nome: self._nova_col(n),
                        bg=BORDER, fg=TXT3, small=True).pack(side=tk.LEFT, padx=3)

            def _on_tipo(event=None, ct=cb_tipo, ec=e_ctx,
                         kts=keys_tipo, lts=labels_tipo):
                idx = lts.index(ct.get()) if ct.get() in lts else 0
                if kts[idx] == "livre":
                    ec.pack(side=tk.LEFT, padx=3, ipady=3)
                else:
                    ec.pack_forget()

            def _on_modo(modo, sf=semi_f):
                if "semi" in modo:
                    sf.pack(side=tk.LEFT, padx=(4, 0))
                else:
                    sf.pack_forget()

            cb_tipo.bind("<<ComboboxSelected>>", _on_tipo)
            _on_tipo()
            modo_dd.on_change = _on_modo

            self.abas_config[nome] = {
                "var": var, "cb_in": cb_in, "cb_out": cb_out,
                "cb_tipo": cb_tipo, "modo_dd": modo_dd,
                "cb_imp": cb_imp, "cb_novo": cb_novo, "e_ctx": e_ctx,
                "keys_tipo": keys_tipo, "labels_tipo": labels_tipo,
            }

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

    def _abrir_arquivo(self):
        path = filedialog.askopenfilename(
            title="Selecionar planilha",
            filetypes=[("Planilhas", "*.xlsx *.csv"), ("Todos", "*.*")])
        if not path:
            return
        self.arquivo_dados = path
        nome = Path(path).name
        try:
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
            self._upd_m("ARQUIVO", Path(nome).stem[:10],
                        f"{len(self.sheet_names)} aba(s)")
            self._upd_m("ABAS", len(self.sheet_names), "selecionadas")
            self._rebuild_abas()
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
                tag.pack(side=tk.LEFT, padx=(0, 5), pady=3)
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
        ctx = self.text_ctx.get("1.0", tk.END).strip()
        threading.Thread(target=self._run_all, args=(sel, ctx), daemon=True).start()

    def _run_all(self, selecionadas, ctx_global):
        total_abas         = len(selecionadas)
        self._fila_revisao = []
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
                cats_imp  = self.codificador.categorias[:]
                resultado = self.codificador.codificar_lote_modo(
                    respostas, tipo=tipo, modo=modo,
                    contexto_custom=ctx_custom,
                    categorias_imputacao=cats_imp,
                    callback_progresso=_make_cb(nome, idx_aba, total_abas))
                if "imputado" in resultado:
                    col_imp  = cfg["cb_imp"].get()  or "col_imputado"
                    col_novo = cfg["cb_novo"].get()  or "col_nova"
                    df[col_imp]  = resultado["imputado"]
                    df[col_novo] = resultado["novo"]
                    self.root.after(0, lambda n=nome, ci=col_imp, cn=col_novo:
                        self._log(f"  '{n}' - imp:'{ci}'  novo:'{cn}'", "OK", "g"))
                else:
                    df[col_out] = resultado["resultado"]
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

    # ── Abrir Tabulação ← v2.0 ───────────────────────────────────────────────
    def _abrir_tabulacao(self):
        TelaTabulacao(self.root, sheets=self.sheets if self.sheets else None)

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