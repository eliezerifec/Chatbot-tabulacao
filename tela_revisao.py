"""
Tela de Revisão — Popup pós-codificação
----------------------------------------
Fluxo:
  1. Aparece automaticamente ao terminar a codificação
  2. Pergunta "Me ajuda a te ajudar!" com Sim / Não
  3. Se Sim → revisão card a card (5 por aba)
  4. Se Não → mostra botão Exportar Planilha
"""

import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import pandas as pd
from aprendizado import BancoAprendizado

# ── Paleta ────────────────────────────────────────────────────────────────────
AZUL      = "#1a3a6b"
AZUL_L    = "#e8f0fb"
OURO      = "#c8971e"
OURO_L    = "#fef7e8"
VERDE     = "#16713d"
VERDE_L   = "#eaf5ee"
ROXO      = "#5b21b6"
BRANCO    = "#ffffff"
BG        = "#f4f6fa"
BORDER    = "#e2e8f0"
TXT_1     = "#0f172a"
TXT_2     = "#475569"
TXT_3     = "#94a3b8"

F_H1    = ("Segoe UI", 13, "bold")
F_H2    = ("Segoe UI", 10, "bold")
F_BODY  = ("Segoe UI", 9)
F_SMALL = ("Segoe UI", 8)
F_BOLD  = ("Segoe UI", 9, "bold")


def _center(win, w, h):
    win.update_idletasks()
    sw = win.winfo_screenwidth()
    sh = win.winfo_screenheight()
    win.geometry(f"{w}x{h}+{(sw-w)//2}+{(sh-h)//2}")


class TelaRevisao:
    """
    Abre automaticamente após codificação.
    parent   : janela principal (tk.Tk)
    banco    : BancoAprendizado
    fila     : lista de {'aba', 'tipo', 'exemplos': [...]}
    sheets   : dict nome->DataFrame (para exportar)
    """

    def __init__(self, parent, banco: BancoAprendizado,
                 fila: list[dict], sheets: dict = None):
        self.parent  = parent
        self.banco   = banco
        self.fila    = fila
        self.sheets  = sheets or {}

        self.idx_aba     = 0
        self.idx_exemplo = 0
        self.salvos      = 0
        self.total_exemplos = sum(len(item["exemplos"]) for item in self.fila)
        self._offsets_abas = []
        acumulado = 0
        for item in self.fila:
            self._offsets_abas.append(acumulado)
            acumulado += len(item["exemplos"])

        self._criar_janela()
        self._tela_convite()   # começa pela tela de convite

    # ── Janela base ───────────────────────────────────────────────────────────

    def _criar_janela(self):
        self.win = tk.Toplevel(self.parent)
        self.win.title("Treinar IA — Revisão Rápida")
        self.win.resizable(False, False)
        self.win.configure(bg=BG)
        self.win.grab_set()
        self.win.focus_set()

    def _limpar(self):
        for w in self.win.winfo_children():
            w.destroy()

    # ══════════════════════════════════════════════════════════════════════════
    # TELA 1 — Convite
    # ══════════════════════════════════════════════════════════════════════════

    def _tela_convite(self):
        self._limpar()
        _center(self.win, 420, 280)

        # Header
        hdr = tk.Frame(self.win, bg=AZUL)
        hdr.pack(fill=tk.X)
        tk.Label(hdr, text="🎓  Treinar IA", bg=AZUL, fg=BRANCO,
                 font=F_H1, pady=14, padx=18).pack(side=tk.LEFT)

        # Corpo
        body = tk.Frame(self.win, bg=BG, padx=30, pady=24)
        body.pack(fill=tk.BOTH, expand=True)

        tk.Label(body, text="Me ajuda a te ajudar!",
                 bg=BG, fg=TXT_1, font=("Segoe UI", 14, "bold")).pack(pady=(0, 8))

        tk.Label(body,
                 text="Quer me ajudar a ficar melhor?\nLeva menos de 1 minuto — só 5 respostas rápidas.",
                 bg=BG, fg=TXT_2, font=F_BODY, justify="center").pack(pady=(0, 22))

        btns = tk.Frame(body, bg=BG)
        btns.pack()

        tk.Button(btns, text="👍  Sim, vamos lá!", bg=VERDE, fg=BRANCO,
                  font=F_H2, relief="flat", padx=22, pady=10,
                  cursor="hand2", activebackground="#0f4f2a",
                  activeforeground=BRANCO,
                  command=self._iniciar_revisao).pack(side=tk.LEFT, padx=(0, 12))

        tk.Button(btns, text="Agora não", bg=BORDER, fg=TXT_2,
                  font=F_BODY, relief="flat", padx=16, pady=10,
                  cursor="hand2", activebackground="#cbd5e1",
                  activeforeground=TXT_1,
                  command=self._tela_exportar).pack(side=tk.LEFT)

    # ══════════════════════════════════════════════════════════════════════════
    # TELA 2 — Revisão card a card
    # ══════════════════════════════════════════════════════════════════════════

    def _iniciar_revisao(self):
        self._limpar()
        _center(self.win, 480, 400)

        # Header com progresso
        hdr = tk.Frame(self.win, bg=AZUL)
        hdr.pack(fill=tk.X)
        tk.Label(hdr, text="🎓  Treinar IA", bg=AZUL, fg=BRANCO,
                 font=F_H1, pady=12, padx=18).pack(side=tk.LEFT)
        self.lbl_hdr_prog = tk.Label(hdr, text="", bg=AZUL, fg="#7fa8d4",
                                     font=F_SMALL, padx=16)
        self.lbl_hdr_prog.pack(side=tk.RIGHT)

        # Corpo
        self.corpo = tk.Frame(self.win, bg=BG, padx=20, pady=14)
        self.corpo.pack(fill=tk.BOTH, expand=True)

        # Info aba
        self.lbl_aba = tk.Label(self.corpo, text="", bg=BG,
                                fg=TXT_3, font=F_SMALL)
        self.lbl_aba.pack(anchor="w")

        # Card resposta
        card = tk.Frame(self.corpo, bg=BRANCO, relief="solid", bd=1,
                        padx=16, pady=14)
        card.pack(fill=tk.X, pady=(6, 0))

        tk.Label(card, text="Resposta do participante:",
                 bg=BRANCO, fg=TXT_3, font=F_SMALL).pack(anchor="w")
        self.lbl_resposta = tk.Label(card, text="", bg=BRANCO, fg=TXT_1,
                                     font=("Segoe UI", 12, "bold"),
                                     wraplength=410, justify="left")
        self.lbl_resposta.pack(anchor="w", pady=(2, 12))

        tk.Label(card, text="IA classificou como:",
                 bg=BRANCO, fg=TXT_3, font=F_SMALL).pack(anchor="w")
        self.lbl_cat_ia = tk.Label(card, text="", bg=AZUL_L, fg=AZUL,
                                   font=F_BOLD, padx=10, pady=4)
        self.lbl_cat_ia.pack(anchor="w", pady=(2, 0))

        # Frame de correção — fica DENTRO do card, sempre presente, escondido
        self.frame_correcao = tk.Frame(card, bg=BRANCO)

        tk.Label(self.frame_correcao, text="Categoria correta:",
                 bg=BRANCO, fg=TXT_2, font=F_SMALL).pack(anchor="w", pady=(10, 2))

        entry_wrap = tk.Frame(self.frame_correcao, bg=BORDER, padx=1, pady=1)
        entry_wrap.pack(fill=tk.X)
        self.entry_correcao = tk.Entry(entry_wrap, font=("Segoe UI", 10),
                                       relief="flat", bd=0, bg=BRANCO)
        self.entry_correcao.pack(fill=tk.X, ipady=7, padx=8)
        self.entry_correcao.bind("<Return>", lambda e: self._salvar_correcao())

        # Botões
        self.btn_frame = tk.Frame(self.win, bg=BG, padx=20, pady=12)
        self.btn_frame.pack(fill=tk.X)

        self.btn_correto = tk.Button(
            self.btn_frame, text="✓  Correto", bg=VERDE, fg=BRANCO,
            font=F_BOLD, relief="flat", padx=16, pady=8, cursor="hand2",
            activebackground="#0f4f2a", activeforeground=BRANCO,
            command=self._aprovar)
        self.btn_correto.pack(side=tk.LEFT, padx=(0, 8))

        self.btn_corrigir = tk.Button(
            self.btn_frame, text="✎  Corrigir", bg=OURO, fg=BRANCO,
            font=F_BOLD, relief="flat", padx=16, pady=8, cursor="hand2",
            activebackground="#a57a14", activeforeground=BRANCO,
            command=self._mostrar_correcao)
        self.btn_corrigir.pack(side=tk.LEFT, padx=(0, 8))

        self.btn_salvar = tk.Button(
            self.btn_frame, text="Salvar →", bg=AZUL, fg=BRANCO,
            font=F_BOLD, relief="flat", padx=16, pady=8, cursor="hand2",
            activebackground="#122a52", activeforeground=BRANCO,
            command=self._salvar_correcao)
        # começa oculto
        self.btn_salvar.pack(side=tk.LEFT)
        self.btn_salvar.pack_forget()

        tk.Button(self.btn_frame, text="Pular →", bg=BORDER, fg=TXT_2,
                  font=F_BODY, relief="flat", padx=12, pady=8,
                  cursor="hand2", activebackground="#cbd5e1",
                  command=self._pular).pack(side=tk.RIGHT)

        # Barra de progresso
        self.progress = ttk.Progressbar(self.win, mode="determinate", maximum=100)
        self.progress.pack(fill=tk.X, padx=20, pady=(0, 14))

        self._mostrar_atual()

    # ── Lógica de navegação ───────────────────────────────────────────────────

    def _mostrar_atual(self):
        if self.idx_aba >= len(self.fila):
            self._tela_conclusao()
            return

        aba_data = self.fila[self.idx_aba]
        exemplos = aba_data["exemplos"]

        if self.idx_exemplo >= len(exemplos):
            self.idx_aba    += 1
            self.idx_exemplo = 0
            self._mostrar_atual()
            return

        exemplo = exemplos[self.idx_exemplo]

        total = self.total_exemplos
        feitos = self._offsets_abas[self.idx_aba] + self.idx_exemplo
        pct = (feitos / total * 100) if total else 0

        self.lbl_aba.config(
            text=f"Aba: {aba_data['aba']}  •  {self.idx_exemplo+1} de {len(exemplos)}")
        self.lbl_resposta.config(text=f'"{exemplo["resposta"]}"')
        self.lbl_cat_ia.config(text=exemplo["categoria"])
        self.lbl_hdr_prog.config(text=f"{feitos+1} / {total}")
        self.progress.config(value=pct)

        # Garante que o frame de correção está oculto e botões no estado inicial
        self._esconder_correcao()
        self.entry_correcao.delete(0, tk.END)
        self.entry_correcao.config(bg=BRANCO)

    def _mostrar_correcao(self):
        """Exibe o campo de texto para correção."""
        self.frame_correcao.pack(fill=tk.X, pady=(8, 0))
        self.btn_salvar.pack(side=tk.LEFT)
        self.btn_corrigir.pack_forget()
        self.btn_correto.pack_forget()
        self.entry_correcao.focus_set()

    def _esconder_correcao(self):
        self.frame_correcao.pack_forget()
        self.btn_salvar.pack_forget()
        if not self.btn_correto.winfo_ismapped():
            self.btn_correto.pack(side=tk.LEFT, padx=(0, 8))
        if not self.btn_corrigir.winfo_ismapped():
            self.btn_corrigir.pack(side=tk.LEFT, padx=(0, 8))

    def _aprovar(self):
        aba_data = self.fila[self.idx_aba]
        exemplo  = aba_data["exemplos"][self.idx_exemplo]
        self.banco.salvar(aba_data["tipo"], exemplo["resposta"],
                          exemplo["categoria"], exemplo["categoria"])
        self.salvos += 1
        self._avancar()

    def _salvar_correcao(self):
        nova = self.entry_correcao.get().strip()
        if not nova:
            self.entry_correcao.config(bg="#fef2f2")
            return
        aba_data = self.fila[self.idx_aba]
        exemplo  = aba_data["exemplos"][self.idx_exemplo]
        self.banco.salvar(aba_data["tipo"], exemplo["resposta"],
                          exemplo["categoria"], nova)
        self.salvos += 1
        self._avancar()

    def _pular(self):
        self._avancar()

    def _avancar(self):
        self.idx_exemplo += 1
        self._mostrar_atual()

    # ══════════════════════════════════════════════════════════════════════════
    # TELA 3 — Conclusão
    # ══════════════════════════════════════════════════════════════════════════

    def _tela_conclusao(self):
        self._limpar()
        _center(self.win, 400, 280)

        hdr = tk.Frame(self.win, bg=VERDE)
        hdr.pack(fill=tk.X)
        tk.Label(hdr, text="✅  Treino concluído!", bg=VERDE, fg=BRANCO,
                 font=F_H1, pady=14, padx=18).pack(side=tk.LEFT)

        body = tk.Frame(self.win, bg=BG, padx=30, pady=20)
        body.pack(fill=tk.BOTH, expand=True)

        try:
            st = self.banco.stats()
            info = (f"Salvei {self.salvos} exemplo(s) novo(s).\n\n"
                    f"📊 Banco coletivo:\n"
                    f"   {st['total']} exemplos  •  {st['taxa_acerto']}% de acerto\n\n"
                    f"Obrigado! Os próximos resultados\nserão mais precisos. 🚀")
        except Exception:
            info = f"Salvei {self.salvos} exemplo(s). Obrigado!"

        tk.Label(body, text=info, bg=BG, fg=TXT_1,
                 font=F_BODY, justify="left").pack(anchor="w", pady=(0, 20))

        btns = tk.Frame(body, bg=BG)
        btns.pack(anchor="w")

        tk.Button(btns, text="⬇  Exportar planilha", bg=AZUL, fg=BRANCO,
                  font=F_BOLD, relief="flat", padx=16, pady=8,
                  cursor="hand2", activebackground="#122a52",
                  activeforeground=BRANCO,
                  command=self._exportar).pack(side=tk.LEFT, padx=(0, 10))

        tk.Button(btns, text="Fechar", bg=BORDER, fg=TXT_2,
                  font=F_BODY, relief="flat", padx=14, pady=8,
                  cursor="hand2", command=self.win.destroy).pack(side=tk.LEFT)

    # ══════════════════════════════════════════════════════════════════════════
    # TELA — "Agora não" → direto para exportar
    # ══════════════════════════════════════════════════════════════════════════

    def _tela_exportar(self):
        self._limpar()
        _center(self.win, 380, 220)

        hdr = tk.Frame(self.win, bg=AZUL)
        hdr.pack(fill=tk.X)
        tk.Label(hdr, text="📊  Exportar resultado", bg=AZUL, fg=BRANCO,
                 font=F_H1, pady=14, padx=18).pack(side=tk.LEFT)

        body = tk.Frame(self.win, bg=BG, padx=30, pady=24)
        body.pack(fill=tk.BOTH, expand=True)

        tk.Label(body, text="Tudo pronto! Exporte a planilha\ncom os resultados da codificação.",
                 bg=BG, fg=TXT_2, font=F_BODY, justify="center").pack(pady=(0, 20))

        btns = tk.Frame(body, bg=BG)
        btns.pack()

        tk.Button(btns, text="⬇  Exportar planilha", bg=AZUL, fg=BRANCO,
                  font=F_BOLD, relief="flat", padx=18, pady=10,
                  cursor="hand2", activebackground="#122a52",
                  activeforeground=BRANCO,
                  command=self._exportar).pack(side=tk.LEFT, padx=(0, 10))

        tk.Button(btns, text="Fechar", bg=BORDER, fg=TXT_2,
                  font=F_BODY, relief="flat", padx=14, pady=10,
                  cursor="hand2", command=self.win.destroy).pack(side=tk.LEFT)

    # ── Exportar ──────────────────────────────────────────────────────────────

    def _exportar(self):
        if not self.sheets:
            messagebox.showinfo("Exportar",
                                "Nenhum dado disponível.\nExporte pela janela principal.",
                                parent=self.win)
            return
        path = filedialog.asksaveasfilename(
            parent=self.win,
            defaultextension=".xlsx",
            filetypes=[("Excel", "*.xlsx"), ("CSV", "*.csv")])
        if not path:
            return
        try:
            if path.endswith(".csv"):
                list(self.sheets.values())[0].to_csv(path, index=False)
            else:
                with pd.ExcelWriter(path, engine="openpyxl") as writer:
                    for nome, df in self.sheets.items():
                        df.to_excel(writer, sheet_name=nome, index=False)
            messagebox.showinfo("Sucesso", f"Planilha salva!", parent=self.win)
            self.win.destroy()
        except Exception as e:
            messagebox.showerror("Erro", str(e), parent=self.win)
