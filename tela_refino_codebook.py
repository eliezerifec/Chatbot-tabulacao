import tkinter as tk
from tkinter import messagebox, ttk


BG = "#f8fafc"
CARD = "#ffffff"
BORDER = "#dbe4ee"
TXT1 = "#0f172a"
TXT2 = "#475569"
TXT3 = "#64748b"
AZUL = "#1d4ed8"
AZUL_L = "#dbeafe"
VERDE = "#059669"
VERDE_L = "#dcfce7"
OURO = "#b45309"
OURO_L = "#ffedd5"

F_H1 = ("Segoe UI", 13, "bold")
F_H2 = ("Segoe UI", 10, "bold")
F_BODY = ("Segoe UI", 9)
F_SMALL = ("Segoe UI", 8)


def _center(win, w, h):
    win.update_idletasks()
    sw = win.winfo_screenwidth()
    sh = win.winfo_screenheight()
    x = max((sw - w) // 2, 10)
    y = max((sh - h) // 2, 10)
    win.geometry(f"{w}x{h}+{x}+{y}")


class TelaRefinoCodebook:
    def __init__(self, parent, sessao, on_apply=None):
        self.parent = parent
        self.sessao = sessao or {"items": []}
        self.items = [self._clone_item(item) for item in self.sessao.get("items", [])]
        self.on_apply = on_apply
        self.categoria_atual = None
        self._categoria_ids = {}

        self.win = tk.Toplevel(parent)
        self.win.title("Refinar Codebook")
        self.win.configure(bg=BG)
        self.win.minsize(1080, 640)
        self.win.grab_set()
        self.win.focus_set()
        _center(self.win, 1280, 760)

        self._build()
        self._refresh_all()

    def _clone_item(self, item):
        return {
            **item,
            "categorias": list(item.get("categorias", [])),
            "origens": dict(item.get("origens", {})),
        }

    def _build(self):
        hdr = tk.Frame(self.win, bg=CARD, padx=18, pady=14)
        hdr.pack(fill=tk.X)
        tk.Label(hdr, text="Refinar Codebook", bg=CARD, fg=TXT1, font=F_H1).pack(anchor="w")
        tk.Label(
            hdr,
            text="Veja as categorias criadas, a frequencia delas e recodifique respostas sem alterar os agentes.",
            bg=CARD, fg=TXT3, font=F_BODY
        ).pack(anchor="w", pady=(4, 0))

        top = tk.Frame(self.win, bg=BG, padx=18, pady=14)
        top.pack(fill=tk.X)
        self.lbl_stats = tk.Label(top, text="", bg=BG, fg=TXT2, font=F_BODY)
        self.lbl_stats.pack(side=tk.LEFT)

        filtro = tk.Frame(top, bg=BG)
        filtro.pack(side=tk.RIGHT)
        tk.Label(filtro, text="Buscar categoria:", bg=BG, fg=TXT2, font=F_SMALL).pack(side=tk.LEFT, padx=(0, 6))
        self.var_busca = tk.StringVar()
        ent = tk.Entry(filtro, textvariable=self.var_busca, font=F_BODY, relief="solid", bd=1)
        ent.pack(side=tk.LEFT, ipadx=40)
        self.var_busca.trace_add("write", lambda *args: self._refresh_categories())

        body = tk.Frame(self.win, bg=BG, padx=18, pady=(0, 18))
        body.pack(fill=tk.BOTH, expand=True)
        body.grid_columnconfigure(0, weight=2)
        body.grid_columnconfigure(1, weight=3)
        body.grid_rowconfigure(0, weight=1)

        self._build_categories_panel(body)
        self._build_responses_panel(body)

    def _build_categories_panel(self, parent):
        panel = tk.Frame(parent, bg=CARD, highlightbackground=BORDER, highlightthickness=1)
        panel.grid(row=0, column=0, sticky="nsew", padx=(0, 10))
        panel.grid_rowconfigure(1, weight=1)
        panel.grid_columnconfigure(0, weight=1)

        head = tk.Frame(panel, bg=CARD, padx=14, pady=12)
        head.grid(row=0, column=0, sticky="ew")
        tk.Label(head, text="Categorias", bg=CARD, fg=TXT1, font=F_H2).pack(side=tk.LEFT)

        btns = tk.Frame(head, bg=CARD)
        btns.pack(side=tk.RIGHT)
        tk.Button(
            btns, text="+ Nova categoria", bg=VERDE, fg="white", font=F_SMALL,
            relief="flat", padx=10, pady=5, cursor="hand2",
            command=self._criar_categoria_manual
        ).pack(side=tk.LEFT, padx=(0, 6))
        tk.Button(
            btns, text="Renomear", bg=AZUL, fg="white", font=F_SMALL,
            relief="flat", padx=10, pady=5, cursor="hand2",
            command=self._renomear_categoria
        ).pack(side=tk.LEFT)

        cols = ("categoria", "qtd", "pct", "novo")
        self.tree_cats = ttk.Treeview(panel, columns=cols, show="headings", height=18)
        self.tree_cats.grid(row=1, column=0, sticky="nsew", padx=14, pady=(0, 12))
        self.tree_cats.heading("categoria", text="Categoria")
        self.tree_cats.heading("qtd", text="Freq.")
        self.tree_cats.heading("pct", text="%")
        self.tree_cats.heading("novo", text="Novas")
        self.tree_cats.column("categoria", width=220, anchor="w")
        self.tree_cats.column("qtd", width=70, anchor="center")
        self.tree_cats.column("pct", width=70, anchor="center")
        self.tree_cats.column("novo", width=70, anchor="center")
        self.tree_cats.bind("<<TreeviewSelect>>", self._on_select_category)

        sb = ttk.Scrollbar(panel, orient="vertical", command=self.tree_cats.yview)
        sb.place(relx=1.0, x=-14, y=64, relheight=0.82, anchor="ne")
        self.tree_cats.configure(yscrollcommand=sb.set)

    def _build_responses_panel(self, parent):
        panel = tk.Frame(parent, bg=CARD, highlightbackground=BORDER, highlightthickness=1)
        panel.grid(row=0, column=1, sticky="nsew")
        panel.grid_rowconfigure(2, weight=1)
        panel.grid_columnconfigure(0, weight=1)

        head = tk.Frame(panel, bg=CARD, padx=14, pady=12)
        head.grid(row=0, column=0, sticky="ew")
        tk.Label(head, text="Respostas codificadas", bg=CARD, fg=TXT1, font=F_H2).pack(anchor="w")
        self.lbl_categoria = tk.Label(head, text="Selecione uma categoria a esquerda.", bg=CARD, fg=TXT3, font=F_BODY)
        self.lbl_categoria.pack(anchor="w", pady=(4, 0))

        acoes = tk.Frame(panel, bg=CARD, padx=14, pady=(0, 10))
        acoes.grid(row=1, column=0, sticky="ew")

        tk.Label(acoes, text="Mover selecionadas para:", bg=CARD, fg=TXT2, font=F_SMALL).pack(side=tk.LEFT)
        self.var_destino = tk.StringVar()
        self.cb_destino = ttk.Combobox(acoes, textvariable=self.var_destino, state="readonly", width=26, font=F_BODY)
        self.cb_destino.pack(side=tk.LEFT, padx=(8, 8))
        tk.Button(
            acoes, text="Mover", bg=AZUL, fg="white", font=F_SMALL,
            relief="flat", padx=10, pady=5, cursor="hand2",
            command=self._mover_para_categoria_existente
        ).pack(side=tk.LEFT, padx=(0, 6))
        tk.Button(
            acoes, text="Nova categoria p/ selecionadas", bg=OURO, fg="white", font=F_SMALL,
            relief="flat", padx=10, pady=5, cursor="hand2",
            command=self._criar_categoria_para_selecionadas
        ).pack(side=tk.LEFT)

        cols = ("aba", "linha", "origem", "categorias", "resposta")
        self.tree_resp = ttk.Treeview(panel, columns=cols, show="headings")
        self.tree_resp.grid(row=2, column=0, sticky="nsew", padx=14, pady=(0, 12))
        self.tree_resp.heading("aba", text="Aba")
        self.tree_resp.heading("linha", text="Linha")
        self.tree_resp.heading("origem", text="Origem")
        self.tree_resp.heading("categorias", text="Categorias")
        self.tree_resp.heading("resposta", text="Resposta")
        self.tree_resp.column("aba", width=110, anchor="w")
        self.tree_resp.column("linha", width=55, anchor="center")
        self.tree_resp.column("origem", width=80, anchor="center")
        self.tree_resp.column("categorias", width=180, anchor="w")
        self.tree_resp.column("resposta", width=520, anchor="w")
        self.tree_resp.configure(selectmode="extended")

        sb = ttk.Scrollbar(panel, orient="vertical", command=self.tree_resp.yview)
        sb.place(relx=1.0, x=-14, y=118, relheight=0.72, anchor="ne")
        self.tree_resp.configure(yscrollcommand=sb.set)

        rodape = tk.Frame(panel, bg=CARD, padx=14, pady=(0, 12))
        rodape.grid(row=3, column=0, sticky="ew")
        tk.Button(
            rodape, text="Fechar", bg="#e2e8f0", fg=TXT2, font=F_SMALL,
            relief="flat", padx=12, pady=7, cursor="hand2",
            command=self.win.destroy
        ).pack(side=tk.RIGHT)
        tk.Button(
            rodape, text="Aplicar alteracoes", bg=VERDE, fg="white", font=F_SMALL,
            relief="flat", padx=12, pady=7, cursor="hand2",
            command=self._aplicar_e_fechar
        ).pack(side=tk.RIGHT, padx=(0, 8))

    def _refresh_all(self):
        self._refresh_categories()
        self._refresh_destinos()
        self._refresh_responses()

    def _refresh_categories(self):
        busca = self.var_busca.get().strip().lower()
        total_itens = max(len(self.items), 1)
        stats = {}

        for item in self.items:
            vistos = set()
            for cat in item.get("categorias", []):
                if cat in vistos:
                    continue
                vistos.add(cat)
                if busca and busca not in cat.lower():
                    continue
                info = stats.setdefault(cat, {"qtd": 0, "novas": 0})
                info["qtd"] += 1
                if item.get("origens", {}).get(cat) == "novo":
                    info["novas"] += 1

        atual = self.categoria_atual if self.categoria_atual in stats else None
        self.categoria_atual = atual
        self._categoria_ids.clear()

        for iid in self.tree_cats.get_children():
            self.tree_cats.delete(iid)

        ordenadas = sorted(stats.items(), key=lambda kv: (-kv[1]["qtd"], kv[0].lower()))
        for idx, (cat, info) in enumerate(ordenadas):
            iid = f"cat::{idx}"
            self._categoria_ids[iid] = cat
            pct = round(info["qtd"] / total_itens * 100, 1)
            self.tree_cats.insert("", "end", iid=iid, values=(cat, info["qtd"], pct, info["novas"]))
            if cat == self.categoria_atual:
                self.tree_cats.selection_set(iid)

        total_categorias = len(stats)
        codificadas = self.sessao.get("respostas_codificadas", len(self.items))
        total_respostas = self.sessao.get("total_respostas", len(self.items))
        self.lbl_stats.config(
            text=f"{total_categorias} categorias  |  {codificadas} respostas codificadas  |  {total_respostas} respostas na rodada"
        )

        if not self.categoria_atual and ordenadas:
            primeiro_iid = next(iter(self._categoria_ids))
            self.tree_cats.selection_set(primeiro_iid)
            self._on_select_category()
        else:
            self._refresh_responses()

    def _refresh_destinos(self):
        categorias = sorted({cat for item in self.items for cat in item.get("categorias", [])}, key=str.lower)
        self.cb_destino["values"] = categorias
        if categorias and self.var_destino.get() not in categorias:
            self.var_destino.set(categorias[0])

    def _refresh_responses(self):
        for iid in self.tree_resp.get_children():
            self.tree_resp.delete(iid)

        categoria = self.categoria_atual
        if categoria:
            filtrados = [item for item in self.items if categoria in item.get("categorias", [])]
            self.lbl_categoria.config(text=f"Categoria selecionada: {categoria} ({len(filtrados)} resposta(s))")
        else:
            filtrados = list(self.items)
            self.lbl_categoria.config(text="Sem categoria selecionada. Mostrando todas as respostas.")

        for item in filtrados:
            cats = ", ".join(item.get("categorias", []))
            origem = self._origem_item(item, categoria)
            resposta = str(item.get("resposta", "")).replace("\n", " ").strip()
            if len(resposta) > 140:
                resposta = resposta[:137] + "..."
            self.tree_resp.insert(
                "", "end", iid=item["id"],
                values=(item["aba"], int(item["linha"]) + 1, origem, cats, resposta)
            )

    def _origem_item(self, item, categoria=None):
        if categoria and categoria in item.get("origens", {}):
            return item["origens"][categoria]
        origens = set(item.get("origens", {}).values())
        if len(origens) == 1:
            return next(iter(origens))
        if "novo" in origens and "imputado" in origens:
            return "misto"
        return "resultado"

    def _on_select_category(self, event=None):
        selecionado = self.tree_cats.selection()
        if not selecionado:
            self.categoria_atual = None
        else:
            self.categoria_atual = self._categoria_ids.get(selecionado[0])
        self._refresh_responses()

    def _categoria_existe(self, nome):
        nome_limpo = nome.strip().lower()
        for item in self.items:
            for cat in item.get("categorias", []):
                if cat.lower() == nome_limpo:
                    return cat
        return None

    def _criar_categoria_manual(self):
        nome = self._prompt_categoria("Nova categoria", "Nome da nova categoria:")
        if not nome:
            return
        existente = self._categoria_existe(nome)
        if existente:
            self.categoria_atual = existente
            self._refresh_all()
            return
        self.var_destino.set(nome)
        self.cb_destino["values"] = list(self.cb_destino["values"]) + [nome]
        messagebox.showinfo(
            "Refinar Codebook",
            "Categoria criada no codebook. Agora selecione respostas e mova para ela."
        )

    def _renomear_categoria(self):
        if not self.categoria_atual:
            messagebox.showwarning("Refinar Codebook", "Selecione uma categoria para renomear.")
            return

        novo_nome = self._prompt_categoria("Renomear categoria", "Novo nome da categoria:", self.categoria_atual)
        if not novo_nome or novo_nome == self.categoria_atual:
            return

        existente = self._categoria_existe(novo_nome)
        destino = existente or novo_nome
        origem_destino = "imputado" if existente else "novo"

        for item in self.items:
            cats = item.get("categorias", [])
            if self.categoria_atual not in cats:
                continue
            novas = []
            novas_origens = {}
            for cat in cats:
                if cat == self.categoria_atual:
                    if destino not in novas:
                        novas.append(destino)
                    novas_origens[destino] = item.get("origens", {}).get(destino, origem_destino)
                else:
                    if cat not in novas:
                        novas.append(cat)
                    novas_origens[cat] = item.get("origens", {}).get(cat, "resultado")
            item["categorias"] = novas
            item["origens"] = novas_origens

        self.categoria_atual = destino
        self._refresh_all()

    def _mover_para_categoria_existente(self):
        destino = self.var_destino.get().strip()
        if not destino:
            messagebox.showwarning("Refinar Codebook", "Escolha a categoria de destino.")
            return
        self._mover_respostas(destino, "imputado")

    def _criar_categoria_para_selecionadas(self):
        nome = self._prompt_categoria("Nova categoria", "Nome da categoria para as respostas selecionadas:")
        if not nome:
            return
        destino = self._categoria_existe(nome) or nome
        self.var_destino.set(destino)
        self._mover_respostas(destino, "novo" if destino == nome else "imputado")

    def _mover_respostas(self, destino, origem_destino):
        selecionadas = list(self.tree_resp.selection())
        if not selecionadas:
            messagebox.showwarning("Refinar Codebook", "Selecione pelo menos uma resposta.")
            return

        for iid in selecionadas:
            item = self._find_item(iid)
            if not item:
                continue
            cats = list(item.get("categorias", []))
            origens = dict(item.get("origens", {}))

            if self.categoria_atual and self.categoria_atual in cats:
                cats = [destino if cat == self.categoria_atual else cat for cat in cats]
                origens.pop(self.categoria_atual, None)
            elif len(cats) <= 1:
                cats = [destino]
                origens = {}
            elif destino not in cats:
                cats.append(destino)

            dedup = []
            novas_origens = {}
            for cat in cats:
                if cat not in dedup:
                    dedup.append(cat)
                if cat == destino:
                    novas_origens[cat] = origens.get(cat, origem_destino)
                else:
                    novas_origens[cat] = origens.get(cat, "resultado")

            item["categorias"] = dedup
            item["origens"] = novas_origens

        self.categoria_atual = destino
        self._refresh_all()

    def _find_item(self, iid):
        for item in self.items:
            if item.get("id") == iid:
                return item
        return None

    def _prompt_categoria(self, titulo, texto, valor_inicial=""):
        top = tk.Toplevel(self.win)
        top.title(titulo)
        top.configure(bg=BG)
        top.resizable(False, False)
        top.transient(self.win)
        top.grab_set()
        _center(top, 360, 150)

        tk.Label(top, text=texto, bg=BG, fg=TXT1, font=F_BODY).pack(anchor="w", padx=18, pady=(18, 8))
        var = tk.StringVar(value=valor_inicial)
        ent = tk.Entry(top, textvariable=var, font=F_BODY, relief="solid", bd=1)
        ent.pack(fill=tk.X, padx=18, ipady=5)
        ent.focus_set()
        ent.selection_range(0, tk.END)

        resultado = {"valor": None}

        def _confirmar():
            valor = var.get().strip()
            if not valor:
                return
            resultado["valor"] = valor
            top.destroy()

        botoes = tk.Frame(top, bg=BG)
        botoes.pack(fill=tk.X, padx=18, pady=18)
        tk.Button(
            botoes, text="Cancelar", bg="#e2e8f0", fg=TXT2, font=F_SMALL,
            relief="flat", padx=10, pady=5, command=top.destroy
        ).pack(side=tk.RIGHT)
        tk.Button(
            botoes, text="Salvar", bg=AZUL, fg="white", font=F_SMALL,
            relief="flat", padx=10, pady=5, command=_confirmar
        ).pack(side=tk.RIGHT, padx=(0, 8))

        ent.bind("<Return>", lambda event: _confirmar())
        top.wait_window()
        return resultado["valor"]

    def _aplicar_e_fechar(self):
        if self.on_apply:
            self.on_apply(self.items)
        self.win.destroy()
