"""
tela_pipeline.py — Pipeline completo: base bruta → base final + tabulação
==========================================================================
Fluxo automatizado em 5 etapas, com 2 checkpoints de revisão humana:

  1. Upload      — base SurveyMonkey (.xlsx/.csv) + questionário (.docx)
  2. Limpeza     — IA extrai os pulos do questionário → você aprova → aplica
  3. Codificação — abertas detectadas automaticamente → IA codifica
  4. Categorias  — você revisa/renomeia as categorias criadas
  5. Base final  — merge da codificação + tabulação Excel e PowerPoint

Reutiliza: tabulador.py (set_header, detectar_perguntas, exportar_excel),
codificador.py (CodificadorIA), gerador_ppt.py e limpeza.py.
"""

from __future__ import annotations

import os
import tempfile
from io import BytesIO
from pathlib import Path

import pandas as pd
import streamlit as st

_MODOS = {"Simples": "simples", "Múltipla": "multipla"}
_SEP_CAT = ", "


# ─────────────────────────────────────────────────────────────────────────────
# HELPERS
# ─────────────────────────────────────────────────────────────────────────────

def _sg(key, default=None):
    return st.session_state.get(key, default)


def _ss(key, value):
    st.session_state[key] = value


def _steps(current: int, total: int = 5) -> None:
    labels = ["Upload", "Limpeza", "Codificação", "Categorias", "Base final"]
    partes = []
    for i in range(1, total + 1):
        cls = "done" if i < current else ("active" if i == current else "")
        partes.append(
            f'<span class="ifec-step {cls}"></span>'
            f'<span style="font-size:0.72rem; color:#475569; margin-right:0.9rem;">'
            f'{labels[i - 1]}</span>'
        )
    st.markdown(
        f'<div class="ifec-steps" style="margin-bottom:1rem;">{"".join(partes)}</div>',
        unsafe_allow_html=True,
    )


def _df_para_excel(sheets: dict[str, pd.DataFrame]) -> bytes:
    bio = BytesIO()
    with pd.ExcelWriter(bio, engine="openpyxl") as writer:
        for nome, df in sheets.items():
            df.to_excel(writer, sheet_name=str(nome)[:31], index=False)
    return bio.getvalue()


def _ler_base(name: str, data: bytes, two_line: bool):
    """Lê a base e retorna (df, tipos_sm, q0_map)."""
    from tabulador import set_header

    bio = BytesIO(data)
    if name.lower().endswith(".csv"):
        raw = pd.read_csv(bio, header=None, dtype=object)
    else:
        raw = pd.read_excel(bio, header=None, sheet_name=0)
    if two_line:
        return set_header(raw)
    df = raw.copy()
    df.columns = [str(c) for c in raw.iloc[0]]
    return df.iloc[1:].reset_index(drop=True), {}, {}


def _sanitize_export_df(df: pd.DataFrame) -> pd.DataFrame:
    clean = df.copy()
    seen: dict[str, int] = {}
    cols = []
    for col in clean.columns:
        nome = str(col)
        if nome in seen:
            seen[nome] += 1
            cols.append(f"{nome}_{seen[nome]}")
        else:
            seen[nome] = 0
            cols.append(nome)
    clean.columns = cols
    for col in clean.columns:
        if isinstance(clean[col], pd.DataFrame):
            clean[col] = clean[col].iloc[:, 0]
        try:
            clean[col] = clean[col].astype(object)
        except Exception:
            pass
    return clean


@st.cache_resource
def _get_codificador():
    from codificador import CodificadorIA
    return CodificadorIA()


def _grupos_abertas(df: pd.DataFrame, perguntas: list[dict]) -> list[dict]:
    """
    Grupos de colunas abertas a codificar, um por pergunta.
    Cada grupo agrupa subcolunas da mesma pergunta ("Outro. Qual?" ou "1:/2:/3:")
    para que recebam o MESMO conjunto de categorias.
    """
    grupos = []
    for p in perguntas:
        cols = [c for c in p.get("cols_outros", []) if c in df.columns]
        if not cols:
            continue
        n_resp = 0
        for c in cols:
            s = df[c]
            n_resp += int((s.notna() & ~s.astype(str).str.strip()
                           .isin(["", "-", "nan"])).sum())
        if n_resp == 0:
            continue
        eh_outro = p["tipo"] != "ABERTA"
        grupos.append({
            "num": p["num"],
            "pergunta": p["pergunta"],
            "cols": cols,
            "n_respostas": n_resp,
            "eh_outro": eh_outro,
        })
    return grupos


def _aplicar_renomeacao(serie: pd.Series, mapa: dict[str, str]) -> pd.Series:
    """Renomeia categorias token a token (células podem ter 'A, B, C')."""
    def _troca(val):
        if pd.isna(val):
            return val
        partes = [t.strip() for t in str(val).split(_SEP_CAT) if t.strip()]
        novas = []
        for t in partes:
            novo = mapa.get(t, t)
            if novo and novo not in novas:
                novas.append(novo)
        return _SEP_CAT.join(novas) if novas else pd.NA
    return serie.map(_troca)


def _reset_pipeline():
    for k in list(st.session_state.keys()):
        if k.startswith("pipe_"):
            st.session_state.pop(k, None)
    _ss("pipe_step", 1)


# ─────────────────────────────────────────────────────────────────────────────
# ETAPA 1 — UPLOAD
# ─────────────────────────────────────────────────────────────────────────────

def _etapa1_upload():
    st.markdown('<div class="ifec-card">', unsafe_allow_html=True)
    st.markdown('<p class="ifec-section-title">Arquivos da pesquisa</p>',
                unsafe_allow_html=True)
    st.markdown(
        '<p class="ifec-section-sub">Suba a base exportada do SurveyMonkey e o '
        'questionário em Word — os pulos descritos nele guiam a limpeza</p>',
        unsafe_allow_html=True,
    )

    base = st.file_uploader(
        "Base do SurveyMonkey (.xlsx ou .csv)",
        type=["xlsx", "csv"], key="pipe_up_base",
    )
    quest = st.file_uploader(
        "Questionário (.docx) — opcional, mas necessário para a limpeza de pulos",
        type=["docx"], key="pipe_up_quest",
    )
    titulo = st.text_input("Título da pesquisa (usado na tabulação)",
                           value=_sg("pipe_titulo", ""), key="pipe_titulo_input")
    two_line = st.checkbox("Cabeçalho em duas linhas (padrão SurveyMonkey)",
                           value=True, key="pipe_two_line")
    st.markdown('</div>', unsafe_allow_html=True)

    if base is not None:
        _ss("pipe_base_name", base.name)
        _ss("pipe_base_bytes", base.getvalue())
    if quest is not None:
        _ss("pipe_quest_name", quest.name)
        _ss("pipe_quest_bytes", quest.getvalue())

    _, col_r = st.columns([2, 1])
    with col_r:
        if st.button("Avançar", type="primary", use_container_width=True,
                     disabled="pipe_base_bytes" not in st.session_state,
                     key="pipe_next1"):
            try:
                df, tipos_sm, q0_map = _ler_base(
                    _sg("pipe_base_name"), _sg("pipe_base_bytes"), two_line
                )
                from tabulador import detectar_perguntas
                perguntas = detectar_perguntas(df, tipos_sm, q0_map)
            except Exception as exc:
                st.error(f"Não foi possível ler a base: {exc}")
                return
            _ss("pipe_df", df)
            _ss("pipe_tipos_sm", tipos_sm)
            _ss("pipe_q0_map", q0_map)
            _ss("pipe_perguntas", perguntas)
            _ss("pipe_titulo", titulo or Path(_sg("pipe_base_name", "Pesquisa")).stem)
            _ss("pipe_step", 2)
            st.rerun()


# ─────────────────────────────────────────────────────────────────────────────
# ETAPA 2 — LIMPEZA (checkpoint 1)
# ─────────────────────────────────────────────────────────────────────────────

def _etapa2_limpeza():
    import limpeza

    df = _sg("pipe_df")
    perguntas = _sg("pipe_perguntas", [])

    st.markdown('<div class="ifec-card">', unsafe_allow_html=True)
    st.markdown('<p class="ifec-section-title">Limpeza da base</p>',
                unsafe_allow_html=True)
    st.markdown(
        '<p class="ifec-section-sub">Revise as regras de pulo extraídas do '
        'questionário antes de aplicar</p>',
        unsafe_allow_html=True,
    )

    c1, c2, c3 = st.columns(3)
    c1.metric("Respondentes", len(df))
    c2.metric("Colunas", len(df.columns))
    c3.metric("Perguntas detectadas", len(perguntas))

    with st.expander("Perguntas detectadas na base"):
        st.dataframe(pd.DataFrame(
            [{"Nº": p["num"], "Tipo": p["tipo"],
              "Pergunta": p["pergunta"][:120],
              "Colunas": len(p["colunas"]) + len(p.get("cols_outros", []))}
             for p in perguntas]
        ), use_container_width=True, hide_index=True)

    # ── Extração das regras via IA ───────────────────────────────────────────
    tem_quest = "pipe_quest_bytes" in st.session_state
    api_ok = bool(os.getenv("OPENAI_API_KEY"))

    if not tem_quest:
        st.info("Sem questionário (.docx): a limpeza fará apenas remoção de "
                "linhas vazias e duplicados. Volte à etapa 1 para anexá-lo.")
    elif not api_ok:
        st.warning("OPENAI_API_KEY não configurada — não é possível extrair "
                   "as regras do questionário.")
    elif "pipe_regras" not in st.session_state:
        if st.button("Extrair regras de pulo do questionário (IA)",
                     type="primary", key="pipe_extrair"):
            with st.spinner("Lendo o questionário e extraindo a lógica de pulos..."):
                try:
                    texto = limpeza.extrair_texto_docx(_sg("pipe_quest_bytes"))
                    regras = limpeza.extrair_regras_questionario(texto, df, perguntas)
                    _ss("pipe_regras", regras)
                    st.rerun()
                except Exception as exc:
                    st.error(f"Erro na extração de regras: {exc}")

    regras = _sg("pipe_regras", [])

    if regras:
        aval = limpeza.avaliar_regras(df, perguntas, regras)
        st.markdown("**Regras extraídas do questionário** — desmarque as que "
                    "não devem ser aplicadas:")
        editado = st.data_editor(
            aval,
            column_config={
                "id": st.column_config.TextColumn("Regra", disabled=True),
                "Ativa": st.column_config.CheckboxColumn("Ativa"),
                "Descrição": st.column_config.TextColumn(disabled=True, width="large"),
                "Condição": st.column_config.TextColumn(disabled=True),
                "Perguntas-alvo": st.column_config.TextColumn(disabled=True),
                "Respondentes na condição": st.column_config.NumberColumn(disabled=True),
                "Violações (linhas)": st.column_config.NumberColumn(disabled=True),
                "Células a limpar": st.column_config.NumberColumn(disabled=True),
            },
            hide_index=True, use_container_width=True, key="pipe_regras_editor",
        )
        ativas = dict(zip(editado["id"], editado["Ativa"]))
        for r in regras:
            r["ativa"] = bool(ativas.get(r["id"], True))

        if st.button("Reextrair regras", key="pipe_reextrair"):
            st.session_state.pop("pipe_regras", None)
            st.rerun()
    elif "pipe_regras" in st.session_state:
        st.info("A IA não encontrou regras de pulo no questionário. "
                "A limpeza fará apenas a parte genérica.")

    st.divider()
    rem_vazias = st.checkbox("Remover linhas sem nenhuma pergunta respondida",
                             value=True, key="pipe_rem_vazias")
    rem_dup = st.checkbox("Remover respondentes duplicados (respondent_id)",
                          value=True, key="pipe_rem_dup")
    st.markdown('</div>', unsafe_allow_html=True)

    col_a, col_b, col_c = st.columns(3)
    with col_a:
        if st.button("Voltar", use_container_width=True, key="pipe_back2"):
            _ss("pipe_step", 1)
            st.rerun()
    with col_b:
        if st.button("Pular limpeza", use_container_width=True, key="pipe_skip2"):
            _ss("pipe_df_limpo", df.copy())
            _ss("pipe_relatorio", pd.DataFrame(
                columns=["respondente", "acao", "regra", "detalhe"]))
            _ss("pipe_resumo", {"linhas_antes": len(df), "linhas_depois": len(df)})
            _ss("pipe_step", 3)
            st.rerun()
    with col_c:
        if st.button("Aplicar limpeza e avançar", type="primary",
                     use_container_width=True, key="pipe_next2"):
            with st.spinner("Aplicando limpeza..."):
                df_limpo, relatorio, resumo = limpeza.aplicar_limpeza(
                    df, perguntas, regras,
                    remover_sem_resposta=rem_vazias,
                    remover_duplicados=rem_dup,
                )
            _ss("pipe_df_limpo", df_limpo)
            _ss("pipe_relatorio", relatorio)
            _ss("pipe_resumo", resumo)
            _ss("pipe_step", 3)
            st.rerun()


# ─────────────────────────────────────────────────────────────────────────────
# ETAPA 3 — CODIFICAÇÃO DAS ABERTAS
# ─────────────────────────────────────────────────────────────────────────────

def _etapa3_codificacao():
    df = _sg("pipe_df_limpo")
    perguntas = _sg("pipe_perguntas", [])
    resumo = _sg("pipe_resumo", {})
    relatorio = _sg("pipe_relatorio")

    st.markdown('<div class="ifec-card">', unsafe_allow_html=True)
    st.markdown('<p class="ifec-section-title">Codificação das abertas</p>',
                unsafe_allow_html=True)
    st.markdown(
        '<p class="ifec-section-sub">Colunas abertas detectadas automaticamente — '
        'subcolunas da mesma pergunta recebem as mesmas categorias</p>',
        unsafe_allow_html=True,
    )

    if resumo:
        acoes = len(relatorio) if relatorio is not None else 0
        st.caption(
            f"Limpeza aplicada: {resumo.get('linhas_antes', '?')} → "
            f"{resumo.get('linhas_depois', '?')} respondentes, "
            f"{resumo.get('celulas_limpas', 0)} célula(s) de pulo limpas, "
            f"{acoes} ação(ões) registradas no relatório."
        )

    grupos = _sg("pipe_grupos")
    if grupos is None:
        grupos = _grupos_abertas(df, perguntas)
        _ss("pipe_grupos", grupos)

    if not grupos:
        st.info("Nenhuma coluna aberta com respostas foi detectada na base.")
        st.markdown('</div>', unsafe_allow_html=True)
        col_a, col_b = st.columns(2)
        with col_a:
            if st.button("Voltar", use_container_width=True, key="pipe_back3"):
                _ss("pipe_step", 2)
                st.rerun()
        with col_b:
            if st.button("Avançar sem codificar", type="primary",
                         use_container_width=True, key="pipe_skip3"):
                _ss("pipe_coded", {})
                _ss("pipe_step", 5)
                st.rerun()
        return

    tabela = pd.DataFrame([{
        "Codificar": True,
        "Nº": g["num"],
        "Pergunta": g["pergunta"][:110],
        "Tipo": "Outro (semiaberta)" if g["eh_outro"] else "Aberta",
        "Subcolunas": len(g["cols"]),
        "Respostas": g["n_respostas"],
        "Modo": "Simples",
    } for g in grupos])

    editado = st.data_editor(
        tabela,
        column_config={
            "Codificar": st.column_config.CheckboxColumn("Codificar"),
            "Nº": st.column_config.TextColumn(disabled=True),
            "Pergunta": st.column_config.TextColumn(disabled=True, width="large"),
            "Tipo": st.column_config.TextColumn(disabled=True),
            "Subcolunas": st.column_config.NumberColumn(disabled=True),
            "Respostas": st.column_config.NumberColumn(disabled=True),
            "Modo": st.column_config.SelectboxColumn(
                "Modo", options=list(_MODOS.keys()),
                help="Múltipla: separa a resposta por vírgula e codifica cada parte"),
        },
        hide_index=True, use_container_width=True, key="pipe_grupos_editor",
    )

    contexto = st.text_area(
        "Contexto da pesquisa para a IA (tema, público, objetivo)",
        value=_sg("pipe_contexto", ""),
        placeholder="Ex.: Pesquisa com responsáveis sobre interesse em Ensino "
                    "Médio com Curso Técnico Integrado no RJ.",
        key="pipe_contexto_input",
    )
    st.markdown('</div>', unsafe_allow_html=True)

    api_ok = bool(os.getenv("OPENAI_API_KEY"))
    if not api_ok:
        st.warning("OPENAI_API_KEY não configurada — a codificação não pode rodar.")

    col_a, col_b = st.columns(2)
    with col_a:
        if st.button("Voltar", use_container_width=True, key="pipe_back3b"):
            _ss("pipe_step", 2)
            st.rerun()
    with col_b:
        if st.button("Iniciar codificação com IA", type="primary",
                     use_container_width=True, disabled=not api_ok,
                     key="pipe_next3"):
            _ss("pipe_contexto", contexto)
            selecionados = [
                (g, _MODOS.get(linha["Modo"], "simples"))
                for g, (_, linha) in zip(grupos, editado.iterrows())
                if linha["Codificar"]
            ]
            if not selecionados:
                st.error("Selecione ao menos uma pergunta para codificar.")
                return
            _rodar_codificacao(df, selecionados, contexto)


def _rodar_codificacao(df: pd.DataFrame, selecionados: list, contexto: str):
    codificador = _get_codificador()
    coded: dict[str, pd.Series] = {}
    info_grupos: list[dict] = []
    progress = st.progress(0.0, text="Preparando codificação...")

    for gi, (grupo, modo) in enumerate(selecionados):
        # Junta as células preenchidas de todas as subcolunas do grupo
        celulas: list[tuple] = []   # (row_idx, col)
        respostas: list[str] = []
        for col in grupo["cols"]:
            serie = df[col]
            mask = serie.notna() & ~serie.astype(str).str.strip().isin(["", "-", "nan"])
            for idx in df.index[mask]:
                celulas.append((idx, col))
                respostas.append(str(serie.at[idx]))

        nome_curto = grupo["pergunta"][:60]

        def on_progress(i, total, resposta, categoria,
                        _gi=gi, _nome=nome_curto, _n=len(selecionados)):
            pct = (_gi + (i + 1) / max(total, 1)) / max(_n, 1)
            progress.progress(min(pct, 1.0),
                              text=f"{_nome}: {i + 1}/{total}")

        contexto_grupo = (
            f"{contexto}\n\nPergunta do questionário: {grupo['pergunta']}\n"
            "Crie categorias temáticas curtas (máx. 4 palavras), consistentes "
            "entre respostas semelhantes."
        ).strip()

        resultado = codificador.codificar_lote_modo(
            respostas,
            tipo="livre",
            modo=modo,
            contexto_custom=contexto_grupo,
            callback_progresso=on_progress,
        )
        valores = resultado.get("resultado", [])

        for col in grupo["cols"]:
            if col not in coded:
                coded[col] = pd.Series(pd.NA, index=df.index, dtype=object)
        for (idx, col), valor in zip(celulas, valores):
            if str(valor).strip() and str(valor).strip().upper() != "SEM_RESPOSTA":
                coded[col].at[idx] = str(valor).strip()

        info_grupos.append({**grupo, "modo": modo})

    progress.progress(1.0, text="Codificação concluída.")
    _ss("pipe_coded", coded)
    _ss("pipe_grupos_codificados", info_grupos)
    _ss("pipe_step", 4)
    st.rerun()


# ─────────────────────────────────────────────────────────────────────────────
# ETAPA 4 — REVISÃO DE CATEGORIAS (checkpoint 2)
# ─────────────────────────────────────────────────────────────────────────────

def _freq_categorias(coded: dict, cols: list[str]) -> pd.Series:
    series = [coded[c] for c in cols if c in coded]
    if not series:
        return pd.Series(dtype=int)
    junto = pd.concat(series, ignore_index=True).dropna().astype(str)
    tokens = junto.str.split(_SEP_CAT).explode().str.strip()
    tokens = tokens[tokens.ne("")]
    return tokens.value_counts()


def _etapa4_categorias():
    df = _sg("pipe_df_limpo")
    coded: dict[str, pd.Series] = _sg("pipe_coded", {})
    grupos = _sg("pipe_grupos_codificados", [])

    st.markdown('<div class="ifec-card">', unsafe_allow_html=True)
    st.markdown('<p class="ifec-section-title">Revisão das categorias</p>',
                unsafe_allow_html=True)
    st.markdown(
        '<p class="ifec-section-sub">Renomeie categorias para corrigir ou unir '
        '(dê o mesmo nome para unificar) — depois aprove para gerar a base final</p>',
        unsafe_allow_html=True,
    )

    nomes = [f'{g["num"]} — {g["pergunta"][:80]}' for g in grupos]
    sel = st.selectbox("Pergunta codificada", nomes, key="pipe_cat_sel")
    grupo = grupos[nomes.index(sel)]

    freq = _freq_categorias(coded, grupo["cols"])
    if freq.empty:
        st.info("Nenhuma categoria gerada para esta pergunta.")
    else:
        base_editor = pd.DataFrame({
            "Categoria": freq.index,
            "Respostas": freq.values,
            "Renomear para": freq.index,
        })
        editado = st.data_editor(
            base_editor,
            column_config={
                "Categoria": st.column_config.TextColumn(disabled=True),
                "Respostas": st.column_config.NumberColumn(disabled=True),
                "Renomear para": st.column_config.TextColumn(
                    help="Edite para renomear; use o mesmo nome em duas linhas "
                         "para uni-las"),
            },
            hide_index=True, use_container_width=True,
            key=f"pipe_cat_editor_{grupo['num']}",
        )

        col_l, col_r = st.columns(2)
        with col_l:
            if st.button("Aplicar renomeações desta pergunta",
                         key=f"pipe_cat_apply_{grupo['num']}"):
                mapa = {
                    str(linha["Categoria"]): str(linha["Renomear para"]).strip()
                    for _, linha in editado.iterrows()
                    if str(linha["Renomear para"]).strip()
                    and str(linha["Renomear para"]).strip() != str(linha["Categoria"])
                }
                if mapa:
                    for c in grupo["cols"]:
                        if c in coded:
                            coded[c] = _aplicar_renomeacao(coded[c], mapa)
                    _ss("pipe_coded", coded)
                    st.success(f"{len(mapa)} categoria(s) renomeada(s).")
                    st.rerun()
                else:
                    st.info("Nenhuma renomeação a aplicar.")
        with col_r:
            cat_ver = st.selectbox("Ver respostas da categoria",
                                   list(freq.index), key=f"pipe_cat_ver_{grupo['num']}")

        if cat_ver:
            linhas = []
            for c in grupo["cols"]:
                if c not in coded:
                    continue
                mask = coded[c].astype(str).str.contains(
                    str(cat_ver), case=False, na=False, regex=False)
                for idx in df.index[mask.reindex(df.index, fill_value=False)]:
                    linhas.append({
                        "Resposta original": df.at[idx, c],
                        "Categoria(s)": coded[c].at[idx],
                    })
            st.dataframe(pd.DataFrame(linhas).head(200),
                         use_container_width=True, hide_index=True, height=260)

    st.markdown('</div>', unsafe_allow_html=True)

    col_a, col_b = st.columns(2)
    with col_a:
        if st.button("Voltar e recodificar", use_container_width=True,
                     key="pipe_back4"):
            _ss("pipe_step", 3)
            st.rerun()
    with col_b:
        if st.button("Aprovar categorias e gerar base final", type="primary",
                     use_container_width=True, key="pipe_next4"):
            _ss("pipe_step", 5)
            st.rerun()


# ─────────────────────────────────────────────────────────────────────────────
# ETAPA 5 — BASE FINAL + TABULAÇÃO
# ─────────────────────────────────────────────────────────────────────────────

def _montar_base_final(df: pd.DataFrame, coded: dict[str, pd.Series]) -> pd.DataFrame:
    """Insere cada coluna codificada imediatamente após a coluna original."""
    df_final = df.copy()
    for col, serie in coded.items():
        if col not in df_final.columns:
            continue
        nome_cod = f"{col}_cod"
        df_final[nome_cod] = serie.reindex(df_final.index)
        cols = list(df_final.columns)
        cols.remove(nome_cod)
        cols.insert(cols.index(col) + 1, nome_cod)
        df_final = df_final[cols]
    return df_final


def _etapa5_final():
    df = _sg("pipe_df_limpo")
    coded: dict[str, pd.Series] = _sg("pipe_coded", {})
    relatorio = _sg("pipe_relatorio")
    titulo = _sg("pipe_titulo", "Pesquisa IFec RJ")

    if "pipe_df_final" not in st.session_state:
        with st.spinner("Montando a base final..."):
            _ss("pipe_df_final", _montar_base_final(df, coded))
    df_final = _sg("pipe_df_final")

    st.markdown('<div class="ifec-card">', unsafe_allow_html=True)
    st.markdown('<p class="ifec-section-title">Base final e tabulação</p>',
                unsafe_allow_html=True)

    c1, c2, c3 = st.columns(3)
    c1.metric("Respondentes", len(df_final))
    c2.metric("Colunas (com codificação)", len(df_final.columns))
    c3.metric("Perguntas codificadas", len(_sg("pipe_grupos_codificados", [])))

    st.download_button(
        "Baixar base final (.xlsx)",
        data=_df_para_excel({"Base final": _sanitize_export_df(df_final)}),
        file_name="base_final.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )
    if relatorio is not None and len(relatorio) > 0:
        st.download_button(
            "Baixar relatório de limpeza (.xlsx)",
            data=_df_para_excel({"Relatório de limpeza": relatorio}),
            file_name="relatorio_limpeza.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

    st.divider()
    st.markdown("**Tabulação automática**")

    if st.button("Gerar tabulação (Excel + PowerPoint)", type="primary",
                 use_container_width=True, key="pipe_gerar_tab"):
        from tabulador import detectar_perguntas, exportar_excel

        df_clean = _sanitize_export_df(df_final)
        with st.spinner("Detectando perguntas e gerando a tabulação..."):
            perguntas_tab = detectar_perguntas(
                df_clean, _sg("pipe_tipos_sm", {}), _sg("pipe_q0_map", {})
            )
            # Excel
            with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
                path_xlsx = tmp.name
            try:
                exportar_excel(df_clean, perguntas_tab, saida=path_xlsx,
                               titulo=titulo, total_respostas=len(df_clean))
                _ss("pipe_tab_xlsx", Path(path_xlsx).read_bytes())
            finally:
                Path(path_xlsx).unlink(missing_ok=True)
            # PowerPoint
            try:
                from gerador_ppt import gerar_ppt
                with tempfile.NamedTemporaryFile(suffix=".pptx", delete=False) as tmp:
                    path_ppt = tmp.name
                try:
                    gerar_ppt(df_clean, perguntas_tab, saida=path_ppt,
                              titulo=titulo)
                    _ss("pipe_tab_ppt", Path(path_ppt).read_bytes())
                finally:
                    Path(path_ppt).unlink(missing_ok=True)
            except Exception as exc:
                _ss("pipe_tab_ppt", None)
                st.warning(f"PowerPoint não gerado: {exc}")
        st.rerun()

    if _sg("pipe_tab_xlsx"):
        st.download_button(
            "Baixar tabulação (.xlsx)",
            data=_sg("pipe_tab_xlsx"),
            file_name="tabulacao.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
    if _sg("pipe_tab_ppt"):
        st.download_button(
            "Baixar apresentação (.pptx)",
            data=_sg("pipe_tab_ppt"),
            file_name="tabulacao.pptx",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            use_container_width=True,
        )

    st.markdown('</div>', unsafe_allow_html=True)

    col_a, col_b = st.columns(2)
    with col_a:
        if st.button("Voltar às categorias", use_container_width=True,
                     key="pipe_back5"):
            st.session_state.pop("pipe_df_final", None)
            st.session_state.pop("pipe_tab_xlsx", None)
            st.session_state.pop("pipe_tab_ppt", None)
            _ss("pipe_step", 4)
            st.rerun()
    with col_b:
        if st.button("Iniciar novo processamento", use_container_width=True,
                     key="pipe_reset"):
            _reset_pipeline()
            st.rerun()


# ─────────────────────────────────────────────────────────────────────────────
# ENTRADA
# ─────────────────────────────────────────────────────────────────────────────

def render_pipeline():
    step = _sg("pipe_step", 1)
    _steps(step)
    if step == 1:
        _etapa1_upload()
    elif step == 2:
        _etapa2_limpeza()
    elif step == 3:
        _etapa3_codificacao()
    elif step == 4:
        _etapa4_categorias()
    elif step == 5:
        _etapa5_final()
    else:
        _ss("pipe_step", 1)
        _etapa1_upload()
