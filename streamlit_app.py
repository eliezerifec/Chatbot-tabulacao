import os
import re
from io import BytesIO
from pathlib import Path
import tempfile

import pandas as pd
import streamlit as st


# =============================================================================
# CONFIGURACAO DA PAGINA
# =============================================================================
st.set_page_config(
    page_title="Codificador IFec",
    page_icon="logo_ifec.png",
    layout="wide",
    initial_sidebar_state="expanded",
)


# =============================================================================
# ESTILO VISUAL (paleta IFec + cards neumorphic + sidebar arredondada)
# =============================================================================
st.markdown(
    """
    <style>
    /* ---------- Reset / fundo geral ---------- */
    .stApp { background: #ececec; }
    .block-container {
        padding-top: 1.2rem;
        padding-bottom: 3rem;
        max-width: 1100px;
    }
    header[data-testid="stHeader"] { background: transparent; }

    /* ---------- Sidebar arredondada estilo cartao ---------- */
    [data-testid="stSidebar"] > div:first-child {
        background: #ffffff;
        border-radius: 0 28px 28px 0;
        box-shadow: 4px 4px 14px rgba(0,0,0,0.08);
        padding-top: 1.2rem;
    }
    [data-testid="stSidebar"] * { color: #111827; }
    [data-testid="stSidebar"] .stButton button {
        width: 100%;
        background: transparent;
        color: #111827;
        border: none;
        border-radius: 999px;
        padding: 0.55rem 0.8rem;
        font-weight: 700;
        font-size: 0.78rem;
        letter-spacing: 0.06em;
        text-transform: uppercase;
        text-align: center;
        transition: background 0.15s ease;
    }
    [data-testid="stSidebar"] .stButton button:hover {
        background: #eef2ff;
        color: #1e3a8a;
    }
    [data-testid="stSidebar"] .stButton button:focus { box-shadow: none; }

    /* Botao de menu ATIVO (azul preenchido) */
    [data-testid="stSidebar"] .nav-active .stButton button,
    [data-testid="stSidebar"] .nav-active .stButton button:hover {
        background: #1e3a8a;
        color: #ffffff;
    }

    /* ---------- Cards neumorphic ---------- */
    .ifec-card {
        background: #ececec;
        border-radius: 28px;
        padding: 2rem 2.2rem;
        box-shadow:
            8px 8px 18px rgba(163,177,198,0.55),
            -8px -8px 18px rgba(255,255,255,0.85);
        margin-bottom: 1.4rem;
    }
    .ifec-card-inner {
        background: #ececec;
        border-radius: 22px;
        padding: 1.4rem 1.6rem;
        box-shadow:
            inset 4px 4px 8px rgba(163,177,198,0.45),
            inset -4px -4px 8px rgba(255,255,255,0.85);
        margin-bottom: 1rem;
    }

    /* Titulo de secao centralizado */
    .ifec-section-title {
        text-align: center;
        font-size: 1.25rem;
        font-weight: 800;
        color: #0f172a;
        margin: 0 0 0.3rem 0;
        letter-spacing: 0.01em;
    }
    .ifec-section-sub {
        text-align: center;
        color: #475569;
        font-size: 0.95rem;
        margin: 0 0 1.2rem 0;
    }

    /* Badge "Pergunta X de Y" */
    .ifec-pill {
        display: flex;
        align-items: center;
        justify-content: center;
        background: #ececec;
        border-radius: 999px;
        padding: 0.65rem 1.6rem;
        box-shadow:
            inset 4px 4px 8px rgba(163,177,198,0.45),
            inset -4px -4px 8px rgba(255,255,255,0.85);
        font-weight: 700;
        color: #0f172a;
        margin: 0.4rem 0 1rem 0;
    }

    /* ---------- Inputs em forma de pilula ---------- */
    .stTextInput input,
    .stSelectbox div[data-baseweb="select"] > div,
    .stMultiSelect div[data-baseweb="select"] > div {
        background: #ffffff !important;
        border-radius: 999px !important;
        border: 1px solid #d1d5db !important;
    }
    .stTextArea textarea {
        background: #ffffff !important;
        border-radius: 18px !important;
        border: 1px solid #d1d5db !important;
    }

    /* File uploader em estilo dropzone arredondado */
    [data-testid="stFileUploader"] section {
        background: #ececec;
        border-radius: 22px;
        border: 2px dashed #cbd5e1;
        padding: 2rem 1rem;
        box-shadow:
            inset 4px 4px 8px rgba(163,177,198,0.35),
            inset -4px -4px 8px rgba(255,255,255,0.85);
    }
    [data-testid="stFileUploader"] section:hover {
        border-color: #1e3a8a;
    }

    /* Botoes primarios e download */
    .stButton button[kind="primary"],
    .stDownloadButton button {
        background: #1e3a8a !important;
        color: #ffffff !important;
        border: none !important;
        border-radius: 999px !important;
        padding: 0.55rem 1.6rem !important;
        font-weight: 700 !important;
        letter-spacing: 0.04em;
    }
    .stButton button[kind="primary"]:hover,
    .stDownloadButton button:hover {
        background: #1d4ed8 !important;
        color: #ffffff !important;
    }

    /* Cabecalho da pagina */
    .ifec-page-title {
        font-size: 1.5rem;
        font-weight: 800;
        color: #0f172a;
        margin: 0.2rem 0 0.4rem 0;
    }
    .ifec-page-sub { color: #475569; margin-bottom: 1rem; }

    /* Step indicator (4 bolinhas) */
    .ifec-steps {
        display: flex;
        justify-content: center;
        gap: 0.6rem;
        margin: 0.2rem 0 1.2rem 0;
    }
    .ifec-step {
        width: 0.6rem; height: 0.6rem; border-radius: 999px;
        background: #cbd5e1;
    }
    .ifec-step.active { background: #1e3a8a; width: 1.6rem; }
    .ifec-step.done   { background: #1e3a8a; }
    </style>
    """,
    unsafe_allow_html=True,
)


# =============================================================================
# FUNCOES DE BACKEND (preservadas do app original)
# =============================================================================
def _secret_to_env() -> None:
    if os.getenv("OPENAI_API_KEY"):
        return
    try:
        key = st.secrets.get("OPENAI_API_KEY", "")
    except Exception:
        key = ""
    if key:
        os.environ["OPENAI_API_KEY"] = key


@st.cache_resource(show_spinner=False)
def _get_codificador():
    _secret_to_env()
    from codificador import CodificadorIA
    return CodificadorIA()


@st.cache_data(show_spinner=False)
def _read_uploaded_file(name: str, data: bytes) -> dict[str, pd.DataFrame]:
    bio = BytesIO(data)
    if name.lower().endswith(".csv"):
        return {"Planilha": pd.read_csv(bio)}
    xl = pd.ExcelFile(bio)
    return {sheet: xl.parse(sheet) for sheet in xl.sheet_names}


def _to_excel(sheets: dict[str, pd.DataFrame]) -> bytes:
    out = BytesIO()
    with pd.ExcelWriter(out, engine="xlsxwriter") as writer:
        for sheet_name, df in sheets.items():
            safe_name = sheet_name[:31] or "Planilha"
            df.to_excel(writer, sheet_name=safe_name, index=False)
    return out.getvalue()


@st.cache_data(show_spinner=False)
def _read_tabulation_file(
    name: str, data: bytes, two_line_header: bool
) -> tuple[pd.DataFrame, dict, dict]:
    """
    Retorna (df, tipos_sm, q0_map).
    tipos_sm : {coluna → "RU"|"RM"|"ABERTA"} — extraído da linha 1 do SM.
    q0_map   : {coluna → texto_linha0}        — chave de agrupamento de RM.
    Ambos vazios quando two_line_header=False ou CSV.
    """
    bio = BytesIO(data)
    if name.lower().endswith(".csv"):
        return pd.read_csv(bio), {}, {}
    if two_line_header:
        from tabulador import set_header
        raw = pd.read_excel(bio, header=None, sheet_name=0)
        df, tipos_sm, q0_map = set_header(raw)   # 3-tupla
        return df, tipos_sm, q0_map
    return pd.read_excel(bio, sheet_name=0), {}, {}


def _sanitize_export_df(df: pd.DataFrame) -> pd.DataFrame:
    clean = df.copy()
    if isinstance(clean.columns, pd.MultiIndex):
        clean.columns = ["_".join(str(c) for c in col).strip("_") for col in clean.columns]
    seen: dict[str, int] = {}
    cols = []
    for col in clean.columns:
        name = str(col)
        if name in seen:
            seen[name] += 1
            cols.append(f"{name}_{seen[name]}")
        else:
            seen[name] = 0
            cols.append(name)
    clean.columns = cols
    for col in clean.columns:
        if isinstance(clean[col], pd.DataFrame):
            clean[col] = clean[col].iloc[:, 0]
        try:
            clean[col] = clean[col].astype(object)
        except Exception:
            pass
    return clean


def _build_tab_excel(df: pd.DataFrame, perguntas: list[dict], titulo: str,
                     aberturas_cols: list[str] | None = None,
                     filtro_col: str | None = None) -> bytes:
    from tabulador import exportar_excel
    with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
        path = tmp.name
    try:
        exportar_excel(
            _sanitize_export_df(df),
            perguntas,
            saida=path,
            titulo=titulo or "Pesquisa IFec RJ",
            total_respostas=len(df),
            aberturas_cols=aberturas_cols or None,
            filtro_col=filtro_col or None,
        )
        return Path(path).read_bytes()
    finally:
        try:
            Path(path).unlink(missing_ok=True)
        except Exception:
            pass


def _candidatas_abertura(df: pd.DataFrame) -> list[str]:
    """
    Colunas candidatas a abertura / filtro:
    - 2 a 60 valores únicos (aceita localizações com muitas cidades)
    - Não muito esparsas (NaN < 90 % das linhas)
    - Texto curto (mediana ≤ 80 chars) — exclui respostas abertas longas
    - Não é coluna de sistema
    """
    try:
        from tabulador import _IGNORAR as _TAB_IGN
    except ImportError:
        _TAB_IGN = set()

    result = []
    n_rows = max(len(df), 1)
    for col in df.columns:
        if col in _TAB_IGN or col.endswith("_cod"):
            continue
        serie = df[col].dropna().astype(str)
        if len(serie) / n_rows < 0.10:      # > 90 % NaN → skip
            continue
        n_unique = serie.nunique()
        if not (2 <= n_unique <= 60):
            continue
        median_len = serie.str.len().median()
        if median_len > 80:                 # respostas abertas longas → skip
            continue
        result.append(col)
    return result


def _build_tab_ppt(df: pd.DataFrame, perguntas: list[dict], titulo: str,
                   subtitulo: str, periodo: str) -> bytes:
    from gerador_ppt import gerar_ppt
    with tempfile.NamedTemporaryFile(suffix=".pptx", delete=False) as tmp:
        path = tmp.name
    try:
        gerar_ppt(
            _sanitize_export_df(df),
            perguntas,
            saida=path,
            titulo=titulo or "Pesquisa IFec RJ",
            subtitulo=subtitulo or "",
            periodo=periodo or "",
        )
        return Path(path).read_bytes()
    finally:
        try:
            Path(path).unlink(missing_ok=True)
        except Exception:
            pass


def _parse_list(raw: str) -> list[str]:
    items = []
    for part in raw.replace("\n", ",").split(","):
        val = part.strip()
        if val and val not in items:
            items.append(val)
    return items


def _load_taxonomies():
    _secret_to_env()
    from codificador import MODOS_RESPOSTA, TIPOS_PERGUNTA
    return TIPOS_PERGUNTA, MODOS_RESPOSTA


def _run_coding(sheets, configs, global_context, previous_categories):
    codificador = _get_codificador()
    result_sheets = {name: df.copy() for name, df in sheets.items()}
    selected = [name for name, cfg in configs.items() if cfg["selected"]]
    progress = st.progress(0, text="Preparando codificacao...")

    for sheet_idx, sheet_name in enumerate(selected):
        cfg = configs[sheet_name]
        df = result_sheets[sheet_name]
        respostas = df[cfg["input_col"]].astype(str).tolist()
        cats_prev = previous_categories.get(sheet_name, [])
        cats_imp = cats_prev or cfg["categories"]
        context = cfg["context"] if cfg["type_key"] == "livre" else global_context

        def on_progress(i_local, total_local, resposta, categoria):
            pct = (sheet_idx + ((i_local + 1) / max(total_local, 1))) / max(len(selected), 1)
            progress.progress(
                min(pct, 1.0),
                text=f"{sheet_name}: {i_local + 1}/{total_local}",
            )

        coded = codificador.codificar_lote_modo(
            respostas,
            tipo=cfg["type_key"],
            modo=cfg["mode_key"],
            contexto_custom=context,
            categorias_imputacao=cats_imp,
            categorias_anteriores=cats_prev,
            callback_progresso=on_progress,
        )

        if "imputado" in coded:
            df[cfg["imputed_col"]] = coded["imputado"]
            df[cfg["new_col"]] = coded["novo"]
        else:
            df[cfg["output_col"]] = coded["resultado"]
        result_sheets[sheet_name] = df

    progress.progress(1.0, text="Codificacao concluida.")
    return result_sheets


# =============================================================================
# HELPERS DE UI
# =============================================================================
def _step_indicator(current: int, total: int = 4) -> None:
    """Bolinhas indicadoras de etapa do wizard."""
    bolas = []
    for i in range(1, total + 1):
        if i < current:
            bolas.append('<span class="ifec-step done"></span>')
        elif i == current:
            bolas.append('<span class="ifec-step active"></span>')
        else:
            bolas.append('<span class="ifec-step"></span>')
    st.markdown(f'<div class="ifec-steps">{"".join(bolas)}</div>', unsafe_allow_html=True)


def _ss_get(key, default=None):
    return st.session_state.get(key, default)


def _ss_set(key, value):
    st.session_state[key] = value


# =============================================================================
# MODO CODIFICADOR (wizard de 4 passos)
# =============================================================================
def _render_codificador(api_ok: bool) -> None:
    if not api_ok:
        st.warning("Configure OPENAI_API_KEY em Secrets para habilitar a codificacao.")

    step = _ss_get("cod_step", 1)

    # Etapa 4: resultado final ja existe -> mostra direto
    if step == 4 and "result_sheets" in st.session_state:
        _step_indicator(4)
        _render_resultado_codificacao()
        return

    _step_indicator(step)

    if step == 1:
        _step1_upload()
    elif step == 2:
        _step2_perguntas()
    elif step == 3:
        _step3_executar()
    else:
        _ss_set("cod_step", 1)
        _step1_upload()


# ---- Etapa 1 ---------------------------------------------------------------
def _step1_upload() -> None:
    st.markdown('<div class="ifec-card">', unsafe_allow_html=True)
    st.markdown(
        '<p class="ifec-section-sub">Importe a base que sera codificada</p>',
        unsafe_allow_html=True,
    )

    uploaded = st.file_uploader(
        "Arraste e solte ou clique para importar a base a ser codificada",
        type=["xlsx", "csv"],
        help="O arquivo deve ser .xlsx ou .csv",
        key="cod_upload_main",
    )

    st.markdown(
        '<p style="text-align:center; color:#475569; font-size:0.86rem; margin-top:0.4rem;">'
        'O arquivo deve ser xlsx</p>',
        unsafe_allow_html=True,
    )

    st.divider()

    recorrente = st.checkbox(
        "Pesquisa recorrente? Deseja inserir categorias de pesquisas anteriores?",
        value=_ss_get("cod_recorrente", False),
        key="cod_recorrente",
    )

    previous_file = None
    if recorrente:
        previous_file = st.file_uploader(
            "Arraste e solte ou clique para importar o dicionario da base anterior",
            type=["xlsx", "csv"],
            help="O arquivo deve ser .xlsx ou .csv",
            key="cod_upload_prev",
        )

    st.markdown('</div>', unsafe_allow_html=True)

    # Salva arquivos no session_state
    if uploaded is not None:
        _ss_set("cod_uploaded_name", uploaded.name)
        _ss_set("cod_uploaded_bytes", uploaded.getvalue())
    if previous_file is not None:
        _ss_set("cod_prev_name", previous_file.name)
        _ss_set("cod_prev_bytes", previous_file.getvalue())
    elif not recorrente:
        st.session_state.pop("cod_prev_name", None)
        st.session_state.pop("cod_prev_bytes", None)
        st.session_state.pop("cod_prev_categories", None)

    # Acoes
    _, col_right = st.columns([1, 1])
    with col_right:
        pode_avancar = "cod_uploaded_bytes" in st.session_state
        if st.button("Avancar", type="primary", disabled=not pode_avancar,
                     use_container_width=True, key="cod_next_1"):
            _ss_set("cod_step", 2)
            _ss_set("cod_q_idx", 0)
            st.rerun()


# ---- Etapa 2 ---------------------------------------------------------------
def _step2_perguntas() -> None:
    # carrega base principal
    try:
        sheets = _read_uploaded_file(
            _ss_get("cod_uploaded_name"),
            _ss_get("cod_uploaded_bytes"),
        )
    except Exception as exc:
        st.error(f"Nao foi possivel ler o arquivo: {exc}")
        if st.button("Voltar", key="back_err"):
            _ss_set("cod_step", 1)
            st.rerun()
        return

    # carrega base anterior (se houver)
    previous_categories: dict[str, list[str]] = {}
    if _ss_get("cod_prev_bytes") is not None:
        with st.expander("Pesquisa anterior — confirme as colunas de categoria", expanded=False):
            prev_sheets = _read_uploaded_file(
                _ss_get("cod_prev_name"),
                _ss_get("cod_prev_bytes"),
            )
            for prev_name, prev_df in prev_sheets.items():
                if prev_df.empty:
                    continue
                default_col = prev_df.columns[-1]
                category_col = st.selectbox(
                    f"Coluna de categorias em '{prev_name}'",
                    list(prev_df.columns),
                    index=list(prev_df.columns).index(default_col),
                    key=f"prev_col_{prev_name}",
                )
                cats = (
                    prev_df[category_col].dropna().astype(str).map(str.strip)
                    .loc[lambda s: s.ne("")].drop_duplicates().tolist()
                )
                previous_categories[prev_name] = cats
        _ss_set("cod_prev_categories", previous_categories)
    else:
        previous_categories = _ss_get("cod_prev_categories", {}) or {}

    tipos_pergunta, modos_resposta = _load_taxonomies()
    type_labels = {data["label"]: key for key, data in tipos_pergunta.items()}
    mode_labels = {data["label"]: key for key, data in modos_resposta.items()}

    sheet_names = list(sheets.keys())
    total_q = len(sheet_names)
    q_idx = max(0, min(_ss_get("cod_q_idx", 0), total_q - 1))
    sheet_name = sheet_names[q_idx]
    df = sheets[sheet_name]

    # ---- Card principal ----
    st.markdown('<div class="ifec-card">', unsafe_allow_html=True)
    st.markdown('<p class="ifec-section-title">Perguntas</p>', unsafe_allow_html=True)
    st.markdown(
        '<p class="ifec-section-sub">Selecione as colunas, o tipo de pergunta '
        'e a biblioteca da base anterior</p>',
        unsafe_allow_html=True,
    )

    # Navegacao Pergunta X de Y
    nav_l, nav_c, nav_r = st.columns([1, 6, 1])
    with nav_l:
        if st.button("‹", key="q_prev", disabled=q_idx == 0,
                     use_container_width=True):
            _ss_set("cod_q_idx", q_idx - 1)
            st.rerun()
    with nav_c:
        st.markdown(
            f'<div class="ifec-pill">Pergunta {q_idx + 1} de {total_q} '
            f'&nbsp;&nbsp;<span style="color:#64748b; font-weight:500;">'
            f'({sheet_name})</span></div>',
            unsafe_allow_html=True,
        )
    with nav_r:
        if st.button("›", key="q_next", disabled=q_idx >= total_q - 1,
                     use_container_width=True):
            _ss_set("cod_q_idx", q_idx + 1)
            st.rerun()

    # Recupera config previa desta aba
    cfg_prev = _ss_get(f"cod_cfg_{sheet_name}", {}) or {}

    cols = list(df.columns)

    selected = st.checkbox(
        "Codificar esta aba",
        value=cfg_prev.get("selected", True),
        key=f"sel_{sheet_name}",
    )

    input_col = st.selectbox(
        "Coluna de Resposta",
        cols,
        index=cols.index(cfg_prev["input_col"]) if cfg_prev.get("input_col") in cols else 0,
        key=f"in_{sheet_name}",
    )

    output_col = st.text_input(
        "Coluna de codificacao",
        value=cfg_prev.get("output_col", "Categorizado"),
        key=f"out_{sheet_name}",
    )

    type_options = list(type_labels.keys())
    default_type_label = next(k for k, v in type_labels.items() if v == "livre")
    type_label = st.selectbox(
        "Tipo de pergunta",
        type_options,
        index=type_options.index(cfg_prev.get("type_label", default_type_label))
              if cfg_prev.get("type_label", default_type_label) in type_options
              else type_options.index(default_type_label),
        key=f"type_{sheet_name}",
    )

    # Aba base anterior (se houver pesquisa anterior carregada)
    aba_anterior = ""
    if previous_categories:
        opcoes_aba_ant = ["(nenhuma)"] + list(previous_categories.keys())
        aba_pref = cfg_prev.get("aba_anterior", "(nenhuma)")
        idx_default = opcoes_aba_ant.index(aba_pref) if aba_pref in opcoes_aba_ant else 0
        aba_anterior = st.selectbox(
            "Aba base anterior",
            opcoes_aba_ant,
            index=idx_default,
            key=f"aba_ant_{sheet_name}",
        )

    mode_options = list(mode_labels.keys())
    mode_pref = cfg_prev.get("mode_label", mode_options[0])
    mode_label = st.selectbox(
        "Modo de resposta",
        mode_options,
        index=mode_options.index(mode_pref) if mode_pref in mode_options else 0,
        key=f"mode_{sheet_name}",
    )
    mode_key = mode_labels[mode_label]

    # Campos extras
    imputed_col = cfg_prev.get("imputed_col", "col_imputado")
    new_col = cfg_prev.get("new_col", "col_nova")
    categories: list[str] = cfg_prev.get("categories", [])
    if "semi" in mode_key:
        sem_a, sem_b = st.columns(2)
        with sem_a:
            imputed_col = st.text_input(
                "Coluna de imputacao",
                value=imputed_col,
                key=f"imp_{sheet_name}",
            )
        with sem_b:
            new_col = st.text_input(
                "Coluna de novas categorias",
                value=new_col,
                key=f"new_{sheet_name}",
            )
        categories = _parse_list(
            st.text_area(
                "Categorias pre-definidas",
                value=", ".join(categories),
                placeholder="Uma categoria por linha ou separadas por virgula.",
                key=f"cats_{sheet_name}",
            )
        )

    custom_context = st.text_area(
        "Contexto especifico desta pergunta",
        value=cfg_prev.get("context", ""),
        placeholder="Use quando o tipo da pergunta for Personalizado.",
        key=f"ctx_{sheet_name}",
        height=80,
    )

    st.markdown('</div>', unsafe_allow_html=True)

    # Configuracoes auxiliares fora do carrossel
    with st.expander("Contexto geral da pesquisa", expanded=False):
        global_context = st.text_area(
            "Descreva o objetivo da pesquisa e os criterios de classificacao",
            value=_ss_get("cod_global_context", ""),
            height=100,
            key="cod_global_context_input",
        )
        _ss_set("cod_global_context", global_context)

    with st.expander("Previa dos dados desta aba", expanded=False):
        st.dataframe(df.head(20), use_container_width=True, hide_index=True)

    # Salva config desta aba
    cfg_atual = {
        "selected": selected,
        "input_col": input_col,
        "output_col": output_col.strip() or "Categorizado",
        "type_label": type_label,
        "type_key": type_labels[type_label],
        "mode_label": mode_label,
        "mode_key": mode_key,
        "categories": categories,
        "imputed_col": imputed_col.strip() or "col_imputado",
        "new_col": new_col.strip() or "col_nova",
        "context": custom_context.strip(),
        "aba_anterior": aba_anterior,
    }
    _ss_set(f"cod_cfg_{sheet_name}", cfg_atual)

    # Acoes da etapa
    bt1, _, bt3 = st.columns([1, 1, 1])
    with bt1:
        if st.button("Voltar", key="cod_back_2", use_container_width=True):
            _ss_set("cod_step", 1)
            st.rerun()
    with bt3:
        todas_configuradas = all(
            _ss_get(f"cod_cfg_{n}") is not None for n in sheet_names
        )
        if st.button("Iniciar codificacao", type="primary",
                     disabled=not todas_configuradas, use_container_width=True,
                     key="cod_next_2"):
            _ss_set("cod_step", 3)
            st.rerun()


# ---- Etapa 3 ---------------------------------------------------------------
def _step3_executar() -> None:
    """Executa a codificacao com spinner — sem log ao vivo."""

    # Guarda: se o resultado ja existe, avanca direto para etapa 4
    # sem recodificar (evita gastar tokens duplos ao clicar em "Ver resultado")
    if "result_sheets" in st.session_state:
        _ss_set("cod_step", 4)
        st.rerun()
        return

    st.markdown('<div class="ifec-card">', unsafe_allow_html=True)
    st.markdown(
        '<p class="ifec-section-title">Codificacao em andamento</p>',
        unsafe_allow_html=True,
    )
    st.markdown(
        '<p class="ifec-section-sub">Aguarde — a IA esta categorizando todas as respostas.</p>',
        unsafe_allow_html=True,
    )

    try:
        sheets = _read_uploaded_file(
            _ss_get("cod_uploaded_name"),
            _ss_get("cod_uploaded_bytes"),
        )
    except Exception as exc:
        st.error(f"Erro ao reler o arquivo: {exc}")
        st.markdown('</div>', unsafe_allow_html=True)
        return

    configs = {}
    previous_categories = _ss_get("cod_prev_categories", {}) or {}
    previous_for_run: dict[str, list[str]] = {}

    for sheet_name in sheets.keys():
        cfg = _ss_get(f"cod_cfg_{sheet_name}")
        if not cfg:
            continue
        configs[sheet_name] = {
            "selected": cfg["selected"],
            "input_col": cfg["input_col"],
            "output_col": cfg["output_col"],
            "type_key": cfg["type_key"],
            "mode_key": cfg["mode_key"],
            "categories": cfg["categories"],
            "imputed_col": cfg["imputed_col"],
            "new_col": cfg["new_col"],
            "context": cfg["context"],
        }
        aba_ant = cfg.get("aba_anterior", "")
        if aba_ant and aba_ant != "(nenhuma)" and aba_ant in previous_categories:
            previous_for_run[sheet_name] = previous_categories[aba_ant]

    global_context = _ss_get("cod_global_context", "")

    with st.spinner("Codificando respostas..."):
        try:
            result_sheets = _run_coding(
                sheets, configs, global_context.strip(), previous_for_run
            )
        except Exception as exc:
            st.error(f"Erro durante a codificacao: {exc}")
            st.markdown('</div>', unsafe_allow_html=True)
            if st.button("Voltar para configuracao", key="back_to_cfg"):
                _ss_set("cod_step", 2)
                st.rerun()
            return

    # Salva resultado e avanca automaticamente para etapa 4
    # (sem botao intermediario, que causava reexecucao da codificacao)
    _ss_set("result_sheets", result_sheets)
    _ss_set("cod_step", 4)
    st.markdown('</div>', unsafe_allow_html=True)
    st.rerun()


# ---- Etapa 4 ---------------------------------------------------------------
def _render_resultado_codificacao() -> None:
    result_sheets: dict[str, pd.DataFrame] = _ss_get("result_sheets", {})
    if not result_sheets:
        st.info("Nenhum resultado disponivel ainda.")
        return

    st.markdown('<div class="ifec-card">', unsafe_allow_html=True)
    st.markdown(
        '<p class="ifec-section-title">Frequencia de categorias</p>',
        unsafe_allow_html=True,
    )

    preview_sheet = st.selectbox(
        "Aba para analisar",
        list(result_sheets.keys()),
        key="preview_sheet",
    )
    df_res = result_sheets[preview_sheet]
    cfg = _ss_get(f"cod_cfg_{preview_sheet}", {}) or {}
    out_col = cfg.get("output_col", "Categorizado")
    if out_col not in df_res.columns:
        for c in [cfg.get("imputed_col"), cfg.get("new_col")]:
            if c and c in df_res.columns:
                out_col = c
                break

    # Frequencia (com explode pra resp multi-categoria separadas por ;)
    freq = pd.Series(dtype=int)
    if out_col in df_res.columns:
        serie = df_res[out_col].dropna().astype(str).map(str.strip)
        serie = serie.str.split(";").explode().map(str.strip)
        serie = serie[serie.ne("")]
        freq = serie.value_counts()

    # Layout: dois quadros lado a lado
    left, right = st.columns(2)
    with left:
        st.markdown('<div class="ifec-card-inner">', unsafe_allow_html=True)
        st.markdown("**Categorias criadas**")
        if not freq.empty:
            cat_df = freq.reset_index()
            cat_df.columns = ["Categoria", "Respostas"]
            cat_selecionada = st.radio(
                "Selecione uma categoria para ver as respostas",
                cat_df["Categoria"].tolist(),
                key=f"cat_sel_{preview_sheet}",
                label_visibility="collapsed",
            )
            st.dataframe(
                cat_df, use_container_width=True, hide_index=True, height=320
            )
        else:
            cat_selecionada = None
            st.caption("Nenhuma categoria detectada na coluna de saida.")
        st.markdown('</div>', unsafe_allow_html=True)

    with right:
        st.markdown('<div class="ifec-card-inner">', unsafe_allow_html=True)
        st.markdown("**Respostas para a categoria selecionada**")
        if cat_selecionada and out_col in df_res.columns:
            input_col = cfg.get("input_col")
            cols_show = [c for c in [input_col, out_col] if c and c in df_res.columns]
            mascara = (
                df_res[out_col].astype(str).str.contains(
                    str(cat_selecionada), case=False, na=False, regex=False
                )
            )
            st.dataframe(
                df_res.loc[mascara, cols_show].head(200),
                use_container_width=True,
                hide_index=True,
                height=320,
            )
        else:
            st.caption("Selecione uma categoria a esquerda.")
        st.markdown('</div>', unsafe_allow_html=True)

    # Download centralizado
    st.markdown(
        '<div style="display:flex; justify-content:center; margin-top:1rem;">',
        unsafe_allow_html=True,
    )
    st.download_button(
        "Baixar planilha codificada",
        data=_to_excel(result_sheets),
        file_name="base_codificada.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
    st.markdown('</div>', unsafe_allow_html=True)
    st.markdown('</div>', unsafe_allow_html=True)

    # Acoes secundarias
    a, b = st.columns(2)
    with a:
        if st.button("Reconfigurar perguntas", use_container_width=True,
                     key="reconfig_btn"):
            _ss_set("cod_step", 2)
            st.rerun()
    with b:
        if st.button("Iniciar nova codificacao", use_container_width=True,
                     key="reset_btn"):
            for k in list(st.session_state.keys()):
                if k.startswith("cod_") or k == "result_sheets":
                    st.session_state.pop(k, None)
            _ss_set("cod_step", 1)
            st.rerun()


# =============================================================================
# MODO TABULADOR (mesmo visual neumorphic)
# =============================================================================
def _render_tabulador() -> None:
    st.markdown('<div class="ifec-card">', unsafe_allow_html=True)
    st.markdown(
        '<p class="ifec-section-title">Tabulacao automatica</p>',
        unsafe_allow_html=True,
    )
    st.markdown(
        '<p class="ifec-section-sub">Detecta perguntas, gera tabulacao em Excel e PowerPoint</p>',
        unsafe_allow_html=True,
    )

    # Limpa cache quando o usuário quiser (útil após atualizações de código)
    if st.button("🔄 Limpar cache e reiniciar leitura", key="tab_clear_cache"):
        st.cache_data.clear()
        for k in list(st.session_state.keys()):
            if k.startswith("tab_"):
                st.session_state.pop(k, None)
        st.rerun()

    fonte_options = ["Arquivo enviado"]
    if "result_sheets" in st.session_state:
        fonte_options.append("Resultado codificado")
    fonte = st.radio("Fonte da base", fonte_options, horizontal=True)

    df_tab: pd.DataFrame | None = None
    source_name = ""

    if fonte == "Resultado codificado":
        sheet_name = st.selectbox(
            "Aba codificada para tabular",
            list(st.session_state["result_sheets"].keys()),
            key="tab_result_sheet",
        )
        df_tab = st.session_state["result_sheets"][sheet_name].copy()
        source_name = sheet_name
    else:
        upload = st.file_uploader(
            "Arraste e solte ou clique para importar a base",
            type=["xlsx", "csv"],
            key="tab_upload",
        )
        if upload is None:
            st.markdown('</div>', unsafe_allow_html=True)
            st.info("Envie uma planilha para iniciar a tabulacao.")
            return
        two_line_header = st.checkbox(
            "Cabecalho em duas linhas (SurveyMonkey/TabIFec)",
            value=True,
            help="Combina as duas primeiras linhas como nomes de colunas.",
        )
        try:
            df_tab, tipos_sm, q0_map = _read_tabulation_file(
                upload.name, upload.getvalue(), two_line_header
            )
            st.session_state["tab_tipos_sm"] = tipos_sm
            st.session_state["tab_q0_map"]   = q0_map
        except Exception as exc:
            st.error(f"Nao foi possivel preparar a base: {exc}")
            st.markdown('</div>', unsafe_allow_html=True)
            return
        source_name = Path(upload.name).stem

    c1, c2, c3 = st.columns(3)
    c1.metric("Base", source_name)
    c2.metric("Respondentes", len(df_tab))
    c3.metric("Colunas", len(df_tab.columns))

    if st.button("Detectar perguntas", type="primary", key="tab_detect"):
        from tabulador import detectar_perguntas
        try:
            tipos_sm = st.session_state.get("tab_tipos_sm", {})
            q0_map   = st.session_state.get("tab_q0_map",   {})
            qs = detectar_perguntas(df_tab, tipos_sm, q0_map)
            st.session_state["tab_questions"]      = qs
            st.session_state["tab_question_order"] = list(range(len(qs)))
            st.session_state["tab_source_name"]    = source_name
            st.session_state.pop("tab_excel_bytes", None)
        except Exception as exc:
            st.error(f"Erro ao detectar perguntas: {exc}")

    perguntas_detectadas = st.session_state.get("tab_questions", [])
    if not perguntas_detectadas:
        with st.expander("Previa da base", expanded=True):
            st.dataframe(df_tab.head(30), use_container_width=True, hide_index=True)
        st.markdown('</div>', unsafe_allow_html=True)
        return

    n_det = len(perguntas_detectadas)
    st.success(f"{n_det} pergunta(s) detectada(s).")

    # ── Aberturas e Filtro ────────────────────────────────────────────────────
    candidatas_ab = _candidatas_abertura(df_tab)
    with st.expander("⚙️ Cruzamentos e filtros por aba", expanded=bool(candidatas_ab)):
        col_ab, col_fil = st.columns(2)
        with col_ab:
            ab_sel = st.multiselect(
                "Cruzar por (aberturas):",
                options=candidatas_ab,
                default=[],
                help="Gera colunas de Total + % para cada valor das colunas selecionadas.",
                key="tab_aberturas",
            )
        with col_fil:
            filtro_options = ["(nenhum)"] + candidatas_ab
            filtro_sel = st.selectbox(
                "Gerar aba por (filtro):",
                options=filtro_options,
                index=0,
                help="Cria uma aba separada no Excel para cada valor único da coluna selecionada.",
                key="tab_filtro",
            )

    titulo = st.text_input("Titulo do relatorio", value="Pesquisa IFec RJ", key="tab_title")
    sub_a, sub_b = st.columns(2)
    with sub_a:
        subtitulo = st.text_input("Subtitulo do PowerPoint", value="", key="tab_subtitle")
    with sub_b:
        periodo = st.text_input("Periodo", value="", key="tab_period")

    st.markdown('</div>', unsafe_allow_html=True)

    from tabulador import TIPOS_LABEL, tabular_pergunta
    tipo_keys   = list(TIPOS_LABEL.keys())
    tipo_labels = [f"{key} - {TIPOS_LABEL[key]}" for key in tipo_keys]

    # ── Ordem atual das perguntas (suporta reordenação) ───────────────────────
    order: list[int] = st.session_state.get(
        "tab_question_order", list(range(n_det))
    )
    # Garante consistência se o número de perguntas mudou
    if len(order) != n_det or set(order) != set(range(n_det)):
        order = list(range(n_det))
        st.session_state["tab_question_order"] = order

    perguntas_config: list[dict] = []

    for pos, orig_idx in enumerate(order):
        pergunta = perguntas_detectadas[orig_idx]
        pnum     = f"P{pos + 1:02d}"
        label    = f"{pnum} - {pergunta.get('pergunta', '')[:80]}"

        # ── Botões ↑ ↓ + expander lado a lado ───────────────────────────────
        col_btn, col_exp = st.columns([0.05, 0.95])
        with col_btn:
            st.write("")  # espaçador vertical
            if st.button("↑", key=f"tab_up_{orig_idx}",
                         help="Mover para cima",
                         disabled=(pos == 0)):
                order[pos], order[pos - 1] = order[pos - 1], order[pos]
                st.session_state["tab_question_order"] = order
                st.rerun()
            if st.button("↓", key=f"tab_dn_{orig_idx}",
                         help="Mover para baixo",
                         disabled=(pos == n_det - 1)):
                order[pos], order[pos + 1] = order[pos + 1], order[pos]
                st.session_state["tab_question_order"] = order
                st.rerun()

        with col_exp:
            with st.expander(label, expanded=(pos < 2)):
                top_a, top_b, top_c, top_d = st.columns([1, 1, 2, 3])
                ativo = top_a.checkbox(
                    "Ativa", value=pergunta.get("ativo", True),
                    key=f"tab_active_{orig_idx}",
                )
                mostrar_outro = top_b.checkbox(
                    "Mostrar Outro",
                    value=pergunta.get("mostrar_outro", True),
                    help="Quando marcado, exibe as codificações do campo 'Outro' acima da linha Outro.",
                    key=f"tab_outro_{orig_idx}",
                )
                tipo_atual = pergunta.get("tipo", "ABERTA")
                tipo_index = tipo_keys.index(tipo_atual) if tipo_atual in tipo_keys \
                             else tipo_keys.index("ABERTA")
                tipo_label = top_c.selectbox(
                    "Tipo", tipo_labels, index=tipo_index,
                    key=f"tab_type_{orig_idx}",
                )
                texto = top_d.text_input(
                    "Pergunta",
                    value=str(pergunta.get("pergunta", "")),
                    key=f"tab_question_{orig_idx}",
                )
                nota = st.text_area(
                    "Nota",
                    value=str(pergunta.get("nota", "")),
                    height=70,
                    key=f"tab_note_{orig_idx}",
                )
                ordem_raw = st.text_input(
                    "Ordem das opções (sep. por ; — vazio = automático)",
                    value="; ".join(pergunta.get("ordem") or []),
                    help="Ex: Sim; Não; Talvez",
                    key=f"tab_ordem_{orig_idx}",
                )
                ordem_lista = [o.strip() for o in re.split(r"[;\n]", ordem_raw) if o.strip()]

                # ── Coluna de codificações (Outro) ───────────────────────────
                # Mostra col_cod detectada e permite troca manual
                auto_cod = pergunta.get("col_cod") or ""
                colunas_cod = ["(nenhuma)"] + [
                    c for c in df_tab.columns
                    if c.endswith("_cod") or (
                        auto_cod and c == auto_cod
                    )
                ]
                # Remove duplicatas mantendo ordem
                seen_cod: set = set()
                colunas_cod_uniq = []
                for x in colunas_cod:
                    if x not in seen_cod:
                        seen_cod.add(x)
                        colunas_cod_uniq.append(x)

                default_cod_idx = (
                    colunas_cod_uniq.index(auto_cod)
                    if auto_cod and auto_cod in colunas_cod_uniq
                    else 0
                )
                col_cod_sel = st.selectbox(
                    "Coluna de codificações do Outro (col_cod)",
                    options=colunas_cod_uniq,
                    index=default_cod_idx,
                    help="Selecione a coluna que contém as categorias codificadas do campo 'Outro'. "
                         "Detectada automaticamente — troque se necessário.",
                    key=f"tab_cod_{orig_idx}",
                )
                col_cod_final = col_cod_sel if col_cod_sel != "(nenhuma)" else None

                if not auto_cod and not col_cod_final:
                    st.caption("⚠️ Coluna _cod não detectada — sub-itens do Outro não serão exibidos.")
                elif col_cod_final:
                    st.caption(f"✓ col_cod: `{col_cod_final}`")

                cfg = dict(pergunta)
                cfg["num"]          = pnum
                cfg["ativo"]        = ativo
                cfg["mostrar_outro"]= mostrar_outro
                cfg["col_cod"]      = col_cod_final      # substitui o valor auto-detectado
                cfg["tipo"]         = tipo_keys[tipo_labels.index(tipo_label)]
                cfg["pergunta"]     = texto.strip() or pergunta.get("pergunta", "")
                cfg["nota"]         = nota.strip()
                cfg["ordem"]        = ordem_lista
                perguntas_config.append(cfg)

                colunas = cfg.get("colunas", [])
                st.caption(
                    "Colunas: " + ", ".join(str(c) for c in colunas[:6])
                    + ("..." if len(colunas) > 6 else "")
                )
                if st.checkbox("Prévia da tabulação", key=f"tab_preview_{orig_idx}"):
                    try:
                        st.dataframe(
                            tabular_pergunta(df_tab, cfg),
                            use_container_width=True, hide_index=True,
                        )
                    except Exception as exc:
                        st.warning(f"Nao foi possivel tabular esta pergunta: {exc}")

    ativas = [p for p in perguntas_config if p.get("ativo") and p.get("tipo") != "IGNORAR"]

    st.markdown('<div class="ifec-card">', unsafe_allow_html=True)
    st.write(f"**{len(ativas)} pergunta(s) ativa(s) para exportacao.**")

    gen_excel, gen_ppt = st.columns(2)
    with gen_excel:
        if st.button("Gerar Excel de tabulacao", disabled=not ativas,
                     use_container_width=True, key="tab_gen_xlsx"):
            with st.spinner("Gerando Excel..."):
                try:
                    _filtro = filtro_sel if filtro_sel != "(nenhum)" else None
                    st.session_state["tab_excel_bytes"] = _build_tab_excel(
                        df_tab, ativas, titulo,
                        aberturas_cols=ab_sel or None,
                        filtro_col=_filtro,
                    )
                except Exception as exc:
                    st.error(f"Erro ao gerar Excel: {exc}")
        if "tab_excel_bytes" in st.session_state:
            st.download_button(
                "Baixar Excel",
                data=st.session_state["tab_excel_bytes"],
                file_name=f"Tabulacao_{source_name}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )

    with gen_ppt:
        if st.button("Gerar PowerPoint", disabled=not ativas,
                     use_container_width=True, key="tab_gen_ppt"):
            with st.spinner("Gerando PowerPoint..."):
                try:
                    st.session_state["tab_ppt_bytes"] = _build_tab_ppt(
                        df_tab, ativas, titulo, subtitulo, periodo
                    )
                except Exception as exc:
                    st.error(f"Erro ao gerar PowerPoint: {exc}")
        if "tab_ppt_bytes" in st.session_state:
            st.download_button(
                "Baixar PowerPoint",
                data=st.session_state["tab_ppt_bytes"],
                file_name=f"Apresentacao_{source_name}.pptx",
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                use_container_width=True,
            )

    st.markdown('</div>', unsafe_allow_html=True)


# =============================================================================
# MAIN
# =============================================================================
def main() -> None:
    _secret_to_env()

    if "active_view" not in st.session_state:
        st.session_state["active_view"] = "codificador"

    # ---- SIDEBAR ----
    with st.sidebar:
        try:
            st.image("logo_ifec_header.png", use_container_width=True)
        except Exception:
            st.markdown("### IFec RJ")

        st.markdown("<div style='height: 1.5rem;'></div>", unsafe_allow_html=True)

        # Botao Codificador
        is_cod = st.session_state["active_view"] == "codificador"
        st.markdown(
            f'<div class="{"nav-active" if is_cod else ""}">',
            unsafe_allow_html=True,
        )
        if st.button("Codificador", key="nav_cod"):
            st.session_state["active_view"] = "codificador"
            st.rerun()
        st.markdown('</div>', unsafe_allow_html=True)

        # Botao Tabulacao
        is_tab = st.session_state["active_view"] == "tabulador"
        st.markdown(
            f'<div class="{"nav-active" if is_tab else ""}">',
            unsafe_allow_html=True,
        )
        if st.button("Tabulacao automatica (em teste)", key="nav_tab"):
            st.session_state["active_view"] = "tabulador"
            st.rerun()
        st.markdown('</div>', unsafe_allow_html=True)

        st.markdown(
            "<div style='flex:1; min-height: 6rem;'></div>",
            unsafe_allow_html=True,
        )

        # Status da API discreto no rodape
        api_ok = bool(os.getenv("OPENAI_API_KEY"))
        try:
            if not api_ok:
                key = st.secrets.get("OPENAI_API_KEY", "")
                if key:
                    os.environ["OPENAI_API_KEY"] = key
                    api_ok = True
        except Exception:
            pass

        st.caption("API: configurada" if api_ok else "API: nao configurada")

    # ---- AREA CENTRAL ----
    if st.session_state["active_view"] == "codificador":
        st.markdown('<p class="ifec-page-title">Codificador</p>', unsafe_allow_html=True)
        st.markdown(
            '<p class="ifec-page-sub">Codificacao de pesquisas com IA</p>',
            unsafe_allow_html=True,
        )
        _render_codificador(api_ok=bool(os.getenv("OPENAI_API_KEY")))
    else:
        st.markdown(
            '<p class="ifec-page-title">Tabulacao automatica</p>',
            unsafe_allow_html=True,
        )
        st.markdown(
            '<p class="ifec-page-sub">Em teste — gera Excel e PowerPoint da base</p>',
            unsafe_allow_html=True,
        )
        _render_tabulador()


if __name__ == "__main__":
    main()
