"""
streamlit_app.py — Codificador de Pesquisas IFec RJ
====================================================
Interface web em Streamlit com identidade visual herdada do app desktop
(Tkinter), paleta IFec, layout denso e cards profissionais.

Fluxo:
  - Sidebar: logo IFec, upload da base, upload opcional de pesquisa anterior,
    status da OPENAI_API_KEY.
  - Aba COD: codificação por aba com IA (CodificadorIA).
  - Aba TAB: detecção de perguntas, tabulação, exportação Excel + PPT.
"""

import os
from io import BytesIO
from pathlib import Path
import tempfile

import pandas as pd
import streamlit as st


# ─── Configuração da página ──────────────────────────────────────────────────
st.set_page_config(
    page_title="Codificador IFec",
    page_icon="logo_ifec.png",
    layout="wide",
    initial_sidebar_state="expanded",
)


# ─── Paleta IFec (espelha o legacy_tkinter_app.py) ───────────────────────────
# NAV_BG=#111827  NAV_ACTIVE=#1d4ed8  AZUL_DARK=#1e3a8a  AZUL_LIGHT=#eff6ff
# VERDE=#059669  ROXO=#7c3aed  OURO=#d97706
# Mantemos as cores via CSS variables para reuso e consistência visual.

_IFEC_CSS = """
<style>
:root {
    --ifec-nav-bg: #111827;
    --ifec-nav-active: #1d4ed8;
    --ifec-azul: #1d4ed8;
    --ifec-azul-dark: #1e3a8a;
    --ifec-azul-light: #eff6ff;
    --ifec-azul-mid: #bfdbfe;
    --ifec-verde: #059669;
    --ifec-verde-light: #ecfdf5;
    --ifec-roxo: #7c3aed;
    --ifec-roxo-light: #f5f3ff;
    --ifec-ouro: #d97706;
    --ifec-ouro-light: #fffbeb;
    --ifec-bg: #f9fafb;
    --ifec-card: #ffffff;
    --ifec-border: #e5e7eb;
    --ifec-border-strong: #d1d5db;
    --ifec-txt1: #111827;
    --ifec-txt2: #374151;
    --ifec-txt3: #6b7280;
    --ifec-txt4: #9ca3af;
}

/* Plano de fundo geral */
.stApp { background: var(--ifec-bg); }

/* Densidade: reduz padding vertical do container principal */
.block-container {
    padding-top: 0.5rem !important;
    padding-bottom: 2rem !important;
    max-width: 1400px;
}

/* Tipografia base mais densa */
html, body, [class*="css"] {
    font-family: "Segoe UI", "Inter", system-ui, -apple-system, sans-serif;
}

/* ── HEADER da aplicação ───────────────────────────────────────────────── */
.ifec-header {
    background: linear-gradient(90deg, var(--ifec-azul-dark) 0%, var(--ifec-nav-active) 100%);
    padding: 14px 22px;
    border-radius: 10px;
    margin-bottom: 16px;
    display: flex;
    align-items: center;
    gap: 14px;
    box-shadow: 0 2px 8px rgba(29, 78, 216, 0.18);
}
.ifec-header-icon {
    width: 38px; height: 38px;
    border-radius: 8px;
    background: rgba(255,255,255,0.12);
    display: flex; align-items: center; justify-content: center;
    color: #ffffff; font-size: 20px; font-weight: 700;
    border: 1px solid rgba(255,255,255,0.18);
}
.ifec-header-title {
    color: #ffffff;
    font-size: 1.15rem;
    font-weight: 700;
    line-height: 1.2;
    margin: 0;
}
.ifec-header-sub {
    color: var(--ifec-azul-mid);
    font-size: 0.78rem;
    margin: 2px 0 0 0;
    letter-spacing: 0.02em;
}

/* ── SIDEBAR ───────────────────────────────────────────────────────────── */
[data-testid="stSidebar"] {
    background: var(--ifec-nav-bg);
    border-right: 1px solid #1f2937;
}
[data-testid="stSidebar"] * { color: #e5e7eb; }
[data-testid="stSidebar"] h1,
[data-testid="stSidebar"] h2,
[data-testid="stSidebar"] h3,
[data-testid="stSidebar"] h4 { color: #ffffff !important; }

/* Logo box no topo da sidebar */
.ifec-sidebar-logo {
    background: var(--ifec-azul-dark);
    padding: 14px 12px;
    border-radius: 8px;
    margin-bottom: 12px;
    text-align: center;
    border: 1px solid #1f2937;
}
.ifec-sidebar-divider {
    height: 1px;
    background: #1f2937;
    margin: 12px 0;
}
.ifec-sidebar-section {
    color: #93c5fd !important;
    font-size: 0.7rem;
    font-weight: 700;
    letter-spacing: 0.1em;
    text-transform: uppercase;
    margin: 4px 0 6px 0;
}

/* Uploaders na sidebar — fundo um tom acima */
[data-testid="stSidebar"] [data-testid="stFileUploaderDropzone"] {
    background: #1f2937;
    border: 1px dashed #374151;
}
[data-testid="stSidebar"] [data-testid="stFileUploaderDropzone"]:hover {
    border-color: var(--ifec-nav-active);
}

/* Badge de status da API */
.ifec-badge {
    display: inline-flex; align-items: center; gap: 6px;
    padding: 6px 10px;
    border-radius: 6px;
    font-size: 0.78rem;
    font-weight: 600;
    width: 100%;
    justify-content: center;
}
.ifec-badge-ok {
    background: rgba(5, 150, 105, 0.18);
    color: #34d399;
    border: 1px solid rgba(5, 150, 105, 0.4);
}
.ifec-badge-warn {
    background: rgba(217, 119, 6, 0.18);
    color: #fbbf24;
    border: 1px solid rgba(217, 119, 6, 0.4);
}

/* ── ABAS principais (COD / TAB) ───────────────────────────────────────── */
.stTabs [data-baseweb="tab-list"] {
    gap: 6px;
    border-bottom: 2px solid var(--ifec-border);
    padding-bottom: 0;
}
.stTabs [data-baseweb="tab"] {
    background: var(--ifec-card);
    border: 1px solid var(--ifec-border);
    border-bottom: none;
    border-radius: 8px 8px 0 0;
    padding: 10px 22px;
    font-weight: 600;
    font-size: 0.92rem;
    color: var(--ifec-txt3);
    transition: all 0.15s ease;
}
.stTabs [data-baseweb="tab"]:hover {
    background: var(--ifec-azul-light);
    color: var(--ifec-azul-dark);
}
.stTabs [aria-selected="true"] {
    background: var(--ifec-nav-active) !important;
    color: #ffffff !important;
    border-color: var(--ifec-nav-active) !important;
}
.stTabs [data-baseweb="tab-highlight"] { display: none; }

/* ── CARDS (section header + bloco de conteúdo) ─────────────────────────── */
.ifec-section-header {
    display: flex; align-items: center; gap: 10px;
    margin: 14px 0 8px 0;
    padding-bottom: 6px;
    border-bottom: 1px solid var(--ifec-border);
}
.ifec-pill {
    background: var(--ifec-azul-light);
    color: var(--ifec-azul-dark);
    padding: 5px 9px;
    border-radius: 6px;
    font-size: 1rem;
    font-weight: 700;
    line-height: 1;
}
.ifec-pill-verde { background: var(--ifec-verde-light); color: var(--ifec-verde); }
.ifec-pill-roxo  { background: var(--ifec-roxo-light);  color: var(--ifec-roxo); }
.ifec-pill-ouro  { background: var(--ifec-ouro-light);  color: var(--ifec-ouro); }
.ifec-section-title {
    font-size: 0.98rem;
    font-weight: 700;
    color: var(--ifec-txt1);
    line-height: 1.1;
    margin: 0;
}
.ifec-section-sub {
    font-size: 0.78rem;
    color: var(--ifec-txt4);
    margin: 2px 0 0 0;
}

/* ── MÉTRICAS estilizadas (cards com borda lateral colorida) ────────────── */
.ifec-stats { display: flex; gap: 12px; margin: 8px 0 14px 0; flex-wrap: wrap; }
.ifec-stat {
    flex: 1 1 0;
    min-width: 160px;
    background: var(--ifec-card);
    border: 1px solid var(--ifec-border);
    border-left: 4px solid var(--ifec-azul);
    border-radius: 8px;
    padding: 12px 16px;
    box-shadow: 0 1px 2px rgba(0,0,0,0.04);
}
.ifec-stat-label {
    font-size: 0.72rem;
    color: var(--ifec-txt3);
    font-weight: 600;
    text-transform: uppercase;
    letter-spacing: 0.06em;
    margin: 0 0 4px 0;
}
.ifec-stat-value {
    font-size: 1.35rem;
    font-weight: 700;
    color: var(--ifec-txt1);
    margin: 0;
    line-height: 1.1;
    word-break: break-word;
}
.ifec-stat-verde { border-left-color: var(--ifec-verde); }
.ifec-stat-roxo  { border-left-color: var(--ifec-roxo); }
.ifec-stat-ouro  { border-left-color: var(--ifec-ouro); }

/* ── BOTÕES primários no padrão IFec ───────────────────────────────────── */
.stButton > button[kind="primary"] {
    background: var(--ifec-nav-active);
    border-color: var(--ifec-nav-active);
    color: #ffffff;
    font-weight: 600;
    border-radius: 6px;
    box-shadow: 0 1px 2px rgba(29, 78, 216, 0.2);
}
.stButton > button[kind="primary"]:hover {
    background: var(--ifec-azul-dark);
    border-color: var(--ifec-azul-dark);
    color: #ffffff;
}
.stButton > button[kind="primary"]:disabled {
    background: #93c5fd;
    border-color: #93c5fd;
}

/* Botões secundários */
.stButton > button[kind="secondary"] {
    background: var(--ifec-card);
    color: var(--ifec-txt1);
    border: 1px solid var(--ifec-border-strong);
    font-weight: 600;
    border-radius: 6px;
}
.stButton > button[kind="secondary"]:hover {
    border-color: var(--ifec-azul);
    color: var(--ifec-azul);
}

.stDownloadButton > button {
    background: var(--ifec-verde);
    color: #ffffff;
    border: none;
    font-weight: 600;
    border-radius: 6px;
}
.stDownloadButton > button:hover {
    background: #047857;
    color: #ffffff;
}

/* ── EXPANDERS (cartões de configuração) ───────────────────────────────── */
[data-testid="stExpander"] {
    border: 1px solid var(--ifec-border) !important;
    border-radius: 8px !important;
    background: var(--ifec-card) !important;
    margin-bottom: 8px;
    box-shadow: 0 1px 2px rgba(0,0,0,0.03);
}
[data-testid="stExpander"] summary {
    font-weight: 600;
    color: var(--ifec-txt1);
    padding: 10px 14px !important;
}
[data-testid="stExpander"] summary:hover {
    background: var(--ifec-azul-light);
    color: var(--ifec-azul-dark);
}

/* ── INPUTS — bordas suaves consistentes ────────────────────────────────── */
.stTextInput input, .stTextArea textarea, .stSelectbox div[data-baseweb="select"] {
    border-radius: 6px !important;
}

/* ── ALERTS mais limpos ────────────────────────────────────────────────── */
.stAlert {
    border-radius: 8px;
    border-width: 1px;
}

/* ── DATAFRAME ──────────────────────────────────────────────────────────── */
[data-testid="stDataFrame"] {
    border: 1px solid var(--ifec-border);
    border-radius: 8px;
    overflow: hidden;
}

/* ── Esconder o "Made with Streamlit" / footer ──────────────────────────── */
footer { visibility: hidden; }
#MainMenu { visibility: hidden; }

/* ── Headings: títulos de seção mais compactos ──────────────────────────── */
h1, h2, h3, h4 { color: var(--ifec-txt1); }
h2 { font-size: 1.25rem !important; margin-top: 0.6rem !important; }
h3 { font-size: 1.05rem !important; margin-top: 0.5rem !important; }

/* Caption mais legível */
.stCaption, [data-testid="stCaptionContainer"] {
    color: var(--ifec-txt3) !important;
    font-size: 0.84rem !important;
}

/* Divider mais sutil */
hr { border-color: var(--ifec-border) !important; margin: 0.8rem 0 !important; }
</style>
"""
st.markdown(_IFEC_CSS, unsafe_allow_html=True)


# ─── Helpers visuais ─────────────────────────────────────────────────────────
def _render_header(title: str, subtitle: str, icon: str = "▣") -> None:
    """Header azul tipo barra de aplicativo."""
    st.markdown(
        f"""
        <div class="ifec-header">
            <div class="ifec-header-icon">{icon}</div>
            <div>
                <p class="ifec-header-title">{title}</p>
                <p class="ifec-header-sub">{subtitle}</p>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def _section(title: str, sub: str = "", icon: str = "●", variant: str = "azul") -> None:
    """Cabeçalho de seção com pílula colorida (estilo sec_header do Tkinter)."""
    pill_class = {
        "azul": "ifec-pill",
        "verde": "ifec-pill ifec-pill-verde",
        "roxo": "ifec-pill ifec-pill-roxo",
        "ouro": "ifec-pill ifec-pill-ouro",
    }.get(variant, "ifec-pill")
    sub_html = f'<p class="ifec-section-sub">{sub}</p>' if sub else ""
    st.markdown(
        f"""
        <div class="ifec-section-header">
            <span class="{pill_class}">{icon}</span>
            <div>
                <p class="ifec-section-title">{title}</p>
                {sub_html}
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def _stats(items: list[tuple[str, str, str]]) -> None:
    """Renderiza cards de estatística. items=[(label, value, variant)]."""
    cards_html = []
    for label, value, variant in items:
        klass = {
            "azul": "ifec-stat",
            "verde": "ifec-stat ifec-stat-verde",
            "roxo": "ifec-stat ifec-stat-roxo",
            "ouro": "ifec-stat ifec-stat-ouro",
        }.get(variant, "ifec-stat")
        cards_html.append(
            f'<div class="{klass}">'
            f'<p class="ifec-stat-label">{label}</p>'
            f'<p class="ifec-stat-value">{value}</p>'
            f"</div>"
        )
    st.markdown(
        f'<div class="ifec-stats">{"".join(cards_html)}</div>',
        unsafe_allow_html=True,
    )


# ─── Utilitários originais (preservados) ─────────────────────────────────────
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
def _read_tabulation_file(name: str, data: bytes, two_line_header: bool) -> pd.DataFrame:
    bio = BytesIO(data)
    if name.lower().endswith(".csv"):
        return pd.read_csv(bio)

    if two_line_header:
        from tabulador import set_header

        raw = pd.read_excel(bio, header=None, sheet_name=0)
        return set_header(raw)

    return pd.read_excel(bio, sheet_name=0)


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


def _build_tab_excel(df: pd.DataFrame, perguntas: list[dict], titulo: str) -> bytes:
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
        )
        return Path(path).read_bytes()
    finally:
        try:
            Path(path).unlink(missing_ok=True)
        except Exception:
            pass


def _build_tab_ppt(
    df: pd.DataFrame,
    perguntas: list[dict],
    titulo: str,
    subtitulo: str,
    periodo: str,
) -> bytes:
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


def _import_previous(uploaded) -> dict[str, list[str]]:
    if uploaded is None:
        return {}

    sheets = _read_uploaded_file(uploaded.name, uploaded.getvalue())
    previous: dict[str, list[str]] = {}
    for sheet_name, df in sheets.items():
        if df.empty:
            continue
        default_col = df.columns[-1]
        category_col = st.selectbox(
            f"Coluna de categorias em {sheet_name}",
            list(df.columns),
            index=list(df.columns).index(default_col),
            key=f"prev_col_{sheet_name}",
        )
        cats = (
            df[category_col]
            .dropna()
            .astype(str)
            .map(str.strip)
            .loc[lambda s: s.ne("")]
            .drop_duplicates()
            .tolist()
        )
        previous[sheet_name] = cats
    return previous


def _run_coding(
    sheets: dict[str, pd.DataFrame],
    configs: dict[str, dict],
    global_context: str,
    previous_categories: dict[str, list[str]],
) -> dict[str, pd.DataFrame]:
    codificador = _get_codificador()
    result_sheets = {name: df.copy() for name, df in sheets.items()}

    selected = [name for name, cfg in configs.items() if cfg["selected"]]
    progress = st.progress(0, text="Preparando codificacao...")
    log_box = st.empty()
    logs: list[str] = []

    for sheet_idx, sheet_name in enumerate(selected):
        cfg = configs[sheet_name]
        df = result_sheets[sheet_name]
        respostas = df[cfg["input_col"]].astype(str).tolist()
        cats_prev = previous_categories.get(sheet_name, [])
        cats_imp = cats_prev or cfg["categories"]
        context = cfg["context"] if cfg["type_key"] == "livre" else global_context

        progress.progress(
            sheet_idx / max(len(selected), 1),
            text=f"Codificando aba: {sheet_name}",
        )

        codes = codificador.codificar_lote(
            respostas,
            cfg["type_key"],
            cfg["mode_key"],
            categorias_imputadas=cats_imp,
            contexto=context,
        )

        if "semi" in cfg["mode_key"]:
            imp_col = cfg["imputed_col"]
            new_col = cfg["new_col"]
            df[imp_col] = [c.get("imputado", "") for c in codes]
            df[new_col] = [c.get("nova", "") for c in codes]
        else:
            df[cfg["output_col"]] = [c.get("codigo", "") for c in codes]

        logs.append(f"[OK] {sheet_name}: {len(codes)} resposta(s) processada(s).")
        log_box.code("\n".join(logs[-12:]))

    progress.progress(1.0, text="Codificacao concluida.")
    return result_sheets


# ─── Render: COD ─────────────────────────────────────────────────────────────
def _render_codificador(uploaded, previous_file, api_ok: bool) -> None:
    _section(
        "Codificador IA",
        "Configure cada aba e codifique respostas abertas com a OpenAI.",
        icon="⌨",
        variant="azul",
    )

    if not uploaded:
        st.info("Envie uma planilha na barra lateral para iniciar a codificação.")
        return

    if not api_ok:
        st.warning(
            "OPENAI_API_KEY não está configurada. Defina a chave em "
            "`.streamlit/secrets.toml` ou como variável de ambiente para liberar a codificação."
        )

    try:
        sheets = _read_uploaded_file(uploaded.name, uploaded.getvalue())
    except Exception as exc:
        st.error(f"Não foi possível ler o arquivo: {exc}")
        return

    if not sheets:
        st.warning("Nenhuma aba encontrada no arquivo.")
        return

    tipos_pergunta, modos_resposta = _load_taxonomies()

    _section("Contexto e referências", icon="◆", variant="roxo")
    global_context = st.text_area(
        "Contexto geral",
        placeholder="Descreva o objetivo da pesquisa e os critérios de classificação.",
        height=100,
        label_visibility="visible",
    )

    with st.expander("Pesquisa anterior (reaproveitar categorias)", expanded=previous_file is not None):
        previous_categories = _import_previous(previous_file)
        if previous_categories:
            st.success(
                f"{sum(len(v) for v in previous_categories.values())} categorias carregadas."
            )
        else:
            st.caption("Opcional: carregue uma pesquisa já codificada para reutilizar categorias.")

    _section(
        "Configuração das abas",
        f"{len(sheets)} aba(s) encontrada(s). Selecione e ajuste cada uma.",
        icon="❑",
        variant="azul",
    )

    configs: dict[str, dict] = {}
    type_labels = {data["label"]: key for key, data in tipos_pergunta.items()}
    mode_labels = {data["label"]: key for key, data in modos_resposta.items()}

    for sheet_name, df in sheets.items():
        with st.expander(f"📄  {sheet_name}  —  {len(df)} linha(s) × {len(df.columns)} coluna(s)", expanded=True):
            cols = list(df.columns)
            left, mid, right = st.columns([1, 2, 2])
            selected = left.checkbox("Codificar", value=True, key=f"sel_{sheet_name}")
            input_col = mid.selectbox("Coluna de entrada", cols, key=f"in_{sheet_name}")
            output_col = right.text_input(
                "Coluna de saída",
                value="codigo_ia",
                key=f"out_{sheet_name}",
            )

            col_a, col_b = st.columns(2)
            type_label = col_a.selectbox(
                "Tipo da pergunta",
                list(type_labels),
                index=list(type_labels).index(next(k for k, v in type_labels.items() if v == "livre")),
                key=f"type_{sheet_name}",
            )
            mode_label = col_b.selectbox(
                "Modo de resposta",
                list(mode_labels),
                key=f"mode_{sheet_name}",
            )

            mode_key = mode_labels[mode_label]
            categories = []
            imputed_col = "col_imputado"
            new_col = "col_nova"
            if "semi" in mode_key:
                sem_a, sem_b = st.columns(2)
                imputed_col = sem_a.text_input(
                    "Coluna de imputação",
                    value="col_imputado",
                    key=f"imp_{sheet_name}",
                )
                new_col = sem_b.text_input(
                    "Coluna de novas categorias",
                    value="col_nova",
                    key=f"new_{sheet_name}",
                )
                categories = _parse_list(
                    st.text_area(
                        "Categorias pré-definidas",
                        placeholder="Uma categoria por linha ou separadas por vírgula.",
                        key=f"cats_{sheet_name}",
                    )
                )

            custom_context = st.text_area(
                "Contexto específico",
                placeholder="Use quando o tipo da pergunta for Personalizado.",
                key=f"ctx_{sheet_name}",
                height=80,
            )

            configs[sheet_name] = {
                "selected": selected,
                "input_col": input_col,
                "output_col": output_col.strip() or "codigo_ia",
                "type_key": type_labels[type_label],
                "mode_key": mode_key,
                "categories": categories,
                "imputed_col": imputed_col.strip() or "col_imputado",
                "new_col": new_col.strip() or "col_nova",
                "context": custom_context.strip(),
            }

            st.dataframe(df.head(20), use_container_width=True, hide_index=True)

    can_run = api_ok and any(cfg["selected"] for cfg in configs.values())
    st.divider()
    if st.button("▶  Iniciar codificação", type="primary", disabled=not can_run, use_container_width=False):
        with st.spinner("Codificando respostas..."):
            try:
                result_sheets = _run_coding(
                    sheets,
                    configs,
                    global_context.strip(),
                    previous_categories,
                )
            except Exception as exc:
                st.error(f"Erro durante a codificação: {exc}")
                return

        st.session_state["result_sheets"] = result_sheets
        st.success("Codificação finalizada.")

    if "result_sheets" in st.session_state:
        _section("Resultado", "Pré-visualize e baixe o Excel codificado.", icon="✓", variant="verde")
        preview_sheet = st.selectbox(
            "Aba para visualizar",
            list(st.session_state["result_sheets"].keys()),
            key="preview_sheet",
        )
        st.dataframe(
            st.session_state["result_sheets"][preview_sheet].head(100),
            use_container_width=True,
            hide_index=True,
        )
        st.download_button(
            "⬇  Baixar Excel codificado",
            data=_to_excel(st.session_state["result_sheets"]),
            file_name="base_codificada.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )


# ─── Render: TAB ─────────────────────────────────────────────────────────────
def _render_tabulador(uploaded) -> None:
    _section(
        "Tabulador",
        "Detecta perguntas, gera tabulação em Excel e monta o PowerPoint.",
        icon="📊",
        variant="roxo",
    )

    fonte = st.radio(
        "Fonte da base",
        ["Arquivo enviado", "Resultado codificado"],
        horizontal=True,
        disabled="result_sheets" not in st.session_state,
    )

    if fonte == "Resultado codificado" and "result_sheets" in st.session_state:
        sheet_name = st.selectbox(
            "Aba para tabular",
            list(st.session_state["result_sheets"].keys()),
            key="tab_result_sheet",
        )
        df_tab = st.session_state["result_sheets"][sheet_name].copy()
        source_name = sheet_name
    else:
        if uploaded is None:
            st.info("Envie uma planilha na barra lateral para usar o tabulador.")
            return
        two_line_header = st.checkbox(
            "Usar cabeçalho em duas linhas do SurveyMonkey/TabIFec",
            value=True,
            help="Combina as duas primeiras linhas como nomes de colunas, igual ao app local.",
        )
        try:
            df_tab = _read_tabulation_file(uploaded.name, uploaded.getvalue(), two_line_header)
        except Exception as exc:
            st.error(f"Não foi possível preparar a base para tabulação: {exc}")
            return
        source_name = Path(uploaded.name).stem

    from tabulador import TIPOS_LABEL, detectar_perguntas, tabular_pergunta

    _stats([
        ("Base", source_name, "azul"),
        ("Respondentes", f"{len(df_tab):,}".replace(",", "."), "verde"),
        ("Colunas", str(len(df_tab.columns)), "roxo"),
    ])

    if st.button("🔍  Detectar perguntas", type="primary"):
        try:
            st.session_state["tab_questions"] = detectar_perguntas(df_tab)
            st.session_state["tab_source_name"] = source_name
        except Exception as exc:
            st.error(f"Erro ao detectar perguntas: {exc}")

    perguntas_detectadas = st.session_state.get("tab_questions", [])
    if not perguntas_detectadas:
        _section("Pré-visualização da base", icon="◐", variant="ouro")
        st.dataframe(df_tab.head(30), use_container_width=True, hide_index=True)
        return

    st.success(f"{len(perguntas_detectadas)} pergunta(s) detectada(s).")

    _section("Identificação do relatório", icon="✎", variant="ouro")
    head_a, head_b, head_c = st.columns(3)
    titulo = head_a.text_input("Título do relatório", value="Pesquisa IFec RJ", key="tab_title")
    subtitulo = head_b.text_input("Subtítulo do PowerPoint", value="", key="tab_subtitle")
    periodo = head_c.text_input("Período", value="", key="tab_period")

    _section(
        "Perguntas detectadas",
        "Revise tipo, texto e nota antes de exportar.",
        icon="❑",
        variant="azul",
    )

    tipo_keys = list(TIPOS_LABEL.keys())
    tipo_labels = [f"{key} - {TIPOS_LABEL[key]}" for key in tipo_keys]
    perguntas_config = []

    for idx, pergunta in enumerate(perguntas_detectadas):
        label = f"{pergunta.get('num', f'P{idx + 1:02d}')} — {pergunta.get('pergunta', '')}"
        with st.expander(label, expanded=idx < 3):
            top_a, top_b, top_c = st.columns([1, 2, 4])
            ativo = top_a.checkbox("Ativa", value=pergunta.get("ativo", True), key=f"tab_active_{idx}")
            tipo_atual = pergunta.get("tipo", "ABERTA")
            tipo_index = tipo_keys.index(tipo_atual) if tipo_atual in tipo_keys else tipo_keys.index("ABERTA")
            tipo_label = top_b.selectbox(
                "Tipo",
                tipo_labels,
                index=tipo_index,
                key=f"tab_type_{idx}",
            )
            texto = top_c.text_input(
                "Pergunta",
                value=str(pergunta.get("pergunta", "")),
                key=f"tab_question_{idx}",
            )
            nota = st.text_area(
                "Nota",
                value=str(pergunta.get("nota", "")),
                height=70,
                key=f"tab_note_{idx}",
            )

            cfg = dict(pergunta)
            cfg["ativo"] = ativo
            cfg["tipo"] = tipo_keys[tipo_labels.index(tipo_label)]
            cfg["pergunta"] = texto.strip() or pergunta.get("pergunta", "")
            cfg["nota"] = nota.strip()
            perguntas_config.append(cfg)

            colunas_str = ", ".join(str(c) for c in cfg.get("colunas", [])[:6])
            if len(cfg.get("colunas", [])) > 6:
                colunas_str += "…"
            st.caption(f"Colunas: {colunas_str}")
            if st.checkbox("Prévia da tabulação", key=f"tab_preview_{idx}"):
                try:
                    st.dataframe(tabular_pergunta(df_tab, cfg), use_container_width=True, hide_index=True)
                except Exception as exc:
                    st.warning(f"Não foi possível tabular esta pergunta: {exc}")

    ativas = [p for p in perguntas_config if p.get("ativo") and p.get("tipo") != "IGNORAR"]
    st.divider()

    _section(
        "Exportação",
        f"{len(ativas)} pergunta(s) ativa(s) para exportação.",
        icon="⬇",
        variant="verde",
    )

    gen_excel, gen_ppt = st.columns(2)
    with gen_excel:
        if st.button("📑  Gerar Excel de tabulação", disabled=not ativas, type="primary"):
            with st.spinner("Gerando Excel..."):
                try:
                    st.session_state["tab_excel_bytes"] = _build_tab_excel(df_tab, ativas, titulo)
                except Exception as exc:
                    st.error(f"Erro ao gerar Excel: {exc}")
        if "tab_excel_bytes" in st.session_state:
            st.download_button(
                "⬇  Baixar Excel",
                data=st.session_state["tab_excel_bytes"],
                file_name=f"Tabulacao_{source_name}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

    with gen_ppt:
        if st.button("🎯  Gerar PowerPoint", disabled=not ativas, type="primary"):
            with st.spinner("Gerando PowerPoint..."):
                try:
                    st.session_state["tab_ppt_bytes"] = _build_tab_ppt(
                        df_tab,
                        ativas,
                        titulo,
                        subtitulo,
                        periodo,
                    )
                except Exception as exc:
                    st.error(f"Erro ao gerar PowerPoint: {exc}")
        if "tab_ppt_bytes" in st.session_state:
            st.download_button(
                "⬇  Baixar PowerPoint",
                data=st.session_state["tab_ppt_bytes"],
                file_name=f"Apresentacao_{source_name}.pptx",
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            )


# ─── Sidebar ─────────────────────────────────────────────────────────────────
def _render_sidebar():
    with st.sidebar:
        # Logo IFec em caixa azul
        logo_path = Path("logo_ifec_header.png")
        if logo_path.exists():
            st.markdown('<div class="ifec-sidebar-logo">', unsafe_allow_html=True)
            st.image(str(logo_path), use_container_width=True)
            st.markdown("</div>", unsafe_allow_html=True)
        else:
            st.markdown(
                '<div class="ifec-sidebar-logo"><h3 style="margin:0;color:#fff;">IFec RJ</h3></div>',
                unsafe_allow_html=True,
            )

        st.markdown(
            '<p class="ifec-sidebar-section">Instituto Fecomércio</p>'
            '<p style="margin:0;font-size:0.78rem;color:#9ca3af;">'
            "Codificação e tabulação de pesquisas via Streamlit."
            "</p>",
            unsafe_allow_html=True,
        )

        st.markdown('<div class="ifec-sidebar-divider"></div>', unsafe_allow_html=True)

        st.markdown('<p class="ifec-sidebar-section">Arquivos</p>', unsafe_allow_html=True)
        uploaded = st.file_uploader(
            "Base de dados",
            type=["xlsx", "csv"],
            help="Planilha SurveyMonkey ou similar (.xlsx ou .csv).",
        )
        previous_file = st.file_uploader(
            "Pesquisa anterior",
            type=["xlsx", "csv"],
            help="Opcional. Use para reaproveitar categorias já validadas.",
        )

        st.markdown('<div class="ifec-sidebar-divider"></div>', unsafe_allow_html=True)

        st.markdown('<p class="ifec-sidebar-section">Status</p>', unsafe_allow_html=True)
        api_ok = bool(os.getenv("OPENAI_API_KEY"))
        if api_ok:
            st.markdown(
                '<div class="ifec-badge ifec-badge-ok">● OPENAI_API_KEY configurada</div>',
                unsafe_allow_html=True,
            )
        else:
            st.markdown(
                '<div class="ifec-badge ifec-badge-warn">⚠ OPENAI_API_KEY ausente</div>',
                unsafe_allow_html=True,
            )

        return uploaded, previous_file, api_ok


# ─── Main ────────────────────────────────────────────────────────────────────
def main() -> None:
    _secret_to_env()

    uploaded, previous_file, api_ok = _render_sidebar()

    _render_header(
        title="Codificador de Pesquisas — IFec RJ",
        subtitle="Codificação com IA · Tabulação · Exportação Excel & PowerPoint",
        icon="◆",
    )

    tab_cod, tab_tab = st.tabs(["⌨   COD · Codificação", "📊   TAB · Tabulação"])
    with tab_cod:
        _render_codificador(uploaded, previous_file, api_ok)
    with tab_tab:
        _render_tabulador(uploaded)


if __name__ == "__main__":
    main()