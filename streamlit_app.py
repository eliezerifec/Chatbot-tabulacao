import os
from io import BytesIO

import pandas as pd
import streamlit as st


st.set_page_config(
    page_title="Codificador IFec",
    page_icon="IFec",
    layout="wide",
    initial_sidebar_state="expanded",
)


st.markdown(
    """
    <style>
    .block-container { padding-top: 1.5rem; padding-bottom: 2rem; }
    [data-testid="stSidebar"] { background: #111827; }
    [data-testid="stSidebar"] * { color: #f9fafb; }
    [data-testid="stSidebar"] .stButton button { width: 100%; }
    .metric-card {
        border: 1px solid #e5e7eb;
        border-radius: 8px;
        padding: 14px 16px;
        background: #ffffff;
    }
    .section-title {
        font-size: 0.9rem;
        font-weight: 700;
        color: #111827;
        margin: 0 0 0.35rem 0;
    }
    .muted { color: #6b7280; font-size: 0.86rem; }
    </style>
    """,
    unsafe_allow_html=True,
)


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

        logs.append(f"{sheet_name}: {len(respostas)} respostas")
        log_box.code("\n".join(logs[-15:]), language="text")

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
        logs.append(f"{sheet_name}: concluida")
        log_box.code("\n".join(logs[-15:]), language="text")

    progress.progress(1.0, text="Codificacao concluida.")
    return result_sheets


def main() -> None:
    _secret_to_env()

    with st.sidebar:
        st.image("logo_ifec_header.png", use_container_width=True)
        st.markdown("### Codificador")
        st.caption("Versao web para uso no Streamlit via GitHub.")
        uploaded = st.file_uploader("Base de dados", type=["xlsx", "csv"])
        previous_file = st.file_uploader(
            "Pesquisa anterior",
            type=["xlsx", "csv"],
            help="Opcional. Use para reaproveitar categorias ja validadas.",
        )

        api_ok = bool(os.getenv("OPENAI_API_KEY"))
        if api_ok:
            st.success("OPENAI_API_KEY configurada")
        else:
            st.warning("Configure OPENAI_API_KEY em Secrets no Streamlit Cloud")

    st.title("Codificador de Pesquisas")
    st.caption("Upload, configuracao por aba, codificacao com IA e exportacao em Excel.")

    if uploaded is None:
        st.info("Envie uma planilha .xlsx ou .csv para comecar.")
        return

    try:
        sheets = _read_uploaded_file(uploaded.name, uploaded.getvalue())
    except Exception as exc:
        st.error(f"Nao foi possivel ler o arquivo: {exc}")
        return

    total_rows = sum(len(df) for df in sheets.values())
    c1, c2, c3 = st.columns(3)
    c1.metric("Arquivo", uploaded.name)
    c2.metric("Abas", len(sheets))
    c3.metric("Linhas", total_rows)

    try:
        tipos_pergunta, modos_resposta = _load_taxonomies()
    except Exception as exc:
        st.error(str(exc))
        st.stop()

    global_context = st.text_area(
        "Contexto geral",
        placeholder="Descreva o objetivo da pesquisa e os criterios de classificacao.",
        height=100,
    )

    with st.expander("Pesquisa anterior", expanded=previous_file is not None):
        previous_categories = _import_previous(previous_file)
        if previous_categories:
            st.success(
                f"{sum(len(v) for v in previous_categories.values())} categorias carregadas."
            )
        else:
            st.caption("Opcional: carregue uma pesquisa ja codificada para reutilizar categorias.")

    st.subheader("Configuracao das abas")
    configs: dict[str, dict] = {}
    type_labels = {data["label"]: key for key, data in tipos_pergunta.items()}
    mode_labels = {data["label"]: key for key, data in modos_resposta.items()}

    for sheet_name, df in sheets.items():
        with st.expander(sheet_name, expanded=True):
            cols = list(df.columns)
            left, mid, right = st.columns([1, 2, 2])
            selected = left.checkbox("Codificar", value=True, key=f"sel_{sheet_name}")
            input_col = mid.selectbox("Coluna de entrada", cols, key=f"in_{sheet_name}")
            output_col = right.text_input(
                "Coluna de saida",
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
                    "Coluna de imputacao",
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
                        "Categorias pre-definidas",
                        placeholder="Uma categoria por linha ou separadas por virgula.",
                        key=f"cats_{sheet_name}",
                    )
                )

            custom_context = st.text_area(
                "Contexto especifico",
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
    if st.button("Iniciar codificacao", type="primary", disabled=not can_run):
        with st.spinner("Codificando respostas..."):
            try:
                result_sheets = _run_coding(
                    sheets,
                    configs,
                    global_context.strip(),
                    previous_categories,
                )
            except Exception as exc:
                st.error(f"Erro durante a codificacao: {exc}")
                return

        st.session_state["result_sheets"] = result_sheets
        st.success("Codificacao finalizada.")

    if "result_sheets" in st.session_state:
        st.subheader("Resultado")
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
            "Baixar Excel codificado",
            data=_to_excel(st.session_state["result_sheets"]),
            file_name="base_codificada.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )


if __name__ == "__main__":
    main()
