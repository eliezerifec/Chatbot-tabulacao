<<<<<<< HEAD
from streamlit_app import main
=======
import io
import json
import sys
from pathlib import Path

import pandas as pd
import streamlit as st


def _find_source_dir() -> Path:
    candidates = [
        Path(__file__).resolve().parent,
        Path(r"C:\Users\Eliezer.Rodrigues\Documents\git"),
    ]
    for candidate in candidates:
        if (candidate / "codificador.py").exists():
            return candidate
    raise FileNotFoundError(
        "Nao encontrei codificador.py. Coloque este app na mesma pasta do projeto "
        "ou ajuste o caminho em _find_source_dir()."
    )


SOURCE_DIR = _find_source_dir()
if str(SOURCE_DIR) not in sys.path:
    sys.path.insert(0, str(SOURCE_DIR))

from codificador import TIPOS_PERGUNTA, CodificadorIA, _get_api_key  # noqa: E402


st.set_page_config(
    page_title="Codificador de Pesquisas",
    page_icon=":bar_chart:",
    layout="wide",
)


MODOS_UI = {
    "Simples": "simples",
    "Multipla": "multipla",
    "Semiaberta - Simples": "semi_simples",
    "Semiaberta - Multipla": "semi_multipla",
}


def init_state():
    defaults = {
        "codificador": CodificadorIA(),
        "uploaded_name": None,
        "uploaded_bytes": None,
        "sheets": {},
        "sheet_names": [],
        "logs": ["Sistema iniciado - aguardando arquivo"],
        "resultado_pronto": False,
        "contexto_global": (
            "Pesquisa de satisfacao de evento.\n"
            "Respostas curtas de participantes."
        ),
    }
    for key, value in defaults.items():
        if key not in st.session_state:
            st.session_state[key] = value


def log(message: str):
    st.session_state.logs.append(message)


def load_uploaded_file(uploaded_file):
    if not uploaded_file:
        return

    file_bytes = uploaded_file.getvalue()
    same_file = (
        st.session_state.uploaded_name == uploaded_file.name
        and st.session_state.uploaded_bytes == file_bytes
    )
    if same_file:
        return

    buffer = io.BytesIO(file_bytes)
    if uploaded_file.name.lower().endswith(".csv"):
        sheets = {"Planilha": pd.read_csv(buffer)}
    else:
        excel = pd.ExcelFile(buffer)
        sheets = {name: excel.parse(name) for name in excel.sheet_names}

    st.session_state.uploaded_name = uploaded_file.name
    st.session_state.uploaded_bytes = file_bytes
    st.session_state.sheets = sheets
    st.session_state.sheet_names = list(sheets.keys())
    st.session_state.resultado_pronto = False
    log(
        f"Arquivo carregado: {uploaded_file.name} | "
        f"{len(sheets)} aba(s) | {sum(len(df) for df in sheets.values())} linhas"
    )


def render_top_metrics():
    codificador = st.session_state.codificador
    banco_total = "-"
    banco_taxa = "sem dados"
    try:
        stats = codificador.banco.stats()
        banco_total = str(stats.get("total", "-"))
        banco_taxa = f"{stats.get('taxa_acerto', 0)}% precisao"
    except Exception:
        pass

    cols = st.columns(4)
    cols[0].metric("Arquivo", st.session_state.uploaded_name or "-")
    cols[1].metric("Abas", len(st.session_state.sheet_names))
    cols[2].metric("Modelo", "gpt-5.4 / gpt-4o")
    cols[3].metric("Banco IA", banco_total, banco_taxa)


def render_sidebar():
    with st.sidebar:
        st.header("Configuracao")
        st.caption(f"Fonte detectada: `{SOURCE_DIR}`")
        if not _get_api_key():
            st.warning(
                "Configure `OPENAI_API_KEY` em `st.secrets`, variavel de ambiente "
                "ou arquivo `.env` para habilitar a codificacao."
            )
        st.text_area(
            "Contexto global",
            key="contexto_global",
            height=100,
            help="Usado principalmente no tipo Personalizado.",
        )

        uploaded_codes = st.file_uploader(
            "Importar mapeamentos",
            type=["json", "xlsx"],
            help="Opcional. Carrega codigos existentes no motor.",
        )
        if uploaded_codes and st.button("Aplicar mapeamentos", use_container_width=True):
            try:
                if uploaded_codes.name.lower().endswith(".json"):
                    data = json.loads(uploaded_codes.getvalue().decode("utf-8"))
                else:
                    df_map = pd.read_excel(io.BytesIO(uploaded_codes.getvalue()))
                    data = dict(
                        zip(
                            df_map.iloc[:, 0].astype(str),
                            df_map.iloc[:, 1].astype(str),
                        )
                    )
                st.session_state.codificador.carregar_codigos(data)
                log(f"Mapeamentos importados: {len(data)} item(ns)")
                st.success("Mapeamentos aplicados.")
            except Exception as exc:
                st.error(f"Erro ao importar mapeamentos: {exc}")

        categories_text = st.text_input(
            "Categorias manuais",
            placeholder="Ex.: Organizacao, Atendimento, Musica",
        )
        if st.button("Adicionar categorias", use_container_width=True):
            added = 0
            for cat in categories_text.split(","):
                clean = cat.strip()
                if clean:
                    st.session_state.codificador.adicionar_categoria(clean)
                    added += 1
            if added:
                log(f"Categorias adicionadas manualmente: {added}")
                st.success(f"{added} categoria(s) adicionada(s).")

        if st.session_state.codificador.categorias:
            st.caption("Categorias atuais")
            st.write(", ".join(st.session_state.codificador.categorias))


def render_upload_section():
    st.subheader("Upload da Base")
    uploaded_file = st.file_uploader(
        "Envie XLSX ou CSV",
        type=["xlsx", "csv"],
        label_visibility="collapsed",
    )
    if uploaded_file:
        try:
            load_uploaded_file(uploaded_file)
        except Exception as exc:
            st.error(f"Erro ao abrir arquivo: {exc}")

    if st.session_state.uploaded_name:
        st.success(
            f"{st.session_state.uploaded_name} carregado com "
            f"{len(st.session_state.sheet_names)} aba(s)."
        )


def render_sheet_config():
    if not st.session_state.sheet_names:
        st.info("Carregue uma planilha para configurar as abas.")
        return

    st.subheader("Configuracao das Abas")
    tipo_labels = {value["label"]: key for key, value in TIPOS_PERGUNTA.items()}
    tipo_options = list(tipo_labels.keys())
    modo_options = list(MODOS_UI.keys())

    for sheet_name in st.session_state.sheet_names:
        df = st.session_state.sheets[sheet_name]
        cols = list(df.columns)
        with st.expander(f"Aba: {sheet_name} ({len(df)} linhas)", expanded=True):
            col1, col2, col3, col4 = st.columns([1, 2, 2, 2])
            col1.checkbox("Processar", value=True, key=f"enabled::{sheet_name}")
            col2.selectbox("Coluna de entrada", cols, key=f"in::{sheet_name}")
            col3.selectbox(
                "Coluna de saida",
                options=cols + ["codigo_ia"],
                index=(cols + ["codigo_ia"]).index("codigo_ia"),
                key=f"out::{sheet_name}",
            )
            col4.selectbox("Tipo", tipo_options, key=f"tipo::{sheet_name}")

            col5, col6 = st.columns([2, 2])
            col5.selectbox("Modo", modo_options, key=f"modo::{sheet_name}")
            modo = st.session_state[f"modo::{sheet_name}"]

            if "Semiaberta" in modo:
                semi1, semi2 = st.columns(2)
                semi1.selectbox(
                    "Coluna imputada",
                    options=cols + ["col_imputado"],
                    index=(cols + ["col_imputado"]).index("col_imputado"),
                    key=f"imp::{sheet_name}",
                )
                semi2.selectbox(
                    "Coluna nova",
                    options=cols + ["col_nova"],
                    index=(cols + ["col_nova"]).index("col_nova"),
                    key=f"novo::{sheet_name}",
                )

            tipo_key = tipo_labels[st.session_state[f"tipo::{sheet_name}"]]
            if tipo_key == "livre":
                col6.text_input(
                    "Contexto personalizado",
                    key=f"ctx::{sheet_name}",
                    placeholder="Descreva como a IA deve categorizar esta aba.",
                )
            else:
                col6.caption(TIPOS_PERGUNTA[tipo_key]["descricao"])

            st.dataframe(df.head(5), use_container_width=True)


def get_selected_configs():
    tipo_labels = {value["label"]: key for key, value in TIPOS_PERGUNTA.items()}
    selected = []

    for sheet_name in st.session_state.sheet_names:
        if not st.session_state.get(f"enabled::{sheet_name}", True):
            continue

        selected.append(
            {
                "sheet_name": sheet_name,
                "col_in": st.session_state[f"in::{sheet_name}"],
                "col_out": st.session_state[f"out::{sheet_name}"],
                "modo": MODOS_UI[st.session_state[f"modo::{sheet_name}"]],
                "tipo": tipo_labels[st.session_state[f"tipo::{sheet_name}"]],
                "ctx": st.session_state.get(f"ctx::{sheet_name}", "").strip(),
                "col_imp": st.session_state.get(f"imp::{sheet_name}", "col_imputado"),
                "col_novo": st.session_state.get(f"novo::{sheet_name}", "col_nova"),
            }
        )

    return selected


def run_processing():
    selected = get_selected_configs()
    if not selected:
        st.warning("Selecione ao menos uma aba para processar.")
        return

    progress = st.progress(0)
    status = st.empty()
    log_box = st.empty()
    total = len(selected)

    for index, cfg in enumerate(selected, start=1):
        sheet_name = cfg["sheet_name"]
        df = st.session_state.sheets[sheet_name].copy()
        respostas = df[cfg["col_in"]].astype(str).tolist()
        status.info(f"Processando aba {index}/{total}: {sheet_name}")
        log(
            f"Iniciando '{sheet_name}' | tipo={cfg['tipo']} | modo={cfg['modo']} | "
            f"{len(respostas)} respostas"
        )

        try:
            resultado = st.session_state.codificador.codificar_lote_modo(
                respostas=respostas,
                tipo=cfg["tipo"],
                modo=cfg["modo"],
                contexto_custom=cfg["ctx"] or st.session_state.contexto_global,
                categorias_imputacao=st.session_state.codificador.categorias[:],
            )

            if "imputado" in resultado:
                df[cfg["col_imp"]] = resultado["imputado"]
                df[cfg["col_novo"]] = resultado["novo"]
                log(
                    f"Aba '{sheet_name}' concluida com colunas "
                    f"'{cfg['col_imp']}' e '{cfg['col_novo']}'"
                )
            else:
                df[cfg["col_out"]] = resultado["resultado"]
                log(f"Aba '{sheet_name}' concluida em '{cfg['col_out']}'")

            st.session_state.sheets[sheet_name] = df
        except Exception as exc:
            log(f"Erro em '{sheet_name}': {exc}")
            st.error(f"Falha ao processar '{sheet_name}': {exc}")

        progress.progress(index / total)
        log_box.code("\n".join(st.session_state.logs[-12:]), language="text")

    st.session_state.resultado_pronto = True
    status.success("Codificacao finalizada.")


def build_excel_bytes():
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        for sheet_name, df in st.session_state.sheets.items():
            df.to_excel(writer, sheet_name=sheet_name[:31], index=False)
    output.seek(0)
    return output.getvalue()


def render_export_section():
    if not st.session_state.sheets:
        return

    st.subheader("Exportacao")
    excel_bytes = build_excel_bytes()
    st.download_button(
        "Baixar XLSX",
        data=excel_bytes,
        file_name="resultado_codificado.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )

    if len(st.session_state.sheets) == 1:
        sheet_name, df = next(iter(st.session_state.sheets.items()))
        csv_bytes = df.to_csv(index=False).encode("utf-8-sig")
        st.download_button(
            f"Baixar CSV ({sheet_name})",
            data=csv_bytes,
            file_name=f"{sheet_name}.csv",
            mime="text/csv",
            use_container_width=True,
        )


def render_logs():
    st.subheader("Log do Sistema")
    st.code("\n".join(st.session_state.logs[-20:]), language="text")


def main():
    init_state()
    render_sidebar()

    st.title("Codificador de Pesquisas")
    st.caption("Interface web em Streamlit para codificacao de pesquisas.")

    render_top_metrics()
    render_upload_section()
    render_sheet_config()

    process_cols = st.columns([1, 1, 2])
    if process_cols[0].button("Iniciar codificacao", type="primary", use_container_width=True):
        run_processing()
    if process_cols[1].button("Limpar logs", use_container_width=True):
        st.session_state.logs = ["Logs reiniciados."]

    render_export_section()

    if st.session_state.resultado_pronto and st.session_state.sheets:
        st.subheader("Previa do Resultado")
        first_sheet = st.selectbox("Escolha uma aba para visualizar", st.session_state.sheet_names)
        st.dataframe(st.session_state.sheets[first_sheet], use_container_width=True)

    render_logs()
>>>>>>> 6f8655e1ba9f62285bca01d4f1b80fe19097f13e


if __name__ == "__main__":
    main()
