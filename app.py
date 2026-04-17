from __future__ import annotations

import io
import tempfile
from pathlib import Path

import pandas as pd
import streamlit as st

from tabifec_py import (
    auto_tabulacao,
    check_tempo,
    para_codificar,
    resposta_abertura2,
    set_header,
)


st.set_page_config(page_title="TabIFec", layout="wide")


def _load_excel(uploaded_file) -> pd.DataFrame:
    return pd.read_excel(io.BytesIO(uploaded_file.getvalue()))


def _download_excel(df: pd.DataFrame, filename: str, label: str) -> None:
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="dados")
    st.download_button(
        label=label,
        data=buffer.getvalue(),
        file_name=filename,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


def _run_check_tempo(df: pd.DataFrame, entrevistador_col: str) -> pd.DataFrame:
    with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
        temp_path = Path(tmp.name)
    try:
        with pd.ExcelWriter(temp_path, engine="openpyxl") as writer:
            df.to_excel(writer, index=False)
        return check_tempo(var=entrevistador_col, arquivo=temp_path)
    finally:
        temp_path.unlink(missing_ok=True)


def _run_auto_tabulacao(df: pd.DataFrame, tipo_resposta: str) -> str:
    with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp_xlsx:
        temp_xlsx = Path(tmp_xlsx.name)
    with tempfile.NamedTemporaryFile(suffix=".Rmd", delete=False) as tmp_rmd:
        temp_rmd = Path(tmp_rmd.name)
    try:
        with pd.ExcelWriter(temp_xlsx, engine="openpyxl") as writer:
            df.to_excel(writer, index=False)
        output_path = auto_tabulacao(
            arquivo=temp_xlsx,
            tipo_resposta=tipo_resposta,
            output=temp_rmd,
        )
        return output_path.read_text(encoding="utf-8")
    finally:
        temp_xlsx.unlink(missing_ok=True)
        temp_rmd.unlink(missing_ok=True)


def main() -> None:
    st.title("TabIFec no Streamlit")
    st.write(
        "Esta versao substitui a interface desktop baseada em `tkinter` por componentes nativos do Streamlit."
    )

    uploaded_file = st.file_uploader("Envie a base SurveyMonkey em Excel", type=["xlsx", "xls"])
    if uploaded_file is None:
        st.info("Carregue um arquivo para habilitar as operacoes.")
        return

    try:
        bruto = _load_excel(uploaded_file)
    except Exception as exc:
        st.error(f"Nao foi possivel ler o arquivo enviado: {exc}")
        return

    with st.spinner("Formatando cabecalho da base..."):
        base = set_header(bruto)

    st.success("Arquivo carregado com sucesso.")

    with st.expander("Visualizar base formatada", expanded=True):
        st.dataframe(base.head(50), use_container_width=True)
        _download_excel(base, "base_formatada.xlsx", "Baixar base formatada")

    st.divider()
    st.subheader("Codificacao de respostas abertas")
    pergunta_cod = st.selectbox(
        "Escolha a pergunta para preparar a codificacao",
        options=list(base.columns),
        index=0,
    )
    if st.button("Gerar planilha para codificacao", use_container_width=True):
        try:
            cod_df = para_codificar(pergunta_cod, base)
        except Exception as exc:
            st.error(f"Falha ao preparar a codificacao: {exc}")
        else:
            st.dataframe(cod_df, use_container_width=True)
            _download_excel(cod_df, "para_codificar.xlsx", "Baixar planilha de codificacao")

    st.divider()
    st.subheader("Tabulacao")
    col1, col2 = st.columns(2)
    with col1:
        pergunta_tab = st.selectbox("Pergunta para tabular", options=list(base.columns), key="pergunta_tab")
    with col2:
        abertura = st.selectbox(
            "Abertura",
            options=["TOTAL"] + list(base.columns),
            index=0,
            key="abertura_tab",
        )

    if st.button("Gerar tabela", use_container_width=True):
        try:
            tabela = resposta_abertura2(base, pergunta_tab, abertura=abertura)
        except Exception as exc:
            st.error(f"Falha ao gerar a tabela: {exc}")
        else:
            st.dataframe(tabela, use_container_width=True)
            _download_excel(tabela, "tabela.xlsx", "Baixar tabela")

    st.divider()
    st.subheader("Checagem de tempo de resposta")
    candidatos_entrevistador = [
        col for col in base.columns if "Entrevistador" in str(col) or "Pesquisador" in str(col)
    ]
    entrevistador_col = st.selectbox(
        "Coluna do entrevistador",
        options=candidatos_entrevistador or list(base.columns),
    )

    if st.button("Analisar tempos suspeitos", use_container_width=True):
        try:
            suspeitos = _run_check_tempo(base, entrevistador_col)
        except Exception as exc:
            st.error(f"Falha na checagem de tempo: {exc}")
        else:
            st.dataframe(suspeitos, use_container_width=True)
            _download_excel(suspeitos, "tempos_suspeitos.xlsx", "Baixar resultado da checagem")

    st.divider()
    st.subheader("RMarkdown automatico")
    tipo_resposta = st.selectbox(
        "Funcao de tabulacao no RMarkdown",
        options=["resposta_abertura2", "resposta_abertura", "resposta_abertura3", "resposta_abertura4"],
        index=0,
    )
    if st.button("Gerar RMarkdown", use_container_width=True):
        try:
            rmd_text = _run_auto_tabulacao(base, tipo_resposta)
        except Exception as exc:
            st.error(f"Falha ao gerar o RMarkdown: {exc}")
        else:
            st.code(rmd_text, language="markdown")
            st.download_button(
                "Baixar RMarkdown",
                data=rmd_text.encode("utf-8"),
                file_name="Tabelas_Cruzamentos.Rmd",
                mime="text/markdown",
            )


if __name__ == "__main__":
    main()
