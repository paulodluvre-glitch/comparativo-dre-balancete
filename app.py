from __future__ import annotations

import streamlit as st

from comparator import build_report_workbook, compare_reports


st.set_page_config(page_title="Comparativo Balancete x DRE", page_icon=":bar_chart:", layout="wide")

st.title("Comparativo Balancete x DRE")
st.caption(
    "Anexe 1 arquivo de balancete `.xls` e 1 arquivo de DRE `.xlsx` para gerar o relatorio de inconsistencias."
)

with st.sidebar:
    st.subheader("Filtro padrao")
    st.write(
        "Receitas, impostos, folha, despesas financeiras e contas sem nota sao filtradas automaticamente."
    )

balancete_file = st.file_uploader("Balancete do Dominio", type=["xls"])
dre_file = st.file_uploader("DRE do Sempre", type=["xlsx"])

if balancete_file and dre_file:
    try:
        dre_rows, balancete_rows, summary, artifacts = compare_reports(
            balancete_file.getvalue(),
            dre_file.getvalue(),
        )
        report_bytes = build_report_workbook(
            dre_rows=dre_rows,
            balancete_rows=balancete_rows,
            summary=summary,
            artifacts=artifacts,
            balancete_filename=balancete_file.name,
            dre_filename=dre_file.name,
        )

        metric_1, metric_2, metric_3, metric_4 = st.columns(4)
        metric_1.metric("Categorias OK", summary["categorias_ok"])
        metric_2.metric("Categorias divergentes", summary["categorias_divergentes"])
        metric_3.metric(
            "DRE sem balancete",
            summary["categorias_nao_localizadas_no_balancete"],
        )
        metric_4.metric(
            "Balancete sem DRE",
            summary["contas_nao_localizadas_na_dre"],
        )

        st.subheader("Resumo rapido")
        st.write(
            {
                "Categorias DRE analisadas": summary["categorias_dre_analisadas"],
                "Contas balancete analisadas": summary["contas_balancete_analisadas"],
                "Registros DRE filtrados": summary["registros_dre_filtrados"],
                "Registros balancete filtrados": summary["registros_balancete_filtrados"],
            }
        )

        st.download_button(
            label="Baixar relatorio em xlsx",
            data=report_bytes,
            file_name="relatorio_inconsistencias.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    except Exception as error:
        st.error(f"Nao foi possivel processar os arquivos: {error}")
else:
    st.info("Envie os dois arquivos para liberar o processamento.")
