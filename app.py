import streamlit as st
import tempfile

from src.excel_utils import (
    load_workbook_data,
    get_sheet_names,
    get_columns,
    generate_filtered_files
)

st.set_page_config(
    page_title="Gerador de Excel Filtrado",
    layout="centered"
)

st.title("📊 Gerador Automático de Arquivos Excel")

st.markdown(
    """
    Faça upload de uma planilha Excel e gere arquivos filtrados automaticamente.
    """
)

uploaded_file = st.file_uploader(
    "Selecione um arquivo Excel",
    type=["xlsx", "xlsm"]
)

if uploaded_file:

    try:

        with tempfile.NamedTemporaryFile(
            delete=False,
            suffix=".xlsx"
        ) as tmp:

            tmp.write(uploaded_file.getbuffer())
            temp_path = tmp.name

        workbook = load_workbook_data(temp_path)

        sheet_names = get_sheet_names(workbook)

        default_sheet_index = (
            sheet_names.index("Ativos")
            if "Ativos" in sheet_names
            else 0
        )

        selected_sheet = st.selectbox(
            "Selecione a aba",
            sheet_names,
            index=default_sheet_index
        )

        columns = get_columns(
            workbook,
            selected_sheet
        )

        default_column_index = 0

        for i, col in enumerate(columns):

            normalized = str(col).strip().upper()

            if normalized in [
                "BASE_NUCLEO",
                "BASE NUCLEO",
                "BASE_NÚCLEO",
                "BASE NÚCLEO"
            ]:
                default_column_index = i
                break

        selected_column = st.selectbox(
            "Selecione a coluna de filtro",
            columns,
            index=default_column_index
        )

        unique_values = get_columns(
            workbook,
            selected_sheet,
            return_unique=True,
            filter_column=selected_column
        )

        generation_mode = st.radio(
            "Modo de geração",
            [
                "Gerar todos os arquivos",
                "Gerar apenas itens selecionados"
            ]
        )

        selected_values = []

        if generation_mode == "Gerar apenas itens selecionados":

            selected_values = st.multiselect(
                "Selecione os valores",
                unique_values
            )

        if st.button("🚀 Gerar Arquivos"):

            progress_bar = st.progress(0)

            status_text = st.empty()

            values_to_generate = (
                unique_values
                if generation_mode == "Gerar todos os arquivos"
                else selected_values
            )

            if not values_to_generate:

                st.warning(
                    "Selecione ao menos um item."
                )

            else:
                st.info(f"{len(values_to_generate)} arquivo(s) serão gerados.")
                zip_path = generate_filtered_files(
                    source_file=temp_path,
                    sheet_name=selected_sheet,
                    filter_column=selected_column,
                    values=values_to_generate,
                    progress_bar=progress_bar,
                    status_text=status_text
                )

                st.success(
                    "Arquivos gerados com sucesso!"
                )

                with open(zip_path, "rb") as f:

                    st.download_button(
                        label="📥 Baixar Arquivos ZIP",
                        data=f,
                        file_name="arquivos_filtrados.zip",
                        mime="application/zip"
                    )

    except Exception as e:

        st.error(
            f"Erro ao processar arquivo: {str(e)}"
        )
