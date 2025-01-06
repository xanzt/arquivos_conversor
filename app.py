import streamlit as st
from docx import Document
import pandas as pd
import io

def excel_to_word(excel_file):
    df = pd.read_excel(excel_file)
    doc = Document()

    table = doc.add_table(rows=(len(df) + 1), cols=len(df.columns))
    table.style = 'Table Grid'

    for i, column in enumerate(df.columns):
        table.cell(0, i).text = str(column)

    for row_index, row in enumerate(df.values):
        for col_index, value in enumerate(row):
            table.cell(row_index + 1, col_index).text = str(value)

    output = io.BytesIO()
    doc.save(output)
    return output.getvalue()

def main():
    st.title("Conversor de Arquivos")
    
    # Opções de conversão no menu lateral
    st.sidebar.header("Opções de Conversão")
    option = st.sidebar.selectbox(
        "Escolha a conversão:",
        ("Excel para Word")
    )

    uploaded_file = st.file_uploader("Envie o arquivo Excel", type=["xlsx"])

    if uploaded_file is not None:
        if option == "Excel para Word":
            if uploaded_file.name.endswith(".xlsx"):
                st.success("Arquivo Excel carregado!")
                word_data = excel_to_word(uploaded_file)
                # Garantindo que o botão de download apareça após a conversão
                st.download_button(
                    label="Baixar Word",
                    data=word_data,
                    file_name="convertido.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
            else:
                st.error("Por favor, envie um arquivo Excel (.xlsx).")

if __name__ == "__main__":
    main()
