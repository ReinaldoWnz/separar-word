import streamlit as st
import tempfile
import os
import pandas as pd
from docx import Document
import zipfile
import copy

st.set_page_config(page_title="Separador DOCX com Formatação", layout="centered")

st.title("📄 Separador de DOCX — Mantendo Formatação")

st.markdown("""
**Como usar:**
1. Envie o arquivo `.docx` com várias procurações (uma por página ou bloco).  
2. Envie a planilha `.csv` ou `.xlsx` com as colunas **Credor** e **Número**.  
3. O app cria um `.docx` separado para cada parte, **mantendo toda a formatação original**.  
---
""")

docx_file = st.file_uploader("📎 Envie o arquivo DOCX", type=["docx"])
table_file = st.file_uploader("📊 Envie a planilha (CSV ou XLSX)", type=["csv", "xlsx"])


def split_docx_with_formatting(doc_path, num_parts):
    """Divide o DOCX em partes iguais, copiando os elementos XML com deepcopy (mantém formatação)."""
    doc = Document(doc_path)
    body_elements = list(doc.element.body)
    total_elements = len(body_elements)

    chunk_size = total_elements // num_parts
    remainder = total_elements % num_parts

    docs = []
    start = 0
    for i in range(num_parts):
        end = start + chunk_size + (1 if i < remainder else 0)
        sub_doc = Document()
        body = sub_doc._element.body

        # Remove parágrafos vazios criados automaticamente
        for _ in range(len(body)):
            body.remove(body[0])

        # Copia os elementos XML com deepcopy (mantém tudo)
        for el in body_elements[start:end]:
            body.append(copy.deepcopy(el))

        docs.append(sub_doc)
        start = end
    return docs


if docx_file and table_file:
    if st.button("🚀 Gerar DOCXs Separados"):
        with st.spinner("Processando documentos..."):
            with tempfile.TemporaryDirectory() as tmpdir:
                # Salva o DOCX original
                docx_path = os.path.join(tmpdir, "entrada.docx")
                with open(docx_path, "wb") as f:
                    f.write(docx_file.read())

                # Lê a planilha
                if table_file.name.endswith(".csv"):
                    df = pd.read_csv(table_file)
                else:
                    df = pd.read_excel(table_file)

                if len(df.columns) < 2:
                    st.error("⚠️ A planilha precisa ter pelo menos duas colunas.")
                    st.stop()

                num_parts = len(df)

                # Divide o DOCX mantendo a formatação
                parts = split_docx_with_formatting(docx_path, num_parts)

                output_dir = os.path.join(tmpdir, "saida_docs")
                os.makedirs(output_dir, exist_ok=True)

                # Salva cada parte com nome da planilha
                for i, sub_doc in enumerate(parts):
                    nome1 = str(df.iloc[i, 0]).strip().replace("/", "-")
                    nome2 = str(df.iloc[i, 1]).strip().replace("/", "-")
                    nome_final = f"PROCURAÇÃO - {nome1} - {nome2}.docx"
                    sub_doc.save(os.path.join(output_dir, nome_final))

                # Compacta tudo em ZIP
                zip_path = os.path.join(tmpdir, "procurações_formatadas.zip")
                with zipfile.ZipFile(zip_path, "w") as zipf:
                    for file in os.listdir(output_dir):
                        zipf.write(os.path.join(output_dir, file), file)

                with open(zip_path, "rb") as f:
                    st.success("✅ Arquivos DOCX gerados com formatação preservada!")
                    st.download_button(
                        "📦 Baixar ZIP",
                        f,
                        file_name="procurações_formatadas.zip",
                        mime="application/zip"
                    )
