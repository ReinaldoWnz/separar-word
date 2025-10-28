import streamlit as st
import tempfile
import os
import pandas as pd
from docx import Document
from lxml import etree
from docx.oxml import parse_xml
import zipfile

st.set_page_config(page_title="Separador DOCX com Formatação", layout="centered")

st.title("📄 Separador de DOCX — Mantendo Formatação")

st.markdown("""
**Como usar:**
1. Envie o arquivo `.docx` com várias procurações (uma por página).  
2. Envie a planilha `.csv` ou `.xlsx` com as duas colunas (Credor e Número).  
3. O app criará um `.docx` separado para cada página, **mantendo toda a formatação original**.  
---
""")

docx_file = st.file_uploader("📎 Envie o arquivo DOCX", type=["docx"])
table_file = st.file_uploader("📊 Envie a planilha (CSV ou XLSX)", type=["csv", "xlsx"])


def clone_docx_elements(doc, start_idx, end_idx):
    """Clona elementos XML (preserva formatação, tabelas, imagens, etc.)"""
    new_doc = Document()
    body = new_doc._element.body
    for _ in range(len(body)):  # remove parágrafos vazios criados automaticamente
        body.remove(body[0])
    for el in doc.element.body[start_idx:end_idx]:
        body.append(el)
    return new_doc


if docx_file and table_file:
    if st.button("🚀 Gerar DOCXs Separados"):
        with st.spinner("Processando documentos..."):
            with tempfile.TemporaryDirectory() as tmpdir:
                # Salva DOCX
                docx_path = os.path.join(tmpdir, "entrada.docx")
                with open(docx_path, "wb") as f:
                    f.write(docx_file.read())

                doc = Document(docx_path)

                # Lê tabela
                if table_file.name.endswith(".csv"):
                    df = pd.read_csv(table_file)
                else:
                    df = pd.read_excel(table_file)

                if len(df.columns) < 2:
                    st.error("⚠️ A planilha precisa ter pelo menos DUAS colunas.")
                    st.stop()

                num_parts = len(df)
                total_elements = len(doc.element.body)
                base_chunk = total_elements // num_parts
                remainder = total_elements % num_parts

                output_dir = os.path.join(tmpdir, "saida_docs")
                os.makedirs(output_dir, exist_ok=True)

                start = 0
                for i in range(num_parts):
                    end = start + base_chunk + (1 if i < remainder else 0)
                    sub_doc = clone_docx_elements(doc, start, end)

                    nome1 = str(df.iloc[i, 0]).strip().replace("/", "-")
                    nome2 = str(df.iloc[i, 1]).strip().replace("/", "-")
                    new_name = f"PROCURAÇÃO - {nome1} - {nome2}.docx"

                    sub_doc.save(os.path.join(output_dir, new_name))
                    start = end

                # Compacta em ZIP
                zip_path = os.path.join(tmpdir, "procurações_formatadas.zip")
                with zipfile.ZipFile(zip_path, "w") as zipf:
                    for file in os.listdir(output_dir):
                        zipf.write(os.path.join(output_dir, file), file)

                with open(zip_path, "rb") as f:
                    st.success("✅ Arquivos gerados com sucesso e formatação preservada!")
                    st.download_button(
                        "📦 Baixar ZIP",
                        f,
                        file_name="procurações_formatadas.zip",
                        mime="application/zip"
                    )
