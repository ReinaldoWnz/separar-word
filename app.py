import streamlit as st
import tempfile
import os
import pandas as pd
import zipfile
from PyPDF2 import PdfReader, PdfWriter
from pdf2docx import Converter

st.set_page_config(page_title="Separar e Renomear PDFs", layout="centered")

st.title("📄 Separar e Renomear PDF por Página")

st.markdown("""
### ⚙️ Como usar:
1. Envie o PDF **já exportado pelo Word** (com todas as procurações).  
2. Envie a planilha `.csv` ou `.xlsx` com duas colunas (Nome e Número).  
3. O app vai separar cada página, renomear conforme a tabela e gerar um `.zip` pronto.  
4. (Opcional) você pode escolher converter cada página também em `.docx`.  
---
""")

pdf_file = st.file_uploader("📎 Envie o arquivo PDF", type=["pdf"])
table_file = st.file_uploader("📊 Envie a planilha (CSV ou XLSX)", type=["csv", "xlsx"])
converter_docx = st.checkbox("📝 Também gerar versão .DOCX de cada página", value=False)

if pdf_file and table_file:
    if st.button("🚀 Gerar arquivos separados"):
        with st.spinner("Processando..."):
            with tempfile.TemporaryDirectory() as tmpdir:
                # Salva PDF
                pdf_path = os.path.join(tmpdir, "entrada.pdf")
                with open(pdf_path, "wb") as f:
                    f.write(pdf_file.read())

                # Lê a planilha
                if table_file.name.endswith(".csv"):
                    df = pd.read_csv(table_file)
                else:
                    df = pd.read_excel(table_file)

                if len(df.columns) < 2:
                    st.error("⚠️ A planilha precisa ter DUAS colunas (Nome e Número).")
                    st.stop()

                reader = PdfReader(pdf_path)
                num_pages = len(reader.pages)
                num_linhas = len(df)
                limite = min(num_pages, num_linhas)

                st.info(f"📄 PDF com {num_pages} páginas — Planilha com {num_linhas} linhas.")
                if num_pages != num_linhas:
                    st.warning("⚠️ Quantidades diferentes! Serão gerados apenas até o número menor entre páginas e linhas.")

                pdf_dir = os.path.join(tmpdir, "pdfs")
                docx_dir = os.path.join(tmpdir, "docxs")
                os.makedirs(pdf_dir, exist_ok=True)
                if converter_docx:
                    os.makedirs(docx_dir, exist_ok=True)

                progress = st.progress(0)
                for i in range(limite):
                    nome1 = str(df.iloc[i, 0]).strip().replace("/", "-")
                    nome2 = str(df.iloc[i, 1]).strip().replace("/", "-")
                    base_name = f"PROCURAÇÃO - {nome1} - {nome2}"

                    # Cria PDF separado
                    writer = PdfWriter()
                    writer.add_page(reader.pages[i])
                    pdf_out = os.path.join(pdf_dir, f"{base_name}.pdf")
                    with open(pdf_out, "wb") as f_out:
                        writer.write(f_out)

                    # Se solicitado, gera .docx também
                    if converter_docx:
                        docx_out = os.path.join(docx_dir, f"{base_name}.docx")
                        cv = Converter(pdf_out)
                        cv.convert(docx_out)
                        cv.close()

                    progress.progress((i + 1) / limite)

                # Cria ZIP
                zip_path = os.path.join(tmpdir, "arquivos_separados.zip")
                with zipfile.ZipFile(zip_path, "w") as zipf:
                    for folder in [pdf_dir, docx_dir] if converter_docx else [pdf_dir]:
                        for file in os.listdir(folder):
                            zipf.write(os.path.join(folder, file), file)

                with open(zip_path, "rb") as f:
                    st.success("✅ Tudo pronto!")
                    st.download_button("📦 Baixar ZIP", f, file_name="procurações.zip", mime="application/zip")
