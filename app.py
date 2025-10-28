import streamlit as st
import tempfile
import os
import pandas as pd
import zipfile
from PyPDF2 import PdfReader, PdfWriter
from docx2pdf import convert
import pypandoc
from pdf2docx import Converter

st.set_page_config(page_title="DOCX → PDF → Split → DOCX", layout="centered")

st.title("📄 Divisor Automático de DOCX (com Reconversão para DOCX)")

st.markdown("""
1️⃣ Envie o arquivo `.docx` principal  
2️⃣ Envie a planilha `.csv` ou `.xlsx` com duas colunas  
3️⃣ O app irá converter o DOCX em PDF, separar página por página,  
renomear conforme sua planilha e reconverter cada uma em `.docx`  
---
""")

docx_file = st.file_uploader("📎 Envie o arquivo DOCX", type=["docx"])
table_file = st.file_uploader("📊 Envie a planilha (CSV ou XLSX)", type=["csv", "xlsx"])


def convert_docx_to_pdf(docx_path, pdf_path):
    """Converte DOCX em PDF mantendo formatação."""
    try:
        convert(docx_path, pdf_path)
    except Exception:
        pypandoc.convert_file(docx_path, 'pdf', outputfile=pdf_path)


if docx_file and table_file:
    if st.button("🚀 Gerar DOCXs Separados"):
        with st.spinner("Processando... isso pode levar alguns minutos."):
            with tempfile.TemporaryDirectory() as tmpdir:
                # Salva DOCX
                docx_path = os.path.join(tmpdir, "entrada.docx")
                with open(docx_path, "wb") as f:
                    f.write(docx_file.read())

                # Converte para PDF
                pdf_path = os.path.join(tmpdir, "saida.pdf")
                convert_docx_to_pdf(docx_path, pdf_path)

                # Lê planilha
                if table_file.name.endswith(".csv"):
                    df = pd.read_csv(table_file)
                else:
                    df = pd.read_excel(table_file)

                if len(df.columns) < 2:
                    st.error("⚠️ A planilha precisa ter DUAS colunas (nome e número).")
                    st.stop()

                reader = PdfReader(pdf_path)
                output_dir = os.path.join(tmpdir, "pdfs")
                os.makedirs(output_dir, exist_ok=True)

                # Divide e salva PDFs
                for i, page in enumerate(reader.pages):
                    writer = PdfWriter()
                    writer.add_page(page)

                    nome1 = str(df.iloc[i, 0]).strip().replace("/", "-")
                    nome2 = str(df.iloc[i, 1]).strip().replace("/", "-")
                    nome_base = f"PROCURAÇÃO - {nome1} - {nome2}"

                    pdf_output = os.path.join(output_dir, f"{nome_base}.pdf")
                    with open(pdf_output, "wb") as f_out:
                        writer.write(f_out)

                # Reconverte cada PDF → DOCX
                docx_output_dir = os.path.join(tmpdir, "docxs")
                os.makedirs(docx_output_dir, exist_ok=True)

                for file in os.listdir(output_dir):
                    if file.endswith(".pdf"):
                        input_pdf = os.path.join(output_dir, file)
                        output_docx = os.path.join(docx_output_dir, file.replace(".pdf", ".docx"))
                        cv = Converter(input_pdf)
                        cv.convert(output_docx, start=0, end=None)
                        cv.close()

                # Compacta todos os DOCX
                zip_path = os.path.join(tmpdir, "procurações_final.zip")
                with zipfile.ZipFile(zip_path, "w") as zipf:
                    for file in os.listdir(docx_output_dir):
                        zipf.write(os.path.join(docx_output_dir, file), file)

                with open(zip_path, "rb") as f:
                    st.success("✅ Tudo pronto! DOCXs separados e renomeados.")
                    st.download_button("📦 Baixar ZIP", f, file_name="procurações_final.zip", mime="application/zip")
