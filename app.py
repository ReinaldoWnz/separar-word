import streamlit as st
import tempfile
import os
import pandas as pd
from PyPDF2 import PdfReader, PdfWriter
from pdf2docx import Converter
import zipfile

st.set_page_config(page_title="Separador DOCX Cloud", layout="centered")

st.title("📄 Separar DOCX por página (Cloud Safe)")

st.markdown("""
### 🔧 Como funciona:
1. Você envia o arquivo `.docx` completo  
2. Ele é convertido internamente em PDF (sem Word nem LibreOffice)  
3. Cada página vira um arquivo separado  
4. Os nomes são gerados conforme sua planilha  
5. Tudo é reconvertido para `.docx` e baixado em ZIP  
---
""")

docx_file = st.file_uploader("📎 Envie o arquivo DOCX", type=["docx"])
table_file = st.file_uploader("📊 Envie a planilha (CSV ou XLSX)", type=["csv", "xlsx"])


def convert_docx_to_pdf_with_pdf2docx(docx_path, pdf_path):
    """Usa pdf2docx para gerar PDF simplificado"""
    from fpdf import FPDF
    from docx import Document

    doc = Document(docx_path)
    pdf = FPDF()
    pdf.set_auto_page_break(auto=True, margin=15)
    pdf.add_page()
    pdf.set_font("Arial", size=12)

    for p in doc.paragraphs:
        pdf.multi_cell(0, 10, p.text)
    pdf.output(pdf_path)


if docx_file and table_file:
    if st.button("🚀 Gerar PDFs e DOCXs Separados"):
        with st.spinner("Processando, aguarde..."):
            with tempfile.TemporaryDirectory() as tmpdir:
                # Salva DOCX
                docx_path = os.path.join(tmpdir, "entrada.docx")
                with open(docx_path, "wb") as f:
                    f.write(docx_file.read())

                # Converte DOCX → PDF (básico, sem LibreOffice)
                pdf_path = os.path.join(tmpdir, "saida.pdf")
                convert_docx_to_pdf_with_pdf2docx(docx_path, pdf_path)

                # Lê planilha
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

                st.info(f"📄 PDF tem {num_pages} páginas; planilha tem {num_linhas} linhas.")

                pdf_dir = os.path.join(tmpdir, "pdfs")
                docx_dir = os.path.join(tmpdir, "docxs")
                os.makedirs(pdf_dir, exist_ok=True)
                os.makedirs(docx_dir, exist_ok=True)

                # Divide PDF por página e reconverte
                progress = st.progress(0)
                for i in range(limite):
                    writer = PdfWriter()
                    writer.add_page(reader.pages[i])
                    nome1 = str(df.iloc[i, 0]).strip().replace("/", "-")
                    nome2 = str(df.iloc[i, 1]).strip().replace("/", "-")
                    base_name = f"PROCURAÇÃO - {nome1} - {nome2}"

                    pdf_out = os.path.join(pdf_dir, f"{base_name}.pdf")
                    with open(pdf_out, "wb") as f_out:
                        writer.write(f_out)

                    # Reconverte para DOCX
                    docx_out = os.path.join(docx_dir, f"{base_name}.docx")
                    cv = Converter(pdf_out)
                    cv.convert(docx_out)
                    cv.close()

                    progress.progress((i + 1) / limite)

                # Gera ZIP final
                zip_path = os.path.join(tmpdir, "procurações.zip")
                with zipfile.ZipFile(zip_path, "w") as zipf:
                    for file in os.listdir(docx_dir):
                        zipf.write(os.path.join(docx_dir, file), file)

                with open(zip_path, "rb") as f:
                    st.success("✅ Arquivos gerados com sucesso!")
                    st.download_button("📦 Baixar ZIP", f, file_name="procurações.zip", mime="application/zip")
