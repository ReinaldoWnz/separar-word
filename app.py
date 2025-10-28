import streamlit as st
import tempfile
import os
import pandas as pd
import zipfile
from PyPDF2 import PdfReader, PdfWriter
from pdf2docx import Converter
import pypandoc

st.set_page_config(page_title="PDF → DOCX Texto Editável", layout="centered")

st.title("📄 Separar PDF e Converter para DOCX (Texto Editável)")

st.markdown("""
### ⚙️ Fluxo:
1️⃣ Envie o **PDF gerado no Word**  
2️⃣ Envie a planilha `.csv` ou `.xlsx` com as colunas (Nome e Número)  
3️⃣ O app:
- Separa o PDF página por página  
- Renomeia conforme a planilha  
- Converte cada uma em `.docx` **com texto real**  
---
""")

pdf_file = st.file_uploader("📎 Envie o arquivo PDF", type=["pdf"])
table_file = st.file_uploader("📊 Envie a planilha (CSV ou XLSX)", type=["csv", "xlsx"])

def safe_pdf_to_docx(pdf_in, docx_out):
    """Converte PDF em DOCX de forma segura, preservando texto"""
    try:
        cv = Converter(pdf_in)
        cv.convert(docx_out, start=0, end=None, graceful=True, single_thread=True)
        cv.close()
        return True
    except Exception:
        try:
            pypandoc.convert_file(pdf_in, 'docx', outputfile=docx_out)
            return True
        except Exception:
            return False

if pdf_file and table_file:
    if st.button("🚀 Gerar DOCXs Separados"):
        with st.spinner("Processando, aguarde..."):
            with tempfile.TemporaryDirectory() as tmpdir:
                # Salva PDF
                pdf_path = os.path.join(tmpdir, "entrada.pdf")
                with open(pdf_path, "wb") as f:
                    f.write(pdf_file.read())

                # Lê planilha
                if table_file.name.endswith(".csv"):
                    df = pd.read_csv(table_file)
                else:
                    df = pd.read_excel(table_file)

                if len(df.columns) < 2:
                    st.error("⚠️ A planilha precisa ter duas colunas (Nome e Número).")
                    st.stop()

                reader = PdfReader(pdf_path)
                num_pages = len(reader.pages)
                limite = min(num_pages, len(df))

                st.info(f"📄 PDF com {num_pages} páginas e {len(df)} linhas na planilha.")
                if num_pages != len(df):
                    st.warning("⚠️ Quantidades diferentes — serão processados até o número menor.")

                pdf_dir = os.path.join(tmpdir, "pdfs")
                docx_dir = os.path.join(tmpdir, "docxs")
                os.makedirs(pdf_dir, exist_ok=True)
                os.makedirs(docx_dir, exist_ok=True)

                progress = st.progress(0)
                sucesso, falhas = 0, []

                for i in range(limite):
                    nome1 = str(df.iloc[i, 0]).strip().replace("/", "-")
                    nome2 = str(df.iloc[i, 1]).strip().replace("/", "-")
                    base_name = f"PROCURAÇÃO - {nome1} - {nome2}"

                    # Salva a página individual
                    writer = PdfWriter()
                    writer.add_page(reader.pages[i])
                    pdf_out = os.path.join(pdf_dir, f"{base_name}.pdf")
                    with open(pdf_out, "wb") as f_out:
                        writer.write(f_out)

                    # Converte para DOCX texto real
                    docx_out = os.path.join(docx_dir, f"{base_name}.docx")
                    ok = safe_pdf_to_docx(pdf_out, docx_out)

                    if ok:
                        sucesso += 1
                    else:
                        falhas.append(base_name)

                    progress.progress((i + 1) / limite)

                # Compacta
                zip_path = os.path.join(tmpdir, "procurações_texto.zip")
                with zipfile.ZipFile(zip_path, "w") as zipf:
                    for file in os.listdir(docx_dir):
                        zipf.write(os.path.join(docx_dir, file), file)

                with open(zip_path, "rb") as f:
                    st.success(f"✅ {sucesso}/{limite} DOCXs convertidos com sucesso!")
                    if falhas:
                        st.warning(f"⚠️ Falharam: {', '.join(falhas[:5])}...")
                    st.download_button("📦 Baixar ZIP", f, file_name="procurações_texto.zip", mime="application/zip")
