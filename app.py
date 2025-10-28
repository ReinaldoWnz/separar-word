import streamlit as st
import tempfile
import os
import pandas as pd
import zipfile
from PyPDF2 import PdfReader, PdfWriter
from pdf2docx import Converter

st.set_page_config(page_title="PDF → DOCX com Renomeação", layout="centered")

st.title("📄 Separar PDF e Converter para DOCX (Renomeando Automaticamente)")

st.markdown("""
### ⚙️ Como usar:
1️⃣ Envie o arquivo **PDF completo** (gerado pelo Word).  
2️⃣ Envie a planilha `.csv` ou `.xlsx` com duas colunas (Nome e Número).  
3️⃣ O app vai:
- Separar cada página do PDF,  
- Renomear conforme a planilha,  
- Converter cada uma em `.docx`,  
- Gerar um `.zip` com tudo pronto.  
---
""")

pdf_file = st.file_uploader("📎 Envie o arquivo PDF", type=["pdf"])
table_file = st.file_uploader("📊 Envie a planilha (CSV ou XLSX)", type=["csv", "xlsx"])

if pdf_file and table_file:
    if st.button("🚀 Gerar DOCXs Separados"):
        with st.spinner("Processando, aguarde..."):
            with tempfile.TemporaryDirectory() as tmpdir:
                # Salva PDF original
                pdf_path = os.path.join(tmpdir, "entrada.pdf")
                with open(pdf_path, "wb") as f:
                    f.write(pdf_file.read())

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

                st.info(f"📄 PDF tem {num_pages} páginas | 📊 Planilha tem {num_linhas} linhas.")
                if num_pages != num_linhas:
                    st.warning("⚠️ Quantidades diferentes! Serão processados até o número menor entre páginas e linhas.")

                # Cria pastas temporárias
                pdf_dir = os.path.join(tmpdir, "pdfs")
                docx_dir = os.path.join(tmpdir, "docxs")
                os.makedirs(pdf_dir, exist_ok=True)
                os.makedirs(docx_dir, exist_ok=True)

                progress = st.progress(0)
                status = st.empty()

                # Divide, renomeia e converte
                for i in range(limite):
                    nome1 = str(df.iloc[i, 0]).strip().replace("/", "-")
                    nome2 = str(df.iloc[i, 1]).strip().replace("/", "-")
                    base_name = f"PROCURAÇÃO - {nome1} - {nome2}"

                    # Salva página individual
                    writer = PdfWriter()
                    writer.add_page(reader.pages[i])
                    pdf_out = os.path.join(pdf_dir, f"{base_name}.pdf")
                    with open(pdf_out, "wb") as f_out:
                        writer.write(f_out)

                    # Converte para DOCX (modo seguro)
                    try:
                        docx_out = os.path.join(docx_dir, f"{base_name}.docx")
                        cv = Converter(pdf_out)
                        cv.convert(docx_out, start=0, end=None, graceful=True)
                        cv.close()
                    except Exception:
                        status.warning(f"⚠️ Erro ao converter {base_name}.pdf — arquivo pulado.")
                        continue

                    progress.progress((i + 1) / limite)
                    status.text(f"Convertendo {i+1}/{limite}: {base_name}")

                # Compacta todos os DOCX
                zip_path = os.path.join(tmpdir, "procurações_final.zip")
                with zipfile.ZipFile(zip_path, "w") as zipf:
                    for file in os.listdir(docx_dir):
                        zipf.write(os.path.join(docx_dir, file), file)

                with open(zip_path, "rb") as f:
                    st.success("✅ Conversão concluída com sucesso!")
                    st.download_button("📦 Baixar ZIP com DOCXs", f, file_name="procurações_final.zip", mime="application/zip")
