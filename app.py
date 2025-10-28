import streamlit as st
import tempfile
import os
import pandas as pd
import zipfile
import subprocess
from PyPDF2 import PdfReader, PdfWriter
from pdf2docx import Converter

st.set_page_config(page_title="DOCX → PDF → Separar → DOCX", layout="centered")

st.title("📄 Conversor e Separador de DOCX com LibreOffice")

st.markdown("""
### 🧠 Como funciona:
1️⃣ Envie o arquivo `.docx` completo  
2️⃣ Envie a planilha `.csv` ou `.xlsx` com duas colunas (Nome e Número)  
3️⃣ O app converte o DOCX para PDF **mantendo toda a formatação**  
4️⃣ Separa o PDF página por página  
5️⃣ Renomeia conforme a planilha  
6️⃣ Reconverte cada página para `.docx`  
7️⃣ Gera um ZIP com todos os arquivos prontos  
---
""")

# ============ Upload ============
docx_file = st.file_uploader("📎 Envie o arquivo DOCX", type=["docx"])
table_file = st.file_uploader("📊 Envie a planilha (CSV ou XLSX)", type=["csv", "xlsx"])


# ============ Função principal ============
def convert_docx_to_pdf(docx_path, output_dir):
    """Converte DOCX em PDF usando LibreOffice headless (mantém formatação)"""
    subprocess.run(["apt-get", "update"], check=True)
    subprocess.run(["apt-get", "install", "-y", "libreoffice"], check=True)
    subprocess.run(
        [
            "libreoffice",
            "--headless",
            "--convert-to",
            "pdf",
            "--outdir",
            output_dir,
            docx_path,
        ],
        check=True,
    )
    pdf_name = os.path.splitext(os.path.basename(docx_path))[0] + ".pdf"
    return os.path.join(output_dir, pdf_name)


if docx_file and table_file:
    if st.button("🚀 Gerar DOCXs Separados"):
        with st.spinner("Processando, aguarde..."):
            with tempfile.TemporaryDirectory() as tmpdir:
                # Salva DOCX
                docx_path = os.path.join(tmpdir, "entrada.docx")
                with open(docx_path, "wb") as f:
                    f.write(docx_file.read())

                # Converte DOCX → PDF com LibreOffice
                pdf_path = convert_docx_to_pdf(docx_path, tmpdir)

                # Lê a planilha
                if table_file.name.endswith(".csv"):
                    df = pd.read_csv(table_file)
                else:
                    df = pd.read_excel(table_file)

                # Remove possíveis cabeçalhos extras
                df = df.iloc[1:].reset_index(drop=True) if len(df.columns) > 2 or "Unnamed" in str(df.columns[0]) else df

                if len(df.columns) < 2:
                    st.error("⚠️ A planilha precisa ter DUAS colunas (Nome e Número).")
                    st.stop()

                reader = PdfReader(pdf_path)
                num_pages = len(reader.pages)
                num_linhas = len(df)
                limite = min(num_pages, num_linhas)

                st.info(f"📄 PDF gerado com {num_pages} páginas. Planilha contém {num_linhas} linhas.")
                if num_pages != num_linhas:
                    st.warning("⚠️ Quantidade diferente! Gerando até o número menor entre páginas e linhas.")

                # Cria diretórios
                pdf_dir = os.path.join(tmpdir, "pdfs_separados")
                docx_dir = os.path.join(tmpdir, "docxs_final")
                os.makedirs(pdf_dir, exist_ok=True)
                os.makedirs(docx_dir, exist_ok=True)

                # Divide o PDF página por página
                for i in range(limite):
                    writer = PdfWriter()
                    writer.add_page(reader.pages[i])

                    nome1 = str(df.iloc[i, 0]).strip().replace("/", "-")
                    nome2 = str(df.iloc[i, 1]).strip().replace("/", "-")
                    base_name = f"PROCURAÇÃO - {nome1} - {nome2}"

                    pdf_output = os.path.join(pdf_dir, f"{base_name}.pdf")
                    with open(pdf_output, "wb") as f_out:
                        writer.write(f_out)

                # Converte cada PDF para DOCX
                for file in os.listdir(pdf_dir):
                    if file.endswith(".pdf"):
                        pdf_input = os.path.join(pdf_dir, file)
                        docx_output = os.path.join(docx_dir, file.replace(".pdf", ".docx"))
                        cv = Converter(pdf_input)
                        cv.convert(docx_output)
                        cv.close()

                # Compacta em ZIP final
                zip_path = os.path.join(tmpdir, "procurações_final.zip")
                with zipfile.ZipFile(zip_path, "w") as zipf:
                    for file in os.listdir(docx_dir):
                        zipf.write(os.path.join(docx_dir, file), file)

                # Download
                with open(zip_path, "rb") as f:
                    st.success("✅ Arquivos gerados com sucesso!")
                    st.download_button(
                        "📦 Baixar ZIP com DOCXs separados",
                        f,
                        file_name="procurações_final.zip",
                        mime="application/zip"
                    )
