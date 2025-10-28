import streamlit as st
import tempfile
import os
import pandas as pd
from PyPDF2 import PdfReader, PdfWriter
from docx import Document
import pypandoc
import zipfile
from docx2pdf import convert

st.set_page_config(page_title="Conversor e Separador de PDFs", layout="centered")

st.title("📄 DOCX → PDF → Separador Automático")

st.markdown("""
**Passos:**
1. Envie o arquivo `.docx` principal.  
2. Envie a planilha `.csv` ou `.xlsx` com duas colunas.  
3. Clique em **Gerar PDFs Separados**.  
4. Baixe o `.zip` com os PDFs renomeados.  
---
""")

docx_file = st.file_uploader("📎 Envie o arquivo DOCX", type=["docx"])
table_file = st.file_uploader("📊 Envie a tabela com nomes (CSV ou XLSX)", type=["csv", "xlsx"])


def convert_docx_to_pdf(docx_path, pdf_path):
    """Converte DOCX em PDF usando docx2pdf (mantém formatação)."""
    try:
        convert(docx_path, pdf_path)
    except Exception:
        # fallback pypandoc
        pypandoc.convert_file(docx_path, 'pdf', outputfile=pdf_path)


if docx_file and table_file:
    if st.button("🚀 Gerar PDFs Separados"):
        with st.spinner("Convertendo e separando páginas..."):
            with tempfile.TemporaryDirectory() as tmpdir:
                # Salva o DOCX
                docx_path = os.path.join(tmpdir, "entrada.docx")
                with open(docx_path, "wb") as f:
                    f.write(docx_file.read())

                # Caminho de saída do PDF
                pdf_path = os.path.join(tmpdir, "saida.pdf")
                convert_docx_to_pdf(docx_path, pdf_path)

                # Lê a tabela
                if table_file.name.endswith(".csv"):
                    df = pd.read_csv(table_file)
                else:
                    df = pd.read_excel(table_file)

                # Remove cabeçalho caso exista
                if "Unnamed" in df.columns[0]:
                    df = df.iloc[1:].reset_index(drop=True)

                if len(df.columns) < 2:
                    st.error("⚠️ A planilha precisa ter DUAS colunas (nome e número).")
                    st.stop()

                # Lê o PDF
                reader = PdfReader(pdf_path)
                num_pages = len(reader.pages)
                num_rows = len(df)

                st.info(f"📄 O PDF tem {num_pages} páginas.")
                st.info(f"📊 A planilha tem {num_rows} linhas.")

                limite = min(num_pages, num_rows)

                output_dir = os.path.join(tmpdir, "pdfs_separados")
                os.makedirs(output_dir, exist_ok=True)

                # Divide e renomeia
                for i in range(limite):
                    writer = PdfWriter()
                    writer.add_page(reader.pages[i])

                    nome1 = str(df.iloc[i, 0]).strip().replace("/", "-")
                    nome2 = str(df.iloc[i, 1]).strip().replace("/", "-")
                    new_name = f"PROCURAÇÃO - {nome1} - {nome2}.pdf"

                    out_path = os.path.join(output_dir, new_name)
                    with open(out_path, "wb") as out_file:
                        writer.write(out_file)

                # Compacta em ZIP
                zip_path = os.path.join(tmpdir, "procurações.zip")
                with zipfile.ZipFile(zip_path, "w") as zipf:
                    for file in os.listdir(output_dir):
                        zipf.write(os.path.join(output_dir, file), file)

                with open(zip_path, "rb") as f:
                    st.success("✅ PDFs separados e renomeados com sucesso!")
                    st.download_button(
                        "📦 Baixar ZIP",
                        f,
                        file_name="procurações.zip",
                        mime="application/zip"
                    )
