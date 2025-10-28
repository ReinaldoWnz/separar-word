import streamlit as st
import tempfile
import os
from docx import Document
import pandas as pd
import zipfile

st.set_page_config(page_title="Separador e Renomeador de DOCX", layout="centered")

st.title("📄 Separador de DOCX com Renomeação Automática")

st.markdown("""
**Como usar:**
1. Envie o arquivo `.docx` original (com várias procurações, cada uma em uma página ou separada por parágrafos).  
2. Envie a planilha `.csv` ou `.xlsx` com duas colunas: **Credor Original** e **Número Atual**.  
3. Clique em **Gerar DOCXs Separados**.  
4. Baixe o `.zip` com os arquivos já nomeados.
---
""")

docx_file = st.file_uploader("📎 Envie o arquivo DOCX", type=["docx"])
table_file = st.file_uploader("📊 Envie a planilha com nomes (CSV ou XLSX)", type=["csv", "xlsx"])

# --- Função para dividir o DOCX ---
def split_docx(doc_path, num_parts):
    """Divide um DOCX em várias partes (por página simulada ou blocos)."""
    doc = Document(doc_path)
    paragraphs = doc.paragraphs
    total_paragraphs = len(paragraphs)
    # Divide o total de parágrafos de forma uniforme
    chunk_size = total_paragraphs // num_parts
    remainder = total_paragraphs % num_parts

    parts = []
    start = 0
    for i in range(num_parts):
        end = start + chunk_size + (1 if i < remainder else 0)
        sub_doc = Document()
        for p in paragraphs[start:end]:
            sub_doc.add_paragraph(p.text, style=p.style)
        parts.append(sub_doc)
        start = end
    return parts


if docx_file and table_file:
    if st.button("🚀 Gerar DOCXs Separados"):
        with st.spinner("Processando documentos..."):
            with tempfile.TemporaryDirectory() as tmpdir:
                # Salva DOCX temporário
                docx_path = os.path.join(tmpdir, "entrada.docx")
                with open(docx_path, "wb") as f:
                    f.write(docx_file.read())

                # Lê a tabela
                if table_file.name.endswith(".csv"):
                    df = pd.read_csv(table_file)
                else:
                    df = pd.read_excel(table_file)

                if len(df.columns) < 2:
                    st.error("⚠️ A planilha precisa ter pelo menos DUAS colunas.")
                    st.stop()

                num_docs = len(df)

                # Divide o DOCX em partes
                docs = split_docx(docx_path, num_docs)

                # Cria pasta de saída
                output_dir = os.path.join(tmpdir, "saida_docs")
                os.makedirs(output_dir, exist_ok=True)

                # Cria DOCXs renomeados
                for i, sub_doc in enumerate(docs):
                    nome1 = str(df.iloc[i, 0]).strip().replace("/", "-")
                    nome2 = str(df.iloc[i, 1]).strip().replace("/", "-")
                    file_name = f"PROCURAÇÃO - {nome1} - {nome2}.docx"
                    out_path = os.path.join(output_dir, file_name)
                    sub_doc.save(out_path)

                # Compacta em ZIP
                zip_path = os.path.join(tmpdir, "procurações.zip")
                with zipfile.ZipFile(zip_path, "w") as zipf:
                    for file in os.listdir(output_dir):
                        zipf.write(os.path.join(output_dir, file), file)

                with open(zip_path, "rb") as f:
                    st.success("✅ Arquivos DOCX gerados com sucesso!")
                    st.download_button(
                        "📦 Baixar ZIP com as procurações",
                        f,
                        file_name="procurações_separadas.zip",
                        mime="application/zip",
                    )
