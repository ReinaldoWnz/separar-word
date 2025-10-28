import streamlit as st
import zipfile
import tempfile
import os
import shutil
import pandas as pd
from lxml import etree
import copy

st.set_page_config(page_title="Separador DOCX com Formatação Original", layout="centered")

st.title("📄 Separador DOCX — 100% Formatação Preservada")

st.markdown("""
**Instruções:**
1. Envie o arquivo `.docx` contendo todas as procurações.  
2. Envie a planilha `.csv` ou `.xlsx` com as colunas **Credor** e **Número**.  
3. O app localizará cada bloco de texto iniciando com a palavra **“PROCURAÇÃO”** e criará um arquivo separado para cada um.  
4. A formatação original é mantida integralmente.
---
""")

docx_file = st.file_uploader("📎 Envie o arquivo DOCX", type=["docx"])
table_file = st.file_uploader("📊 Envie a planilha (CSV ou XLSX)", type=["csv", "xlsx"])

def split_docx_by_keyword(docx_path, keyword, df, output_dir):
    """Divide o DOCX por palavra-chave mantendo toda a estrutura original."""
    # Cria diretório temporário
    extract_dir = os.path.join(output_dir, "extract")
    os.makedirs(extract_dir, exist_ok=True)

    with zipfile.ZipFile(docx_path, 'r') as zip_ref:
        zip_ref.extractall(extract_dir)

    # Lê o XML principal
    document_xml = os.path.join(extract_dir, "word", "document.xml")
    tree = etree.parse(document_xml)
    root = tree.getroot()
    body = root.find("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}body")

    # Divide o XML em blocos
    keyword_elements = []
    for i, p in enumerate(body.findall("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}p")):
        if keyword.lower() in ''.join(p.itertext()).lower():
            keyword_elements.append(i)

    keyword_elements.append(len(body))  # último bloco

    # Gera documentos separados
    for i in range(len(keyword_elements) - 1):
        start, end = keyword_elements[i], keyword_elements[i + 1]
        new_root = copy.deepcopy(root)
        new_body = new_root.find("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}body")

        # Remove conteúdo antigo e insere só o bloco desejado
        for el in list(new_body):
            new_body.remove(el)
        for el in body[start:end]:
            new_body.append(copy.deepcopy(el))

        # Atualiza XML
        new_tree = etree.ElementTree(new_root)

        # Cria nova pasta
        part_dir = os.path.join(output_dir, f"part_{i+1}")
        shutil.copytree(extract_dir, part_dir)

        # Substitui o XML pelo novo conteúdo
        new_tree.write(os.path.join(part_dir, "word", "document.xml"), xml_declaration=True, encoding='utf-8')

        # Compacta em novo DOCX
        nome1 = str(df.iloc[i, 0]).strip().replace("/", "-")
        nome2 = str(df.iloc[i, 1]).strip().replace("/", "-")
        output_docx = os.path.join(output_dir, f"PROCURAÇÃO - {nome1} - {nome2}.docx")

        with zipfile.ZipFile(output_docx, "w", zipfile.ZIP_DEFLATED) as docx_zip:
            for folder, _, files in os.walk(part_dir):
                for file in files:
                    full_path = os.path.join(folder, file)
                    rel_path = os.path.relpath(full_path, part_dir)
                    docx_zip.write(full_path, rel_path)

def process_files(docx_file, table_file):
    with tempfile.TemporaryDirectory() as tmpdir:
        docx_path = os.path.join(tmpdir, "entrada.docx")
        with open(docx_path, "wb") as f:
            f.write(docx_file.read())

        # Lê planilha
        if table_file.name.endswith(".csv"):
            df = pd.read_csv(table_file)
        else:
            df = pd.read_excel(table_file)

        if len(df.columns) < 2:
            st.error("⚠️ A planilha precisa ter pelo menos duas colunas.")
            st.stop()

        output_dir = os.path.join(tmpdir, "saida")
        os.makedirs(output_dir, exist_ok=True)

        split_docx_by_keyword(docx_path, "PROCURAÇÃO", df, output_dir)

        # Compacta resultados
        zip_path = os.path.join(tmpdir, "procurações_preservadas.zip")
        with zipfile.ZipFile(zip_path, "w") as zipf:
            for file in os.listdir(output_dir):
                if file.endswith(".docx"):
                    zipf.write(os.path.join(output_dir, file), file)

        with open(zip_path, "rb") as f:
            st.success("✅ Arquivos gerados com formatação 100% preservada!")
            st.download_button("📦 Baixar ZIP", f, file_name="procurações_preservadas.zip", mime="application/zip")


if docx_file and table_file:
    if st.button("🚀 Gerar DOCXs Separados (com formatação original)"):
        with st.spinner("Processando documento..."):
            process_files(docx_file, table_file)
