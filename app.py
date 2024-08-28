import streamlit as st
import pandas as pd
from PyPDF2 import PdfReader, PdfWriter
import re
import os
import zipfile
from io import BytesIO
from difflib import SequenceMatcher
from PyPDF2.errors import PdfReadError

# Função para salvar cada página como um PDF separado
def save_certificate(pdf_reader, page_number, output_path):
    pdf_writer = PdfWriter()
    pdf_writer.add_page(pdf_reader.pages[page_number])
    
    with open(output_path, 'wb') as output_pdf:
        pdf_writer.write(output_pdf)

# Função para encontrar o nome mais próximo no mapa usando SequenceMatcher para maior precisão
def find_best_match_sequence(name, valid_employee_names):
    best_match = None
    highest_ratio = 0
    for employee_name in valid_employee_names:
        ratio = SequenceMatcher(None, name, employee_name).ratio()
        if ratio > highest_ratio:
            highest_ratio = ratio
            best_match = employee_name
    return best_match if highest_ratio > 0.6 else None

# Configuração da página do Streamlit
st.title("Divisão de Certificados PDF por Funcionário")

# Upload do arquivo Excel com o mapa cliente_funcionário
mapa_file = st.file_uploader("Carregar o arquivo Excel com o Mapa Cliente-Funcionário", type="xlsx")

# Upload de múltiplos arquivos PDF
pdf_files = st.file_uploader("Carregar os arquivos PDF dos Certificados", type="pdf", accept_multiple_files=True)

# Processo após carregar os arquivos
if mapa_file and pdf_files:
    # Carregar o mapa Excel
    mapa_df = pd.read_excel(mapa_file)
    mapa_df['Formando'] = mapa_df['Formando'].astype(str).str.strip()
    mapa_df['Cliente'] = mapa_df['Cliente'].astype(str).fillna('').str.strip()
    valid_employee_names = mapa_df['Formando'].dropna().unique()

    # Diretório temporário para salvar os PDFs divididos
    temp_dir = 'temp_certificates/'
    if not os.path.exists(temp_dir):
        os.makedirs(temp_dir)

    certificates = {}

    # Processar cada PDF carregado
    for pdf_file in pdf_files:
        try:
            reader = PdfReader(pdf_file)
            st.write(f"Processando arquivo: {pdf_file.name}")

            # Loop para extrair nomes e dividir os certificados
            for i in range(len(reader.pages)):
                text = reader.pages[i].extract_text().strip()
                # Ajustar a regex para capturar nomes até encontrar "natural", "nascido", ou outras palavras indicativas
                name_pattern = re.compile(r"Certifica-se que ([A-ZÀ-ÖØ-öø-ÿa-zà-öø-ÿ\s]+)")
                match = name_pattern.search(text)

                if match:
                    # Capturar apenas o nome do funcionário e ignorar partes extras
                    employee_name = match.group(1).split('natural')[0].split('nascido')[0].strip()
                    st.write(f"Nome encontrado: {employee_name}")
                    best_match_name = find_best_match_sequence(employee_name, valid_employee_names)
                    if best_match_name:
                        client_name = mapa_df[mapa_df['Formando'] == best_match_name]['Cliente'].values[0].strip()
                        date = "25-06-2024"  # Data fixa para este exemplo
                        # Adicionando um índice ao nome do arquivo para diferenciar múltiplos certificados
                        index = len(certificates.get(client_name, [])) + 1
                        file_name = f"{client_name}_{best_match_name}_{date}_{index}.pdf".replace(" ", "_")
                        output_path = os.path.join(temp_dir, file_name)

                        save_certificate(reader, i, output_path)
                        if client_name not in certificates:
                            certificates[client_name] = []
                        certificates[client_name].append(output_path)
                        st.write(f"Arquivo adicionado para {client_name}: {output_path}")
                    else:
                        st.write(f"Nome não encontrado no mapa: {employee_name}")

        except PdfReadError:
            st.error(f"O arquivo {pdf_file.name} está criptografado e não pode ser processado.")

    # Criar o arquivo ZIP com pastas para cada cliente
    zip_buffer = BytesIO()
    with zipfile.ZipFile(zip_buffer, "w") as zf:
        for client_name, files in certificates.items():
            for file_path in files:
                # Adicionar cada arquivo PDF à sua respectiva pasta no ZIP
                zf.write(file_path, os.path.join(client_name, os.path.basename(file_path)))

    zip_buffer.seek(0)

    # Fornecer o arquivo ZIP para download
    st.download_button(
        label="Baixar Todos em ZIP",
        data=zip_buffer,
        file_name="certificados.zip",
        mime="application/zip"
    )
