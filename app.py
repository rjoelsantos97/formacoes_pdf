import streamlit as st
import pandas as pd
from PyPDF2 import PdfReader, PdfWriter
import re
from difflib import SequenceMatcher
import os

# Função para encontrar o nome mais próximo no mapa usando SequenceMatcher para maior precisão
def find_best_match_sequence(name, valid_employee_names):
    best_match = None
    highest_ratio = 0
    for employee_name in valid_employee_names:
        ratio = SequenceMatcher(None, name, employee_name).ratio()
        if ratio > highest_ratio:
            highest_ratio = ratio
            best_match = employee_name
    return best_match if highest_ratio > 0.6 else None  # Usar cutoff de 0.6 para considerar correspondências aproximadas

# Função para salvar cada página como um PDF separado
def save_certificate(pdf_reader, page_number, output_path):
    pdf_writer = PdfWriter()
    pdf_writer.add_page(pdf_reader.pages[page_number])
    
    with open(output_path, 'wb') as output_pdf:
        pdf_writer.write(output_pdf)

# Configuração da página do Streamlit
st.title("Divisão de Certificados PDF por Funcionário")

# Upload do arquivo Excel com o mapa cliente_funcionario
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
    output_dir = 'certificados_divididos/'
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)

    # Processar cada PDF carregado
    for pdf_file in pdf_files:
        reader = PdfReader(pdf_file)

        # Loop para extrair nomes e dividir os certificados
        certificates = {}
        for i in range(len(reader.pages)):
            text = reader.pages[i].extract_text().strip()
            name_pattern = re.compile(r"Certifica-se que (.+)")
            match = name_pattern.search(text)

            if match:
                employee_name = match.group(1).strip()
                best_match_name = find_best_match_sequence(employee_name, valid_employee_names)
                if best_match_name:
                    client_name = mapa_df[mapa_df['Formando'] == best_match_name]['Cliente'].values[0].strip()
                    date = "25-06-2024"  # Pode-se ajustar para extrair dinamicamente, se necessário
                    file_name = f"{client_name}_{best_match_name}_{date}.pdf".replace(" ", "_")
                    certificates[file_name] = i

        # Salvar os certificados divididos
        for file_name, page_number in certificates.items():
            save_certificate(reader, page_number, output_dir + file_name)

    # Listar os arquivos gerados para download
    st.success(f"Processo concluído! Os certificados foram salvos na pasta '{output_dir}'.")

    # Exibir links para download dos arquivos gerados
    for file_name in os.listdir(output_dir):
        st.markdown(f"[{file_name}](./{output_dir}{file_name})")

