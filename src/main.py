import pandas as pd
from docx import Document
from docx.shared import Inches
import os
from docx2pdf import convert

# Pastas
FOTOS_DIR = 'assets' 
RELATORIOS_DIR = 'reports'

# Cria pasta de relatórios se não existir
os.makedirs(RELATORIOS_DIR, exist_ok=True)

# Lê a planilha
planilha = pd.read_excel('planilha_fiscalizacao.xlsx')

# Cria o relatório 
for index, row in planilha.iterrows():
    doc = Document()
    doc.add_heading('Relatório de Fiscalização', 0)
    doc.add_paragraph(f"Data: {row['Data']}")
    doc.add_paragraph(f"Local: {row['Local']}")
    doc.add_paragraph(f"Fiscal: {row['Pessoal Responsável']}")
    doc.add_paragraph(f"Descrição: {row['Não conformidade']}")

    # Fotos separadas por ponto e vírgula na planilha
    fotos = row['Fotos'].split(';')  

    for foto in fotos:
        foto_path = os.path.join(FOTOS_DIR, foto.strip())
        if os.path.exists(foto_path):
            doc.add_paragraph(f"Foto: {foto}")
            doc.add_picture(foto_path, width=Inches(4))
        else:
            doc.add_paragraph(f"Foto não encontrada: {foto}")

    print(doc)

# Nome do arquivo
nome_relatorio = f"relatorio_{index+1}.docx"
caminho_docx = os.path.join(RELATORIOS_DIR, nome_relatorio)

# Salva o Word
doc.save(caminho_docx)

# Converte para PDF
convert(caminho_docx, caminho_docx.replace('.docx', '.pdf'))


print("Relatório gerado com sucesso!")