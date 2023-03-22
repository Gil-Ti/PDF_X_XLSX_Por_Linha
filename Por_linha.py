import PyPDF2
from openpyxl import Workbook
import re # importar a biblioteca re

# Abra o arquivo PDF e crie um objeto de leitura
pdf_file = open('PDF\BM.pdf', 'rb')
pdf_reader = PyPDF2.PdfFileReader(pdf_file)

# Crie um novo arquivo XLSX e uma planilha
workbook = Workbook()
worksheet = workbook.active

# Itere sobre todas as páginas do PDF
for page_num in range(pdf_reader.numPages):
    page = pdf_reader.getPage(page_num)
    text = page.extractText()

    # Busque por ocorrências da palavra "Real"
    for match in re.finditer(r'Real', text):
        # Extraia o texto que vem logo após a palavra "Real"
        data = text[match.end():match.end()+50].strip()

        # Salve os dados na planilha
        worksheet.append([data])

# Salve o arquivo XLSX
workbook.save('Excel/final.xlsx')
