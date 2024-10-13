import os
import openpyxl
from PyPDF2 import PdfReader, PdfWriter
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from io import BytesIO

# Function to add a name to top right conner 
def add_text_to_pdf(input_pdf, output_pdf, text):
    reader = PdfReader(input_pdf)
    writer = PdfWriter()

    for i in range(len(reader.pages)):
        page = reader.pages[i]

        # Cria um PDF temporário com o texto
        packet = BytesIO()
        can = canvas.Canvas(packet, pagesize=A4)
        can.setFont("Helvetica", 15)
        can.drawString(450, 820, text)  # Posição do texto no canto superior direito
        can.save()

        packet.seek(0)
        temp_reader = PdfReader(packet)
        temp_page = temp_reader.pages[0]

        # Mescla o texto com a página do PDF
        page.merge_page(temp_page)
        writer.add_page(page)

    # Salva o novo PDF com o nome desejado
    with open(output_pdf, "wb") as output_file:
        writer.write(output_file)

# Função para ler o Excel e processar os PDFs
def process_pdfs_from_excel(folder_path, excel_filename):
    # Abre o arquivo Excel
    workbook = openpyxl.load_workbook(excel_filename)
    sheet = workbook.active

    # Itera pelas linhas do Excel (ignorando o cabeçalho)
    for row in range(2, sheet.max_row + 1):
        current_name = sheet[f"A{row}"].value  # Nome atual do PDF
        new_name = sheet[f"B{row}"].value      # Nome desejado (modificado)

        if current_name and new_name:
            input_pdf = os.path.join(folder_path, current_name)
            output_pdf = os.path.join(folder_path, f"{new_name}.pdf")

            # Verifica se o PDF existe
            if os.path.exists(input_pdf):
                # Adiciona o novo nome no canto superior direito do PDF
                add_text_to_pdf(input_pdf, output_pdf, new_name)
                print(f"Processado: {current_name} -> {new_name}.pdf")
            else:
                print(f"Arquivo não encontrado: {current_name}")

# Caminho da pasta onde os PDFs estão
folder_path = r"C:\Users\ruben\Documents\teses"
# Nome do arquivo Excel
excel_filename = "pdf_names.xlsx"

# Processa os PDFs com base no arquivo Excel
process_pdfs_from_excel(folder_path, excel_filename)
