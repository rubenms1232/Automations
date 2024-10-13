import os
import openpyxl
from PyPDF2 import PdfReader, PdfWriter
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from io import BytesIO

def add_text_to_pdf(input_pdf, output_pdf, text, apply_to_all_pages=True):
    reader = PdfReader(input_pdf)
    writer = PdfWriter()

    for i in range(len(reader.pages)):
        page = reader.pages[i]

        if not apply_to_all_pages and i > 0:
            writer.add_page(page)
            continue

        packet = BytesIO()
        can = canvas.Canvas(packet, pagesize=A4)
        can.setFont("Helvetica", 15)
        can.drawString(450, 820, text)
        can.save()

        packet.seek(0)
        temp_reader = PdfReader(packet)
        temp_page = temp_reader.pages[0]

        page.merge_page(temp_page)
        writer.add_page(page)

    with open(output_pdf, "wb") as output_file:
        writer.write(output_file)

def process_pdfs_from_excel(folder_path, excel_filename, output_folder, apply_to_all_pages):
    workbook = openpyxl.load_workbook(excel_filename)
    sheet = workbook.active

    for row in range(2, sheet.max_row + 1):
        current_name = sheet[f"A{row}"].value
        new_name = sheet[f"B{row}"].value

        if current_name and new_name:
            input_pdf = os.path.join(folder_path, current_name)
            output_pdf = os.path.join(output_folder, f"{new_name}.pdf")

            if os.path.exists(input_pdf):
                add_text_to_pdf(input_pdf, output_pdf, new_name, apply_to_all_pages)
                print(f"Processado: {current_name} -> {new_name}.pdf")
            else:
                print(f"Arquivo não encontrado: {current_name}")

folder_path = r"C:\Users\ruben\Documents\teses"
output_folder = r"C:\Users\ruben\Documents\teses\output"
excel_filename = "exceloutput/pdf_names.xlsx"

if not os.path.exists(output_folder):
    os.makedirs(output_folder)

user_choice = input("Deseja aplicar o nome em todas as páginas ou apenas na primeira? (digite 'todas' ou 'primeira'): ").strip().lower()

apply_to_all_pages = True if user_choice == 'todas' else False

process_pdfs_from_excel(folder_path, excel_filename, output_folder, apply_to_all_pages)
