import os
import openpyxl

#Function
def create_excel_with_filenames(folder_path, excel_filename):
    # Excelsheet
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.title = "PDF Names"

    # Cabeçalho
    sheet["A1"] = "Nome Atual"
    sheet["B1"] = "Nome Desejado"

    # Listar todos os ficheiros
    row = 2
    for filename in os.listdir(folder_path):
        if filename.endswith(".pdf"):
            sheet[f"A{row}"] = filename
            sheet[f"B{row}"] = ""  # Coluna vazia para o usuário preencher o nome desejado
            row += 1

    # Save
    workbook.save(excel_filename)
    print(f"Arquivo Excel '{excel_filename}' criado com sucesso.")



#Settings for the Excel Script

folder_path = r"C:\Users\ruben\Documents\teses"

excel_filename = "pdf_names.xlsx"


create_excel_with_filenames(folder_path, excel_filename)


#pip install openpyxl
