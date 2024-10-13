import os
import openpyxl

# Função
def create_excel_with_filenames(folder_path, excel_filename):
    # Criar planilha Excel
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
            # Colocar o nome atual
            sheet[f"A{row}"] = filename
            
            # Tentar separar o nome em duas partes
            split_name = filename.split()

            # Verifica se há pelo menos duas partes no nome
            if len(split_name) >= 2:
                new_name = split_name[0] + " " + split_name[1]
            else:
                # Se não houver duas partes, usa o nome original sem modificações
                new_name = filename

            sheet[f"B{row}"] = new_name  # Nome desejado

            row += 1

    # Salvar arquivo Excel
    workbook.save(excel_filename)
    print(f"Arquivo Excel '{excel_filename}' criado com sucesso.")

# Configurações para o caminho da pasta e arquivo Excel
folder_path = r"C:\Users\ruben\Documents\teses"
excel_filename = "exceloutput/pdf_names.xlsx"

# Chamar função para criar Excel
create_excel_with_filenames(folder_path, excel_filename)
