import pandas as pd
import os
import shutil
import subprocess
from openpyxl import Workbook
from dotenv import load_dotenv

load_dotenv()
BASE = os.getenv('ONEDRIVE_BASE')

folder_paths = [
    os.path.join(BASE, 'Remanejamentos'),
    os.path.join(BASE, 'Remanejamentos (SEGOV_DCAF_FONTE 95_ FONTE 80)'),
]
output_file      = os.path.join(BASE, 'Robo', 'copia.xlsx')
processed_folder = os.path.join(BASE, 'Realizados')


def create_output_file_if_not_exists(output_file):
    if not os.path.exists(output_file):
        wb = Workbook()
        ws = wb.active
        ws.title = "PREENCHER AQUI"
        wb.save(output_file)


def combine_excel_files(folder_paths, output_file, processed_folder):
    if not os.path.exists(processed_folder):
        os.makedirs(processed_folder)

    create_output_file_if_not_exists(output_file)

    dataframes = []

    for folder_path in folder_paths:
        files_in_folder = os.listdir(folder_path)
        xlsm_files = [f for f in files_in_folder if f.endswith('.xlsm')]

        for filename in xlsm_files:
            file_path = os.path.join(folder_path, filename)
            try:
                df = pd.read_excel(file_path, sheet_name='PREENCHER AQUI')
                dataframes.append(df)
                print(f"Arquivo {filename} lido com sucesso.")
                shutil.move(file_path, os.path.join(processed_folder, filename))
            except Exception as e:
                print(f"Erro ao ler o arquivo {filename}: {e}")

    if not dataframes:
        print("Nenhum dado foi lido das planilhas. Verifique se os arquivos contêm a aba 'PREENCHER AQUI'.")
        return

    combined_df = pd.concat(dataframes, ignore_index=True)

    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        combined_df.to_excel(writer, sheet_name='PREENCHER AQUI', index=False)

    print("Dados consolidados e salvos com sucesso em um novo arquivo .xlsx.")


combine_excel_files(folder_paths=folder_paths, output_file=output_file, processed_folder=processed_folder)
