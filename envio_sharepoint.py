from office365.runtime.auth.user_credential import UserCredential
from office365.sharepoint.client_context import ClientContext
from office365.sharepoint.files.file import File

import pandas as pd
import os
import logging
import time

logging.basicConfig(level=logging.INFO)

sharepoint_url = "https://anttgov.sharepoint.com/sites/teste/"
target_list_title = "DB de CONTROLE"


username = input("Enter your SharePoint username: ")
password = input("Enter your SharePoint password: ")

credentials = UserCredential(username=username, password=password)
ctx = ClientContext(sharepoint_url).with_credentials(credentials=credentials)

def upload_excel_to_sharepoint(excel_path):
    try:
        df = pd.read_excel(excel_path)
        target_list = ctx.web.lists.get_by_tittle(target_list_title)
        
        for index, row in df.iterrows():
            item_data ={
                'Title': str(row['Title']), # Example columns
                'Column1': str(row['Column1']), # Subistitua pelos nomes reais
                'Column2': str(row['Column2']),
            }
            
            target_list.add_item(item_data).execute_query()
            logging.info(f"linha {index + 1} adicionada à lista do SharePoint.")
            
        logging.info("Upload concluído com sucesso.")
    except Exception as e:
        logging.error(f"Erro ao fazer upload: {e}")
        
if __name__ in "__main__":
    excel_file_path = "caminho/para/seu/arquivo.xlsx"  # Substitua pelo caminho real
    if os.path.exists(excel_file_path):
        upload_excel_to_sharepoint(excel_file_path)
    else:
        logging.error("Arquivo Excel não encontrado.")