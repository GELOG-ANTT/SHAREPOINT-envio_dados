from office365.sharepoint.client_context import ClientContext

import msal
import os
import logging
import json

from config import SHAREPOINT_CONFIG

class SharePointManager:
    def __init__(self):
        logging.basicConfig(level=logging.INFO)
        self.config = SHAREPOINT_CONFIG
        self.cache_file = os.path.join(os.path.dirname(__file__), "token_cache.json")
        
    def load_cache(self):
        if os.path.exists(self.cache_file):
            with open(self.cache_file, "r") as f:
                return json.load(f)
        return None

    def save_cache(self, cache):
        with open(self.cache_file, "w") as f:
            json.dump(cache, f)

    def get_access_token(self):
        try:
            cache = msal.SerializableTokenCache()
            cached_data = self.load_cache()
    
            if cached_data:
                cache.deserialize(cached_data)

            app = msal.ConfidentialClientApplication(
                self.config["client_id"],
                authority=f"https://login.microsoftonline.com/{self.config['tenant_id']}",
                token_cache=cache
            )

            accounts = app.get_accounts()
            if accounts:
                token_result = app.acquire_token_silent(
                    self.config["scopes"],
                    account=accounts[0]
                )
                if token_result:
                    if cache.has_state_changed:
                        self.save_cache(cache.serialize())
                    return token_result['access_token']

            token_result = app.acquire_token_for_client(scopes=self.config['scopes'])
            if "access_token" not in token_result:
                raise Exception(
                    f"Falha ao obter token: {token_result.get('error_description')}"
                )
                
            if cache.has_state_changed:
                self.save_cache(cache.serialize())
            
            return token_result['access_token']
        
        except Exception as e:
            logging.error(f"Erro ao obter token de acesso: {e}")
            raise

    def send_data_to_sharepoint(self, data_dict):
        try:
            access_token = self.get_access_token()
            ctx = ClientContext(self.config["sharepoint_url"]).with_access_token(access_token)

            target_list = ctx.web.lists.get_by_title(self.config["target_list_title"])
            ctx.load(target_list)
            ctx.execute_query()
            
            target_list.add_item(data_dict).execute_query()
            logging.info("Dados enviados com sucesso para o SharePoint.")
        except Exception as e:
            logging.error(f"Erro ao enviar dados para o SharePoint: {e}")
            raise