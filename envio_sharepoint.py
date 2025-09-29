from office365.sharepoint.client_context import ClientContext
import msal
import os
import logging
import json
from datetime import datetime

class SharePointManager:
    """
    Classe responsável por gerenciar o envio de dados para o SharePoint.
    Gerencia autenticação, cache de tokens e envio de dados.
    """
    def __init__(self, config=None):
        """
        Inicializa o gerenciador do SharePoint com configurações necessárias
        
        Args:
            config: Dicionário com configurações do SharePoint (opcional)
        """
        # Configuração de logging
        logging.basicConfig(level=logging.INFO)
        
        # Carrega configurações
        self.config = config or self._load_config()
        
        # Define caminho do cache
        self.cache_file = os.path.join(os.path.dirname(__file__), "token_cache.json")
        
        # Inicializa cache de token
        self.token_cache = msal.SerializableTokenCache()
        
        # Carrega cache existente
        self._load_token_cache()
        
    def _load_config(self):
        """Carrega configurações do SharePoint"""
        try:
            config_path = os.path.join(os.path.dirname(__file__), 'config.json')
            with open(config_path, 'r') as f:
                return json.load(f)
        except Exception as e:
            logging.error(f"Erro ao carregar configurações: {e}")
            raise

    def _load_token_cache(self):
        """Carrega cache de token existente"""
        try:
            if os.path.exists(self.cache_file):
                with open(self.cache_file, "r") as f:
                    self.token_cache.deserialize(json.load(f))
        except Exception as e:
            logging.warning(f"Erro ao carregar cache: {e}")

    def _save_token_cache(self):
        """Salva cache de token"""
        try:
            if self.token_cache.has_state_changed:
                with open(self.cache_file, "w") as f:
                    json.dump(self.token_cache.serialize(), f)
        except Exception as e:
            logging.warning(f"Erro ao salvar cache: {e}")

    def get_access_token(self):
        """Obtém token de acesso para o SharePoint"""
        try:
            app = msal.ConfidentialClientApplication(
                self.config["client_id"],
                authority=f"https://login.microsoftonline.com/{self.config['tenant_id']}",
                client_credential=self.config["client_secret"],
                token_cache=self.token_cache
            )

            # Tenta obter token do cache
            accounts = app.get_accounts()
            if accounts:
                token_result = app.acquire_token_silent(
                    self.config["scopes"],
                    account=accounts[0]
                )
                if token_result:
                    self._save_token_cache()
                    return token_result['access_token']

            # Obtém novo token
            token_result = app.acquire_token_for_client(scopes=self.config['scopes'])
            if "access_token" not in token_result:
                raise Exception(f"Falha ao obter token: {token_result.get('error_description')}")
                
            self._save_token_cache()
            return token_result['access_token']
            
        except Exception as e:
            logging.error(f"Erro ao obter token de acesso: {e}")
            raise

    def send_to_sharepoint(self, data_dict):
        """
        Envia dados para uma lista do SharePoint
        
        Args:
            data_dict: Dicionário com dados a serem enviados
        """
        try:
            # Obtém token e cria contexto
            access_token = self.get_access_token()
            ctx = ClientContext(self.config["sharepoint_url"]).with_access_token(access_token)

            # Obtém lista do SharePoint
            target_list = ctx.web.lists.get_by_title(self.config["target_list_title"])
            
            # Prepara dados formatados
            formatted_data = self._format_data_for_sharepoint(data_dict)
            
            # Envia dados
            target_list.add_item(formatted_data).execute_query()
            logging.info(f"Dados enviados com sucesso: {formatted_data.get('Title', '')}")
            
            return True
            
        except Exception as e:
            logging.error(f"Erro ao enviar dados para SharePoint: {e}")
            return False

    def _format_data_for_sharepoint(self, data):
        """
        Formata dados para o padrão do SharePoint
        Args:
            data: Dicionário com dados originais
        Returns:
            dict: Dados formatados para o SharePoint
        """
        try:
            # Formatação de datas
            if 'DATA' in data:
                data['DATA'] = datetime.strptime(data['DATA'], '%Y-%m-%d').strftime('%Y-%m-%dT%H:%M:%SZ')
            
            if 'DATA_VENCIMENTO' in data:
                data['DATA_VENCIMENTO'] = datetime.strptime(data['DATA_VENCIMENTO'], '%Y-%m-%d').strftime('%Y-%m-%dT%H:%M:%SZ')
            
            # Garante que Title existe (requerido pelo SharePoint)
            if 'PROCESSO' in data and 'Title' not in data:
                data['Title'] = data['PROCESSO']
                
            return data
            
        except Exception as e:
            logging.error(f"Erro ao formatar dados: {e}")
            return data