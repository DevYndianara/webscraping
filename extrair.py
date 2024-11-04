# Instalar:
# pip install requests
# pip install requests beautifulsoup4 pyth-docx

# O script é para efetuar autenticação do html e especificar o setor do user, posteriormente, acessar a url que deseja criar o arquivo em .docx (word)

import requests
from bs4 import BeautifulSoup
from docx import Document
import logging
import json
from requests.adapters import HTTPAdapter
from requests.packages.urllib3.util.retry import Retry
import time

class WebTextExtractor:
    def __init__(self, 
                 login_url=None, 
                 username=None, 
                 password=None,
                 setor=None,
                 login_data=None,
                 auth_method='post'):
        """
        Inicializa o extrator com configurações de autenticação e setor
        """
        # Configuração de logging
        logging.basicConfig(
            level=logging.DEBUG,
            format='%(asctime)s - %(levelname)s: %(message)s'
        )
        self.logger = logging.getLogger(__name__)
        
        # Configuração da sessão
        self.session = requests.Session()
        
        # Headers padrão
        self.session.headers.update({
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36',
            'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,*/*;q=0.8',
            'Accept-Language': 'pt-BR,pt;q=0.8,en-US;q=0.5,en;q=0.3',
            'Content-Type': 'application/x-www-form-urlencoded',
            'X-Requested-With': 'XMLHttpRequest'
        })
        
        # Configurações de autenticação
        self.login_url = login_url
        self.username = username
        self.password = password
        self.setor = setor
        self.auth_method = auth_method
        self.login_data = login_data or {}

    def obter_setores_disponiveis(self, url_setores):
        """
        Obtém lista de setores disponíveis na página
        """
        try:
            response = self.session.get(url_setores)
            soup = BeautifulSoup(response.text, 'html.parser')
            
            # Procura o select/div de setores
            setor_select = soup.find('select', {'name': 'setor'})  # Ajuste o seletor conforme necessário
            if setor_select:
                setores = {}
                for option in setor_select.find_all('option'):
                    setores[option.text.strip()] = option['value']
                return setores
            
            # Se não encontrar select, procura por div
            setor_divs = soup.find_all('div', {'class': 'setor-item'})  # Ajuste a classe conforme necessário
            if setor_divs:
                setores = {}
                for div in setor_divs:
                    setor_id = div.get('data-setor-id')  # Ajuste o atributo conforme necessário
                    setor_nome = div.text.strip()
                    setores[setor_nome] = setor_id
                return setores
            
            return None
        except Exception as e:
            self.logger.error(f"Erro ao obter setores: {e}")
            return None

    def selecionar_setor(self, setor_id):
        """
        Seleciona um setor específico após o login
        """
        try:
            # URL para seleção de setor (ajuste conforme necessário)
            setor_url = f"{self.login_url}/selecionar-setor"  # Ajuste a URL conforme necessário
            
            dados_setor = {
                'setor_id': setor_id,
                # Adicione outros campos necessários
            }
            
            response = self.session.post(setor_url, data=dados_setor)
            
            if response.ok:
                self.logger.info(f"Setor selecionado com sucesso: {setor_id}")
                return True
            else:
                self.logger.error(f"Erro ao selecionar setor. Status: {response.status_code}")
                return False
                
        except Exception as e:
            self.logger.error(f"Erro ao selecionar setor: {e}")
            return False

    def autenticar(self):
        """
        Realiza autenticação e seleção de setor
        """
        if not self.login_url:
            self.logger.warning("Nenhuma URL de login fornecida.")
            return True
        
        try:
            # Prepara dados de login
            if not self.login_data:
                self.login_data = {
                    'username': self.username,
                    'password': self.password,
                    'setor': self.setor
                }
            
            # Faz o login
            if self.auth_method == 'post':
                response = self.session.post(self.login_url, data=self.login_data)
            elif self.auth_method == 'json':
                response = self.session.post(self.login_url, json=self.login_data)
            else:
                response = self.session.get(self.login_url, params=self.login_data)
            
            # Log da resposta para debug
            self.logger.debug(f"Status do login: {response.status_code}")
            self.logger.debug(f"Resposta: {response.text[:500]}...")
            
            # Verifica se login foi bem sucedido
            if response.ok:
                self.logger.info("Login realizado com sucesso")
                
                # Se tiver setor específico, seleciona
                if self.setor:
                    return self.selecionar_setor(self.setor)
                return True
            else:
                self.logger.error(f"Falha no login. Status: {response.status_code}")
                return False
                
        except Exception as e:
            self.logger.error(f"Erro durante autenticação: {e}")
            return False

    def extrair_texto_web(self, url):
        """
        Extrai texto da página web após autenticação e seleção de setor
        """
        try:
            # Aguarda um momento para garantir que o setor foi selecionado
            time.sleep(2)
            
            # Faz a requisição
            response = self.session.get(url, timeout=10)
            response.raise_for_status()
            
            # Parse do HTML
            soup = BeautifulSoup(response.text, 'html.parser')
            
            # Remove scripts e styles
            for tag in soup(["script", "style"]):
                tag.decompose()
            
            # Extrai texto
            texto = soup.get_text(separator='\n', strip=True)
            
            self.logger.info(f"Texto extraído com sucesso da URL: {url}")
            return texto
            
        except Exception as e:
            self.logger.error(f"Erro ao extrair texto: {e}")
            return None

    def salvar_documento(self, texto, caminho_arquivo):
        """
        Salva o texto extraído em documento
        """
        if not texto:
            self.logger.warning("Sem texto para salvar")
            return
        
        try:
            doc = Document()
            doc.add_paragraph(texto)
            doc.save(caminho_arquivo)
            self.logger.info(f"Documento salvo em: {caminho_arquivo}")
            
        except Exception as e:
            self.logger.error(f"Erro ao salvar documento: {e}")

def main():
    # Dados de autenticação
    login_url = input("URL de login: ").strip()
    username = input("Usuário: ").strip()
    password = input("Senha: ").strip()
    
    # Inicializa extrator
    extrator = WebTextExtractor(
        login_url=login_url,
        username=username,
        password=password
    )
    
    # Obtém setores disponíveis
    print("\nObtendo setores disponíveis...")
    setores = extrator.obter_setores_disponiveis(login_url)
    
    if setores:
        print("\nSetores disponíveis:")
        for nome, id_setor in setores.items():
            print(f"- {nome} (ID: {id_setor})")
        
        setor_escolhido = input("\nDigite o nome do setor desejado: ").strip()
        if setor_escolhido in setores:
            extrator.setor = setores[setor_escolhido]
        else:
            print("Setor não encontrado!")
            return
    
    # Autenticação
    if not extrator.autenticar():
        print("Falha na autenticação!")
        return
    
    # URL para extração
    url_extracao = input("\nURL da página para extração: ").strip()
    arquivo_saida = input("Nome do arquivo de saída (.docx): ").strip()
    
    # Extrai e salva
    texto = extrator.extrair_texto_web(url_extracao)
    if texto:
        extrator.salvar_documento(texto, arquivo_saida)
        print("Extração concluída com sucesso!")
    else:
        print("Falha na extração do texto!")

if __name__ == '__main__':
    main()