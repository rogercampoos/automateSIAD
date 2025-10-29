import pandas as pd
import logging
import time
from typing import List

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.options import Options

# Classe de exceção personalizada para erros de automação
class AutomationError(Exception):
    pass

class SIADAutomation:
    def __init__(self, excel_path='C:/Users/p0134255/Documents/Rogério/Backup/Tj/Projetos/Phyton/Automação-SIAD/UNIDADES_DIVIDIDAS.xlsx', log_file='siad_automation.log'):
        """
        Initialize the SIAD Automation with logging and browser configuration
        
        :param excel_path: Path to the Excel file with unit data
        :param log_file: Path for the log file
        """
        # Configure logging
        logging.basicConfig(
            level=logging.INFO,
            format='%(asctime)s - %(levelname)s - %(message)s',
            filename=log_file,
            filemode='w'
        )
        self.logger = logging.getLogger(__name__)

        # Browser configuration for sandbox environment
        chrome_options = Options()
        # chrome_options.add_argument("--headless") # Opcional: Descomente se quiser rodar sem interface gráfica
        chrome_options.add_argument("--no-sandbox")
        chrome_options.add_argument("--disable-dev-shm-usage")
        chrome_options.add_argument("--window-size=1920,1080")
        chrome_options.add_experimental_option("prefs", {"credentials_enable_service": False, "profile.password_manager_enabled": False})
        
        # Para uso local, o ChromeDriverManager gerencia o driver
        try:
            service = Service(ChromeDriverManager().install())
            self.driver = webdriver.Chrome(service=service, options=chrome_options)
        except Exception as e:
            self.logger.error(f"WebDriver initialization failed: {e}")
            raise

        # Configuration parameters
        self.TIMEOUT = 60 # Aumentado para 60s devido à lentidão do sistema
        self.excel_path = excel_path
        self.base_url = 'https://www.siad.mg.gov.br/jasi-frontend/'
        
        # Credentials (A serem substituídas pelo usuário)
        self.usuario = 'x0159191'
        self.senha = 'jl1542'
        
        # Mapeamento de XPaths Robustos
        self.XPATHS = {
            'input_usuario': "//input[@placeholder='Usuário']",
            'input_senha': "//input[@placeholder='Senha']",
            'btn_entrar': "//button[normalize-space(text())='Entrar']",
            'input_digite_unidade': "//input[@placeholder='Digite a Unidade']",
            'btn_selecionar': "//button[normalize-space(text())='Selecionar']",
            'menu_principal_icon': "//div[contains(@class, 'menuicon ibars')]", # Corrigido para div, usando contains para maior robustez
            'menu_item_relatorios': "//span[normalize-space(text())='Relatórios']",
            'menu_item_inventario': "//span[normalize-space(text())='Relatório de inventário de bens']",
            'btn_pesquisar': "//button[normalize-space(text())='Pesquisar']",
            'relatorio_link': "//span[text()='INVENTARIO DE PATRIMONIOS']", # Link/Linha do relatório gerado
            'input_unidade_relatorio': "//span[normalize-space(text())='Unidade emitente:']/following-sibling::input", # Corrigido para usar o rótulo fixo 'Unidade emitente:'
            'btn_solicitar_geracao': "//button[normalize-space(text())='Solicitar geração']",
            'btn_ok': "//button[normalize-space(text())='OK']", # Botão OK em caixas de diálogo
            'menu_usuario_icon': "//i[@class='fas fa-user-circle']", # Ícone do menu de usuário (canto superior direito)
            'menu_item_alterar_unidade': "//span[normalize-space(text())='Alterar Unidade']",
            'btn_alterar': "//button[normalize-space(text())='Alterar']",
        }

    def wait_and_interact(self, xpath_key: str, interaction_type: str = 'click', 
                           input_text: str = None, timeout: int = None):
        """
        Wait for element and perform interaction with robust error handling
        
        :param xpath_key: Key from self.XPATHS
        :param interaction_type: Type of interaction ('click', 'send_keys')
        :param input_text: Text to input if interaction is 'send_keys'
        :param timeout: Custom timeout duration
        :return: WebElement if interaction is successful
        """
        xpath = self.XPATHS.get(xpath_key, xpath_key) # Permite passar a chave ou o XPath direto
        
        try:
            timeout = timeout or self.TIMEOUT
            
            # Se for um clique, espera que o elemento esteja clicável
            if interaction_type == 'click':
                condition = EC.element_to_be_clickable((By.XPATH, xpath))
            # Se for send_keys, espera que o elemento esteja presente no DOM e visível
            elif interaction_type == 'send_keys':
                condition = EC.visibility_of_element_located((By.XPATH, xpath))
            else:
                raise ValueError(f"Tipo de interação desconhecido: {interaction_type}")

            element = WebDriverWait(self.driver, timeout).until(condition)

            if interaction_type == 'click':
                element.click()
                self.logger.info(f"Clicou em: {xpath_key} ({xpath})")
            elif interaction_type == 'send_keys':
                element.clear()
                element.send_keys(input_text)
                self.logger.info(f"Digitou '{input_text}' em: {xpath_key} ({xpath})")
            
            return element

        except Exception as e:
            self.logger.error(f"Erro ao interagir com {xpath_key} ({xpath}): {e}")
            self.driver.save_screenshot(f'erro_{xpath_key}.png')
            raise AutomationError(f"Falha na interação com {xpath_key}. Verifique a screenshot.")

    def login(self):
        """Perform login to SIAD system"""
        self.logger.info("Iniciando login...")
        self.driver.get(self.base_url)
        
        self.wait_and_interact('input_usuario', 'send_keys', self.usuario)
        self.wait_and_interact('input_senha', 'send_keys', self.senha)
        self.wait_and_interact('btn_entrar')
        
        self.logger.info("Login realizado. Aguardando tela de seleção de unidade.")
        time.sleep(2) # Pequena pausa para garantir o carregamento do modal

    def select_unit_initial(self, unit_code: str):
        """Select the unit on the initial modal screen"""
        self.logger.info(f"Selecionando unidade inicial: {unit_code}")
        
        # 1. Digitar a unidade no campo do modal via Selenium (mais confiável para digitação)
        # O elemento é retornado para que possamos enviar as teclas
        input_element = self.wait_and_interact('input_digite_unidade', 'send_keys', unit_code)
        
        # 2. Usar navegação por teclado (TAB) para focar no botão "Selecionar"
        input_element.send_keys(Keys.TAB)
        
        # 3. Forçar o clique no botão "Selecionar" via JS (solução mais robusta para o ZK)
        xpath_button = self.XPATHS['btn_selecionar']
        script_click = f"document.evaluate(\"{xpath_button}\", document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue.click();"
        self.driver.execute_script(script_click)
        self.logger.info("Botão 'Selecionar' clicado via TAB e JS.")
        self.logger.info(f"Unidade inicial {unit_code} selecionada.")

    def generate_inventory_report(self, unit_code: str):
        """Navigate and generate the inventory report for a single unit"""
        self.logger.info(f"Iniciando geração de relatório para unidade: {unit_code}")
        
        # 1. Esperar a página principal carregar (ícone do menu principal)
        self.wait_and_interact('menu_principal_icon')
        
        # 2. Navegar: Menu Principal (ícone) -> Relatórios -> Relatório de inventário de bens
        self.wait_and_interact('menu_item_relatorios')
        self.wait_and_interact('menu_item_relatorios')
        self.wait_and_interact('menu_item_inventario')
        
        self.logger.info("Navegação para a tela de relatório concluída.")
        
        # 2. ESPERA PELA ESTABILIDADE DA TELA: Espera pelo rótulo "Unidade emitente:"
        # Isso garante que a tela de filtro carregou antes de tentar clicar em Pesquisar.
        WebDriverWait(self.driver, self.TIMEOUT).until(
            EC.presence_of_element_located((By.XPATH, "//span[normalize-space(text())='Unidade emitente:']"))
        )
        self.logger.info("Tela de filtro de relatório estabilizada.")
        
        # 3. Clicar em "Pesquisar" (a unidade já deve estar preenchida pelo sistema)
        self.wait_and_interact('btn_pesquisar')
        
        self.logger.info("Pesquisa de inventário realizada.")
        
        # 4. Clicar no relatório gerado (link/linha "INVENTARIO DE PATRIMONIOS")
        self.wait_and_interact('relatorio_link')
        
        self.logger.info("Relatório selecionado.")
        
        # 5. Clicar em "Solicitar geração"
        self.wait_and_interact('btn_solicitar_geracao')
        
        # 6. Clicar em "OK" na caixa de diálogo de confirmação
        self.wait_and_interact('btn_ok')
        
        self.logger.info(f"Solicitação de geração de relatório para {unit_code} concluída.")
        time.sleep(5) # Pausa para garantir que a solicitação foi processada antes de mudar de unidade

    def change_unit_and_loop(self, unit_code: str):
        """Change the current unit to the next one in the loop"""
        self.logger.info(f"Iniciando alteração de unidade para a próxima: {unit_code}")
        
        # 1. Clicar no menu de usuário (ícone superior direito)
        self.wait_and_interact('menu_usuario_icon')
        
        # 2. Clicar em "Alterar Unidade"
        self.wait_and_interact('menu_item_alterar_unidade')
        
        # 3. Clicar em "OK" na caixa de diálogo (para confirmar a alteração)
        self.wait_and_interact('btn_ok')
        
        # 4. Inserir a próxima unidade no campo do modal
        self.wait_and_interact('input_digite_unidade', 'send_keys', unit_code)
        
        # 5. Clicar em "Alterar" (Este botão deve aparecer no modal após digitar a unidade)
        self.wait_and_interact('btn_alterar')
        
        self.logger.info(f"Unidade alterada com sucesso para: {unit_code}")
        time.sleep(2)

    def execute_automation(self):
        """
        Main automation execution method
        Reads Excel, processes report generation for each unit
        """
        try:
            # 1. Ler o arquivo Excel
            # A coluna 'Unidade' é a primeira (índice 0)
            df = pd.read_excel(self.excel_path, sheet_name=0, header=0, dtype=str)
            unit_codes = df.iloc[:, 0].dropna().unique().tolist()
            
            if not unit_codes:
                self.logger.warning("Nenhuma unidade encontrada na primeira coluna do Excel. Encerrando.")
                return

            self.logger.info(f"Total de {len(unit_codes)} unidades para processar.")
            
            # 2. Login
            self.login()
            
            # 3. Loop de Automação
            for i, unit_code in enumerate(unit_codes):
                self.logger.info(f"--- Processando unidade {i+1}/{len(unit_codes)}: {unit_code} ---")
                
                # A primeira unidade usa a função select_unit_initial
                if i == 0:
                    self.select_unit_initial(unit_code)
                # As unidades subsequentes usam a função change_unit_and_loop
                else:
                    self.change_unit_and_loop(unit_code)
                
                # Gerar o relatório para a unidade atual
                self.generate_inventory_report(unit_code)
                
            self.logger.info("Automação concluída com sucesso para todas as unidades.")
        
        except AutomationError as e:
            self.logger.error(f"Automação interrompida devido a um erro de interação: {e}")
        except Exception as e:
            self.logger.error(f"Erro inesperado durante a execução: {e}")
        
        finally:
            # Garantir que o driver feche
            try:
                self.driver.quit()
            except:
                pass
            self.logger.info("WebDriver fechado.")

def main():
    try:
        # ATENÇÃO: Substitua o caminho do Excel se necessário e insira suas credenciais na classe SIADAutomation
        automation = SIADAutomation()
        automation.execute_automation()
    except Exception as e:
        print(f"A automação falhou: {e}")

if __name__ == "__main__":
    main()