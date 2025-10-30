"""
Refatoração solicitada (opção 1 - modo direto/preferencial)

Principais mudanças:
- Antes de abrir o menu de usuário, o script tenta detectar se o campo modal
  "Digite a Unidade" já está presente e editável. Se estiver, ele cola/insere a
  próxima unidade diretamente nesse campo e aciona "Selecionar" (fluxo curto).
- Apenas se o campo NÃO estiver disponível, o script executa o fluxo antigo:
  abrir menu_usuario_icon -> Alterar Unidade -> preencher modal -> Alterar.
- _click foi tornado tolerante: faz retries, scrollIntoView, JS fallback e
  retorna False em falhas não-críticas (raise_on_fail controlável).
- Quando ocorre modal "NAO EXISTE PERFIL AUTORIZADO" o código registra a unidade
  e limpa/fecha o modal, mas agora PULA para a próxima unidade (não encerra).
- Parada imediata apenas para erros fatais (AutomationFatalError) nos pontos
  realmente críticos (por exemplo, falha ao acionar botão 'Alterar' após abrir modal).
- Logs mais claros para seguir o fluxo de tentativa direta vs. menu.

Observação: salve este arquivo como .py sem cabeçalhos extras. Recomendo instalar
pyperclip (pip install pyperclip) para melhor confiabilidade do Ctrl+V.
"""
import logging
import time
from typing import Optional

import pandas as pd
import openpyxl

# clipboard helper
try:
    import pyperclip
except Exception:
    pyperclip = None

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.options import Options
from selenium.common.exceptions import WebDriverException, TimeoutException

# Exceção que marca erros fatais que devem encerrar execução
class AutomationFatalError(Exception):
    pass

class SIADAutomation:
    def __init__(self,
                 excel_path: str = 'C:/Users/p0134255/Documents/Rogério/Backup/Tj/Projetos/Phyton/Automação-SIAD/UNIDADES_DIVIDIDAS.xlsx',
                 log_file: str = 'siad_automation.log'):
        logging.basicConfig(
            level=logging.INFO,
            format='%(asctime)s - %(levelname)s - %(message)s',
            filename=log_file,
            filemode='w'
        )
        self.logger = logging.getLogger(__name__)

        chrome_options = Options()
        # chrome_options.add_argument("--headless")  # descomente se desejar rodar sem UI
        chrome_options.add_argument("--no-sandbox")
        chrome_options.add_argument("--disable-dev-shm-usage")
        chrome_options.add_argument("--window-size=1920,1080")
        chrome_options.add_experimental_option("prefs", {"credentials_enable_service": False, "profile.password_manager_enabled": False})

        try:
            service = Service(ChromeDriverManager().install())
            self.driver = webdriver.Chrome(service=service, options=chrome_options)
        except Exception as e:
            self.logger.error(f"WebDriver initialization failed: {e}")
            raise

        self.TIMEOUT = 30
        self.excel_path = excel_path
        self.base_url = 'https://www.siad.mg.gov.br/jasi-frontend/'

        # Credenciais - substituir por mecanismo seguro
        self.usuario = 'x0159191'
        self.senha = 'jl1542'

        self.XPATHS = {
            'input_usuario': "//input[@placeholder='Usuário']",
            'input_senha': "//input[@placeholder='Senha']",
            'btn_entrar': "//button[normalize-space(text())='Entrar']",
            'input_digite_unidade': "//input[@placeholder='Digite a Unidade']",
            'btn_selecionar': "//button[normalize-space(text())='Selecionar']",
            'menu_principal_icon': "//div[contains(@class, 'menuicon ibars')]",
            'menu_item_relatorios': "//span[normalize-space(text())='Relatórios']",
            'menu_item_inventario': "//span[normalize-space(text())='Relatório de inventário de bens']",
            'btn_pesquisar': "//button[normalize-space(text())='Pesquisar']",
            'relatorio_link': "//div[normalize-space(text())='INVENTARIO DE PATRIMONIOS']",
            'input_unidade_tarefa': "//span[contains(text(), 'UNID. ADMINISTRATIVA')]/following-sibling::div[1]//input",
            'btn_solicitar_geracao': "//button[normalize-space(text())='Solicitar geração']",
            'btn_ok': "//button[normalize-space(text())='OK']",
            'menu_usuario_icon': "//i[@class='fas fa-user-circle']",
            'menu_item_alterar_unidade': "//span[normalize-space(text())='Alterar Unidade']",
            'btn_alterar': "//button[normalize-space(text())='Alterar']",
            'error_unidade_nao_autorizada': "//span[contains(translate(., 'abcdefghijklmnopqrstuvwxyz', 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'), 'NAO EXISTE PERFIL AUTORIZADO')]",
            'btn_sair_modal_erro': "//button[normalize-space(text())='SAIR']",
        }

    # -------------------------
    # Helpers
    # -------------------------
    def _set_clipboard(self, text: str) -> bool:
        if not pyperclip:
            self.logger.debug("pyperclip não disponível; não copiei para clipboard.")
            return False
        try:
            pyperclip.copy(str(text))
            self.logger.debug("Texto copiado para clipboard.")
            return True
        except Exception as e:
            self.logger.warning(f"Falha ao copiar para clipboard: {e}")
            return False

    def _safe_js(self, script: str, *args):
        try:
            return self.driver.execute_script(script, *args)
        except Exception as e:
            self.logger.debug(f"JS execution failed: {e}")
            return None

    def _wait_visible(self, xpath: str, timeout: Optional[int] = None):
        timeout = timeout or self.TIMEOUT
        return WebDriverWait(self.driver, timeout).until(EC.visibility_of_element_located((By.XPATH, xpath)))

    def _wait_clickable(self, xpath: str, timeout: Optional[int] = None):
        timeout = timeout or self.TIMEOUT
        return WebDriverWait(self.driver, timeout).until(EC.element_to_be_clickable((By.XPATH, xpath)))

    def _screenshot(self, name: str):
        try:
            self.driver.save_screenshot(name)
        except Exception:
            pass

    # Robust click with retries, scrollIntoView and JS fallback
    def _click(self, xpath_key: str, js_fallback: bool = True, raise_on_fail: bool = True) -> bool:
        xpath = self.XPATHS.get(xpath_key, xpath_key)
        last_exc = None
        for attempt in range(1, 3):  # 2 attempts
            try:
                el = WebDriverWait(self.driver, 8).until(EC.element_to_be_clickable((By.XPATH, xpath)))
                try:
                    # ensure visible
                    self.driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", el)
                except Exception:
                    pass
                try:
                    el.click()
                except Exception:
                    if js_fallback:
                        try:
                            self._safe_js("document.evaluate(arguments[0], document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue.click();", xpath)
                        except Exception as e:
                            raise
                    else:
                        raise
                self.logger.info(f"Clicou em {xpath_key}")
                return True
            except Exception as e:
                last_exc = e
                self.logger.debug(f"Attempt {attempt} to click {xpath_key} failed: {e}")
                # try remove overlays and touch body before retry
                try:
                    self._safe_js("document.querySelectorAll('.z-modal, .z-shadow, .overlay, .ui-widget-overlay').forEach(function(el){el.parentNode && el.parentNode.removeChild(el);});")
                    self._safe_js("document.querySelector('body').click();")
                except Exception:
                    pass
                time.sleep(0.4)
        # after retries
        self._screenshot(f'erro_click_{xpath_key}.png')
        self.logger.error(f"Falha ao clicar em {xpath_key}: {last_exc}")
        if raise_on_fail:
            raise AutomationFatalError(f"Erro ao clicar em {xpath_key}: {last_exc}")
        return False

    # Preenche campo com validação (Ctrl+V, send_keys, JS set)
    def _fill_field_guaranteed(self, xpath_key: str, text: str, allow_clipboard: bool = True):
        xpath = self.XPATHS.get(xpath_key, xpath_key)
        if allow_clipboard:
            did_clip = self._set_clipboard(text)
        else:
            did_clip = False

        el = self._wait_visible(xpath, timeout=10)

        # Try Ctrl+V
        if did_clip:
            try:
                try:
                    el.click()
                except Exception:
                    pass
                el.send_keys(Keys.CONTROL, 'v')
                time.sleep(0.25)
                self._safe_js("arguments[0].dispatchEvent(new Event('input')); arguments[0].dispatchEvent(new Event('change'));", el)
                val = self._safe_js("return arguments[0].value", el)
                if str(val).strip() == str(text).strip():
                    self.logger.info(f"Colado (Ctrl+V) '{text}' em {xpath_key}")
                    return
                else:
                    self.logger.debug(f"Valor após Ctrl+V difere: '{val}' (esperado '{text}').")
            except Exception as e:
                self.logger.debug(f"Ctrl+V falhou: {e}")

        # Try send_keys
        try:
            try:
                el.clear()
            except Exception:
                pass
            el.click()
            el.send_keys(text)
            time.sleep(0.2)
            self._safe_js("arguments[0].dispatchEvent(new Event('input')); arguments[0].dispatchEvent(new Event('change'));", el)
            val = self._safe_js("return arguments[0].value", el)
            if str(val).strip() == str(text).strip():
                self.logger.info(f"Digitado '{text}' em {xpath_key}")
                return
            else:
                self.logger.debug(f"Valor após send_keys difere: '{val}' (esperado '{text}').")
        except Exception as e:
            self.logger.debug(f"send_keys falhou: {e}")

        # Try JS set
        try:
            safe_text = str(text).replace("'", "\\'")
            self._safe_js("arguments[0].value = arguments[1]; arguments[0].dispatchEvent(new Event('input')); arguments[0].dispatchEvent(new Event('change'));", el, safe_text)
            time.sleep(0.15)
            val = self._safe_js("return arguments[0].value", el)
            if str(val).strip() == str(text).strip():
                self.logger.info(f"Set via JS '{text}' em {xpath_key}")
                return
            else:
                self.logger.debug(f"Valor após JS set difere: '{val}' (esperado '{text}').")
        except Exception as e:
            self.logger.debug(f"JS set falhou: {e}")

        self._screenshot(f'erro_preencher_{xpath_key}.png')
        raise AutomationFatalError(f"Não foi possível preencher o campo {xpath_key} com '{text}'")

    # -------------------------
    # NEW: attempt direct fill in currently-open modal (preferential flow)
    # -------------------------
    def attempt_fill_in_current_modal(self, unit_code: str) -> Optional[bool]:
        """
        If the modal input 'input_digite_unidade' is present and visible, try to fill it
        and click 'Selecionar'. Returns:
          - True if we attempted and completed the action (selected or recorded unauthorized)
          - False if the field wasn't present / couldn't be used (caller should fallback to menu)
        """
        try:
            el = WebDriverWait(self.driver, 2).until(EC.visibility_of_element_located((By.XPATH, self.XPATHS['input_digite_unidade'])))
        except Exception:
            return False  # field not present -> fallback required

        self.logger.info("Campo modal 'Digite a Unidade' está presente: tentando colar diretamente sem abrir menu.")
        try:
            # prepare clipboard and fill
            self._set_clipboard(unit_code)
            self._fill_field_guaranteed('input_digite_unidade', unit_code, allow_clipboard=True)

            # try to trigger Selecionar (JS click more robust)
            try:
                self._safe_js("document.evaluate(arguments[0], document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue.click();", self.XPATHS['btn_selecionar'])
            except Exception:
                # fallback clickable click
                try:
                    self._click('btn_selecionar')
                except Exception:
                    self.logger.debug("Falha ao acionar botão Selecionar após colar no modal.")
                    raise

            time.sleep(0.6)

            # check unauthorized modal
            if self._detect_and_record_unauthorized_and_cleanup(unit_code):
                self.logger.info("Unidade sem acesso detectada após tentativa direta no modal (registrada e limpa).")
            else:
                self.logger.info("Unidade selecionada via modal direto com sucesso.")

            return True
        except Exception as e:
            self.logger.debug(f"Tentativa direta no modal falhou: {e}")
            # do not raise here — fallback will handle via opening menu
            return False

    # -------------------------
    # Core flows
    # -------------------------
    def login(self):
        self.logger.info("Iniciando login")
        self.driver.get(self.base_url)
        self._fill_field_guaranteed('input_usuario', self.usuario, allow_clipboard=True)
        self._fill_field_guaranteed('input_senha', self.senha, allow_clipboard=True)
        self._click('btn_entrar')
        time.sleep(2)

    def select_unit_initial(self, unit_code: str) -> bool:
        """
        Select unit on initial modal. Returns True if selected/processed.
        If unauthorized, records and returns False so caller can decide (skip).
        """
        self.logger.info(f"Selecionando unidade inicial: {unit_code}")

        # Prefer direct fill if modal input visible
        tried_direct = self.attempt_fill_in_current_modal(unit_code)
        if tried_direct:
            # if attempted direct, check if unauthorized was recorded
            # _detect_and_record_unauthorized_and_cleanup already recorded and cleaned when necessary,
            # but attempt_fill_in_current_modal returns True even when unauthorized occurred.
            return not self._is_last_unit_unauthorized(unit_code)

        # If direct attempt not possible, fallback to fill normally (shouldn't happen on initial, but safe)
        self._set_clipboard(unit_code)
        self._fill_field_guaranteed('input_digite_unidade', unit_code, allow_clipboard=True)
        try:
            self._safe_js("document.evaluate(arguments[0], document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue.click();", self.XPATHS['btn_selecionar'])
            time.sleep(0.6)
        except Exception as e:
            self._screenshot('erro_selecionar_btn.png')
            raise AutomationFatalError(f"Falha ao acionar botão Selecionar: {e}")

        if self._detect_and_record_unauthorized_and_cleanup(unit_code):
            return False
        return True

    def change_unit_and_loop(self, unit_code: str) -> bool:
        """
        Change unit for subsequent iterations.
        First tries direct modal fill (if modal present). If not present, opens menu and uses 'Alterar Unidade'.
        Returns True if unit changed successfully, False if unit had no access and was recorded/cleaned (skip).
        """
        self.logger.info(f"Iniciando alteração de unidade para: {unit_code}")

        # 1) attempt direct fill in current modal (preferred)
        tried_direct = self.attempt_fill_in_current_modal(unit_code)
        if tried_direct:
            # direct attempt either selected unit or recorded unauthorized -> decide skip based on detection
            return not self._is_last_unit_unauthorized(unit_code)

        # 2) fallback: open menu and use Alterar Unidade flow
        ok = self._click('menu_usuario_icon', raise_on_fail=False)
        if not ok:
            self.logger.warning("Não foi possível abrir o menu do usuário; pulando esta unidade para continuar execução.")
            return False

        ok = self._click('menu_item_alterar_unidade', raise_on_fail=False)
        if not ok:
            self.logger.warning("Não foi possível clicar em 'Alterar Unidade'; pulando esta unidade.")
            return False

        # some flows require clicking OK to proceed
        try:
            self._click('btn_ok', raise_on_fail=False)
        except Exception:
            pass

        # prepare clipboard and fill the modal input
        self._set_clipboard(unit_code)
        self._fill_field_guaranteed('input_digite_unidade', unit_code, allow_clipboard=True)

        # click 'Alterar' - if this fails, treat as fatal (can't proceed reliably)
        try:
            self._safe_js("document.evaluate(arguments[0], document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue.click();", self.XPATHS['btn_alterar'])
            time.sleep(0.6)
        except Exception as e:
            self._screenshot('erro_alterar_btn.png')
            raise AutomationFatalError(f"Falha ao acionar botão Alterar: {e}")

        # check unauthorized
        if self._detect_and_record_unauthorized_and_cleanup(unit_code):
            self.logger.warning(f"Unidade {unit_code} sem acesso detectada. Registrada e pulada.")
            return False

        return True

    def generate_inventory_report(self, unit_code: str):
        self.logger.info(f"Iniciando geração de relatório para unidade: {unit_code}")
        self._click('menu_principal_icon')
        self._click('menu_item_relatorios')
        self._click('menu_item_relatorios')
        self._click('menu_item_inventario')

        WebDriverWait(self.driver, self.TIMEOUT).until(
            EC.presence_of_element_located((By.XPATH, "//span[normalize-space(text())='Unidade emitente:']"))
        )

        self._click('btn_pesquisar')
        self._click('relatorio_link')
        self._fill_field_guaranteed('input_unidade_tarefa', unit_code, allow_clipboard=True)
        self._click('btn_solicitar_geracao')
        self._click('btn_ok')
        time.sleep(1.2)

    # -------------------------
    # Unauthorized detection & record
    # -------------------------
    def write_unauthorized_unit(self, unit_code: str):
        unauthorized_file = 'C:/Users/p0134255/Documents/Rogério/Backup/Tj/Projetos/Phyton/Automação-SIAD/unidades_sem_acesso.xlsx'
        try:
            try:
                df = pd.read_excel(unauthorized_file, engine='openpyxl')
            except FileNotFoundError:
                df = pd.DataFrame(columns=['Unidade', 'Data_Registro', 'Motivo'])
            new_row = pd.DataFrame({
                'Unidade': [unit_code],
                'Data_Registro': [pd.Timestamp.now().strftime('%Y-%m-%d %H:%M:%S')],
                'Motivo': ['NAO EXISTE PERFIL AUTORIZADO']
            })
            df = pd.concat([df, new_row], ignore_index=True)
            df.to_excel(unauthorized_file, index=False, engine='openpyxl')
            self.logger.warning(f"Unidade não autorizada {unit_code} registrada em {unauthorized_file}")
        except Exception as e:
            self.logger.error(f"ERRO ao escrever unidade não autorizada: {e}")

    def _detect_and_record_unauthorized_and_cleanup(self, unit_code: str) -> bool:
        """
        Detects 'NAO EXISTE PERFIL AUTORIZADO' modal. If found:
         - records the unit
         - tries to close the modal
         - attempts to remove overlays and clean the input field (value='')
         - returns True indicating unit had no access (caller should skip it)
        """
        try:
            WebDriverWait(self.driver, 1.5).until(
                EC.presence_of_element_located((By.XPATH, self.XPATHS['error_unidade_nao_autorizada']))
            )
            self.logger.warning(f"Erro de acesso detectado para a unidade: {unit_code}")
            self.write_unauthorized_unit(unit_code)

            # try to click 'SAIR'
            try:
                btn = WebDriverWait(self.driver, 2).until(EC.element_to_be_clickable((By.XPATH, self.XPATHS['btn_sair_modal_erro'])))
                try:
                    btn.click()
                except Exception:
                    try:
                        self._safe_js("document.evaluate(arguments[0], document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue.click();", self.XPATHS['btn_sair_modal_erro'])
                    except Exception:
                        pass
            except Exception:
                self.logger.debug("'SAIR' button not found/clickable")

            # remove overlays
            try:
                self._safe_js("document.querySelectorAll('.z-modal, .z-shadow, .overlay, .ui-widget-overlay').forEach(function(el){el.parentNode && el.parentNode.removeChild(el);});")
            except Exception:
                pass

            # clear and focus the input if present
            try:
                el = self.driver.find_element(By.XPATH, self.XPATHS['input_digite_unidade'])
                try:
                    self._safe_js("arguments[0].removeAttribute('readonly'); arguments[0].removeAttribute('disabled');", el)
                except Exception:
                    pass
                try:
                    self._safe_js("arguments[0].value=''; arguments[0].dispatchEvent(new Event('input')); arguments[0].dispatchEvent(new Event('change'));", el)
                except Exception:
                    pass
                try:
                    el.click()
                except Exception:
                    try:
                        self._safe_js("arguments[0].focus();", el)
                    except Exception:
                        pass
            except Exception:
                self.logger.debug("Campo de digitar unidade não encontrado para limpeza.")

            time.sleep(0.6)
            return True
        except Exception:
            return False

    # Helper to detect whether the last processed unit_code was registered as unauthorized
    # (we simply check the unidades_sem_acesso.xlsx last row if exists and matches unit_code)
    def _is_last_unit_unauthorized(self, unit_code: str) -> bool:
        unauthorized_file = 'C:/Users/p0134255/Documents/Rogério/Backup/Tj/Projetos/Phyton/Automação-SIAD/unidades_sem_acesso.xlsx'
        try:
            df = pd.read_excel(unauthorized_file, engine='openpyxl', dtype=str)
            if not df.empty:
                last = str(df.iloc[-1]['Unidade']).strip()
                return last == str(unit_code).strip()
            return False
        except Exception:
            return False

    # -------------------------
    # Main orchestration
    # -------------------------
    def execute_automation(self):
        try:
            df = pd.read_excel(self.excel_path, sheet_name=0, header=0, dtype=str)
            unit_codes = df.iloc[:, 0].dropna().unique().tolist()
            if not unit_codes:
                self.logger.warning("Nenhuma unidade encontrada na planilha. Encerrando.")
                return

            self.logger.info(f"Total de {len(unit_codes)} unidades para processar.")
            self.login()

            for i, unit_code in enumerate(unit_codes):
                self.logger.info(f"--- Processando unidade {i+1}/{len(unit_codes)}: {unit_code} ---")
                try:
                    # prepare clipboard BEFORE interacting
                    self._set_clipboard(unit_code)

                    if i == 0:
                        ok = self.select_unit_initial(unit_code)
                    else:
                        ok = self.change_unit_and_loop(unit_code)

                    if not ok:
                        # skip unit and continue with next (was unauthorized or menu open failed)
                        self.logger.info(f"Pulando unidade {unit_code} e seguindo para próxima.")
                        time.sleep(0.6)
                        continue

                    # generate report for this unit
                    self.generate_inventory_report(unit_code)
                    time.sleep(0.6)

                except AutomationFatalError as e:
                    self.logger.error(f"Erro inesperado (fatal). Parando execução. Detalhes: {e}")
                    self._screenshot(f'erro_fatal_unidade_{unit_code}.png')
                    raise

            self.logger.info("Execução finalizada (todas unidades processadas).")

        except Exception as e:
            self.logger.error(f"Execução interrompida com erro: {e}")
        finally:
            try:
                self.driver.quit()
            except Exception:
                pass
            self.logger.info("WebDriver fechado.")

def main():
    try:
        automation = SIADAutomation()
        automation.execute_automation()
    except Exception as e:
        print(f"A automação falhou: {e}")

if __name__ == "__main__":
    main()