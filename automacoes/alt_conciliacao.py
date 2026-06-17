import os
import time
import random
import traceback
import pandas as pd
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchWindowException, NoSuchElementException, TimeoutException
from webdriver_manager.chrome import ChromeDriverManager
import pyautogui
import keyboard

automation_status = {
    'running': False,
    'total': 0,
    'processed': 0,
    'current_item': None,
    'logs': [],
    'start_time': None,
    'end_time': None,
    'request_stop': False,
    'relatorio_path': None
}

class WebAppAlteradorConciliacao:
    def __init__(self, dados_login, lista_cnpjs, nome_conciliacao, tipo_pesquisa='CNPJ'):
        self.dados_login = dados_login
        self.lista_cnpjs = lista_cnpjs
        self.nome_conciliacao = nome_conciliacao
        self.tipo_pesquisa = tipo_pesquisa.upper()
        self.driver = None
        self.wait = None
        self.actions = None
        self.resultados = []

    def add_log(self, message, level="INFO"):
        timestamp = datetime.now().strftime('%H:%M:%S')
        prefixes = {"ERROR": "ERRO", "WARNING": "AVISO", "SUCCESS": "SUCESSO"}
        prefix = prefixes.get(level, "")
        log_entry = f"[{timestamp}] {f'{prefix}: ' if prefix else ''}{message}"
        automation_status['logs'].append(log_entry)
        print(log_entry)
        try:
            with open('ativador_log.txt', 'a', encoding='utf-8') as f:
                f.write(log_entry + '\n')
        except:
            pass

    def handle_exception(self, e, context=""):
        error_str = str(e)
        if "stale element reference" in error_str:
            msg = "Elemento desatualizado"
        elif "no such element" in error_str:
            msg = "Elemento não encontrado"
        elif "timeout" in error_str.lower():
            msg = "Tempo de espera excedido"
        elif "element click intercepted" in error_str:
            msg = "Clique interceptado"
        elif "element not interactable" in error_str:
            msg = "Elemento não interagível"
        else:
            msg = error_str.split('\n')[0][:120]
        self.add_log(f"{context}: {msg}", level="ERROR")
        try:
            with open('debug_ativador.txt', 'a', encoding='utf-8') as f:
                f.write(f"\n--- {datetime.now()} ---\n{context}\n{error_str}\n")
                traceback.print_exc(file=f)
        except:
            pass

    def _iniciar_driver(self):
        self.add_log("Iniciando navegador...")
        try:
            user_data_dir = os.path.join(os.getcwd(), "selenium-profile")
            os.makedirs(user_data_dir, exist_ok=True)
            opts = Options()
            opts.add_argument(f"--user-data-dir={user_data_dir}")
            opts.add_argument("--start-maximized")
            opts.add_argument("--disable-extensions")
            opts.add_argument("--disable-popup-blocking")
            opts.add_argument("--disable-blink-features=AutomationControlled")
            opts.binary_location = r"C:\BACKUP_EMANUEL\Program Files\Google\Chrome\Application\chrome.exe"
            service = ChromeService(ChromeDriverManager().install())
            self.driver = webdriver.Chrome(service=service, options=opts)
            self.driver.get("about:blank")
            self.wait = WebDriverWait(self.driver, 15)
            self.actions = ActionChains(self.driver)
            self.driver.implicitly_wait(8)
            time.sleep(2)
            self._navegar_pyautogui("https://online.fwcard.com.br/fwcard/f")
            time.sleep(3)
            self.add_log("Navegador iniciado.")
        except Exception as e:
            self.handle_exception(e, "Erro ao iniciar navegador")
            raise e

    def _navegar_pyautogui(self, url):
        try:
            pyautogui.click(100, 100)
            time.sleep(.5)
            pyautogui.hotkey('ctrl', 'l')
            time.sleep(.5)
            pyautogui.hotkey('ctrl', 'a')
            time.sleep(.3)
            pyautogui.press('delete')
            time.sleep(.3)
            pyautogui.write(url, interval=0.05)
            keyboard.press_and_release('shift+/')
            keyboard.write('p=380:100:')
            pyautogui.press('enter')
            time.sleep(3)
        except Exception as e:
            self.handle_exception(e, "Erro PyAutoGUI")

    def _realizar_login(self):
        self.add_log("Realizando login...")
        try:
            u = WebDriverWait(self.driver, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="P100_USERNAME"]')))
            self._type(self.dados_login['login_infox'], u)
            p = WebDriverWait(self.driver, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="P100_PASSWORD"]')))
            self._type(self.dados_login['senha_infox'], p)
            btn = WebDriverWait(self.driver, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="LOGIN"]/span')))
            try:
                cap = WebDriverWait(self.driver, 5).until(EC.presence_of_element_located((By.XPATH, '//*[@id="uLogin"]/div[5]/div/div/div[1]')))
                self._click(cap)
            except:
                pass
            self._click(btn)
            self.add_log("Login realizado.")
            time.sleep(2)
        except Exception as e:
            self.handle_exception(e, "Erro no login")
            raise e

    def _type(self, text, el):
        el.clear()
        time.sleep(.2)
        for c in text:
            el.send_keys(c)
            time.sleep(random.uniform(.03, .1))

    def _click(self, el):
        self.driver.execute_script("arguments[0].scrollIntoView({block:'center'});", el)
        time.sleep(.2)
        ActionChains(self.driver).move_to_element(el).pause(.1).perform()
        el.click()
        time.sleep(.2)

    def acessar_pagina_pesquisa(self):
        try:
            self.wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="t_MenuNav_11i"]'))).click()
            time.sleep(0.5)
            self.wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="t_MenuNav_11_1i"]'))).click()
            time.sleep(0.5)
            self.wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="t_MenuNav_11_1_2i"]'))).click()
            time.sleep(0.5)
            self.wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="t_MenuNav_11_1_2_1i"]'))).click()
            time.sleep(1)
            return True
        except Exception as e:
            self.handle_exception(e, "Erro ao acessar página de pesquisa")
            return False

    def realizar_alteracao(self, cnpj):
        try:
            self.add_log("Alterando conciliação...")
            btn_alterar = WebDriverWait(self.driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, '//*[@id="ALTERAR_CFG_EDI"]/span'))
            )
            self._click(btn_alterar)
            time.sleep(1)

            btn_conciliacoes = WebDriverWait(self.driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, '//*[@id="P2310_ID_EDIFNC_CONC_VENDA"]'))
            )
            self._click(btn_conciliacoes)
            time.sleep(1)

            nova_conciliacao = WebDriverWait(self.driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, f"//select[@id='P2310_ID_EDIFNC_CONC_VENDA']/option[text()='{self.nome_conciliacao}']"))
            )
            nova_conciliacao.click()
            time.sleep(1)

            btn_salvar = WebDriverWait(self.driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, '//*[@id="OK"]/span'))
            )
            self._click(btn_salvar)
            time.sleep(2)

            return True, "Alteração realizada com sucesso"
        except Exception as e:
            return False, str(e)

    def iniciar(self):
        try:
            self.add_log("Iniciando processo de alteração de conciliação...")
            automation_status['running'] = True
            automation_status['start_time'] = datetime.now()
            automation_status['total'] = len(self.lista_cnpjs)
            automation_status['processed'] = 0

            self._iniciar_driver()
            self._realizar_login()

            for idx, item in enumerate(self.lista_cnpjs):
                if automation_status['request_stop']:
                    self.add_log("Interrupção solicitada pelo usuário.")
                    break

                cnpj = item.get('cod', '').strip()
                automation_status['current_item'] = cnpj
                self.add_log("-" * 40)
                self.add_log(f"Processando [{idx+1}/{len(self.lista_cnpjs)}]: {cnpj}")

                if not self.acessar_pagina_pesquisa():
                    self.resultados.append({'CNPJ': cnpj, 'Status': 'ERRO', 'Mensagem': 'Falha ao acessar página de pesquisa'})
                    continue

                try:
                    xpath_input = '//*[@id="P2100_CNPJ"]' if self.tipo_pesquisa == 'CNPJ' else '//*[@id="P2100_ID_CREDENCIADO"]'
                    input_campo = WebDriverWait(self.driver, 10).until(
                        EC.presence_of_element_located((By.XPATH, xpath_input))
                    )
                    input_campo.clear()
                    input_campo.send_keys(cnpj)

                    self.driver.find_element(By.ID, 'PESQUISAR').click()
                    time.sleep(2)

                    sucesso, msg = self.realizar_alteracao(cnpj)
                    self.resultados.append({
                        'CNPJ': cnpj,
                        'Status': 'SUCESSO' if sucesso else 'ERRO',
                        'Mensagem': msg
                    })

                    if sucesso:
                        self.add_log(f"CNPJ {cnpj} alterado com sucesso!", level="SUCCESS")
                    else:
                        self.add_log(f"Erro ao alterar CNPJ {cnpj}: {msg}", level="ERROR")

                except Exception as e:
                    self.handle_exception(e, f"Erro no processamento do CNPJ {cnpj}")
                    self.resultados.append({'CNPJ': cnpj, 'Status': 'ERRO', 'Mensagem': str(e)})

                automation_status['processed'] = idx + 1

            # Salva o relatório
            if self.resultados:
                ts = datetime.now().strftime('%Y%m%d_%H%M%S')
                relatorio_dir = os.path.join(os.getcwd(), 'uploads')
                os.makedirs(relatorio_dir, exist_ok=True)
                caminho_rel = os.path.join(relatorio_dir, f"resultado_conciliacao_{ts}.xlsx")
                df = pd.DataFrame(self.resultados)
                df.to_excel(caminho_rel, index=False)
                automation_status['relatorio_path'] = caminho_rel
                self.add_log(f"Relatório salvo em {caminho_rel}", level="SUCCESS")

        except Exception as e:
            self.handle_exception(e, "Erro crítico na automação")
        finally:
            if self.driver:
                try:
                    self.driver.quit()
                except:
                    pass
            automation_status['running'] = False
            automation_status['end_time'] = datetime.now()
