from datetime import datetime
import os
import random
import re
import time
import traceback

import keyboard
import pandas as pd
import pyautogui
from selenium import webdriver
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from webdriver_manager.chrome import ChromeDriverManager


extraction_status = {
    "running": False,
    "total": 0,
    "processed": 0,
    "current_item": None,
    "logs": [],
    "start_time": None,
    "end_time": None,
    "request_stop": False,
    "output_file": None,
}


UF_INDEX = {
    "": None,
    "AC": 2,
    "AL": 3,
    "AM": 4,
    "AP": 5,
    "BA": 6,
    "CE": 7,
    "DF": 8,
    "ES": 9,
    "GO": 10,
    "MA": 11,
    "MG": 12,
    "MS": 13,
    "MT": 14,
    "PA": 15,
    "PB": 16,
    "PE": 17,
    "PI": 18,
    "PR": 19,
    "RJ": 20,
    "RN": 21,
    "RO": 22,
    "RR": 23,
    "RS": 24,
    "SC": 25,
    "SE": 26,
    "SP": 27,
    "TO": 28,
}


class WebAppExtrator:
    def __init__(self, dados_login, filtros):
        self.dados_login = dados_login
        self.filtros = filtros
        self.modo_consulta_transacoes = bool(self.filtros.get("consultas_transacoes"))
        self.funcionalidade_transacao = self.filtros.get("funcionalidade_transacao", "valor_venda_total")
        self.driver = None
        self.wait = None
        self.actions = None
        self.dados_coletados = []
        self.colunas_saida = (
            [
                "CODIGO",
                "DATA_INICIAL",
                "DATA_FINAL",
                "DATA_ULTIMA_TRANSACAO",
            ]
            if self.modo_consulta_transacoes and self.funcionalidade_transacao == "data_ultima_transacao"
            else [
                "CODIGO",
                "DATA_INICIAL",
                "DATA_FINAL",
                "QTD_TRANSACOES",
                "VALOR_TOTAL_TRANSACOES",
            ]
            if self.modo_consulta_transacoes
            else [
                "CODIGO",
                "CNPJ",
                "NOME_FANTASIA",
                "RAZAO_SOCIAL",
                "CIDADE",
                "ENDERECO",
                "BAIRRO",
                "CEP",
                "EMAIL",
                "CONCILIACAO",
                "CELULAR",
                "TELEFONE",
                "DATA_CADASTRO",
            ]
        )
        self.xpaths = {
            "campo_codigo": '//*[@id="P2100_ID_CREDENCIADO"]',
            "campo_cnpj": '//*[@id="P2100_CNPJ"]',
            "campo_nome_fantasia": '//*[@id="P2100_NOME_FANTASIA"]',
            "campo_razao_social": '//*[@id="P2100_RAZAO_SOCIAL"]',
            "campo_cidade": '//*[@id="P2100_CIDADE"]',
            "botao_pesquisar": '//*[@id="PESQUISAR"]/span',
            "linhas_resultado": '//*[@id="LISTA"]//tr[td//a]',
            "primeiro_resultado": '(//*[@id="LISTA"]//tr[td//a]//td[1]//a)[1]',
            "link_resultado_por_indice": '(//*[@id="LISTA"]//tr[td//a]//td[1]//a)[{idx}]',
            "codigo_resultado_por_indice": '(//*[@id="LISTA"]//tr[td//a]//td[3])[{idx}]',
            "campo_detalhe_codigo": '//*[@id="P2101_CODIGO"]',
            "campo_detalhe_cnpj": '//*[@id="P2101_CNPJ"]',
            "campo_detalhe_nome_fantasia": '//*[@id="P2101_FANTASIA"]',
            "campo_detalhe_razao_social": '//*[@id="P2101_RAZAO"]',
            "campo_detalhe_cidade": '//*[@id="report_ENDERECOS"]/tbody[2]/tr/td/table/tbody/tr/td[6]/a',
            "campo_detalhe_endereco": '//*[@id="report_ENDERECOS"]/tbody[2]/tr/td/table/tbody/tr/td[3]/a',
            "campo_detalhe_bairro": '//*[@id="report_ENDERECOS"]/tbody[2]/tr/td/table/tbody/tr/td[4]/a',
            "campo_detalhe_cep": '//*[@id="report_ENDERECOS"]/tbody[2]/tr/td/table/tbody/tr/td[5]/a',
            "campo_detalhe_email": '//*[@id="P2101_EMAIL"]',
            "campo_detalhe_conciliacao": '//*[@id="P2101_EDIFNC_CONC_VENDA"]',
            "campo_detalhe_celular": '//*[@id="P2101_CELULAR"]',
            "campo_detalhe_telefone": '//*[@id="P2101_TELEFONE"]',
            "campo_detalhe_data_cadastro": '//*[@id="P2101_DATA_CADASTRO"]',
            "menu_consulta_principal": '//*[@id="t_MenuNav_11i"]',
            "menu_consulta_secundario": '//*[@id="t_MenuNav_11_2"]',
            "menu_consulta_transacoes": '//*[@id="t_MenuNav_11_2_2"]',
            "campo_transacao_data_inicial": '//*[@id="P5320_DATA_INICIAL_input"]',
            "campo_transacao_data_final": '//*[@id="P5320_DATA_FINAL_input"]',
            "campo_transacao_codigo": '//*[@id="P5320_CREDENCIADO_PESQ"]',
            "botao_transacao_pesquisar": '//*[@id="PESQUISAR"]/span',
            "campo_transacao_total_qtd": '//*[@id="P5320_TOTAL_TRANSACAO"]',
            "campo_transacao_total_valor": '//*[@id="P5320_VALOR_TOTAL_TRANSACAO"]',
            "coluna_transacao_data_lancamento": '//*[@id="DATA_LANCTO"]/a',
            "botao_transacao_ordenacao_desc": '//*[@id="LISTA_sort_widget_action_down"]/button',
            "campo_transacao_data_ultima": '//*[@id="5853054184045474152"]/tbody/tr[2]/td[7]',
        }

    def add_log(self, message, level="INFO"):
        timestamp = datetime.now().strftime("%H:%M:%S")
        prefixes = {"ERROR": "ERRO", "WARNING": "AVISO", "SUCCESS": "SUCESSO"}
        prefix = prefixes.get(level, "")
        log_entry = f"[{timestamp}] {f'{prefix}: ' if prefix else ''}{message}"
        extraction_status["logs"].append(log_entry)
        print(log_entry)
        try:
            with open("extrator_log.txt", "a", encoding="utf-8") as f:
                f.write(log_entry + "\n")
        except Exception:
            pass

    def handle_exception(self, e, context=""):
        error_str = str(e)
        lowered = error_str.lower()
        if "stale element reference" in lowered:
            msg = "Elemento desatualizado"
        elif "no such element" in lowered:
            msg = "Elemento nao encontrado"
        elif "timeout" in lowered:
            msg = "Tempo de espera excedido"
        elif "element click intercepted" in lowered:
            msg = "Clique interceptado"
        elif "element not interactable" in lowered:
            msg = "Elemento nao interagivel"
        elif "no such window" in lowered:
            msg = "Janela do navegador foi fechada"
        else:
            msg = error_str.split("\n")[0][:160]
        self.add_log(f"{context}: {msg}", level="ERROR")
        try:
            with open("debug_extrator.txt", "a", encoding="utf-8") as f:
                f.write(f"\n--- {datetime.now()} ---\n{context}\n{error_str}\n")
                traceback.print_exc(file=f)
        except Exception:
            pass

    def _iniciar_driver(self):
        self.add_log("Iniciando navegador...")
        user_data_dir = os.path.join(os.getcwd(), "selenium-profile-extrator")
        os.makedirs(user_data_dir, exist_ok=True)

        opts = Options()
        opts.add_argument(f"--user-data-dir={user_data_dir}")
        opts.add_argument("--start-maximized")
        opts.add_argument("--disable-extensions")
        opts.add_argument("--disable-popup-blocking")
        opts.add_argument("--disable-blink-features=AutomationControlled")
        chrome_binary = os.environ.get("CHROME_BINARY_PATH", "").strip()
        fallback_chrome_binary = r"C:\BACKUP_EMANUEL\Program Files\Google\Chrome\Application\chrome.exe"
        if chrome_binary:
            opts.binary_location = chrome_binary
        elif os.path.exists(fallback_chrome_binary):
            opts.binary_location = fallback_chrome_binary

        service = ChromeService(ChromeDriverManager().install())
        self.driver = webdriver.Chrome(service=service, options=opts)
        self.wait = WebDriverWait(self.driver, 12)
        self.actions = ActionChains(self.driver)
        self.driver.implicitly_wait(1)

        self.driver.get("about:blank")
        time.sleep(1)
        self._navegar_pyautogui("https://online.fwcard.com.br/fwcard/f")
        time.sleep(2)
        self.add_log("Navegador iniciado.")

    def _navegar_pyautogui(self, url):
        try:
            pyautogui.click(100, 100)
            time.sleep(0.3)
            pyautogui.hotkey("ctrl", "l")
            time.sleep(0.2)
            pyautogui.hotkey("ctrl", "a")
            time.sleep(0.1)
            pyautogui.press("delete")
            time.sleep(0.1)
            pyautogui.write(url, interval=0.02)
            keyboard.press_and_release("shift+/")
            keyboard.write("p=380:100:")
            pyautogui.press("enter")
            time.sleep(2)
        except Exception as e:
            self.handle_exception(e, "Erro ao navegar via PyAutoGUI")

    def _realizar_login(self):
        self.add_log("Realizando login...")
        usuario = WebDriverWait(self.driver, 10).until(
            EC.presence_of_element_located((By.XPATH, '//*[@id="P100_USERNAME"]'))
        )
        senha = WebDriverWait(self.driver, 10).until(
            EC.presence_of_element_located((By.XPATH, '//*[@id="P100_PASSWORD"]'))
        )
        botao = WebDriverWait(self.driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, '//*[@id="LOGIN"]/span'))
        )

        self._type(self.dados_login["login_infox"], usuario)
        self._type(self.dados_login["senha_infox"], senha)

        try:
            captcha = WebDriverWait(self.driver, 3).until(
                EC.presence_of_element_located((By.XPATH, '//*[@id="uLogin"]/div[5]/div/div/div[1]'))
            )
            self._click(captcha)
        except Exception:
            pass

        self._click(botao)
        time.sleep(1)
        self.add_log("Login realizado.")

    def _type(self, text, el):
        el.clear()
        time.sleep(0.1)
        for c in text:
            el.send_keys(c)
            time.sleep(random.uniform(0.01, 0.04))

    def _click(self, el):
        self.driver.execute_script("arguments[0].scrollIntoView({block:'center'});", el)
        ActionChains(self.driver).move_to_element(el).pause(0.05).perform()
        el.click()

    def _is_cnpj_value(self, value):
        if not value:
            return False
        digits = re.sub(r"\D", "", str(value))
        return len(digits) == 14

    def _esperar_tela_credenciado(self, timeout=8):
        try:
            WebDriverWait(self.driver, timeout).until(
                EC.presence_of_element_located((By.XPATH, self.xpaths["campo_codigo"]))
            )
            WebDriverWait(self.driver, timeout).until(
                EC.element_to_be_clickable((By.XPATH, self.xpaths["botao_pesquisar"]))
            )
            return True
        except Exception:
            return False

    def _esperar_detalhe_credenciado(self, timeout=8):
        try:
            WebDriverWait(self.driver, timeout).until(
                EC.presence_of_element_located(
                    (
                        By.XPATH,
                        f'{self.xpaths["campo_detalhe_cnpj"]} | {self.xpaths["campo_detalhe_codigo"]}',
                    )
                )
            )
            return True
        except Exception:
            return False

    def _limpar_campos_codigo_cnpj(self):
        for xpath in (self.xpaths["campo_codigo"], self.xpaths["campo_cnpj"]):
            try:
                self.driver.find_element(By.XPATH, xpath).clear()
            except Exception:
                pass
        time.sleep(0.1)

    def _pesquisar_por_planilha(self, valor):
        self._limpar_campos_codigo_cnpj()
        xpath = self.xpaths["campo_cnpj"] if self._is_cnpj_value(valor) else self.xpaths["campo_codigo"]
        try:
            campo = self.wait.until(EC.presence_of_element_located((By.XPATH, xpath)))
            campo.clear()
            campo.send_keys(valor)
            self.wait.until(EC.element_to_be_clickable((By.XPATH, self.xpaths["botao_pesquisar"]))).click()
            return self._aguardar_resultado_pesquisa_planilha(valor)
        except Exception as e:
            self.handle_exception(e, f"Erro ao pesquisar {valor}")
            return False

    def _aguardar_resultado_pesquisa_planilha(self, valor, timeout=8):
        fim = time.time() + timeout
        while time.time() < fim:
            if self._esperar_detalhe_credenciado(timeout=1):
                self.add_log(f"Pesquisa de {valor} abriu diretamente a tela de dados.")
                return True

            try:
                if self.driver.find_elements(By.XPATH, self.xpaths["primeiro_resultado"]):
                    self.add_log(f"Pesquisa de {valor} retornou lista de resultados.")
                    return True
            except Exception:
                pass

            try:
                if self.driver.find_elements(By.XPATH, '//*[@id="LISTA"]//*[contains(@class,"a-IRR-noDataMsg-text")]'):
                    self.add_log(f"Nenhum resultado encontrado para {valor}.", level="WARNING")
                    return False
            except Exception:
                pass

            time.sleep(0.2)

        self.add_log(f"Pesquisa de {valor} nao carregou resultado nem tela de dados.", level="WARNING")
        return False

    def _abrir_primeiro_resultado(self):
        try:
            link = WebDriverWait(self.driver, 4).until(
                EC.presence_of_element_located((By.XPATH, self.xpaths["primeiro_resultado"]))
            )
            return self._abrir_link_resultado(link)
        except TimeoutException:
            return False
        except Exception as e:
            self.handle_exception(e, "Erro ao abrir o primeiro resultado")
            return False

    def _abrir_link_resultado(self, link):
        try:
            href = (link.get_attribute("href") or "").strip()
            if not href:
                return False
            if href.startswith("f?p="):
                href = f"https://online.fwcard.com.br/fwcard/{href}"
            self.driver.get(href)
            return self._esperar_detalhe_credenciado(timeout=6)
        except Exception as e:
            self.handle_exception(e, "Erro ao abrir o link do resultado")
            return False

    def _preparar_proximo_item_planilha(self):
        self.add_log("Voltando para a tela de credenciado para pesquisar o proximo item...")
        if self._ir_para_tela_credenciado():
            return True
        self.add_log(
            "Nao foi possivel retornar a tela de credenciado para continuar a planilha.",
            level="ERROR",
        )
        return False

    def _processar_lista_planilha(self):
        lista = [str(item).strip() for item in self.filtros.get("planilha_codigos", []) if str(item).strip()]
        total = len(lista)
        extraction_status["total"] = total
        self.add_log(f"Total de credenciados na planilha: {total}")

        if total == 0:
            self.add_log("Nenhum item valido encontrado na planilha.", level="WARNING")
            self._salvar_planilha(final=True)
            return

        try:
            for idx, valor in enumerate(lista, start=1):
                if extraction_status["request_stop"]:
                    self.add_log("Interrupcao solicitada.")
                    break

                if idx > 1 and not self._esperar_tela_credenciado(timeout=2):
                    self.add_log("Recuperando a tela de credenciado antes do proximo item...")
                    if not self._preparar_proximo_item_planilha():
                        self.add_log(
                            "Nao foi possivel retornar a tela de credenciado antes do proximo item. Salvando e finalizando.",
                            level="ERROR",
                        )
                        self._salvar_planilha(final=True)
                        return

                try:
                    self.add_log("-" * 40)
                    self.add_log(f"Planilha [{idx}/{total}]: pesquisando {valor}...")
                    extraction_status["current_item"] = f"{idx}/{total}"

                    if not self._pesquisar_por_planilha(valor):
                        self.add_log(f"Nenhum resultado encontrado para {valor}.", level="WARNING")
                        extraction_status["processed"] = idx
                        continue

                    if not self._esperar_detalhe_credenciado(timeout=2):
                        if not self._abrir_primeiro_resultado():
                            self.add_log(f"Nao foi possivel abrir os dados para {valor}.", level="WARNING")
                            extraction_status["processed"] = idx
                            continue

                    dados = self._extrair_dados_credenciado()
                    if not self._dados_foram_extraidos(dados):
                        self.add_log(
                            f"Detalhe de {valor} abriu, mas os campos vieram vazios.",
                            level="WARNING",
                        )
                        extraction_status["processed"] = idx
                        if not self._preparar_proximo_item_planilha():
                            self._salvar_planilha(final=True)
                            return
                        continue

                    self.dados_coletados.append(dados)
                    self.add_log(
                        f"Dados coletados para {valor}: {len(dados)} campos. Total coletados: {len(self.dados_coletados)}"
                    )
                    self.add_log(
                        f"[{idx}] {dados.get('NOME_FANTASIA') or dados.get('RAZAO_SOCIAL') or valor} coletado.",
                        level="SUCCESS",
                    )

                    self._salvar_planilha(final=False)

                    if not self._preparar_proximo_item_planilha():
                        self.add_log(
                            "Nao foi possivel retornar a tela de credenciado. Salvando e finalizando.",
                            level="ERROR",
                        )
                        self._salvar_planilha(final=True)
                        return

                    extraction_status["processed"] = idx
                except Exception as e:
                    self.handle_exception(e, f"Erro no processamento do item [{idx}] {valor}")
                    extraction_status["processed"] = idx - 1
                    self._salvar_planilha(final=True)
                    continue
        except Exception as e:
            self.handle_exception(e, "Erro critico no processamento da planilha")
        finally:
            self._salvar_planilha(final=True)

    def _abrir_pagina_inicial(self):
        self.add_log("Abrindo pagina inicial do sistema...")
        try:
            self.driver.get("https://online.fwcard.com.br/fwcard/f")
            time.sleep(2)
            self.add_log("Pagina inicial aberta.")
            return True
        except Exception as e:
            self.handle_exception(e, "Erro ao abrir pagina inicial")
            return False

    def _navegar_para_credenciado(self):
        self.add_log("Navegando para tela de credenciado...")
        try:
            self.wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="t_MenuNav_11i"]'))).click()
            time.sleep(0.2)
            self.wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="t_MenuNav_11_1i"]'))).click()
            time.sleep(0.2)
            self.wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="t_MenuNav_11_1_2i"]'))).click()
            time.sleep(0.2)
            self.wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="t_MenuNav_11_1_2_1i"]'))).click()
            if not self._esperar_tela_credenciado(timeout=8):
                raise TimeoutException("Tela de credenciado nao carregou")
            self.add_log("Tela de credenciado acessada.")
            return True
        except Exception as e:
            self.handle_exception(e, "Erro ao navegar para tela de credenciado")
            return False

    def _ir_para_tela_credenciado(self):
        if self._navegar_para_credenciado():
            return True
        self.add_log("Tentando recuperar a tela de credenciado pela pagina inicial...", level="WARNING")
        if self._abrir_pagina_inicial():
            return self._navegar_para_credenciado()
        return False

    def _esperar_tela_consulta_transacoes(self, timeout=12):
        try:
            WebDriverWait(self.driver, timeout).until(
                EC.presence_of_element_located((By.XPATH, self.xpaths["campo_transacao_data_inicial"]))
            )
            WebDriverWait(self.driver, timeout).until(
                EC.presence_of_element_located((By.XPATH, self.xpaths["campo_transacao_data_final"]))
            )
            WebDriverWait(self.driver, timeout).until(
                EC.presence_of_element_located((By.XPATH, self.xpaths["campo_transacao_codigo"]))
            )
            WebDriverWait(self.driver, timeout).until(
                EC.element_to_be_clickable((By.XPATH, self.xpaths["botao_transacao_pesquisar"]))
            )
            return True
        except Exception:
            return False

    def _navegar_para_consulta_transacoes(self):
        self.add_log("Navegando para a tela de consulta de transacoes...")
        try:
            self._click(
                self.wait.until(
                    EC.element_to_be_clickable((By.XPATH, self.xpaths["menu_consulta_principal"]))
                )
            )
            time.sleep(0.2)
            self._click(
                self.wait.until(
                    EC.element_to_be_clickable((By.XPATH, self.xpaths["menu_consulta_secundario"]))
                )
            )
            time.sleep(0.2)
            self._click(
                self.wait.until(
                    EC.element_to_be_clickable((By.XPATH, self.xpaths["menu_consulta_transacoes"]))
                )
            )
            if not self._esperar_tela_consulta_transacoes(timeout=15):
                raise TimeoutException("Tela de consulta de transacoes nao carregou")
            self.add_log("Tela de consulta de transacoes acessada.")
            return True
        except Exception as e:
            self.handle_exception(e, "Erro ao navegar para tela de consulta de transacoes")
            return False

    def _ir_para_tela_consulta_transacoes(self):
        if self._navegar_para_consulta_transacoes():
            return True
        self.add_log("Tentando recuperar a tela de consulta de transacoes pela pagina inicial...", level="WARNING")
        if self._abrir_pagina_inicial():
            return self._navegar_para_consulta_transacoes()
        return False

    def _limpar_e_preencher(self, xpath, valor):
        campo = self.wait.until(EC.presence_of_element_located((By.XPATH, xpath)))
        texto = str(valor)
        campo.click()
        time.sleep(0.1)
        campo.clear()
        time.sleep(0.1)
        campo.send_keys(Keys.CONTROL, "a")
        time.sleep(0.05)
        campo.send_keys(Keys.DELETE)
        time.sleep(0.1)
        self._type(texto, campo)
        return campo

    def _preencher_campo_com_confirmacao(self, xpath, valor, nome_campo, tentativas=3):
        texto = str(valor).strip()
        for tentativa in range(1, tentativas + 1):
            campo = self._limpar_e_preencher(xpath, texto)
            time.sleep(0.3)
            atual = (
                campo.get_attribute("value")
                or campo.get_attribute("textContent")
                or campo.text
                or ""
            ).strip()
            if atual == texto:
                return campo
            self.add_log(
                f"Campo {nome_campo} nao confirmou o valor na tentativa {tentativa}. Tentando novamente...",
                level="WARNING",
            )

        raise TimeoutException(f"Campo {nome_campo} nao manteve o valor informado")

    def _obter_texto_campo(self, xpath):
        try:
            el = self.driver.find_element(By.XPATH, xpath)
            val = (
                el.get_attribute("value")
                or el.text
                or el.get_attribute("textContent")
                or el.get_attribute("innerText")
                or ""
            )
            return val.strip()
        except Exception:
            return ""

    def _ha_processamento_ativo(self):
        try:
            return bool(
                self.driver.execute_script(
                    """
                    const selectors = [
                      '.u-Processing',
                      '.a-Spinner',
                      '.apex_loading_indicator',
                      '#apex_wait_overlay'
                    ];
                    return selectors.some((selector) =>
                      [...document.querySelectorAll(selector)].some((el) => !!(el.offsetWidth || el.offsetHeight || el.getClientRects().length))
                    );
                    """
                )
            )
        except Exception:
            return False

    def _aguardar_fim_processamento(self, timeout=90, exigir_inicio=False):
        fim = time.time() + timeout
        iniciou_processamento = not exigir_inicio
        estabilidade_sem_loading = 0

        while time.time() < fim:
            carregando = self._ha_processamento_ativo()
            if carregando:
                iniciou_processamento = True
                estabilidade_sem_loading = 0
            elif iniciou_processamento:
                estabilidade_sem_loading += 1
                if estabilidade_sem_loading >= 3:
                    return True

            time.sleep(0.5)

        if not exigir_inicio and not self._ha_processamento_ativo():
            return True
        raise TimeoutException("Tela nao finalizou o carregamento a tempo")

    def _elemento_visivel(self, xpath):
        try:
            elementos = self.driver.find_elements(By.XPATH, xpath)
            return any(el.is_displayed() for el in elementos)
        except Exception:
            return False

    def _obter_texto_se_visivel(self, xpath):
        try:
            el = self.driver.find_element(By.XPATH, xpath)
            if not el.is_displayed():
                return ""
            texto = (
                el.get_attribute("value")
                or el.text
                or el.get_attribute("textContent")
                or el.get_attribute("innerText")
                or ""
            )
            return texto.strip()
        except Exception:
            return ""

    def _sem_dados_transacoes(self):
        try:
            return bool(
                self.driver.execute_script(
                    """
                    const candidateSelectors = [
                      '#LISTA .nodatafound',
                      '.nodatafound',
                      'td.nodatafound',
                      '.a-IRR-noDataMsg'
                    ];
                    return candidateSelectors.some((selector) =>
                      [...document.querySelectorAll(selector)].some((el) => {
                        const text = (el.innerText || el.textContent || '').trim().toLowerCase();
                        return !!text && (text.includes('nenhum') || text.includes('no data'));
                      })
                    );
                    """
                )
            )
        except Exception:
            return False

    def _aguardar_totais_transacoes(self, timeout=90):
        fim = time.time() + timeout
        inicio = time.time()
        ultimo_qtd = None
        ultimo_valor = None
        estabilidade = 0
        iniciou_processamento = False

        while time.time() < fim:
            if self._ha_processamento_ativo():
                iniciou_processamento = True

            qtd = self._obter_texto_campo(self.xpaths["campo_transacao_total_qtd"])
            valor = self._obter_texto_campo(self.xpaths["campo_transacao_total_valor"])

            if qtd and valor and (iniciou_processamento or (time.time() - inicio) >= 2):
                if qtd == ultimo_qtd and valor == ultimo_valor:
                    estabilidade += 1
                else:
                    estabilidade = 1
                    ultimo_qtd = qtd
                    ultimo_valor = valor

                if estabilidade >= 3 and not self._ha_processamento_ativo():
                    return qtd, valor

            time.sleep(1)

        raise TimeoutException("Totais de transacoes nao carregaram a tempo")

    def _aguardar_data_ultima_transacao(self, timeout=90):
        fim = time.time() + timeout
        while time.time() < fim:
            if self._ha_processamento_ativo():
                time.sleep(0.5)
                continue

            data_ultima = self._obter_texto_se_visivel(self.xpaths["campo_transacao_data_ultima"])
            if data_ultima:
                return data_ultima

            if self._sem_dados_transacoes():
                return "Nao houve transacoes no periodo"

            time.sleep(0.5)

        raise TimeoutException("Resultado da ultima transacao nao carregou a tempo")

    def _consultar_data_ultima_transacao(self, codigo, data_inicial, data_final):
        self.add_log("Pesquisa enviada. Aguardando carregamento da consulta...")
        self._aguardar_fim_processamento(timeout=90, exigir_inicio=False)

        if self._sem_dados_transacoes():
            return {
                "CODIGO": codigo,
                "DATA_INICIAL": data_inicial,
                "DATA_FINAL": data_final,
                "DATA_ULTIMA_TRANSACAO": "Nao houve transacoes no periodo",
            }

        self._click(
            self.wait.until(
                EC.element_to_be_clickable((By.XPATH, self.xpaths["coluna_transacao_data_lancamento"]))
            )
        )
        self.add_log("Ordenando pela coluna Data Lancamento...")
        self._aguardar_fim_processamento(timeout=60, exigir_inicio=False)

        self._click(
            self.wait.until(
                EC.element_to_be_clickable((By.XPATH, self.xpaths["botao_transacao_ordenacao_desc"]))
            )
        )
        self.add_log("Aplicando ordenacao do mais recente para o mais antigo...")
        self._aguardar_fim_processamento(timeout=60, exigir_inicio=False)

        data_ultima = self._aguardar_data_ultima_transacao(timeout=60)
        return {
            "CODIGO": codigo,
            "DATA_INICIAL": data_inicial,
            "DATA_FINAL": data_final,
            "DATA_ULTIMA_TRANSACAO": data_ultima,
        }

    def _consultar_transacoes_item(self, consulta):
        codigo = str(consulta.get("codigo", "")).strip()
        data_inicial = str(consulta.get("data_inicial", "")).strip()
        data_final = str(consulta.get("data_final", "")).strip()

        self._preencher_campo_com_confirmacao(
            self.xpaths["campo_transacao_data_inicial"], data_inicial, "data inicial"
        )
        self._preencher_campo_com_confirmacao(
            self.xpaths["campo_transacao_data_final"], data_final, "data final"
        )
        self._preencher_campo_com_confirmacao(
            self.xpaths["campo_transacao_codigo"], codigo, "codigo do credenciado"
        )
        time.sleep(0.2)

        self._click(
            self.wait.until(
                EC.element_to_be_clickable((By.XPATH, self.xpaths["botao_transacao_pesquisar"]))
            )
        )

        if self.funcionalidade_transacao == "data_ultima_transacao":
            return self._consultar_data_ultima_transacao(codigo, data_inicial, data_final)

        self.add_log("Pesquisa enviada. Aguardando carregamento dos totais...")
        qtd, valor = self._aguardar_totais_transacoes()
        return {
            "CODIGO": codigo,
            "DATA_INICIAL": data_inicial,
            "DATA_FINAL": data_final,
            "QTD_TRANSACOES": qtd,
            "VALOR_TOTAL_TRANSACOES": valor,
        }

    def _processar_consultas_transacoes(self):
        consultas = self.filtros.get("consultas_transacoes", []) or []
        extraction_status["total"] = len(consultas)
        self.add_log(f"Total de consultas de transacoes: {len(consultas)}")

        if not consultas:
            self.add_log("Nenhuma consulta de transacoes informada.", level="WARNING")
            self._salvar_planilha(final=True)
            return

        for idx, consulta in enumerate(consultas, start=1):
            if extraction_status["request_stop"]:
                self.add_log("Interrupcao solicitada.")
                break

            if idx > 1:
                self.add_log("Abrindo novamente Produção > Consulta > Transação para o proximo codigo...")
                if not self._ir_para_tela_consulta_transacoes():
                    raise RuntimeError("Nao foi possivel retornar para a tela de consulta de transacoes")

            codigo = consulta.get("codigo", "")
            extraction_status["current_item"] = f"{idx}/{len(consultas)} - {codigo}"
            self.add_log("-" * 40)
            self.add_log(f"Consulta [{idx}/{len(consultas)}]: codigo {codigo}")

            try:
                dados = self._consultar_transacoes_item(consulta)
                self.dados_coletados.append(dados)
                if self.funcionalidade_transacao == "data_ultima_transacao":
                    self.add_log(
                        f"[{idx}] Codigo {codigo}: ultima transacao={dados['DATA_ULTIMA_TRANSACAO']}",
                        level="SUCCESS",
                    )
                else:
                    self.add_log(
                        f"[{idx}] Codigo {codigo}: qtd={dados['QTD_TRANSACOES']} | valor={dados['VALOR_TOTAL_TRANSACOES']}",
                        level="SUCCESS",
                    )
                extraction_status["processed"] = idx
                self._salvar_planilha(final=False)
            except Exception as e:
                self.handle_exception(e, f"Erro na consulta de transacoes do codigo {codigo}")
                extraction_status["processed"] = idx

        self._salvar_planilha(final=True)

    def _voltar_para_lista(self):
        try:
            self.driver.back()
            if self._esperar_tela_credenciado(timeout=6):
                return True
        except Exception as e:
            self.handle_exception(e, "Erro ao voltar para a lista")
        return self._ir_para_tela_credenciado()

    def _aplicar_filtros(self):
        self.add_log("Aplicando filtros...")

        def fill(xpath, value):
            if not value:
                return
            try:
                el = self.wait.until(EC.presence_of_element_located((By.XPATH, xpath)))
                el.clear()
                el.send_keys(str(value))
            except Exception as e:
                self.handle_exception(e, f"Erro preenchendo {xpath}")

        def select_by_xpath(option_xpath):
            try:
                self.driver.find_element(By.XPATH, option_xpath).click()
                time.sleep(0.1)
            except Exception as e:
                self.handle_exception(e, f"Erro selecionando {option_xpath}")

        f = self.filtros
        fill(self.xpaths["campo_codigo"], f.get("codigo"))
        fill(self.xpaths["campo_cnpj"], f.get("cnpj"))
        fill(self.xpaths["campo_nome_fantasia"], f.get("nome_fantasia"))
        fill(self.xpaths["campo_razao_social"], f.get("razao_social"))
        fill(self.xpaths["campo_cidade"], f.get("cidade"))

        sit = f.get("situacao", "")
        if sit == "APROVADO":
            select_by_xpath('//*[@id="P2100_ID_ESTADO_CREDENCIADO"]/option[2]')
        elif sit == "BLOQUEADO":
            select_by_xpath('//*[@id="P2100_ID_ESTADO_CREDENCIADO"]/option[3]')
        elif sit == "CADASTRO":
            select_by_xpath('//*[@id="P2100_ID_ESTADO_CREDENCIADO"]/option[4]')

        prod_idx = f.get("produto_index")
        if prod_idx:
            select_by_xpath(f'//*[@id="P2100_ID_PRODUTO"]/option[{prod_idx}]')

        uf = f.get("uf", "")
        if uf and uf in UF_INDEX and UF_INDEX[uf]:
            select_by_xpath(f'//*[@id="P2100_UF"]/option[{UF_INDEX[uf]}]')

        time.sleep(0.2)
        self.add_log("Filtros aplicados. Pesquisando...")
        self.wait.until(EC.element_to_be_clickable((By.XPATH, self.xpaths["botao_pesquisar"]))).click()
        time.sleep(0.8)

    def _aumentar_para_mil(self):
        try:
            self.add_log("Aumentando para 1000 linhas...")
            self.wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="LISTA_row_select"]')))
            self.driver.find_element(By.XPATH, '//*[@id="LISTA_row_select"]/option[9]').click()
            time.sleep(0.8)
            self.add_log("Linhas aumentadas para 1000.")
        except Exception as e:
            self.handle_exception(e, "Erro ao aumentar linhas")

    def _contar_credenciados(self):
        try:
            rows = self.driver.find_elements(By.XPATH, self.xpaths["linhas_resultado"])
            return len([row for row in rows if row.text.strip()])
        except Exception:
            return 0

    def _extrair_dados_credenciado(self):
        dados = {}
        self._esperar_detalhe_credenciado(timeout=8)

        def get(key, xpath):
            try:
                el = self.driver.find_element(By.XPATH, xpath)
                val = (
                    el.get_attribute("value")
                    or el.text
                    or el.get_attribute("textContent")
                    or el.get_attribute("innerText")
                    or ""
                )
                dados[key] = val.strip()
            except Exception:
                dados[key] = ""

        get("CODIGO", self.xpaths["campo_detalhe_codigo"])
        get("CNPJ", self.xpaths["campo_detalhe_cnpj"])
        get("NOME_FANTASIA", self.xpaths["campo_detalhe_nome_fantasia"])
        get("RAZAO_SOCIAL", self.xpaths["campo_detalhe_razao_social"])
        get("CIDADE", self.xpaths["campo_detalhe_cidade"])
        get("ENDERECO", self.xpaths["campo_detalhe_endereco"])
        get("BAIRRO", self.xpaths["campo_detalhe_bairro"])
        get("CEP", self.xpaths["campo_detalhe_cep"])
        get("EMAIL", self.xpaths["campo_detalhe_email"])
        get("CONCILIACAO", self.xpaths["campo_detalhe_conciliacao"])
        get("CELULAR", self.xpaths["campo_detalhe_celular"])
        get("TELEFONE", self.xpaths["campo_detalhe_telefone"])
        get("DATA_CADASTRO", self.xpaths["campo_detalhe_data_cadastro"])
        self.add_log(
            "Resumo da coleta: "
            f"CODIGO={dados.get('CODIGO','')[:20]} | "
            f"CNPJ={dados.get('CNPJ','')[:25]} | "
            f"FANTASIA={dados.get('NOME_FANTASIA','')[:40]}",
            level="INFO",
        )
        return dados

    def _dados_foram_extraidos(self, dados):
        chaves_principais = ("CODIGO", "CNPJ", "RAZAO_SOCIAL", "NOME_FANTASIA", "EMAIL", "TELEFONE", "CELULAR")
        return any((dados.get(chave) or "").strip() for chave in chaves_principais)

    def iniciar(self):
        try:
            extraction_status["running"] = True
            extraction_status["start_time"] = datetime.now()
            extraction_status["end_time"] = None
            extraction_status["logs"] = []
            self.add_log("Automacao Extrator iniciada.")

            self._iniciar_driver()
            self._realizar_login()
            self.add_log("Aguardando carregamento da pagina inicial...")
            time.sleep(2)

            if self.modo_consulta_transacoes:
                if not self._navegar_para_consulta_transacoes():
                    raise RuntimeError("Nao foi possivel abrir a tela de consulta de transacoes")
                self._processar_consultas_transacoes()
                return

            if not self._navegar_para_credenciado():
                raise RuntimeError("Nao foi possivel abrir a tela de credenciado")

            if self.filtros.get("planilha_codigos"):
                self._processar_lista_planilha()
            else:
                self._aplicar_filtros()
                self._aumentar_para_mil()

                total = self._contar_credenciados()
                extraction_status["total"] = total
                self.add_log(f"Total de credenciados encontrados: {total}")

                if total == 0:
                    self.add_log("Nenhum credenciado encontrado com os filtros aplicados.", level="WARNING")
                    self._salvar_planilha(final=True)
                    return

                for idx in range(1, total + 1):
                    if extraction_status["request_stop"]:
                        self.add_log("Interrupcao solicitada.")
                        break

                    try:
                        self.add_log("-" * 40)
                        self.add_log(f"Credenciado [{idx}/{total}]...")
                        extraction_status["current_item"] = f"{idx}/{total}"

                        xp_link = self.xpaths["link_resultado_por_indice"].format(idx=idx)
                        link = self.wait.until(EC.element_to_be_clickable((By.XPATH, xp_link)))

                        cod_text = ""
                        try:
                            cod_text = self.driver.find_element(
                                By.XPATH, self.xpaths["codigo_resultado_por_indice"].format(idx=idx)
                            ).text.strip()
                        except Exception:
                            pass

                        self.add_log(f"Abrindo credenciado: {cod_text or idx}")
                        if not self._abrir_link_resultado(link):
                            self.add_log(
                                f"Nao foi possivel abrir o detalhe do credenciado {cod_text or idx}.",
                                level="WARNING",
                            )
                            extraction_status["processed"] = idx
                            continue

                        dados = self._extrair_dados_credenciado()
                        if not self._dados_foram_extraidos(dados):
                            self.add_log(
                                f"Detalhe do credenciado {cod_text or idx} abriu, mas os campos vieram vazios.",
                                level="WARNING",
                            )
                            if not self._voltar_para_lista():
                                raise RuntimeError("Nao foi possivel retornar para a lista apos detalhe vazio")
                            extraction_status["processed"] = idx
                            continue

                        self.dados_coletados.append(dados)
                        self.add_log(
                            f"[{idx}] {dados.get('NOME_FANTASIA') or dados.get('RAZAO_SOCIAL') or cod_text or idx} coletado.",
                            level="SUCCESS",
                        )

                        if not self._voltar_para_lista():
                            raise RuntimeError("Nao foi possivel retornar para a lista apos a coleta")

                        extraction_status["processed"] = idx
                    except Exception as e:
                        self.handle_exception(e, f"Erro no credenciado [{idx}]")
                        self._voltar_para_lista()

                self._salvar_planilha(final=True)
        except Exception as e:
            self.handle_exception(e, "Erro critico")
            self._salvar_planilha(final=True)
        finally:
            extraction_status["end_time"] = datetime.now()
            self._finalizar()

    def _salvar_planilha(self, final=False):
        try:
            self.add_log(f"Tentando salvar planilha. Dados coletados: {len(self.dados_coletados)}")
            os.makedirs("outputs_planilha", exist_ok=True)

            if final:
                ts = datetime.now().strftime("%Y%m%d_%H%M%S")
                arquivo_xlsx = f"outputs_planilha/credenciados_{ts}.xlsx"
                arquivo_csv = f"outputs_planilha/credenciados_{ts}.csv"
            else:
                arquivo_xlsx = "outputs_planilha/credenciados_temp.xlsx"
                arquivo_csv = "outputs_planilha/credenciados_temp.csv"

            df = pd.DataFrame(self.dados_coletados, columns=self.colunas_saida)

            try:
                df.to_excel(arquivo_xlsx, index=False)
                arquivo_salvo = arquivo_xlsx
            except Exception as e_excel:
                self.add_log(f"Falha ao salvar XLSX, usando CSV: {str(e_excel)[:140]}", level="WARNING")
                df.to_csv(arquivo_csv, index=False, encoding="utf-8-sig")
                arquivo_salvo = arquivo_csv

            if final:
                extraction_status["output_file"] = arquivo_salvo
                self.add_log(f"Planilha final salva: {arquivo_salvo}", level="SUCCESS")
            else:
                self.add_log(f"Salvamento intermediario: {arquivo_salvo}", level="INFO")
        except Exception as e:
            self.handle_exception(e, "Erro ao salvar planilha")

    def _finalizar(self):
        if self.driver:
            try:
                self.driver.quit()
            except Exception:
                pass

        tempo = "N/A"
        if extraction_status["start_time"] and extraction_status["end_time"]:
            tempo = str(extraction_status["end_time"] - extraction_status["start_time"]).split(".")[0]

        self.add_log(
            f"Concluido. Extraidos: {extraction_status['processed']}/{extraction_status['total']} | Tempo: {tempo}",
            level="SUCCESS",
        )
        extraction_status["running"] = False
        extraction_status["request_stop"] = False
