from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException, TimeoutException
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service as ChromeService
from webdriver_manager.chrome import ChromeDriverManager
from datetime import datetime
import pandas as pd
import threading
import traceback
import random
import time
import os

# ─────────────────────────────────────────────────────────────────────────────
# TEMPOS (segundos) — ajuste aqui para calibrar velocidade vs estabilidade
# ─────────────────────────────────────────────────────────────────────────────
T = {
    'after_nav_click':   0.5,   # após cada clique no menu de navegação
    'after_login':       1.5,   # após submeter login
    'after_search':      0.8,   # após clicar em PESQUISAR
    'after_link_click':  1.5,   # após abrir contrato do produto
    'after_plan_click':  1.0,   # após clicar no plano existente
    'after_save_plan':   1.5,   # após salvar plano (OK/CRIAR)
    'after_approve':     1.0,   # após aprovar
    'after_back':        1.0,   # após clicar em VOLTAR
    'after_back_cred':   0.8,   # após voltar tela credenciado
    'type_min':          0.02,  # digitação mínima por char
    'type_max':          0.06,  # digitação máxima por char
    'field_clear':       0.15,  # após clear()
    'field_after':       0.2,   # após preencher campo
    'select_after':      0.2,   # após selecionar option
}
# ─────────────────────────────────────────────────────────────────────────────

def _novo_status():
    return {
        'running': False, 'total': 0, 'processed': 0,
        'current_item': None, 'logs': [],
        'start_time': None, 'end_time': None,
        'request_stop': False,
        'output_file': None,
        'log_file': None,
        'error': None,
        'resultado_linhas': [],
    }

automation_status  = _novo_status()
automation_status2 = _novo_status()
_STATUS_SLOTS = [automation_status, automation_status2]


class WebAppAlteradorTaxa:
    def __init__(self, slot=0, single_mode=True, dados_credenciado=None,
                 dados_planilha=None, arquivo_entrada=None, arquivo_saida='resultado',
                 tipo_pesquisa='COD', dados_login=None):
        self.slot              = slot
        self.status            = _STATUS_SLOTS[slot]
        self.single_mode       = single_mode
        self.dados_credenciado = dados_credenciado or {}
        self.dados_planilha    = dados_planilha
        self.arquivo_entrada   = arquivo_entrada
        self.arquivo_saida     = arquivo_saida
        self.tipo_pesquisa     = tipo_pesquisa.upper()
        self.dados_login       = dados_login or {}
        self.driver            = None
        self.wait              = None
        self.actions           = None
        self._produto_atual    = None

    # ── LOG ──────────────────────────────────────────────────────────────────
    def add_log(self, message, level="INFO"):
        ts     = datetime.now().strftime('%H:%M:%S')
        prefix = {"ERROR": "ERRO", "WARNING": "AVISO", "SUCCESS": "SUCESSO"}.get(level, "")
        entry  = f"[{ts}] {f'{prefix}: ' if prefix else ''}{message}"
        self.status['logs'].append(entry)
        print(f"[Slot{self.slot}] {entry}")
        try:
            if self.status.get('log_file'):
                with open(self.status['log_file'], 'a', encoding='utf-8') as f:
                    f.write(entry + '\n')
        except:
            pass

    def handle_exception(self, e, context=""):
        s = str(e)
        if   "stale element reference"   in s: msg = "Elemento desatualizado na página"
        elif "no such element"           in s: msg = "Elemento não encontrado"
        elif "timeout"          in s.lower():  msg = "Tempo de espera excedido"
        elif "element click intercepted" in s: msg = "Clique interceptado"
        elif "element not interactable"  in s: msg = "Elemento não interagível"
        else: msg = s.split('\n')[0][:120]
        self.add_log(f"{context}: {msg}", level="ERROR")
        try:
            with open(f"debug_taxa_slot{self.slot}.txt", 'a', encoding='utf-8') as f:
                f.write(f"\n--- {datetime.now()} ---\n{context}\n{s}\n")
                traceback.print_exc(file=f)
        except:
            pass

    # ── DRIVER ───────────────────────────────────────────────────────────────
    def _iniciar_driver(self):
        self.add_log("Iniciando navegador...")
        try:
            profile = os.path.join(os.getcwd(), f"selenium-profile-taxa{self.slot}")
            os.makedirs(profile, exist_ok=True)
            opts = Options()
            opts.add_argument(f"--user-data-dir={profile}")
            opts.add_argument("--start-maximized")
            opts.add_argument("--disable-extensions")
            opts.add_argument("--disable-popup-blocking")
            opts.add_argument("--disable-blink-features=AutomationControlled")
            opts.binary_location = r"C:\BACKUP_EMANUEL\Program Files\Google\Chrome\Application\chrome.exe"
            self.driver  = webdriver.Chrome(service=ChromeService(ChromeDriverManager().install()), options=opts)
            self.wait    = WebDriverWait(self.driver, 15)
            self.actions = ActionChains(self.driver)
            self.driver.implicitly_wait(5)          # reduzido de 8 → 5
            self.driver.get("about:blank")
            time.sleep(1.5)                         # reduzido de 2 → 1.5
            self._navegar_para_url_com_pyautogui("https://online.fwcard.com.br/fwcard/f?p=380:100:")
            time.sleep(2.5)                         # reduzido de 3 → 2.5
            self.add_log("Navegador iniciado.")
        except Exception as e:
            self.handle_exception(e, "Erro ao iniciar navegador")
            raise

    def _navegar_para_url_com_pyautogui(self, url):
        try:
            import pyautogui, keyboard
            pyautogui.click(100, 100); time.sleep(.4)
            pyautogui.hotkey('ctrl', 'l'); time.sleep(.4)
            pyautogui.hotkey('ctrl', 'a'); time.sleep(.2)
            pyautogui.press('delete'); time.sleep(.2)
            for char in url:
                keyboard.write(char)
                time.sleep(random.uniform(0.02, 0.06))
            time.sleep(.2)
            pyautogui.press('enter')
            time.sleep(2.5)
        except Exception as e:
            self.handle_exception(e, "Erro ao navegar")

    # ── LOGIN ─────────────────────────────────────────────────────────────────
    def _realizar_login(self):
        self.add_log("Realizando login...")
        try:
            u = self.wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="P100_USERNAME"]')))
            self._type(self.dados_login['login_infox'], u)
            p = self.wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="P100_PASSWORD"]')))
            self._type(self.dados_login['senha_infox'], p)
            btn = self.wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="LOGIN"]/span')))
            try:
                cap = WebDriverWait(self.driver, 4).until(
                    EC.presence_of_element_located((By.XPATH, '//*[@id="uLogin"]/div[5]/div/div/div[1]')))
                self._click(cap)
            except: pass
            self._click(btn)
            self.add_log("Login realizado.")
            time.sleep(T['after_login'])
        except Exception as e:
            self.handle_exception(e, "Erro no login")
            raise

    def _type(self, text, el):
        el.clear(); time.sleep(T['field_clear'])
        for c in str(text):
            el.send_keys(c)
            time.sleep(random.uniform(T['type_min'], T['type_max']))

    def _click(self, el):
        self.driver.execute_script("arguments[0].scrollIntoView({block:'center'});", el)
        time.sleep(.15)
        ActionChains(self.driver).move_to_element(el).pause(.08).perform()
        el.click(); time.sleep(.15)

    # ── INICIAR ───────────────────────────────────────────────────────────────
    def iniciar(self):
        ts = datetime.now().strftime('%Y%m%d_%H%M%S')
        os.makedirs('logs_taxa', exist_ok=True)
        os.makedirs('resultados_taxa', exist_ok=True)

        self.status['log_file']         = f"logs_taxa/taxa_slot{self.slot}_{ts}.txt"
        self.status['output_file']      = None
        self.status['error']            = None
        self.status['resultado_linhas'] = []

        try:
            self.add_log('=' * 50)
            self.add_log(f"Alterador de Taxa iniciado — Slot {self.slot}")
            self.add_log('=' * 50)
            self.status['running']    = True
            self.status['start_time'] = datetime.now()

            if not self.single_mode:
                if self.dados_planilha is None:
                    self.dados_planilha = pd.read_excel(self.arquivo_entrada, dtype=str)
                self.status['total'] = len(self.dados_planilha)
            else:
                self.status['total'] = 1

            self._iniciar_driver()
            self._realizar_login()

            # CORREÇÃO: _navegar_para_pesquisa_credenciado() é chamada
            # UMA vez aqui para single_mode, ou dentro do loop para lote.
            # Não é mais chamada em dois lugares.
            if self.single_mode:
                self._navegar_para_pesquisa_credenciado()
                self._processar_individual()
            else:
                self._processar_lote()

        except Exception as e:
            self._registrar_erro_critico(e)
        finally:
            self.status['running']  = False
            self.status['end_time'] = datetime.now()
            self._salvar_resultado_final(ts)
            self._finalizar_driver()

    # ── ERRO CRÍTICO ──────────────────────────────────────────────────────────
    def _registrar_erro_critico(self, e):
        self.handle_exception(e, "ERRO CRÍTICO — automação encerrada")
        proc  = self.status['processed']
        total = self.status['total']
        cods  = [str(r.get('cod', '')) for r in self.status['resultado_linhas']]
        ordem = self._detectar_ordem(cods)
        descricao = (
            f"ERRO FATAL no credenciado #{proc+1} de {total}. "
            f"Último produto: '{self._produto_atual or 'desconhecido'}'. "
            f"Ordem: {ordem}. Erro: {str(e).split(chr(10))[0][:200]}"
        )
        self.status['error'] = descricao
        self.add_log(f"⚠ {descricao}", level="ERROR")

    def _detectar_ordem(self, cods):
        nums = []
        for c in cods:
            try: nums.append(int(c))
            except: pass
        if len(nums) < 2:
            return "indeterminada (poucos registros)"
        if all(nums[i] <= nums[i+1] for i in range(len(nums)-1)):
            return f"crescente (do {nums[0]} ao {nums[-1]})"
        amostra = ' → '.join(str(n) for n in nums[:5])
        return f"embaralhada (exemplo: {amostra}...)"

    # ── NAVEGAÇÃO ─────────────────────────────────────────────────────────────
    def _navegar_para_pesquisa_credenciado(self):
        """
        Percorre o menu até chegar na tela de pesquisa de credenciado.
        Chamada UMA única vez por credenciado no modo lote,
        ou UMA vez antes de _processar_individual no modo single.
        """
        self.add_log("Navegando para pesquisa...")
        try:
            for xp in ['//*[@id="t_MenuNav_11i"]', '//*[@id="t_MenuNav_11_1i"]',
                       '//*[@id="t_MenuNav_11_1_2i"]', '//*[@id="t_MenuNav_11_1_2_1i"]']:
                self.wait.until(EC.element_to_be_clickable((By.XPATH, xp))).click()
                time.sleep(T['after_nav_click'])
            self.add_log("Navegação concluída.")
        except Exception as e:
            self.handle_exception(e, "Erro na navegação")
            raise

    def _pesquisar_credenciado(self, codigo):
        self.add_log(f"Pesquisando: {codigo} ({self.tipo_pesquisa})")
        try:
            if self.tipo_pesquisa == "COD":
                xp = '/html/body/div[1]/div/form/div/div[1]/section/div[2]/section[1]/div[2]/table/tbody/tr[1]/td[1]/span/div/input'
            else:
                xp = '/html/body/div[1]/div/form/div/div[1]/section/div[2]/section[1]/div[2]/table/tbody/tr[1]/td[2]/span/input'
            campo = self.wait.until(EC.presence_of_element_located((By.XPATH, xp)))
            campo.clear()
            campo.send_keys(str(codigo))
            self.wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="PESQUISAR"]/span'))).click()
            time.sleep(T['after_search'])
        except Exception as e:
            self.handle_exception(e, f"Erro ao pesquisar {codigo}")
            raise

    def _acessar_contratos(self):
        try:
            self.wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="VISUALIZAR_CONTRATOS"]/span'))).click()
            time.sleep(0.8)
            self.add_log("Contratos abertos.")
        except Exception as e:
            self.handle_exception(e, "Erro ao acessar contratos")
            raise

    def _voltar_tela_credenciado(self):
        try:
            btn = self.wait.until(EC.element_to_be_clickable((By.XPATH,
                '/html/body/div[2]/div/form/div/div[1]/section/div[1]/span/button/span')))
            self.driver.execute_script("arguments[0].click();", btn)
            time.sleep(T['after_back_cred'])
        except Exception as e:
            self.handle_exception(e, "Erro ao voltar")

    def _voltar_para_contratos(self):
        """Volta da tela do plano para a lista de contratos do produto."""
        try:
            btn = self.wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="VOLTAR"]/span')))
            self.driver.execute_script("arguments[0].click();", btn)
            time.sleep(T['after_back'])
            self.add_log("Retornou para lista de contratos.")
        except Exception as e:
            self.handle_exception(e, "Erro ao voltar para contratos")

    # ── MODO INDIVIDUAL ───────────────────────────────────────────────────────
    def _processar_individual(self):
        cod = self.dados_credenciado['cod_cnpj']
        self.status['current_item'] = cod
        self.add_log('─' * 40)
        self.add_log(f"Processando individual: {cod}")
        try:
            self._pesquisar_credenciado(cod)
            nome = self._obter_nome_fantasia()
            self._acessar_contratos()
            periodicidade = self.dados_credenciado.get('periodicidade', '').upper()
            dados = {
                'cod': cod, 'n_fantasia': nome,
                'taxa': self._formatar_taxa(self.dados_credenciado.get('taxa', '')),
                'periodicidade': periodicidade,
                'pag_dias': int(self.dados_credenciado['pag_dias']),
                'dia1': int(self.dados_credenciado.get('dia1') or 0),
                'dia2': self.dados_credenciado.get('dia2', '0'),
                'bancarizar': self.dados_credenciado.get('bancarizar', 'NÃO').upper(),
                'tipo_dia': self.dados_credenciado.get('tipo_dia', 'INDIFERENTE').upper(),
                'dias_semana': self.dados_credenciado.get('dias_semana', []),
            }
            self._processar_produtos(dados)
            self._voltar_tela_credenciado()
            self._registrar_linha(dados, sucesso=True)
            self.status['processed'] = 1
            self.add_log(f"✓ Concluído: {cod}", level="SUCCESS")
        except Exception as e:
            self.handle_exception(e, f"Erro no individual {cod}")
            self._registrar_linha({'cod': cod, 'n_fantasia': 'N/A'}, sucesso=False, erro=str(e))
            raise

    # ── MODO LOTE ─────────────────────────────────────────────────────────────
    def _processar_lote(self):
        df = self.dados_planilha
        for idx in range(len(df)):
            if self.status['request_stop']:
                self.add_log("Interrupção solicitada.")
                break
            try:
                dados = self._extrair_dados_registro(idx)
            except Exception as e:
                self.handle_exception(e, f"Erro ao extrair linha {idx}")
                continue

            cod  = str(dados['cod'])
            nome = dados.get('n_fantasia', '')
            self.status['current_item'] = cod
            self.add_log('─' * 40)
            self.add_log(f"[{idx+1}/{self.status['total']}] Credenciado: {cod} | {nome}")

            try:
                # CORREÇÃO: navega UMA única vez por credenciado no modo lote
                self._navegar_para_pesquisa_credenciado()
                self._pesquisar_credenciado(cod)
                self._acessar_contratos()
                self._processar_produtos(dados)
                self._voltar_tela_credenciado()
                self._registrar_linha(dados, sucesso=True)
                self.status['processed'] += 1
                self.add_log(f"✓ Credenciado {cod} concluído!", level="SUCCESS")
            except Exception as e:
                self.handle_exception(e, f"Erro ao processar {cod}")
                self._registrar_linha(dados, sucesso=False, erro=str(e).split('\n')[0])
                raise

    # ── PRODUTOS ──────────────────────────────────────────────────────────────
    def _processar_produtos(self, dados):
        contador = 1
        total_ok = 0
        self._produto_atual = None
        self.add_log("Processando produtos...")

        while True:
            if self.status['request_stop']:
                break
            nome_el = self.driver.find_elements(By.XPATH,
                f'/html/body/div[2]/div/form/div/div[1]/section/div[2]/section[2]/div[2]/table'
                f'/tbody[2]/tr/td/table/tbody/tr[{contador}]/td[3]/a')
            if not nome_el or not nome_el[0].text:
                break

            nome_prod = nome_el[0].text
            self._produto_atual = nome_prod
            self.add_log(f"Produto [{contador}]: {nome_prod}")

            link = self.driver.find_elements(By.XPATH,
                f'/html/body/div[2]/div/form/div/div[1]/section/div[2]/section[2]/div[2]/table'
                f'/tbody[2]/tr/td/table/tbody/tr[{contador}]/td[2]/a')
            if not link:
                break
            try:
                link[0].click()
                time.sleep(T['after_link_click'])

                # ── Verificar estado ANTES de qualquer alteração ──────────────
                estado = self._obter_estado_produto()
                self.add_log(f"Estado do produto: {estado}")

                if estado == 'BLOQUEADO':
                    self.add_log(f"Produto [{contador}] '{nome_prod}' BLOQUEADO — pulando sem alteração.", level="WARNING")
                    self._voltar_para_contratos()
                    contador += 1
                    continue

                # ── Alterar plano de compra ───────────────────────────────────
                plano_clicado = self._clicar_plano_existente()
                if plano_clicado:
                    self._atualizar_plano_existente(dados)
                else:
                    self._criar_novo_plano_compra(dados)

                # ── Verificar estado APÓS salvar o plano ─────────────────────
                # - DEFINIÇÃO → plano foi alterado, precisa aprovar
                # - APROVADO  → plano não mudou (já estava correto), pula aprovação
                estado = self._obter_estado_produto()
                if estado == 'DEFINIÇÃO':
                    self.add_log('Estado: DEFINIÇÃO — aprovando...')
                    self._aprovar_produto()
                else:
                    self.add_log(f'Estado: {estado} — já aprovado, sem alteração necessária.')

                self._voltar_para_contratos()

                total_ok += 1
                self.add_log(f"✓ Produto [{contador}] '{nome_prod}' alterado!", level="SUCCESS")
                contador += 1

            except Exception as e:
                self.handle_exception(e, f"Erro produto #{contador} '{self._produto_atual}'")
                raise

        self.add_log(f"Produtos: {total_ok} alterados de {contador-1} encontrados.")

    def _obter_estado_produto(self):
        """
        Lê o campo P2111_ESTADO após salvar o plano de compra.
        Retorna 'APROVADO', 'DEFINIÇÃO' ou o texto que estiver no campo.
        """
        try:
            el = WebDriverWait(self.driver, 5).until(
                EC.presence_of_element_located((By.XPATH, '//*[@id="P2111_ESTADO"]'))
            )
            estado = (el.get_attribute('value') or el.text or '').strip().upper()
            self.add_log(f"Estado do produto: '{estado}'")
            return estado
        except Exception as e:
            self.handle_exception(e, "Não foi possível ler P2111_ESTADO — assumindo DEFINIÇÃO")
            return 'DEFINIÇÃO'  # seguro: tenta aprovar se não conseguir ler

    def _clicar_plano_existente(self):
        """
        Na tela de APROVAÇÃO CONTRATO, clica no link do primeiro plano de compra da tabela.
        Retorna True se encontrou e clicou, False se não há plano (precisa criar novo).
        """
        xp = ('//*[@id="report_PLANOS_DE_COMPRA"]/tbody[2]/tr/td/table/tbody/tr/td[2]/a')
        try:
            plano = WebDriverWait(self.driver, 5).until(
                EC.element_to_be_clickable((By.XPATH, xp))
            )
            self.driver.execute_script("arguments[0].click();", plano)
            time.sleep(T['after_plan_click'])
            self.add_log("Plano existente encontrado — acessando.")
            return True
        except (NoSuchElementException, TimeoutException):
            self.add_log("Nenhum plano na lista — será criado novo.")
            return False

    # ── CAMPOS DE FORMULÁRIO ──────────────────────────────────────────────────
    def _atualizar_campo(self, xpath, valor):
        el = self.wait.until(EC.element_to_be_clickable((By.XPATH, xpath)))
        el.clear(); time.sleep(T['field_clear'])
        el.send_keys(str(valor))
        time.sleep(T['field_after'])

    def _selecionar_combo(self, xpath, valor=None, texto=None, contexto="combo"):
        sel = self.wait.until(EC.presence_of_element_located((By.XPATH, xpath)))
        self.driver.execute_script("arguments[0].scrollIntoView({block:'center'});", sel)
        time.sleep(0.2)

        try:
            self.driver.execute_script("arguments[0].click();", sel)
            time.sleep(0.2)
        except Exception:
            pass

        select = Select(sel)
        selecionado = False

        if valor is not None:
            valor_str = str(valor).strip()
            try:
                select.select_by_value(valor_str)
                selecionado = True
            except Exception:
                for opt in sel.find_elements(By.TAG_NAME, 'option'):
                    opt_value = (opt.get_attribute('value') or '').strip()
                    opt_texto = opt.text.strip().upper()
                    if opt_value == valor_str or opt_texto == valor_str.upper():
                        opt.click()
                        selecionado = True
                        break

        if not selecionado and texto is not None:
            texto_str = str(texto).strip()
            try:
                select.select_by_visible_text(texto_str)
                selecionado = True
            except Exception:
                for opt in sel.find_elements(By.TAG_NAME, 'option'):
                    if texto_str.upper() in opt.text.strip().upper():
                        opt.click()
                        selecionado = True
                        break

        if not selecionado:
            alvo = str(valor if valor is not None else texto).strip()
            raise ValueError(f"Valor '{alvo}' não encontrado em {contexto}")

        self.driver.execute_script("""
            const el = arguments[0];
            el.dispatchEvent(new Event('input', { bubbles: true }));
            el.dispatchEvent(new Event('change', { bubbles: true }));
            el.dispatchEvent(new Event('blur', { bubbles: true }));
            if (window.apex && el.id) {
                try { apex.item(el.id).setValue(el.value); } catch (e) {}
            }
        """, sel)
        time.sleep(T['select_after'])
        return sel

    def _definir_item_apex(self, item_id, valor):
        try:
            return self.driver.execute_script("""
                const itemId = arguments[0];
                const value = String(arguments[1]);
                if (!window.apex || !apex.item) return false;
                try {
                    apex.item(itemId).setValue(value);
                    const el = document.getElementById(itemId);
                    if (el) {
                        el.dispatchEvent(new Event('input', { bubbles: true }));
                        el.dispatchEvent(new Event('change', { bubbles: true }));
                        el.dispatchEvent(new Event('blur', { bubbles: true }));
                    }
                    return true;
                } catch (e) {
                    return false;
                }
            """, item_id, valor)
        except Exception:
            return False

    def _ler_valor_item(self, item_id):
        try:
            valor = self.driver.execute_script("""
                const itemId = arguments[0];
                const el = document.getElementById(itemId);
                if (el) {
                    const raw = (el.value ?? el.getAttribute('value') ?? el.textContent ?? '').toString().trim();
                    if (raw) return raw;
                }
                if (window.apex && apex.item) {
                    try {
                        const raw = apex.item(itemId).getValue();
                        return (raw ?? '').toString().trim();
                    } catch (e) {}
                }
                return '';
            """, item_id)
            return str(valor).strip()
        except Exception:
            return ''

    def _listar_elementos_relacionados(self, item_id):
        try:
            relacionados = self.driver.execute_script("""
                const itemId = arguments[0].toUpperCase();
                return Array.from(document.querySelectorAll('[id],[name]'))
                    .filter(node => {
                        const id = (node.id || '').toUpperCase();
                        const name = (node.getAttribute('name') || '').toUpperCase();
                        return id.includes(itemId) || name.includes(itemId);
                    })
                    .slice(0, 12)
                    .map(node => ({
                        id: node.id || '',
                        name: node.getAttribute('name') || '',
                        tag: (node.tagName || '').toLowerCase(),
                        type: node.getAttribute('type') || '',
                        value: (node.value ?? node.getAttribute('value') ?? '').toString().trim()
                    }));
            """, item_id)
            return relacionados or []
        except Exception:
            return []

    def _forcar_valor_em_relacionados(self, item_id, valor):
        try:
            return self.driver.execute_script("""
                const itemId = arguments[0].toUpperCase();
                const originalId = arguments[0];
                const value = String(arguments[1]);
                let alterou = false;

                const candidatos = Array.from(document.querySelectorAll('[id],[name]')).filter(node => {
                    const id = (node.id || '').toUpperCase();
                    const name = (node.getAttribute('name') || '').toUpperCase();
                    return id.includes(itemId) || name.includes(itemId);
                });

                for (const el of candidatos) {
                    const tag = (el.tagName || '').toLowerCase();
                    if (!['input', 'select', 'textarea'].includes(tag)) continue;

                    try {
                        if (tag === 'select') {
                            const opts = Array.from(el.options || []);
                            const match = opts.find(opt => {
                                const optValue = (opt.value || '').trim();
                                const optText = (opt.text || '').trim();
                                const optDigits = optText.replace(/\\D/g, '');
                                const valueDigits = value.replace(/\\D/g, '');
                                return optValue === value || optText === value || (optDigits && valueDigits && optDigits === valueDigits);
                            });
                            if (match) {
                                el.value = match.value;
                                alterou = true;
                            }
                        } else {
                            el.value = value;
                            alterou = true;
                        }

                        el.dispatchEvent(new Event('input', { bubbles: true }));
                        el.dispatchEvent(new Event('change', { bubbles: true }));
                        el.dispatchEvent(new Event('blur', { bubbles: true }));
                    } catch (e) {}
                }

                if (window.apex && apex.item) {
                    try { apex.item(originalId).setValue(value); } catch (e) {}
                }
                return alterou;
            """, item_id, valor)
        except Exception:
            return False

    def _selecionar_combo_por_rotulo(self, rotulo, valor, contexto):
        valor_str = str(valor).strip()
        xpaths = [
            f"//td[normalize-space()='{rotulo}:' or normalize-space()='* {rotulo}:' or contains(normalize-space(), '{rotulo}:')]/following-sibling::td[1]//select[1]",
            f"//label[normalize-space()='{rotulo}' or normalize-space()='{rotulo}:']/following-sibling::select[1]",
        ]

        for xpath in xpaths:
            try:
                sel = WebDriverWait(self.driver, 2).until(
                    EC.presence_of_element_located((By.XPATH, xpath))
                )
                self.driver.execute_script("arguments[0].scrollIntoView({block:'center'});", sel)
                time.sleep(0.2)

                select = Select(sel)
                selecionado = False

                try:
                    select.select_by_value(valor_str)
                    selecionado = True
                except Exception:
                    pass

                if not selecionado:
                    try:
                        select.select_by_visible_text(valor_str)
                        selecionado = True
                    except Exception:
                        pass

                if not selecionado:
                    for opt in sel.find_elements(By.TAG_NAME, 'option'):
                        opt_value = (opt.get_attribute('value') or '').strip()
                        opt_text = (opt.text or '').strip()
                        opt_digits = ''.join(ch for ch in opt_text if ch.isdigit())
                        value_digits = ''.join(ch for ch in valor_str if ch.isdigit())
                        if opt_value == valor_str or opt_text == valor_str or (opt_digits and value_digits and opt_digits == value_digits):
                            opt.click()
                            selecionado = True
                            break

                if not selecionado:
                    continue

                self.driver.execute_script("""
                    const el = arguments[0];
                    el.dispatchEvent(new Event('input', { bubbles: true }));
                    el.dispatchEvent(new Event('change', { bubbles: true }));
                    el.dispatchEvent(new Event('blur', { bubbles: true }));
                """, sel)
                time.sleep(T['select_after'])

                selecionada = Select(sel).first_selected_option
                atual_valor = (selecionada.get_attribute('value') or '').strip()
                atual_texto = (selecionada.text or '').strip()

                if (
                    atual_valor == valor_str or
                    atual_texto == valor_str or
                    (''.join(ch for ch in atual_texto if ch.isdigit()) and ''.join(ch for ch in atual_texto if ch.isdigit()) == ''.join(ch for ch in valor_str if ch.isdigit()))
                ):
                    self.add_log(
                        f"{contexto} definido via rótulo '{rotulo}' no select '{sel.get_attribute('id') or '-'}' com opção '{atual_texto or atual_valor}'."
                    )
                    return {
                        'ok': True,
                        'id': sel.get_attribute('id') or '',
                        'value': atual_valor,
                        'text': atual_texto,
                    }
            except Exception:
                continue

        return {'ok': False}

    def _valor_confere(self, atual, esperado):
        atual_norm = str(atual).strip()
        esperado_norm = str(esperado).strip()
        if not atual_norm:
            return False
        if atual_norm == esperado_norm:
            return True
        atual_num = ''.join(ch for ch in atual_norm if ch.isdigit())
        esperado_num = ''.join(ch for ch in esperado_norm if ch.isdigit())
        return bool(atual_num and esperado_num and atual_num == esperado_num)

    def _aguardar_apex_processar(self, item_id=None, timeout=8):
        def pronto(_):
            try:
                return self.driver.execute_script("""
                    const itemId = arguments[0];
                    const jq = window.apex && apex.jQuery ? apex.jQuery : (window.jQuery || null);
                    const ativo = jq ? jq(':visible .u-Processing, :visible .t-Processing').length : 0;
                    const requisicoes = jq ? jq.active : 0;
                    if (ativo > 0 || requisicoes > 0) return false;
                    if (!itemId) return true;

                    const el = document.getElementById(itemId);
                    if (!el) {
                        const relacionados = Array.from(document.querySelectorAll('[id],[name]')).filter(node => {
                            const id = (node.id || '').toUpperCase();
                            const name = (node.getAttribute('name') || '').toUpperCase();
                            return id.includes(itemId.toUpperCase()) || name.includes(itemId.toUpperCase());
                        });
                        return relacionados.length > 0 || ativo === 0;
                    }

                    const style = window.getComputedStyle(el);
                    const visivel = style.display !== 'none' && style.visibility !== 'hidden';
                    return visivel;
                """, item_id)
            except Exception:
                return False

        WebDriverWait(self.driver, timeout).until(pronto)
        time.sleep(0.3)

    def _configurar_campo_dia(self, item_id, valor, contexto):
        valor_str = str(valor).strip()
        if valor_str in ('', '0', 'nan'):
            return

        self.add_log(f"Tentando definir {contexto} com valor {valor_str}...")
        self._aguardar_apex_processar(item_id=item_id, timeout=10)
        relacionados = self._listar_elementos_relacionados(item_id)
        if relacionados:
            resumo = ', '.join(
                f"{r.get('tag')}#{r.get('id') or '-'}[{r.get('name') or '-'}]='{r.get('value') or ''}'"
                for r in relacionados[:6]
            )
            self.add_log(f"Elementos relacionados a {contexto}: {resumo}")

        if item_id == 'P2114_DIA1':
            resultado_rotulo = self._selecionar_combo_por_rotulo('Dia 1', valor_str, contexto)
            if resultado_rotulo.get('ok'):
                self.add_log(
                    f"{contexto} confirmado pelo select do painel. Opção atual: '{resultado_rotulo.get('text') or resultado_rotulo.get('value')}'"
                )
                return
            candidatos = [
                f'//*[@id="{item_id}" and self::select]',
                f'//*[@name="{item_id}" and self::select]',
            ]
        else:
            candidatos = [
                f'//*[@id="{item_id}" and self::select]',
                f'//*[@id="{item_id}" and (self::input or self::textarea)]',
                f'//*[@id="{item_id}"]',
                f'//*[contains(@id, "{item_id}") and (self::select or self::input)]',
            ]

        if self._definir_item_apex(item_id, valor_str):
            valor_lido = self._ler_valor_item(item_id)
            if self._valor_confere(valor_lido, valor_str):
                self.add_log(f"{contexto} definido via apex.item(). Valor atual: {valor_lido}")
                time.sleep(T['select_after'])
                return
            self.add_log(
                f"{contexto} não confirmou via apex.item(). Valor atual: '{valor_lido}'",
                level="WARNING"
            )
            if item_id != 'P2114_DIA1' and self._forcar_valor_em_relacionados(item_id, valor_str):
                valor_lido = self._ler_valor_item(item_id)
                if self._valor_confere(valor_lido, valor_str):
                    self.add_log(f"{contexto} confirmado ao forçar elementos relacionados. Valor atual: {valor_lido}")
                    time.sleep(T['select_after'])
                    return
                self.add_log(
                    f"{contexto} ainda não confirmou após forçar relacionados. Valor atual: '{valor_lido}'",
                    level="WARNING"
                )

        for xpath in candidatos:
            try:
                el = WebDriverWait(self.driver, 2).until(
                    EC.presence_of_element_located((By.XPATH, xpath))
                )
                tag = (el.tag_name or '').lower()
                self.driver.execute_script("arguments[0].scrollIntoView({block:'center'});", el)
                time.sleep(0.2)

                if tag == 'select':
                    self._selecionar_combo(xpath, valor=valor_str, texto=valor_str, contexto=contexto)
                    valor_lido = self._ler_valor_item(item_id)
                    if self._valor_confere(valor_lido, valor_str):
                        self.add_log(f"{contexto} definido em select. Valor atual: {valor_lido}")
                        return
                    self.add_log(
                        f"{contexto} não confirmou em select. Valor atual: '{valor_lido}'",
                        level="WARNING"
                    )
                    continue

                try:
                    el.clear()
                except Exception:
                    pass
                try:
                    el.click()
                except Exception:
                    pass
                try:
                    el.send_keys(valor_str)
                except Exception:
                    self.driver.execute_script("arguments[0].value = arguments[1];", el, valor_str)

                self.driver.execute_script("""
                    const el = arguments[0];
                    el.dispatchEvent(new Event('input', { bubbles: true }));
                    el.dispatchEvent(new Event('change', { bubbles: true }));
                    el.dispatchEvent(new Event('blur', { bubbles: true }));
                """, el)

                if self._definir_item_apex(item_id, valor_str):
                    valor_lido = self._ler_valor_item(item_id)
                    if self._valor_confere(valor_lido, valor_str):
                        self.add_log(f"{contexto} sincronizado via apex.item(). Valor atual: {valor_lido}")
                        time.sleep(T['select_after'])
                        return
                    self.add_log(
                        f"{contexto} não confirmou após sincronizar apex. Valor atual: '{valor_lido}'",
                        level="WARNING"
                    )
                else:
                    valor_lido = self._ler_valor_item(item_id)
                    if self._valor_confere(valor_lido, valor_str):
                        self.add_log(f"{contexto} definido em input. Valor atual: {valor_lido}")
                        time.sleep(T['select_after'])
                        return
                    self.add_log(
                        f"{contexto} não confirmou em input. Valor atual: '{valor_lido}'",
                        level="WARNING"
                    )
                continue
            except Exception:
                continue

        botoes = [
            f'//*[@id="{item_id}_lov_btn"]',
            f'//*[@id="{item_id}_BUTTON"]',
            f'//*[@id="{item_id}"]/following-sibling::*[self::button or self::span][1]',
            f'//*[contains(@id, "{item_id}") and (contains(@id, "lov_btn") or contains(@class, "button"))]',
        ]
        for xpath_btn in botoes:
            try:
                btn = WebDriverWait(self.driver, 2).until(
                    EC.element_to_be_clickable((By.XPATH, xpath_btn))
                )
                self.driver.execute_script("arguments[0].scrollIntoView({block:'center'});", btn)
                self.driver.execute_script("arguments[0].click();", btn)
                time.sleep(0.4)
                if self._definir_item_apex(item_id, valor_str):
                    valor_lido = self._ler_valor_item(item_id)
                    if self._valor_confere(valor_lido, valor_str):
                        self.add_log(f"{contexto} definido após clicar no botão auxiliar. Valor atual: {valor_lido}")
                        time.sleep(T['select_after'])
                        return
            except Exception:
                continue

        valor_final = self._ler_valor_item(item_id)
        raise TimeoutException(
            f"Não foi possível confirmar {contexto}. Valor esperado: '{valor_str}' | valor atual: '{valor_final}'"
        )

    def config_periodicidade(self, periodicidade):
        try:
            self._selecionar_combo(
                '//*[@id="P2114_PERIODICIDADE"]',
                texto=periodicidade,
                contexto=f"periodicidade '{periodicidade}'"
            )
            self._aguardar_apex_processar(timeout=10)
        except Exception as e:
            self.handle_exception(e, f"Erro periodicidade '{periodicidade}'")

    def _configurar_dia1(self, dia1):
        try:
            self._configurar_campo_dia('P2114_DIA_1', dia1, f"dia1 '{dia1}'")
        except Exception as e:
            self.handle_exception(e, f"Erro dia1 '{dia1}'")
            raise

    def _configurar_dia2(self, dia2):
        try:
            self._configurar_campo_dia('P2114_DIA_2', dia2, f"dia2 '{dia2}'")
        except Exception as e:
            self.handle_exception(e, f"Erro dia2 '{dia2}'")
            raise

    def _configurar_tipo_dia(self, tipo_dia):
        try:
            v = str(tipo_dia).strip().upper()
            if v == 'IMPAR': v = 'ÍMPAR'
            if v not in ('ÍMPAR', 'PAR', 'INDIFERENTE'): return
            self._selecionar_combo(
                '//*[@id="P2114_TIPO_DIA"]',
                texto=v,
                contexto=f"tipo_dia '{tipo_dia}'"
            )
        except Exception as e:
            self.handle_exception(e, f"Erro tipo_dia '{tipo_dia}'")

    def _configurar_bancarizar(self, bancarizar='NÃO'):
        try:
            v = str(bancarizar).strip().upper()
            if v == 'NAO': v = 'NÃO'
            if v not in ('SIM', 'NÃO'): v = 'NÃO'
            self._selecionar_combo(
                '//*[@id="P2114_BANCARIZACAO"]',
                texto=v,
                contexto=f"bancarizar '{bancarizar}'"
            )
        except Exception as e:
            self.handle_exception(e, f"Erro bancarizar '{bancarizar}'")

    def _configurar_dias_semana(self, dados):
        dias = dados.get('dias_semana', []) if self.single_mode else [
            dados.get(f'dia_sem{i}', '') for i in range(1, 8)
        ]
        dias = [d for d in dias if d and str(d).strip().upper() not in ('', 'NAN', 'NONE')]
        if not dias:
            self.add_log("Nenhum dia da semana informado.", level="WARNING"); return
        validos = ['DOMINGO', 'SEGUNDA', 'TERCA', 'QUARTA', 'QUINTA', 'SEXTA', 'SABADO']
        for i, dia in enumerate(dias, 1):
            d = dia.upper().strip()
            if d not in validos:
                self.add_log(f"Dia inválido: '{dia}'", level="WARNING"); continue
            try:
                self._selecionar_combo(
                    f'//*[@id="P2114_DIA_SEMANA{i}"]',
                    texto=d,
                    contexto=f"dia_semana{i} '{dia}'"
                )
            except Exception as e:
                self.handle_exception(e, f"Erro dia_semana{i}")
        self._configurar_tipo_dia(dados.get('tipo_dia', 'INDIFERENTE'))

    def _atualizar_plano_existente(self, dados):
        self.add_log("Atualizando plano existente...")
        periodicidade = dados.get('periodicidade', '').upper()
        self._atualizar_campo('//*[@id="P2114_PAGTO_APOS"]', dados['pag_dias'])
        self.config_periodicidade(periodicidade)
        if periodicidade == 'SEMANAL':
            self._configurar_dias_semana(dados)
        else:
            self._configurar_dia1(dados['dia1'])
            self._configurar_dia2(dados['dia2'])
        self._atualizar_taxa_se_informada(dados.get('taxa'))
        self._configurar_bancarizar(dados.get('bancarizar', 'NÃO'))
        ok = self.wait.until(EC.element_to_be_clickable((By.XPATH,
            '/html/body/div[1]/div/form/div/div[1]/section/div[1]/span/button[1]/span')))
        self.driver.execute_script("arguments[0].click();", ok)
        time.sleep(T['after_save_plan'])
        self.add_log("Plano atualizado.")

    def _criar_novo_plano_compra(self, dados):
        self.add_log("Criando novo plano de compra...")
        periodicidade = dados.get('periodicidade', '').upper()
        novo = self.wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="NOVO"]/span')))
        self.driver.execute_script("arguments[0].click();", novo); time.sleep(0.8)
        sel1 = self.wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="P2114_PLANO_COMPRA"]')))
        sel1.click(); time.sleep(.4)
        self.wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="P2114_PLANO_COMPRA"]/option[2]'))).click()
        self.add_log("Plano A VISTA selecionado. Aguardando carregamento dos demais campos...")
        time.sleep(2.0)
        self._aguardar_apex_processar(timeout=10)
        self.config_periodicidade(periodicidade)
        if periodicidade == 'SEMANAL':
            self._configurar_dias_semana(dados)
        else:
            self._configurar_dia1(dados['dia1'])
            self._configurar_dia2(dados['dia2'])
        self._atualizar_taxa_se_informada(dados.get('taxa'))
        self._configurar_bancarizar(dados.get('bancarizar', 'NÃO'))
        ok = self.wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="OK"]/span')))
        self.driver.execute_script("arguments[0].click();", ok)
        time.sleep(T['after_save_plan'])
        self.add_log("Plano criado.")

    # ── HELPERS ───────────────────────────────────────────────────────────────
    def _obter_estado_produto(self):
        """
        Lê o campo P2111_ESTADO após salvar o plano de compra.
        Retorna o texto do estado em maiúsculas (ex: 'APROVADO', 'DEFINIÇÃO').
        """
        try:
            el = self.driver.find_element(By.XPATH, '//*[@id="P2111_ESTADO"]')
            estado = el.text.strip().upper()
            if not estado:
                # Alguns campos APEX usam value em vez de text
                estado = el.get_attribute('value').strip().upper()
            self.add_log(f'Estado do produto: {estado}')
            return estado
        except Exception as e:
            self.handle_exception(e, 'Erro ao ler estado do produto')
            return 'DESCONHECIDO'

    def _obter_estado_produto(self):
        """
        Lê o campo P2111_ESTADO na tela de APROVAÇÃO CONTRATO.
        Retorna o texto em maiúsculo (ex: 'APROVADO', 'DEFINIÇÃO').
        Em caso de erro retorna 'DEFINIÇÃO' para garantir que a aprovação ocorra.
        """
        try:
            el = self.driver.find_element(By.XPATH, '//*[@id="P2111_ESTADO"]')
            estado = el.text.strip() or el.get_attribute('value') or ''
            return estado.upper()
        except Exception as e:
            self.handle_exception(e, "Erro ao obter estado — assumindo DEFINIÇÃO")
            return 'DEFINIÇÃO'

    def _obter_nome_fantasia(self):
        try:
            return self.driver.find_element(By.XPATH, '//*[@id="P2101_FANTASIA"]').text
        except:
            return "Não identificado"

    def _formatar_taxa(self, taxa):
        try:
            t = str(taxa).strip().replace(',', '.')
            if not t:
                return ''
            return str(float(t)).replace('.', ',')
        except:
            return str(taxa)

    def _atualizar_taxa_se_informada(self, taxa):
        taxa_formatada = self._formatar_taxa(taxa)
        if not str(taxa_formatada).strip():
            self.add_log("Taxa em branco; mantendo a taxa atual do produto.", level="WARNING")
            return
        self._atualizar_campo('//*[@id="P2114_TAXA_REEMBOLSO"]', taxa_formatada)

    def _extrair_dados_registro(self, idx):
        row = self.dados_planilha.iloc[idx]
        def val(*cols, default=''):
            v = default
            for col in cols:
                if col in row.index:
                    v = row.get(col, default)
                    break
            v = str(v).strip()
            return '' if v.lower() in ('nan', 'none') else v
        cod_raw = val('COD', 'CODIGO', 'COD_CNPJ', 'CNPJ', 'CPF_CNPJ')
        if self.tipo_pesquisa == 'CNPJ':
            cod_raw = cod_raw.split('.')[0].zfill(14)
        dados = {
            'cod':           cod_raw,
            'taxa':          self._formatar_taxa(val('TAXA', 'TAXA_REEMBOLSO', 'TAXA (%)')),
            'periodicidade': val('PERIODICIDADE').upper(),
            'pag_dias':      int(val('PAG', 'PAG_DIAS', 'DIAS_PAGAMENTO', 'PAGAMENTO_DIAS') or 0),
            'dia1':          int(val('DIA1', 'DIA_1') or 0),
            'dia2':          val('DIA2', 'DIA_2'),
            'n_fantasia':    val('FANTASIA'),
            'bancarizar':    val('BANCARIZAR') or 'NÃO',
            'tipo_dia':      val('TIPO_DIA', 'TIPO DIA') or 'INDIFERENTE',
        }
        for i in range(1, 8):
            dados[f'dia_sem{i}'] = val(f'DIA_SEM{i}')
        dias_semana = val('DIAS_SEMANA', 'DIAS SEMANA')
        if dias_semana:
            partes = [
                parte.strip()
                for parte in dias_semana.replace(';', ',').replace('/', ',').split(',')
                if parte.strip()
            ]
            for i, dia in enumerate(partes[:7], start=1):
                dados[f'dia_sem{i}'] = dia
        return dados

    def _aprovar_produto(self):
        """
        Clica em OPÇÕES → Aprovar → confirma na tela de confirmação.
        O menu do APEX só abre com clique físico (ActionChains).
        """
        try:
            self.add_log("Aprovando produto...")

            # 1) Localizar e clicar em OPÇÕES com ActionChains (clique físico real)
            opcoes = self.wait.until(EC.element_to_be_clickable((By.ID, "CONTRATOS")))
            self.driver.execute_script("arguments[0].scrollIntoView({block:'center'});", opcoes)
            time.sleep(0.3)
            ActionChains(self.driver).move_to_element(opcoes).pause(0.2).click().perform()
            time.sleep(1.0)  # aguardar menu renderizar

            # 2) Aguardar item Aprovar ficar clicável e clicar
            aprovar = self.wait.until(
                EC.element_to_be_clickable((By.ID, "CONTRATOS_MENU_menu_1i"))
            )
            ActionChains(self.driver).move_to_element(aprovar).pause(0.2).click().perform()
            time.sleep(1.0)  # aguardar tela de confirmação carregar

            # 3) Confirmar na tela de confirmação — XPATH confirmado: //*[@id="OK"]/span
            ok = self.wait.until(
                EC.element_to_be_clickable((By.XPATH, '//*[@id="OK"]/span'))
            )
            self.driver.execute_script("arguments[0].click();", ok)
            time.sleep(T['after_approve'])

            self.add_log("Produto aprovado.", level="SUCCESS")
        except Exception as e:
            self.handle_exception(e, "Erro ao aprovar produto")
            raise

    # ── RESULTADO ─────────────────────────────────────────────────────────────
    def _registrar_linha(self, dados, sucesso, erro=None):
        if sucesso:
            resultado_col = "Alterado com sucesso!"
        else:
            proc  = self.status['processed']
            ordem = self._detectar_ordem([str(r.get('cod','')) for r in self.status['resultado_linhas']])
            resultado_col = (
                f"Erro! Feito do 1 ao {proc} (ordem: {ordem}) — {str(erro or '')[:80]}"
            )
        self.status['resultado_linhas'].append({**dados, 'RESULTADO': resultado_col})

    def _salvar_resultado_final(self, ts):
        linhas = self.status['resultado_linhas']
        if not linhas:
            return
        try:
            os.makedirs('resultados_taxa', exist_ok=True)
            nome = f"resultados_taxa/resultado_slot{self.slot}_{ts}.xlsx"
            df   = pd.DataFrame(linhas)
            cols = [c for c in df.columns if c != 'RESULTADO'] + ['RESULTADO']
            df[cols].to_excel(nome, index=False)
            self.status['output_file'] = nome
            self.add_log(f"Planilha salva: {nome}", level="SUCCESS")
        except Exception as e:
            self.handle_exception(e, "Erro ao salvar planilha")

    def _finalizar_driver(self):
        if self.driver:
            try: self.driver.quit()
            except: pass
        tempo = "N/A"
        if self.status['start_time'] and self.status['end_time']:
            tempo = str(self.status['end_time'] - self.status['start_time']).split('.')[0]
        self.add_log('=' * 50)
        self.add_log(f"Finalizado. {self.status['processed']}/{self.status['total']} | Tempo: {tempo}")
        self.add_log('=' * 50)
        self.status['request_stop'] = False


# ── MODO DUPLO PARALELO ───────────────────────────────────────────────────────
def iniciar_duplo(dados_login, arquivo_entrada, tipo_pesquisa='COD'):
    df   = pd.read_excel(arquivo_entrada, dtype=str)
    meio = len(df) // 2
    df_a = df.iloc[:meio].reset_index(drop=True)
    df_b = df.iloc[meio:].reset_index(drop=True)

    for s in _STATUS_SLOTS:
        s.update(_novo_status())

    def run(slot, dados):
        time.sleep(slot * 5)
        WebAppAlteradorTaxa(
            slot=slot, single_mode=False,
            dados_planilha=dados,
            tipo_pesquisa=tipo_pesquisa,
            dados_login=dados_login
        ).iniciar()

    t0 = threading.Thread(target=run, args=(0, df_a), daemon=True)
    t1 = threading.Thread(target=run, args=(1, df_b), daemon=True)
    t0.start(); t1.start()
    return t0, t1
