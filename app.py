from flask import Flask, render_template, request, redirect, url_for, flash, jsonify, send_file
import os
import threading
from pathlib import Path
from werkzeug.utils import secure_filename

# ── Importações das automações ─────────────────────────────────────────────────
from automacoes import (
    WebAppAtivador,
    WebAppAlteradorConciliacao,
    WebAppAlteradorTaxa,
    WebAppExtrator,
    automation_status,
    conciliacao_status,
    taxa_status,
    taxa_status2,
    extraction_status,
)


app = Flask(__name__)
app.secret_key = 'ativadorcredenciado2025'

BASE_DIR = Path(__file__).parent.resolve()
UPLOAD_FOLDER = str(BASE_DIR / 'uploads')
ALLOWED_EXTENSIONS = {'xlsx', 'xls'}
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
if not os.path.exists(UPLOAD_FOLDER):
    os.mkdir(UPLOAD_FOLDER)

# ── Status global ativo: indica qual automação está rodando ───────────────────
# None = nenhuma; 'ativador' | 'conciliacao' | 'taxa' | 'extrator'
_active_automation = None


def _get_active_status():
    """Retorna o dicionário de status da automação ativa."""
    if _active_automation == 'ativador':
        return automation_status
    if _active_automation == 'conciliacao':
        return conciliacao_status
    if _active_automation == 'taxa':
        return taxa_status
    if _active_automation == 'extrator':
        return extraction_status
    # fallback: mostra o do ativador mesmo sem estar rodando
    return automation_status


def _active_running_status():
    for name, status in (
        ('ativador', automation_status),
        ('conciliacao', conciliacao_status),
        ('taxa', taxa_status),
        ('taxa', taxa_status2),
        ('extrator', extraction_status),
    ):
        if status.get('running'):
            return name, status
    return None, None


def _can_start_automation():
    active_name, _ = _active_running_status()
    if active_name:
        flash('Ja existe uma automacao em execucao. Aguarde finalizar ou solicite a parada antes de iniciar outra.')
        return False
    return True


def reset_status(status, total=0):
    status.update({
        'running': False, 'total': total, 'processed': 0,
        'current_item': None, 'logs': [],
        'start_time': None, 'end_time': None,
        'request_stop': False,
        'relatorio_path': None,
        'output_file': None,
        'log_file': None,
        'error': None,
        'resultado_linhas': [],
    })


def allowed_file(f):
    return '.' in f and f.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS


def tecnologia_para_int(v):
    v = str(v).strip().upper()
    return 1 if v == 'CIELO' else 2 if v == 'CARDSE' else 3 if v == 'AMBOS' else 0


def acao_valida(v):
    v = str(v).strip().upper()
    return v if v in ('ATIVAR', 'DESATIVAR') else 'ATIVAR'


def normalizar_cpf_cnpj(valor: str) -> str:
    s = str(valor).strip().split('.')[0]
    digits = ''.join(c for c in s if c.isdigit())
    if not digits:
        return s
    n = len(digits)
    if n <= 11:
        return digits.zfill(11)
    elif n <= 13:
        return digits.zfill(14)
    else:
        return digits


def reset_automation_status(total):
    reset_status(automation_status, total)


def preparar_planilha_taxa_lote(df):
    df = df.copy()
    df.columns = [str(col).strip().upper() for col in df.columns]

    aliases = {
        'CODIGO': 'COD',
        'COD_CNPJ': 'COD',
        'CNPJ': 'COD',
        'CPF_CNPJ': 'COD',
        'TAXA_REEMBOLSO': 'TAXA',
        'TAXA (%)': 'TAXA',
        'PAG_DIAS': 'PAG',
        'DIAS_PAGAMENTO': 'PAG',
        'PAGAMENTO_DIAS': 'PAG',
        'DIA_1': 'DIA1',
        'DIA_2': 'DIA2',
        'TIPO DIA': 'TIPO_DIA',
        'DIAS_SEMANA': 'DIA_SEM1',
    }
    df.rename(columns={col: aliases.get(col, col) for col in df.columns}, inplace=True)

    if 'COD' not in df.columns:
        raise ValueError('Coluna COD nao encontrada. Use COD, CODIGO, CNPJ ou COD_CNPJ.')

    df = df[df['COD'].astype(str).str.strip().str.lower().ne('nan')]
    df = df[df['COD'].astype(str).str.strip().ne('')]
    if df.empty:
        raise ValueError('Nenhum registro valido na planilha.')

    for col in ('TAXA', 'PERIODICIDADE', 'PAG', 'DIA1', 'DIA2', 'BANCARIZAR', 'TIPO_DIA'):
        if col not in df.columns:
            df[col] = ''

    return df.reset_index(drop=True)


# ─────────────────────────────────────────────────────────────────────────────
# HOME
# ─────────────────────────────────────────────────────────────────────────────

@app.route('/')
def index():
    return render_template('index.html')


# ─── Ativador Individual ──────────────────────────────────────────────────────

@app.route('/individual', methods=['GET', 'POST'])
def individual():
    global _active_automation
    if request.method == 'POST':
        if not _can_start_automation():
            return redirect(url_for('status'))

        login_infox      = request.form.get('login_infox', '').strip()
        senha_infox      = request.form.get('senha_infox', '').strip()
        tipo_pesquisa    = request.form.get('tipo_pesquisa', 'COD')
        cod_cnpj         = request.form.get('cod_cnpj', '').strip()
        tecnologia_str   = request.form.get('tecnologia', 'CIELO')
        acao_cielo       = acao_valida(request.form.get('acao_cielo', 'ATIVAR'))
        acao_cardse      = acao_valida(request.form.get('acao_cardse', 'ATIVAR'))
        modo_verificacao = request.form.get('modo_verificacao') == '1'
        produto_inicio   = int(request.form.get('produto_inicio') or 1)
        produto_unico    = request.form.get('produto_unico', '').strip()
        cardse_statuses  = request.form.getlist('cardse_statuses')

        if not cod_cnpj:
            flash('Informe o código ou CNPJ.')
            return redirect(url_for('individual'))
        if not login_infox or not senha_infox:
            flash('Login e senha são obrigatórios.')
            return redirect(url_for('individual'))

        tecnologia = tecnologia_para_int(tecnologia_str)
        reset_automation_status(1)
        _active_automation = 'ativador'

        def run():
            WebAppAtivador(
                dados_login={'login_infox': login_infox, 'senha_infox': senha_infox},
                lista_credenciados=[{'cod': cod_cnpj, 'tecnologia': tecnologia,
                                     'acao_cielo': acao_cielo, 'acao_cardse': acao_cardse,
                                     'produto_unico': produto_unico}],
                tipo_pesquisa=tipo_pesquisa,
                tecnologia_padrao=tecnologia,
                acao_cielo=acao_cielo,
                acao_cardse=acao_cardse,
                cardse_statuses=cardse_statuses,
                modo_verificacao=modo_verificacao,
                produto_inicio=produto_inicio,
            ).iniciar()

        threading.Thread(target=run, daemon=True).start()
        return redirect(url_for('status'))

    return render_template('individual.html')


# ─── Ativador Lote ────────────────────────────────────────────────────────────

@app.route('/lote', methods=['GET', 'POST'])
def lote():
    global _active_automation
    if request.method == 'POST':
        if not _can_start_automation():
            return redirect(url_for('status'))

        login_infox      = request.form.get('login_infox', '').strip()
        senha_infox      = request.form.get('senha_infox', '').strip()
        tipo_pesquisa    = request.form.get('tipo_pesquisa', 'COD')
        tecnologia_str   = request.form.get('tecnologia', 'CIELO')
        acao_cielo       = acao_valida(request.form.get('acao_cielo', 'ATIVAR'))
        acao_cardse      = acao_valida(request.form.get('acao_cardse', 'ATIVAR'))
        modo_verificacao = request.form.get('modo_verificacao') == '1'
        modo_coleta      = request.form.get('modo_coleta') == '1'
        origem_transacao_alvo = request.form.get('origem_transacao_alvo', 'CARDSE').strip()
        produto_inicio   = int(request.form.get('produto_inicio') or 1)
        produto_unico_front = request.form.get('produto_unico', '').strip()
        cardse_statuses  = request.form.getlist('cardse_statuses')

        if not login_infox or not senha_infox:
            flash('Login e senha são obrigatórios.')
            return redirect(url_for('lote'))
        if 'file' not in request.files or request.files['file'].filename == '':
            flash('Selecione uma planilha.')
            return redirect(url_for('lote'))

        file = request.files['file']
        if not allowed_file(file.filename):
            flash('Use .xlsx ou .xls')
            return redirect(url_for('lote'))

        upload_dir = app.config['UPLOAD_FOLDER']
        os.makedirs(upload_dir, exist_ok=True)
        filepath = os.path.join(upload_dir, secure_filename(file.filename))
        file.save(filepath)

        import pandas as pd
        df = pd.read_excel(filepath, dtype=str)
        df.columns = [str(col).strip().upper() for col in df.columns]
        if 'COD' not in df.columns:
            flash('Coluna COD não encontrada.')
            return redirect(url_for('lote'))

        tecnologia_padrao = tecnologia_para_int(tecnologia_str)
        tem_tec = 'TECNOLOGIA'  in df.columns
        tem_ac  = 'ACAO_CIELO'  in df.columns
        tem_as  = 'ACAO_CARDSE' in df.columns
        tem_pu  = 'PRODUTO'     in df.columns

        lista = []
        for _, row in df.iterrows():
            raw = str(row['COD']).strip()
            if not raw or raw.lower() == 'nan':
                continue
            if tipo_pesquisa == 'CNPJ':
                cod = normalizar_cpf_cnpj(raw)
            else:
                cod = raw.split('.')[0]

            tec = tecnologia_para_int(row.get('TECNOLOGIA', '')) if tem_tec else 0
            tec = tec if tec > 0 else tecnologia_padrao

            ac_raw = str(row.get('ACAO_CIELO', '')).strip().upper() if tem_ac else ''
            as_raw = str(row.get('ACAO_CARDSE', '')).strip().upper() if tem_as else ''
            ac  = ac_raw if ac_raw in ('ATIVAR', 'DESATIVAR') else acao_cielo
            as_ = as_raw if as_raw in ('ATIVAR', 'DESATIVAR') else acao_cardse

            pu_raw = str(row.get('PRODUTO', '')).strip() if tem_pu else ''
            pu = produto_unico_front if not pu_raw or pu_raw.lower() == 'nan' else pu_raw

            lista.append({'cod': cod, 'tecnologia': tec, 'acao_cielo': ac, 'acao_cardse': as_, 'produto_unico': pu})

        if not lista:
            flash('Nenhum registro válido.')
            return redirect(url_for('lote'))

        reset_automation_status(len(lista))
        _active_automation = 'ativador'

        def run():
            WebAppAtivador(
                dados_login={'login_infox': login_infox, 'senha_infox': senha_infox},
                lista_credenciados=lista,
                tipo_pesquisa=tipo_pesquisa,
                tecnologia_padrao=tecnologia_padrao,
                acao_cielo=acao_cielo,
                acao_cardse=acao_cardse,
                cardse_statuses=cardse_statuses,
                modo_verificacao=modo_verificacao,
                modo_coleta_estabelecimento=modo_coleta,
                origem_transacao_alvo=origem_transacao_alvo,
                produto_inicio=produto_inicio,
            ).iniciar()

        threading.Thread(target=run, daemon=True).start()
        return redirect(url_for('status'))

    return render_template('lote.html')


# ─── Coleta Código Estabelecimento ───────────────────────────────────────────

@app.route('/coleta-codigo-estabelecimento', methods=['GET', 'POST'])
def coleta_codigo_estabelecimento():
    global _active_automation
    if request.method == 'POST':
        if not _can_start_automation():
            return redirect(url_for('status'))

        login_infox   = request.form.get('login_infox', '').strip()
        senha_infox   = request.form.get('senha_infox', '').strip()
        tipo_pesquisa = request.form.get('tipo_pesquisa', 'COD')
        origem_alvo   = request.form.get('origem_transacao', '').strip()

        if not login_infox or not senha_infox:
            flash('Login e senha são obrigatórios.')
            return redirect(url_for('coleta_codigo_estabelecimento'))
        if not origem_alvo:
            flash('Informe a origem da transação que deseja coletar.')
            return redirect(url_for('coleta_codigo_estabelecimento'))
        if 'file' not in request.files or request.files['file'].filename == '':
            flash('Selecione uma planilha.')
            return redirect(url_for('coleta_codigo_estabelecimento'))

        file = request.files['file']
        if not allowed_file(file.filename):
            flash('Use .xlsx ou .xls')
            return redirect(url_for('coleta_codigo_estabelecimento'))

        upload_dir = app.config['UPLOAD_FOLDER']
        os.makedirs(upload_dir, exist_ok=True)
        filepath = os.path.join(upload_dir, secure_filename(file.filename))
        file.save(filepath)

        import pandas as pd
        df = pd.read_excel(filepath, dtype=str)
        if 'COD' not in df.columns:
            flash('Coluna COD não encontrada.')
            return redirect(url_for('coleta_codigo_estabelecimento'))

        lista = []
        for _, row in df.iterrows():
            raw = str(row['COD']).strip()
            if not raw or raw.lower() == 'nan':
                continue
            cod = normalizar_cpf_cnpj(raw) if tipo_pesquisa == 'CNPJ' else raw.split('.')[0]
            lista.append({'cod': cod})

        if not lista:
            flash('Nenhum registro válido.')
            return redirect(url_for('coleta_codigo_estabelecimento'))

        reset_automation_status(len(lista))
        _active_automation = 'ativador'

        def run():
            WebAppAtivador(
                dados_login={'login_infox': login_infox, 'senha_infox': senha_infox},
                lista_credenciados=lista,
                tipo_pesquisa=tipo_pesquisa,
                modo_coleta_estabelecimento=True,
                origem_transacao_alvo=origem_alvo,
            ).iniciar()

        threading.Thread(target=run, daemon=True).start()
        return redirect(url_for('status'))

    return render_template('coleta_codigo_estabelecimento.html')


# ─── Alterar Conciliação ──────────────────────────────────────────────────────

@app.route('/alt-conciliacao', methods=['GET', 'POST'])
def alt_conciliacao():
    global _active_automation
    if request.method == 'POST':
        if not _can_start_automation():
            return redirect(url_for('status'))

        login_infox      = request.form.get('login_infox', '').strip()
        senha_infox      = request.form.get('senha_infox', '').strip()
        tipo_pesquisa    = request.form.get('tipo_pesquisa', 'CNPJ')
        nome_conciliacao = request.form.get('nome_conciliacao', '').strip()

        if not login_infox or not senha_infox:
            flash('Login e senha são obrigatórios.')
            return redirect(url_for('alt_conciliacao'))
        if not nome_conciliacao:
            flash('Selecione o nome da conciliação.')
            return redirect(url_for('alt_conciliacao'))
        if 'file' not in request.files or request.files['file'].filename == '':
            flash('Selecione uma planilha.')
            return redirect(url_for('alt_conciliacao'))

        file = request.files['file']
        if not allowed_file(file.filename):
            flash('Use .xlsx ou .xls')
            return redirect(url_for('alt_conciliacao'))

        upload_dir = app.config['UPLOAD_FOLDER']
        os.makedirs(upload_dir, exist_ok=True)
        filepath = os.path.join(upload_dir, secure_filename(file.filename))
        file.save(filepath)

        import pandas as pd
        df = pd.read_excel(filepath, dtype=str)
        df.columns = [str(col).strip().upper() for col in df.columns]
        if 'COD' not in df.columns:
            flash('Coluna COD não encontrada na planilha.')
            return redirect(url_for('alt_conciliacao'))

        lista = []
        for _, row in df.iterrows():
            raw = str(row['COD']).strip()
            if not raw or raw.lower() == 'nan':
                continue
            cod = normalizar_cpf_cnpj(raw) if tipo_pesquisa == 'CNPJ' else raw.split('.')[0]
            lista.append({'cod': cod})

        if not lista:
            flash('Nenhum registro válido na planilha.')
            return redirect(url_for('alt_conciliacao'))

        # Reseta o status da conciliação
        reset_status(conciliacao_status, len(lista))
        _active_automation = 'conciliacao'

        def run():
            WebAppAlteradorConciliacao(
                dados_login={'login_infox': login_infox, 'senha_infox': senha_infox},
                lista_cnpjs=lista,
                nome_conciliacao=nome_conciliacao,
                tipo_pesquisa=tipo_pesquisa,
            ).iniciar()

        threading.Thread(target=run, daemon=True).start()
        return redirect(url_for('status'))

    return render_template('alt_conciliacao.html')


# ─── Alterar Taxas ────────────────────────────────────────────────────────────

@app.route('/alt-taxa', methods=['GET', 'POST'])
def alt_taxa():
    global _active_automation
    if request.method == 'POST':
        if not _can_start_automation():
            return redirect(url_for('status'))

        login_infox   = request.form.get('login_infox', '').strip()
        senha_infox   = request.form.get('senha_infox', '').strip()
        tipo_pesquisa = request.form.get('tipo_pesquisa', 'COD')
        mode          = request.form.get('mode', 'individual')

        if not login_infox or not senha_infox:
            flash('Login e senha são obrigatórios.')
            return redirect(url_for('alt_taxa'))

        dados_login = {'login_infox': login_infox, 'senha_infox': senha_infox}

        if mode == 'individual':
            cod_cnpj      = request.form.get('cod_cnpj', '').strip()
            taxa          = request.form.get('taxa', '').strip()
            periodicidade = request.form.get('periodicidade', 'MENSAL').strip()
            pag_dias      = int(request.form.get('pag_dias') or 0)
            dia1          = int(request.form.get('dia1') or 0)
            dia2          = request.form.get('dia2', '0').strip()
            bancarizar    = request.form.get('bancarizar', 'NÃO').strip()
            tipo_dia      = request.form.get('tipo_dia', 'INDIFERENTE').strip()
            dias_semana   = request.form.getlist('dias_semana')

            if not cod_cnpj:
                flash('Informe o código ou CNPJ.')
                return redirect(url_for('alt_taxa'))

            reset_status(taxa_status, 1)
            _active_automation = 'taxa'

            def run():
                WebAppAlteradorTaxa(
                    slot=0,
                    single_mode=True,
                    dados_credenciado={
                        'cod_cnpj':      cod_cnpj,
                        'taxa':          taxa,
                        'periodicidade': periodicidade,
                        'pag_dias':      pag_dias,
                        'dia1':          dia1,
                        'dia2':          dia2,
                        'bancarizar':    bancarizar,
                        'tipo_dia':      tipo_dia,
                        'dias_semana':   dias_semana,
                    },
                    tipo_pesquisa=tipo_pesquisa,
                    dados_login=dados_login,
                ).iniciar()

            threading.Thread(target=run, daemon=True).start()

        else:  # lote
            if 'file' not in request.files or request.files['file'].filename == '':
                flash('Selecione uma planilha.')
                return redirect(url_for('alt_taxa'))

            file = request.files['file']
            if not allowed_file(file.filename):
                flash('Use .xlsx ou .xls')
                return redirect(url_for('alt_taxa'))

            upload_dir = app.config['UPLOAD_FOLDER']
            os.makedirs(upload_dir, exist_ok=True)
            filepath = os.path.join(upload_dir, secure_filename(file.filename))
            file.save(filepath)

            import pandas as pd
            df = pd.read_excel(filepath, dtype=str)
            try:
                df = preparar_planilha_taxa_lote(df)
            except ValueError as exc:
                flash(str(exc))
                return redirect(url_for('alt_taxa'))

            reset_status(taxa_status, len(df))
            _active_automation = 'taxa'

            def run():
                WebAppAlteradorTaxa(
                    slot=0,
                    single_mode=False,
                    dados_planilha=df,
                    tipo_pesquisa=tipo_pesquisa,
                    dados_login=dados_login,
                ).iniciar()

            threading.Thread(target=run, daemon=True).start()

        return redirect(url_for('status'))

    return render_template('alt_taxa.html')


# ─── Extrator de Dados ────────────────────────────────────────────────────────

@app.route('/extrator', methods=['GET', 'POST'])
def extrator():
    global _active_automation
    if request.method == 'POST':
        if not _can_start_automation():
            return redirect(url_for('status'))

        login_infox             = request.form.get('login_infox', '').strip()
        senha_infox             = request.form.get('senha_infox', '').strip()
        tipo_pesquisa           = request.form.get('tipo_pesquisa', 'COD')
        consultas_transacoes    = 'consultas_transacoes' in request.form
        funcionalidade_transacao = request.form.get('funcionalidade_transacao', 'valor_venda_total')
        data_inicial            = request.form.get('data_inicial', '').strip()
        data_final              = request.form.get('data_final', '').strip()

        if not login_infox or not senha_infox:
            flash('Login e senha são obrigatórios.')
            return redirect(url_for('extrator'))
        if 'file' not in request.files or request.files['file'].filename == '':
            flash('Selecione uma planilha.')
            return redirect(url_for('extrator'))

        file = request.files['file']
        if not allowed_file(file.filename):
            flash('Use .xlsx ou .xls')
            return redirect(url_for('extrator'))

        upload_dir = app.config['UPLOAD_FOLDER']
        os.makedirs(upload_dir, exist_ok=True)
        filepath = os.path.join(upload_dir, secure_filename(file.filename))
        file.save(filepath)

        import pandas as pd
        df = pd.read_excel(filepath, dtype=str)
        df.columns = [str(col).strip().upper() for col in df.columns]
        if 'COD' not in df.columns:
            flash('Coluna COD não encontrada.')
            return redirect(url_for('extrator'))

        codigos = []
        for _, row in df.iterrows():
            raw = str(row['COD']).strip()
            if not raw or raw.lower() == 'nan':
                continue
            codigos.append(normalizar_cpf_cnpj(raw) if tipo_pesquisa == 'CNPJ' else raw.split('.')[0])

        if not codigos:
            flash('Nenhum registro válido na planilha.')
            return redirect(url_for('extrator'))

        # Monta filtros para o extrator
        filtros = {
            'planilha_codigos':        codigos,
            'consultas_transacoes':    consultas_transacoes,
            'funcionalidade_transacao': funcionalidade_transacao,
        }
        if consultas_transacoes and data_inicial and data_final:
            filtros['consultas_transacoes_list'] = [
                {'codigo': c, 'data_inicial': data_inicial, 'data_final': data_final}
                for c in codigos
            ]

        # Reseta status do extrator
        reset_status(extraction_status, len(codigos))
        _active_automation = 'extrator'

        def run():
            WebAppExtrator(
                dados_login={'login_infox': login_infox, 'senha_infox': senha_infox},
                filtros=filtros,
            ).iniciar()

        threading.Thread(target=run, daemon=True).start()
        return redirect(url_for('status'))

    return render_template('extrator.html')


# ─── Status & API ─────────────────────────────────────────────────────────────

@app.route('/status')
def status():
    return render_template('status.html', status=_get_active_status())


@app.route('/api/status')
def api_status():
    st = _get_active_status()
    total     = st['total']
    processed = st['processed']

    total_safe     = max(1, total)
    processed_safe = min(processed, total_safe)
    progress       = int((processed_safe / total_safe) * 100)
    relatorio      = st.get('relatorio_path') or st.get('output_file')

    tem_relatorio = False
    if relatorio:
        try:
            tem_relatorio = os.path.isfile(os.path.abspath(relatorio))
        except Exception:
            tem_relatorio = False

    return jsonify({
        'running':       st['running'],
        'total':         total,
        'processed':     processed,
        'current_item':  st['current_item'],
        'logs':          st['logs'],
        'progress':      progress,
        'tem_relatorio': tem_relatorio,
        'automacao':     _active_automation or 'nenhuma',
    })


@app.route('/download_relatorio')
def download_relatorio():
    st = _get_active_status()
    caminho = st.get('relatorio_path') or st.get('output_file')
    if not caminho:
        flash('Nenhum relatório disponível.')
        return redirect(url_for('status'))

    caminho_abs = os.path.abspath(caminho)
    if not os.path.isfile(caminho_abs):
        flash('Arquivo de relatório não encontrado.')
        return redirect(url_for('status'))

    return send_file(
        caminho_abs,
        as_attachment=True,
        download_name=os.path.basename(caminho_abs),
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        conditional=False,
    )


@app.route('/stop')
def stop():
    st = _get_active_status()
    if st['running']:
        st['request_stop'] = True
        flash('Interrupção solicitada.')
    else:
        flash('Nenhuma automação em execução.')
    return redirect(url_for('status'))


if __name__ == '__main__':
    app.run(host='0.0.0.0', port=8081, debug=True)
