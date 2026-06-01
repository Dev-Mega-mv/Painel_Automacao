from flask import Flask, render_template, request, redirect, url_for, flash, jsonify, send_file
import os
import threading
from pathlib import Path
from werkzeug.utils import secure_filename
from automacao_ativador import WebAppAtivador, automation_status


app = Flask(__name__)
app.secret_key = 'ativadorcredenciado2025'

BASE_DIR = Path(__file__).parent.resolve()
UPLOAD_FOLDER = str(BASE_DIR / 'uploads')
ALLOWED_EXTENSIONS = {'xlsx', 'xls'}
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
if not os.path.exists(UPLOAD_FOLDER):
    os.mkdir(UPLOAD_FOLDER)

def allowed_file(f): return '.' in f and f.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

def tecnologia_para_int(v):
    v = str(v).strip().upper()
    return 1 if v == 'CIELO' else 2 if v == 'CARDSE' else 3 if v == 'AMBOS' else 0

def acao_valida(v):
    v = str(v).strip().upper()
    return v if v in ('ATIVAR', 'DESATIVAR') else 'ATIVAR'

def reset_automation_status(total):
    automation_status.update({
        'running': False, 'total': total, 'processed': 0,
        'current_item': None, 'logs': [],
        'start_time': None, 'end_time': None,
        'request_stop': False, 'relatorio_path': None
    })


@app.route('/')
def index():
    return render_template('index.html')


# ─── Individual ───────────────────────────────────────────────────────────────

@app.route('/individual', methods=['GET', 'POST'])
def individual():
    if request.method == 'POST':
        login_infox      = request.form.get('login_infox', '').strip()
        senha_infox      = request.form.get('senha_infox', '').strip()
        tipo_pesquisa    = request.form.get('tipo_pesquisa', 'COD')
        cod_cnpj         = request.form.get('cod_cnpj', '').strip()
        tecnologia_str   = request.form.get('tecnologia', 'CIELO')
        acao_cielo       = acao_valida(request.form.get('acao_cielo', 'ATIVAR'))
        acao_cardse      = acao_valida(request.form.get('acao_cardse', 'ATIVAR'))
        modo_verificacao = request.form.get('modo_verificacao') == '1'
        produto_inicio   = int(request.form.get('produto_inicio') or 1)

        if not cod_cnpj:
            flash('Informe o código ou CNPJ.')
            return redirect(url_for('individual'))
        if not login_infox or not senha_infox:
            flash('Login e senha são obrigatórios.')
            return redirect(url_for('individual'))

        tecnologia = tecnologia_para_int(tecnologia_str)

        reset_automation_status(1)

        def run():
            WebAppAtivador(
                dados_login={'login_infox': login_infox, 'senha_infox': senha_infox},
                lista_credenciados=[{'cod': cod_cnpj, 'tecnologia': tecnologia,
                                     'acao_cielo': acao_cielo, 'acao_cardse': acao_cardse}],
                tipo_pesquisa=tipo_pesquisa,
                tecnologia_padrao=tecnologia,
                acao_cielo=acao_cielo,
                acao_cardse=acao_cardse,
                modo_verificacao=modo_verificacao,
                produto_inicio=produto_inicio,
            ).iniciar()

        threading.Thread(target=run, daemon=True).start()
        # Sempre redireciona para status independente do modo
        return redirect(url_for('status'))

    return render_template('individual.html')


# ─── Lote ─────────────────────────────────────────────────────────────────────

@app.route('/lote', methods=['GET', 'POST'])
def lote():
    if request.method == 'POST':
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
        if not os.path.exists(upload_dir):
            os.mkdir(upload_dir)
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

        lista = []
        for _, row in df.iterrows():
            cod = str(row['COD']).strip().split('.')[0]
            if not cod or cod.lower() == 'nan': continue
            if tipo_pesquisa == 'CNPJ': cod = cod.zfill(14)

            tec = tecnologia_para_int(row.get('TECNOLOGIA', '')) if tem_tec else 0
            tec = tec if tec > 0 else tecnologia_padrao

            ac_raw = str(row.get('ACAO_CIELO', '')).strip().upper() if tem_ac else ''
            as_raw = str(row.get('ACAO_CARDSE', '')).strip().upper() if tem_as else ''
            ac  = ac_raw if ac_raw in ('ATIVAR','DESATIVAR') else acao_cielo
            as_ = as_raw if as_raw in ('ATIVAR','DESATIVAR') else acao_cardse

            lista.append({'cod': cod, 'tecnologia': tec, 'acao_cielo': ac, 'acao_cardse': as_})

        if not lista:
            flash('Nenhum registro válido.')
            return redirect(url_for('lote'))

        reset_automation_status(len(lista))

        def run():
            WebAppAtivador(
                dados_login={'login_infox': login_infox, 'senha_infox': senha_infox},
                lista_credenciados=lista,
                tipo_pesquisa=tipo_pesquisa,
                tecnologia_padrao=tecnologia_padrao,
                acao_cielo=acao_cielo,
                acao_cardse=acao_cardse,
                modo_verificacao=modo_verificacao,                modo_coleta_estabelecimento=modo_coleta,
                origem_transacao_alvo=origem_transacao_alvo,                produto_inicio=produto_inicio,
            ).iniciar()

        threading.Thread(target=run, daemon=True).start()
        # Sempre redireciona para status independente do modo
        return redirect(url_for('status'))

    return render_template('lote.html')


@app.route('/coleta-codigo-estabelecimento', methods=['GET', 'POST'])
def coleta_codigo_estabelecimento():
    if request.method == 'POST':
        login_infox    = request.form.get('login_infox', '').strip()
        senha_infox    = request.form.get('senha_infox', '').strip()
        tipo_pesquisa  = request.form.get('tipo_pesquisa', 'COD')
        origem_alvo    = request.form.get('origem_transacao', '').strip()

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
        if not os.path.exists(upload_dir):
            os.mkdir(upload_dir)
        filepath = os.path.join(upload_dir, secure_filename(file.filename))
        file.save(filepath)

        import pandas as pd
        df = pd.read_excel(filepath, dtype=str)
        if 'COD' not in df.columns:
            flash('Coluna COD não encontrada.')
            return redirect(url_for('coleta_codigo_estabelecimento'))

        lista = []
        for _, row in df.iterrows():
            cod = str(row['COD']).strip().split('.')[0]
            if not cod or cod.lower() == 'nan':
                continue
            if tipo_pesquisa == 'CNPJ':
                cod = cod.zfill(14)
            lista.append({'cod': cod})

        if not lista:
            flash('Nenhum registro válido.')
            return redirect(url_for('coleta_codigo_estabelecimento'))

        reset_automation_status(len(lista))

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


# ─── Status & API ─────────────────────────────────────────────────────────────

@app.route('/status')
def status():
    return render_template('status.html', status=automation_status)

@app.route('/api/status')
def api_status():
    total     = automation_status['total']
    processed = automation_status['processed']
    
    # Garante que total nunca seja 0 ou negativo para evitar divisão inválida
    total_safe = max(1, total)
    # Garante que processed não ultrapasse total para evitar >100%
    processed_safe = min(processed, total_safe)
    
    progress  = int((processed_safe / total_safe) * 100)
    relatorio = automation_status.get('relatorio_path')

    # Verifica se o arquivo existe com caminho absoluto
    tem_relatorio = False
    if relatorio:
        try:
            caminho_abs = os.path.abspath(relatorio)
            tem_relatorio = os.path.isfile(caminho_abs)
        except:
            tem_relatorio = False

    return jsonify({
        'running':        automation_status['running'],
        'total':          total,
        'processed':      processed,
        'current_item':   automation_status['current_item'],
        'logs':           automation_status['logs'],
        'progress':       progress,
        'tem_relatorio':  tem_relatorio,
    })

@app.route('/download_relatorio')
def download_relatorio():
    caminho = automation_status.get('relatorio_path')
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
    if automation_status['running']:
        automation_status['request_stop'] = True
        flash('Interrupção solicitada.')
    else:
        flash('Nenhuma automação em execução.')
    return redirect(url_for('status'))


if __name__ == '__main__':
    app.run(host='0.0.0.0', port=8081, debug=True)
