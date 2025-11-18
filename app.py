# app.py — VERSÃO FINAL CORRIGIDA PARA RENDER FREE (SEM PERMISSIONERROR + PLAYWRIGHT FUNCIONANDO)
from flask import Flask, request, jsonify, Response, stream_with_context
from flask_cors import CORS
import os
from dotenv import load_dotenv
from playwright.sync_api import sync_playwright
import pandas as pd
from datetime import datetime, timedelta
import openpyxl
from openpyxl.styles import Font, PatternFill
from openpyxl.chart import BarChart, Reference
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from bs4 import BeautifulSoup
import json
import base64
from urllib.parse import urlparse, parse_qs
from apscheduler.schedulers.background import BackgroundScheduler
import atexit
import tempfile

app = Flask(__name__)
CORS(app)
load_dotenv()

# Configs
EMAIL_SEU = os.getenv('EMAIL_SEU')
EMAIL_SENHA = os.getenv('EMAIL_SENHA')
USUARIO_DEFAULT = 'Guilherme Duarte'
SENHA_DEFAULT = '13072006'

# Detecção segura do disco /data
DATA_DIR = '/data'
if not os.path.exists(DATA_DIR) or not os.access(DATA_DIR, os.W_OK):
    DATA_DIR = tempfile.mkdtemp(prefix='lavanderia_data_')
    print(f"[FALLBACK] /data não acessível. Usando temp: {DATA_DIR}")

# Set Playwright browsers path para 0 (cache padrão do Render)
os.environ['PLAYWRIGHT_BROWSERS_PATH'] = '0'

# Caminhos
config_file = os.path.join(DATA_DIR, 'config.json')
hospitals_file = os.path.join(DATA_DIR, 'hospitals.json')
CAMINHO_RELATORIO = os.path.join(DATA_DIR, 'RelatorioLavanderiaSemanal.xlsx')

config = {}
hospitals = []

# Carregar dados
def load_data():
    global config, hospitals
    if os.path.exists(config_file):
        with open(config_file, 'r') as f:
            config = json.load(f)
    else:
        config = {}
    if os.path.exists(hospitals_file):
        with open(hospitals_file, 'r') as f:
            hospitals = json.load(f)
    else:
        hospitals = []
load_data()

# Salvar dados (sem makedirs para evitar error — tempdir já é writable)
def save_data():
    with open(config_file, 'w') as f:
        json.dump(config, f, indent=2)
    with open(hospitals_file, 'w') as f:
        json.dump(hospitals, f, indent=2)

# Cálculo semana anterior
def calcular_semana_anterior():
    hoje = datetime.now()
    dias_desde_segunda = hoje.weekday()
    inicio = hoje - timedelta(days=dias_desde_segunda + 7)
    fim = inicio + timedelta(days=6)
    return inicio, fim

# Extração com Playwright
def extrair_dados_semana_anterior(page, hospital):
    inicio, fim = calcular_semana_anterior()
    periodo_text = f"{inicio.strftime('%d/%m/%Y')} a {fim.strftime('%d/%m/%Y')}"

    parsed = urlparse(hospital['url'])
    params = parse_qs(parsed.query)
    cliente_id = params.get('cliente', [None])[0]
    if not cliente_id:
        return 0.0, [], periodo_text

    base_url = "https://sistemasaogeraldoservice.com.br/sistema/ListagemLavanderia.aspx"
    dados_por_mes = {}
    current = inicio
    while current <= fim:
        mes_ano = current.strftime("%m/%Y")
        dia_str = current.strftime("%d/%m/%Y")
        dados_por_mes.setdefault(mes_ano, []).append(dia_str)
        current += timedelta(days=1)

    total_kg = 0.0
    todos_dados = []

    for mes_ano, dias in dados_por_mes.items():
        page.goto(f"{base_url}?cliente={cliente_id}&periodo={mes_ano}")
        page.wait_for_load_state('networkidle')

        soup = BeautifulSoup(page.content(), 'html.parser')
        tabela = soup.find('table', id='tabpedidos')
        if not tabela:
            for t in soup.find_all('table'):
                if t.find('th') and 'DIA' in t.get_text(strip=True).upper():
                    tabela = t
                    break
        if not tabela:
            continue

        for linha in tabela.find_all('tr')[1:]:
            cols = linha.find_all('td')
            if len(cols) < 2: continue
            data_str = cols[0].get_text(strip=True)
            if data_str not in dias: continue
            kg_text = cols[1].get_text(strip=True).replace('.', '').replace(' ', '').replace(',', '.')
            try: kg = float(kg_text)
            except: kg = 0.0
            todos_dados.append({'data': data_str, 'kg': kg})
            total_kg += kg

    return total_kg, todos_dados, periodo_text

# Relatório
def gerar_relatorio(resultados):
    df = pd.DataFrame([{'Hospital': r['hospital'], 'Período': r['periodo'], 'Total (Kg)': r['total']} for r in resultados])
    df.to_excel(CAMINHO_RELATORIO, index=False)
    wb = openpyxl.load_workbook(CAMINHO_RELATORIO)
    ws = wb.active
    for cell in ws[1]:
        cell.font = Font(bold=True)
        cell.fill = PatternFill(start_color='CCCCCC', end_color='CCCCCC', fill_type='solid')
    total = sum(r['total'] for r in resultados)
    ws.append(['Total Geral', '', total])
    ws[f'C{ws.max_row}'].font = Font(bold=True)
    if len(resultados) > 1:
        chart = BarChart()
        data = Reference(ws, min_col=3, min_row=1, max_row=ws.max_row-1)
        cats = Reference(ws, min_col=1, min_row=2, max_row=ws.max_row-1)
        chart.add_data(data, titles_from_data=True)
        chart.set_categories(cats)
        chart.title = "Totais por Hospital"
        ws.add_chart(chart, "E2")
    wb.save(CAMINHO_RELATORIO)

# Envio de email
def enviar_email(email_dest):
    if not EMAIL_SEU or not EMAIL_SENHA or not email_dest: return
    msg = MIMEMultipart()
    msg['From'] = EMAIL_SEU
    msg['To'] = email_dest
    msg['Subject'] = f"Relatório Lavanderia {datetime.now().strftime('%d/%m/%Y')}"
    msg.attach(MIMEText("Segue o relatório anexado.", 'plain'))
    with open(CAMINHO_RELATORIO, "rb") as att:
        part = MIMEBase('application', 'octet-stream')
        part.set_payload(att.read())
        encoders.encode_base64(part)
        part.add_header('Content-Disposition', f'attachment; filename={os.path.basename(CAMINHO_RELATORIO)}')
        msg.attach(part)
    server = smtplib.SMTP('smtp.gmail.com', 587)
    server.starttls()
    server.login(EMAIL_SEU, EMAIL_SENHA)
    server.sendmail(EMAIL_SEU, email_dest, msg.as_string())
    server.quit()

# Execução agendada
def executar_relatorio_agendado():
    if not hospitals or not config.get('email'): return
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        page = browser.new_page()
        page.goto('https://sistemasaogeraldoservice.com.br/sistema/Login.aspx')
        page.fill('#txtUsuario', config.get('username', USUARIO_DEFAULT))
        page.fill('#txtSenha', config.get('password', SENHA_DEFAULT))
        page.click('#Button1')
        page.wait_for_load_state('networkidle')

        resultados = []
        for h in hospitals:
            total_kg, dados, periodo = extrair_dados_semana_anterior(page, h)
            resultados.append({'hospital': h['name'], 'periodo': periodo, 'total': total_kg, 'dados': dados})

        gerar_relatorio(resultados)
        enviar_email(config['email'])

        browser.close()

# Agendador
scheduler = BackgroundScheduler()
scheduler.start()
atexit.register(lambda: scheduler.shutdown())

def reagendar():
    scheduler.remove_all_jobs()
    horario = config.get('schedule')
    if horario:
        if horario.startswith('cron['):
            try:
                params = horario[5:-1].strip()
                day_str, time_str = params.split(' ', 1)
                hour, minute = map(int, time_str.split(':'))
                scheduler.add_job(executar_relatorio_agendado, 'cron', day_of_week=day_str, hour=hour, minute=minute, id='recorrente')
            except: pass
        else:
            try:
                dt = datetime.fromisoformat(horario)
                if dt > datetime.now():
                    scheduler.add_job(executar_relatorio_agendado, 'date', run_date=dt, id='unico')
            except: pass

# Rotas
@app.route('/api/data', methods=['GET'])
def get_data():
    return jsonify({'hospitals': hospitals, 'config': config})

@app.route('/api/config', methods=['POST'])
def update_config():
    global config
    config = request.json or {}
    save_data()
    reagendar()
    return jsonify({'status': 'ok'})

@app.route('/api/hospitals', methods=['POST'])
def add_hospital():
    global hospitals
    data = request.json or {}
    hospitals.append(data)
    save_data()
    return jsonify({'status': 'ok'})

@app.route('/api/hospitals/<int:index>', methods=['DELETE'])
def remove_hospital(index):
    global hospitals
    if 0 <= index < len(hospitals):
        del hospitals[index]
        save_data()
    return jsonify({'status': 'ok'})

@app.route('/api/run-stream', methods=['GET'])
def run_stream():
    def event_msg(obj): return f"data: {json.dumps(obj, default=str)}\n\n"
    @stream_with_context
    def gen():
        total = len(hospitals)
        yield event_msg({'type': 'meta', 'total': total})
        if total == 0:
            yield event_msg({'type': 'error', 'error': 'Nenhum hospital'})
            return

        with sync_playwright() as p:
            browser = p.chromium.launch(headless=True)
            page = browser.new_page()
            page.goto('https://sistemasaogeraldoservice.com.br/sistema/Login.aspx')
            page.fill('#txtUsuario', config.get('username', USUARIO_DEFAULT))
            page.fill('#txtSenha', config.get('password', SENHA_DEFAULT))
            page.click('#Button1')
            page.wait_for_load_state('networkidle')

            resultados = []
            for idx, h in enumerate(hospitals, 1):
                yield event_msg({'type': 'progress', 'idx': idx, 'current': idx-1, 'total': total, 'hospital': h['name'], 'status': 'Extraindo...'})
                total_kg, dados, periodo = extrair_dados_semana_anterior(page, h)
                resultados.append({'hospital': h['name'], 'periodo': periodo, 'total': total_kg, 'dados': dados})
                yield event_msg({'type': 'progress', 'idx': idx, 'current': idx, 'total': total, 'hospital': h['name'], 'status': f'{total_kg:.2f} kg'})

            gerar_relatorio(resultados)
            enviar_email(config.get('email', ''))

            with open(CAMINHO_RELATORIO, 'rb') as f:
                excel_b64 = base64.b64encode(f.read()).decode('utf-8')
                yield event_msg({'type': 'excel', 'data': excel_b64, 'filename': 'RelatorioLavanderia.xlsx'})

            yield event_msg({'type': 'done', 'results': resultados})

            browser.close()
    return Response(gen(), mimetype='text/event-stream')

if __name__ == '__main__':
    reagendar()
    app.run(host='0.0.0.0', port=int(os.getenv('PORT', 5000)))
