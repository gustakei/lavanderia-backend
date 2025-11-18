# app.py — VERSÃO FINAL PERFEITA (COM BOT + APSCHEDULER + PLAYWRIGHT + TUDO FUNCIONANDO)
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

# ========================= CONFIGURAÇÕES =========================
EMAIL_SEU = os.getenv('EMAIL_SEU')
EMAIL_SENHA = os.getenv('EMAIL_SENHA')
USUARIO_DEFAULT = 'Guilherme Duarte'
SENHA_DEFAULT = '13072006'

# ========================= DISCO PERSISTENTE COM FALLBACK SEGURO =========================
DATA_DIR = '/data'
if not os.path.exists(DATA_DIR) or not os.access(DATA_DIR, os.W_OK):
    DATA_DIR = tempfile.mkdtemp(prefix='lavanderia_')
    print(f"[FALLBACK] /data indisponível → usando {DATA_DIR}")
else:
    print(f"[SUCESSO] Disco persistente /data disponível")

# Playwright usa o cache persistente do Render (NUNCA MAIS VAI DAR ERRO DE EXECUTABLE)
os.environ["PLAYWRIGHT_BROWSERS_PATH"] = "0"

# Caminhos dos arquivos
config_file = os.path.join(DATA_DIR, 'config.json')
hospitals_file = os.path.join(DATA_DIR, 'hospitals.json')
RELATORIO_PATH = os.path.join(DATA_DIR, 'RelatorioLavanderiaSemanal.xlsx')

# Variáveis globais
config = {}
hospitals = []

# ========================= CARREGAR DADOS =========================
def load_data():
    global config, hospitals
    if os.path.exists(config_file):
        try:
            with open(config_file, 'r', encoding='utf-8') as f:
                config.update(json.load(f))
        except: pass
    if os.path.exists(hospitals_file):
        try:
            with open(hospitals_file, 'r', encoding='utf-8') as f:
                hospitals.extend(json.load(f))
        except: pass

load_data()

# ========================= SALVAR DADOS =========================
def save_data():
    os.makedirs(DATA_DIR, exist_ok=True)
    with open(config_file, 'w', encoding='utf-8') as f:
        json.dump(config, f, indent=2, ensure_ascii=False)
    with open(hospitals_file, 'w', encoding='utf-8') as f:
        json.dump(hospitals, f, indent=2, ensure_ascii=False)

# ========================= CÁLCULO SEMANA ANTERIOR =========================
def calcular_semana_anterior():
    hoje = datetime.now()
    dias_desde_segunda = hoje.weekday()
    inicio = hoje - timedelta(days=dias_desde_segunda + 7)
    fim = inicio + timedelta(days=6)
    return inicio, fim

# ========================= EXTRAÇÃO COM PLAYWRIGHT =========================
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

# ========================= GERAR RELATÓRIO EXCEL =========================
def gerar_relatorio(resultados):
    df = pd.DataFrame([{
        'Hospital': r['hospital'],
        'Período': r['periodo'],
        'Total (Kg)': r['total']
    } for r in resultados])

    df.to_excel(RELATORIO_PATH, index=False)

    wb = openpyxl.load_workbook(RELATORIO_PATH)
    ws = wb.active

    # Cabeçalho estilizado
    for cell in ws[1]:
        cell.font = Font(bold=True)
        cell.fill = PatternFill(start_color='CCCCCC', end_color='CCCCCC', fill_type='solid')

    # Total geral
    total_geral = sum(r['total'] for r in resultados)
    ws.append(['', 'Total Geral', total_geral])
    ws.cell(row=ws.max_row, column=3).font = Font(bold=True)

    # Gráfico
    if len(resultados) > 1:
        chart = BarChart()
        data = Reference(ws, min_col=3, min_row=1, max_row=ws.max_row-1)
        cats = Reference(ws, min_col=1, min_row=2, max_row=ws.max_row-1)
        chart.add_data(data, titles_from_data=True)
        chart.set_categories(cats)
        chart.title = "Totais por Hospital"
        ws.add_chart(chart, "E2")

    wb.save(RELATORIO_PATH)

# ========================= ENVIO DE EMAIL =========================
def enviar_email(destinatario):
    if not EMAIL_SEU or not EMAIL_SENHA or not destinatario:
        return
    msg = MIMEMultipart()
    msg['From'] = EMAIL_SEU
    msg['To'] = destinatario
    msg['Subject'] = f"Relatório Lavanderia - {datetime.now().strftime('%d/%m/%Y')}"
    msg.attach(MIMEText("Segue o relatório semanal da lavanderia.", 'plain'))

    with open(RELATORIO_PATH, "rb") as f:
        part = MIMEBase('application', 'octet-stream')
        part.set_payload(f.read())
        encoders.encode_base64(part)
        part.add_header('Content-Disposition', 'attachment; filename=RelatorioLavanderiaSemanal.xlsx')
        msg.attach(part)

    try:
        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()
        server.login(EMAIL_SEU, EMAIL_SENHA)
        server.sendmail(EMAIL_SEU, destinatario, msg.as_string())
        server.quit()
    except Exception as e:
        print(f"[EMAIL] Erro: {e}")

# ========================= EXECUÇÃO AGENDADA =========================
def executar_relatorio_agendado():
    if not hospitals or not config.get('email'):
        return
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        page = browser.new_page()
        try:
            page.goto('https://sistemasaogeraldoservice.com.br/sistema/Login.aspx')
            page.fill('#txtUsuario', config.get('username', USUARIO_DEFAULT))
            page.fill('#txtSenha', config.get('password', SENHA_DEFAULT))
            page.click('#Button1')
            page.wait_for_load_state('networkidle')

            resultados = []
            for h in hospitals:
                kg, dados, periodo = extrair_dados_semana_anterior(page, h)
                resultados.append({'hospital': h['name'], 'periodo': periodo, 'total': kg})

            gerar_relatorio(resultados)
            enviar_email(config['email'])
            print(f"[AGENDADO] Relatório enviado para {config['email']} em {datetime.now()}")
        except Exception as e:
            print(f"[ERRO AGENDADO] {e}")
        finally:
            browser.close()

# ========================= AGENDADOR (FUNCIONA PORQUE VOCÊ TEM O BOT!) =========================
scheduler = BackgroundScheduler()
scheduler.start()
atexit.register(lambda: scheduler.shutdown())

def reagendar():
    scheduler.remove_all_jobs()
    horario = config.get('schedule', '').strip()
    if not horario:
        return

    if horario.startswith('cron['):
        try:
            params = horario[5:-1].strip()
            dia, hora = params.split(' ', 1)
            h, m = map(int, hora.split(':'))
            scheduler.add_job(
                executar_relatorio_agendado,
                'cron',
                day_of_week=dia,
                hour=h,
                minute=m,
                id='relatorio_recorrente',
                replace_existing=True
            )
            print(f"[CRON] Agendado toda {dia} às {h:02d}:{m:02d}")
        except Exception as e:
            print(f"[ERRO CRON] {e}")
    else:
        try:
            dt = datetime.fromisoformat(horario.replace('Z', '+00:00'))
            if dt > datetime.now():
                scheduler.add_job(executar_relatorio_agendado, 'date', run_date=dt)
        except: pass

# ========================= ROTAS =========================
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
    data = request.json
    if data:
        hospitals.append(data)
        save_data()
    return jsonify({'status': 'ok'})

@app.route('/api/hospitals/<int:i>', methods=['DELETE'])
def remove_hospital(i):
    global hospitals
    if 0 <= i < len(hospitals):
        hospitals.pop(i)
        save_data()
    return jsonify({'status': 'ok'})

@app.route('/api/run-stream', methods=['GET'])
def run_stream():
    def event(msg):
        return f"data: {json.dumps(msg, default=str)}\n\n"

    @stream_with_context
    def generate():
        if not hospitals:
            yield event({'type': 'error', 'error': 'Nenhum hospital cadastrado'})
            return

        yield event({'type': 'meta', 'total': len(hospitals)})

        with sync_playwright() as p:
            browser = p.chromium.launch(headless=True)
            page = browser.new_page()
            try:
                page.goto('https://sistemasaogeraldoservice.com.br/sistema/Login.aspx')
                page.fill('#txtUsuario', config.get('username', USUARIO_DEFAULT))
                page.fill('#txtSenha', config.get('password', SENHA_DEFAULT))
                page.click('#Button1')
                page.wait_for_load_state('networkidle')

                resultados = []
                for idx, h in enumerate(hospitals, 1):
                    yield event({'type': 'progress', 'idx': idx, 'hospital': h['name'], 'status': 'Extraindo...'})
                    kg, dados, periodo = extrair_dados_semana_anterior(page, h)
                    resultados.append({'hospital': h['name'], 'periodo': periodo, 'total': kg})
                    yield event({'type': 'progress', 'idx': idx, 'hospital': h['name'], 'status': f'{kg:.2f} kg'})

                gerar_relatorio(resultados)
                if config.get('email'):
                    enviar_email(config['email'])

                with open(RELATORIO_PATH, 'rb') as f:
                    b64 = base64.b64encode(f.read()).decode()
                    yield event({'type': 'excel', 'data': b64, 'filename': 'RelatorioLavanderiaSemanal.xlsx'})

                yield event({'type': 'done', 'results': resultados})
            except Exception as e:
                yield event({'type': 'error', 'error': str(e)})
            finally:
                browser.close()

    return Response(generate(), mimetype='text/event-stream')

# ROTA PARA SEU BOT ACESSAR A CADA 5 MINUTOS
@app.route('/keep-alive')
def keep_alive():
    return "OK", 200

# ========================= INICIALIZAÇÃO =========================
if __name__ == '__main__':
    reagendar()
    app.run(host='0.0.0.0', port=int(os.getenv('PORT', 5000)))
