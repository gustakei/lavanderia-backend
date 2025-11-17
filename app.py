# app.py — VERSÃO FINAL 100% FUNCIONAL NO RENDER FREE
from flask import Flask, request, jsonify, Response, stream_with_context
from flask_cors import CORS
import os
from dotenv import load_dotenv
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
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
import time
import json
import traceback
from urllib.parse import urlparse, parse_qs
from apscheduler.schedulers.background import BackgroundScheduler
import atexit

# === INICIALIZAÇÃO ===
app = Flask(__name__)
CORS(app)
load_dotenv()

# === CONFIGURAÇÕES ===
EMAIL_SEU = os.getenv('EMAIL_SEU')
EMAIL_SENHA = os.getenv('EMAIL_SENHA')
CHROMEDRIVER_PATH = os.getenv('CHROMEDRIVER_PATH', '/usr/local/bin/chromedriver')

# Caminho persistente (com disco /data ou fallback)
DATA_DIR = os.getenv('DISK_MOUNT_PATH', '.')
os.makedirs(DATA_DIR, exist_ok=True)
CAMINHO_RELATORIO = os.path.join(DATA_DIR, 'RelatorioLavanderiaSemanal.xlsx')
config_file = os.path.join(DATA_DIR, 'config.json')
hospitals_file = os.path.join(DATA_DIR, 'hospitals.json')

# Credenciais padrão
USUARIO_DEFAULT = 'Guilherme Duarte'
SENHA_DEFAULT = '13072006'

# Dados globais
config = {}
hospitals = []

# === CARREGAR DADOS ===
def load_data():
    global config, hospitals
    if os.path.exists(config_file):
        with open(config_file, 'r', encoding='utf-8') as f:
            try: config = json.load(f)
            except: config = {}
    if os.path.exists(hospitals_file):
        with open(hospitals_file, 'r', encoding='utf-8') as f:
            try: hospitals = json.load(f)
            except: hospitals = []
load_data()

def save_data():
    with open(config_file, 'w', encoding='utf-8') as f:
        json.dump(config, f, indent=2, ensure_ascii=False)
    with open(hospitals_file, 'w', encoding='utf-8') as f:
        json.dump(hospitals, f, indent=2, ensure_ascii=False)

# === CÁLCULO DE SEMANA ===
def calcular_semana_anterior():
    hoje = datetime.now()
    dias_desde_segunda = hoje.weekday()
    inicio = hoje - timedelta(days=dias_desde_segunda + 7)
    fim = inicio + timedelta(days=6)
    return inicio, fim

# === DRIVER ===
def make_driver():
    options = Options()
    options.add_argument("--headless=new")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--disable-gpu")
    service = Service(CHROMEDRIVER_PATH)
    return webdriver.Chrome(service=service, options=options)

# === LOGIN ===
def fazer_login(driver):
    driver.get('https://sistemasaogeraldoservice.com.br/sistema/Login.aspx')
    wait = WebDriverWait(driver, 30)
    usuario_field = wait.until(EC.presence_of_element_located((By.ID, 'txtUsuario')))
    senha_field = driver.find_element(By.ID, 'txtSenha')
    botao = driver.find_element(By.ID, 'Button1')
    usuario_field.clear(); usuario_field.send_keys(config.get('username', USUARIO_DEFAULT))
    senha_field.clear(); senha_field.send_keys(config.get('password', SENHA_DEFAULT))
    botao.click()
    time.sleep(5)

# === EXTRAÇÃO ===
def extrair_dados_semana_anterior(driver, hospital):
    inicio, fim = calcular_semana_anterior()
    periodo_text = f"{inicio.strftime('%d/%m/%Y')} a {fim.strftime('%d/%m/%Y')}"
    parsed = urlparse(hospital['url'])
    params = parse_qs(parsed.query)
    cliente_id = params.get('cliente', [None])[0]
    if not cliente_id: return 0.0, [], periodo_text

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
        driver.get(f"{base_url}?cliente={cliente_id}&periodo={mes_ano}")
        time.sleep(5)
        soup = BeautifulSoup(driver.page_source, 'html.parser')
        tabela = soup.find('table', id='tabpedidos')
        if not tabela:
            tabela = next((t for t in soup.find_all('table') if 'DIA' in t.get_text(strip=True).upper()), None)
        if not tabela: continue
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

# === RELATÓRIO (CORRIGIDO GLOBAL) ===
def gerar_relatorio(resultados):
    global CAMINHO_RELATORIO
    df = pd.DataFrame([{'Hospital': r['hospital'], 'Período': r['periodo'], 'Total (Kg)': r['total']} for r in resultados])
    try:
        df.to_excel(CAMINHO_RELATORIO, index=False)
    except Exception as e:
        print(f"[ERRO SALVAR EXCEL] {e}")
        CAMINHO_RELATORIO = os.path.join(DATA_DIR, f"Relatorio_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx")
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
        ws.add_chart(chart, "E2")
    wb.save(CAMINHO_RELATORIO)

# === E-MAIL ===
def enviar_email(email_dest):
    if not EMAIL_SEU or not EMAIL_SENHA or not email_dest: return
    msg = MIMEMultipart()
    msg['From'] = EMAIL_SEU
    msg['To'] = email_dest
    msg['Subject'] = f"Relatório Lavanderia {datetime.now().strftime('%d/%m/%Y')}"
    msg.attach(MIMEText("Segue o relatório em anexo.", 'plain'))
    try:
        with open(CAMINHO_RELATORIO, "rb") as f:
            part = MIMEBase('application', 'octet-stream')
            part.set_payload(f.read())
            encoders.encode_base64(part)
            part.add_header('Content-Disposition', f'attachment; filename={os.path.basename(CAMINHO_RELATORIO)}')
            msg.attach(part)
        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()
        server.login(EMAIL_SEU, EMAIL_SENHA)
        server.sendmail(EMAIL_SEU, email_dest, msg.as_string())
        server.quit()
    except Exception as e:
        print(f"[ERRO ENVIO EMAIL] {e}")

# === EXECUÇÃO AUTOMÁTICA ===
def executar_relatorio_agendado():
    if not hospitals or not config.get('email'): return
    driver = make_driver()
    try:
        fazer_login(driver)
        resultados = []
        for h in hospitals:
            total_kg, dados, periodo = extrair_dados_semana_anterior(driver, h)
            resultados.append({'hospital': h['name'], 'periodo': periodo, 'total': total_kg, 'dados': dados})
        gerar_relatorio(resultados)
        enviar_email(config['email'])
    except Exception as e:
        print(f"[ERRO AGENDADO] {e}")
    finally:
        try: driver.quit()
        except: pass

# === AGENDADOR ===
scheduler = BackgroundScheduler()
scheduler.start()
atexit.register(lambda: scheduler.shutdown())

def reagendar():
    scheduler.remove_all_jobs()
    horario = config.get('schedule')
    if horario:
        try:
            dt = datetime.fromisoformat(horario.replace('Z', '+00:00') if 'Z' in horario else horario)
            if dt > datetime.now():
                scheduler.add_job(
                    executar_relatorio_agendado,
                    'date',
                    run_date=dt,
                    id='relatorio_semanal'
                )
        except Exception as e:
            print(f"[ERRO AGENDAMENTO] {e}")

# === ROTAS ===
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
        driver = make_driver()
        try:
            fazer_login(driver)
            for idx, h in enumerate(hospitals, start=1):
                yield event_msg({'type': 'progress', 'idx': idx, 'current': idx-1, 'hospital': h['name'], 'status': 'Extraindo...'})
                total_kg, dados, periodo = extrair_dados_semana_anterior(driver, h)
                yield event_msg({'type': 'progress', 'idx': idx, 'current': idx, 'hospital': h['name'], 'status': f'{total_kg:.2f} kg'})
                time.sleep(1)
            resultados = []
            for h in hospitals:
                total_kg, dados, periodo = extrair_dados_semana_anterior(driver, h)
                resultados.append({'hospital': h['name'], 'periodo': periodo, 'total': total_kg, 'dados': dados})
            gerar_relatorio(resultados)
            enviar_email(config.get('email', ''))
            yield event_msg({'type': 'done', 'results': resultados})
        except Exception as e:
            yield event_msg({'type': 'error', 'error': str(e)})
        finally:
            try: driver.quit()
            except: pass
    return Response(gen(), mimetype='text/event-stream')

# === INICIALIZAÇÃO ===
if __name__ == '__main__':
    reagendar()
    app.run(host='0.0.0.0', port=int(os.getenv('PORT', 5000)))
