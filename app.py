# app.py — VERSÃO FINAL DEFINITIVA (Playwright + Cron Recorrente + Download Excel)
from flask import Flask, request, jsonify, Response, stream_with_context
from flask_cors import CORS
import os
import tempfile
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
import traceback

# ========================= Configuráveis / Paths =========================
# Prefer /data (montado no Render). Se não for possível, faz fallback para tempdir.
DEFAULT_DATA_DIR = os.environ.get("DATA_DIR", "/data")

# Variáveis globais de caminho (serão atualizadas na init)
DATA_DIR = None
BROWSERS_DIR = None
CONFIG_FILE = None
HOSPITALS_FILE = None
CAMINHO_RELATORIO = None

# ========================= Inicialização segura de storage =========================
def init_storage():
    global DATA_DIR, BROWSERS_DIR, CONFIG_FILE, HOSPITALS_FILE, CAMINHO_RELATORIO

    # tenta usar /data, senão usa tempdir
    try:
        # preferir o DEFAULT_DATA_DIR (normalmente /data no Render)
        os.makedirs(DEFAULT_DATA_DIR, exist_ok=True)
        # testar permissão escrita
        testfile = os.path.join(DEFAULT_DATA_DIR, ".touch_test")
        with open(testfile, "w") as f:
            f.write("ok")
        os.remove(testfile)
        DATA_DIR = DEFAULT_DATA_DIR
        print(f"[INIT] Usando DATA_DIR = {DATA_DIR}")
    except Exception as e:
        # fallback
        tmp = tempfile.mkdtemp(prefix="lavanderia_data_")
        DATA_DIR = tmp
        print(f"[INIT] Não foi possível usar {DEFAULT_DATA_DIR} ({e}). Fazendo fallback para {DATA_DIR}")

    # browsers (para Playwright)
    BROWSERS_DIR = os.path.join(DATA_DIR, "browsers")
    try:
        os.makedirs(BROWSERS_DIR, exist_ok=True)
    except Exception as e:
        # fallback para temp, caso permissão falhe
        tmp = tempfile.mkdtemp(prefix="lavanderia_browsers_")
        BROWSERS_DIR = tmp
        print(f"[INIT] Falha criando browsers dir em {BROWSERS_DIR}: {e}. Usando {tmp}")

    # ajusta variável de ambiente para Playwright (importante)
    os.environ["PLAYWRIGHT_BROWSERS_PATH"] = BROWSERS_DIR
    print(f"[INIT] PLAYWRIGHT_BROWSERS_PATH = {BROWSERS_DIR}")

    # arquivos de config
    CONFIG_FILE = os.path.join(DATA_DIR, "config.json")
    HOSPITALS_FILE = os.path.join(DATA_DIR, "hospitals.json")
    CAMINHO_RELATORIO = os.path.join(DATA_DIR, "RelatorioLavanderiaSemanal.xlsx")

    # cria arquivos vazios se não existirem (em try/except)
    try:
        if not os.path.exists(CONFIG_FILE):
            with open(CONFIG_FILE, "w", encoding="utf-8") as f:
                json.dump({}, f)
            print(f"[INIT] Criado arquivo default: {CONFIG_FILE}")
    except Exception as e:
        print(f"[INIT] Erro criando {CONFIG_FILE}: {e}")
    try:
        if not os.path.exists(HOSPITALS_FILE):
            with open(HOSPITALS_FILE, "w", encoding="utf-8") as f:
                json.dump([], f)
            print(f"[INIT] Criado arquivo default: {HOSPITALS_FILE}")
    except Exception as e:
        print(f"[INIT] Erro criando {HOSPITALS_FILE}: {e}")

# roda inicialização de storage já no import (robusto)
init_storage()

# ========================= App & env =========================
app = Flask(__name__)
CORS(app)
load_dotenv()  # carrega .env se houver

# Carrega configurações do ambiente (podem depender do load_dotenv)
EMAIL_SEU = os.getenv('EMAIL_SEU')
EMAIL_SENHA = os.getenv('EMAIL_SENHA')
USUARIO_DEFAULT = 'Guilherme Duarte'
SENHA_DEFAULT = '13072006'

# Usa as paths definidas na init_storage (variáveis globais definidas lá)
# Se algo não estiver definido por qualquer motivo, define defaults mínimos
if 'DATA_DIR' not in globals() or DATA_DIR is None:
    DATA_DIR = tempfile.mkdtemp(prefix="lavanderia_data_fallback_")
if 'BROWSERS_DIR' not in globals() or BROWSERS_DIR is None:
    BROWSERS_DIR = tempfile.mkdtemp(prefix="lavanderia_browsers_fallback_")
if 'CAMINHO_RELATORIO' not in globals() or CAMINHO_RELATORIO is None:
    CAMINHO_RELATORIO = os.path.join(DATA_DIR, "RelatorioLavanderiaSemanal.xlsx")

CONFIG_FILE = globals().get('CONFIG_FILE', os.path.join(DATA_DIR, "config.json"))
HOSPITALS_FILE = globals().get('HOSPITALS_FILE', os.path.join(DATA_DIR, "hospitals.json"))

print(f"[BOOT] CONFIG_FILE={CONFIG_FILE}, HOSPITALS_FILE={HOSPITALS_FILE}, RELATORIO={CAMINHO_RELATORIO}")

# ========================= Estado em memória =========================
config = {}
hospitals = []

# ========================= Funções de persistência =========================
def load_data():
    global config, hospitals
    try:
        if os.path.exists(CONFIG_FILE):
            with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
                try:
                    config = json.load(f)
                except Exception:
                    print(f"[LOAD] arquivo {CONFIG_FILE} inválido — resetando config")
                    config = {}
        else:
            config = {}
    except Exception as e:
        print(f"[LOAD] Erro lendo {CONFIG_FILE}: {e}")
        config = {}

    try:
        if os.path.exists(HOSPITALS_FILE):
            with open(HOSPITALS_FILE, 'r', encoding='utf-8') as f:
                try:
                    hospitals = json.load(f)
                except Exception:
                    print(f"[LOAD] arquivo {HOSPITALS_FILE} inválido — resetando hospitals")
                    hospitals = []
        else:
            hospitals = []
    except Exception as e:
        print(f"[LOAD] Erro lendo {HOSPITALS_FILE}: {e}")
        hospitals = []

def save_data():
    try:
        with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
            json.dump(config, f, indent=2, ensure_ascii=False)
    except Exception as e:
        print(f"[SAVE] Falha ao salvar {CONFIG_FILE}: {e}")
    try:
        with open(HOSPITALS_FILE, 'w', encoding='utf-8') as f:
            json.dump(hospitals, f, indent=2, ensure_ascii=False)
    except Exception as e:
        print(f"[SAVE] Falha ao salvar {HOSPITALS_FILE}: {e}")

# carregar estado na inicialização do módulo (agora que paths estão garantidos/fallback)
load_data()

# ========================= Utilidades de datas =========================
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

    # cuidado se hospital não tiver 'url'
    url = hospital.get('url', '') if isinstance(hospital, dict) else ''
    parsed = urlparse(url)
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
            if len(cols) < 2:
                continue
            data_str = cols[0].get_text(strip=True)
            if data_str not in dias:
                continue
            kg_text = cols[1].get_text(strip=True).replace('.', '').replace(' ', '').replace(',', '.')
            try:
                kg = float(kg_text)
            except:
                kg = 0.0
            todos_dados.append({'data': data_str, 'kg': kg})
            total_kg += kg

    return total_kg, todos_dados, periodo_text

# ========================= RELATÓRIO EXCEL =========================
def gerar_relatorio(resultados):
    global CAMINHO_RELATORIO
    try:
        df = pd.DataFrame([{'Hospital': r['hospital'], 'Período': r['periodo'], 'Total (Kg)': r['total']} for r in resultados])
    except Exception as e:
        print(f"[REPORT] Erro criando DataFrame: {e}")
        df = pd.DataFrame(columns=['Hospital', 'Período', 'Total (Kg)'])

    # tenta escrever no caminho definido (normalmente em /data)
    try:
        df.to_excel(CAMINHO_RELATORIO, index=False)
    except Exception as e:
        # fallback: arquivo em temp
        fallback = os.path.join(tempfile.gettempdir(), f"Relatorio_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx")
        try:
            df.to_excel(fallback, index=False)
            CAMINHO_RELATORIO = fallback
            print(f"[REPORT] Falha salvando em {CAMINHO_RELATORIO} — salvo em fallback {fallback}")
        except Exception as e2:
            print(f"[REPORT] Falha dupla salvando Excel: {e} | {e2}")
            return

    try:
        wb = openpyxl.load_workbook(CAMINHO_RELATORIO)
        ws = wb.active
        for cell in ws[1]:
            cell.font = Font(bold=True)
            cell.fill = PatternFill(start_color='CCCCCC', end_color='CCCCCC', fill_type='solid')
        total = sum(r.get('total', 0) for r in resultados)
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
    except Exception as e:
        print(f"[REPORT] Erro ao formatar/salvar workbook: {e}\n{traceback.format_exc()}")

# ========================= ENVIO DE EMAIL =========================
def enviar_email(email_dest):
    if not EMAIL_SEU or not EMAIL_SENHA or not email_dest:
        print("[EMAIL] Credenciais ou destinatário ausentes, pulando envio.")
        return
    try:
        msg = MIMEMultipart()
        msg['From'] = EMAIL_SEU
        msg['To'] = email_dest
        msg['Subject'] = f"Relatório Lavanderia {datetime.now().strftime('%d/%m/%Y')}"
        msg.attach(MIMEText("Segue o relatório semanal em anexo.", 'plain'))
        with open(CAMINHO_RELATORIO, "rb") as f:
            part = MIMEBase('application', 'octet-stream')
            part.set_payload(f.read())
            encoders.encode_base64(part)
            part.add_header('Content-Disposition', f'attachment; filename={os.path.basename(CAMINHO_RELATORIO)}')
            msg.attach(part)
        server = smtplib.SMTP('smtp.gmail.com', 587, timeout=30)
        server.starttls()
        server.login(EMAIL_SEU, EMAIL_SENHA)
        server.sendmail(EMAIL_SEU, email_dest, msg.as_string())
        server.quit()
        print(f"[EMAIL] Relatório enviado para {email_dest}")
    except Exception as e:
        print(f"[EMAIL] Falha ao enviar email: {e}\n{traceback.format_exc()}")

# ========================= EXECUÇÃO AUTOMÁTICA (Playwright) =========================
def executar_relatorio_agendado():
    if not hospitals or not config.get('email'):
        print("[AGENDADO] Sem hospitais ou email configurado — pulando.")
        return

    # garante browsers dir existe antes de usar playwright
    try:
        os.makedirs(BROWSERS_DIR, exist_ok=True)
    except Exception as e:
        print(f"[AGENDADO] Falha criando BROWSERS_DIR: {e}")

    try:
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
                    total_kg, dados, periodo = extrair_dados_semana_anterior(page, h)
                    resultados.append({'hospital': h.get('name', '---'), 'periodo': periodo, 'total': total_kg, 'dados': dados})

                gerar_relatorio(resultados)
                enviar_email(config.get('email'))
            except Exception as e:
                print(f"[ERRO AGENDADO] {e}\n{traceback.format_exc()}")
            finally:
                try:
                    browser.close()
                except:
                    pass
    except Exception as e:
        print(f"[AGENDADO] Falha inicializando Playwright: {e}\n{traceback.format_exc()}")

# ========================= AGENDADOR =========================
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
            params = horario[5:-1].strip()  # ex: "2 08:00"
            day_str, time_str = params.split(' ', 1)
            hour, minute = map(int, time_str.split(':'))
            scheduler.add_job(
                executar_relatorio_agendado,
                'cron',
                day_of_week=day_str,
                hour=hour,
                minute=minute,
                id='relatorio_recorrente',
                replace_existing=True
            )
            print(f"[CRON] Agendado recorrente: toda {day_str} às {hour:02d}:{minute:02d}")
        except Exception as e:
            print(f"[ERRO CRON] {e}\n{traceback.format_exc()}")
    else:
        try:
            dt = datetime.fromisoformat(horario.replace('Z', '+00:00') if 'Z' in horario else horario)
            if dt > datetime.now():
                scheduler.add_job(executar_relatorio_agendado, 'date', run_date=dt, id='relatorio_unico')
        except Exception as e:
            print(f"[ERRO CRON DATE] {e}\n{traceback.format_exc()}")

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
            yield event_msg({'type': 'error', 'error': 'Nenhum hospital cadastrado'})
            return

        # garante browsers dir existe
        try:
            os.makedirs(BROWSERS_DIR, exist_ok=True)
        except Exception as e:
            print(f"[STREAM] Falha criando BROWSERS_DIR: {e}")

        try:
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
                    for idx, h in enumerate(hospitals, start=1):
                        yield event_msg({'type': 'progress', 'idx': idx, 'current': idx-1, 'total': total, 'hospital': h.get('name', ''), 'status': 'Extraindo...'})
                        total_kg, dados, periodo = extrair_dados_semana_anterior(page, h)
                        resultados.append({'hospital': h.get('name', ''), 'periodo': periodo, 'total': total_kg, 'dados': dados})
                        yield event_msg({'type': 'progress', 'idx': idx, 'current': idx, 'total': total, 'hospital': h.get('name', ''), 'status': f'{total_kg:.2f} kg'})

                    gerar_relatorio(resultados)
                    enviar_email(config.get('email', ''))

                    # ENVIA O EXCEL PARA DOWNLOAD NO FRONTEND
                    try:
                        with open(CAMINHO_RELATORIO, 'rb') as f:
                            excel_b64 = base64.b64encode(f.read()).decode('utf-8')
                            yield event_msg({'type': 'excel', 'data': excel_b64, 'filename': os.path.basename(CAMINHO_RELATORIO)})
                    except Exception as e:
                        yield event_msg({'type': 'warning', 'warning': f"Relatório não encontrado: {e}"})

                    yield event_msg({'type': 'done', 'results': resultados})
                except Exception as e:
                    yield event_msg({'type': 'error', 'error': str(e)})
                finally:
                    try:
                        browser.close()
                    except:
                        pass
        except Exception as e:
            yield event_msg({'type': 'error', 'error': f"Playwright failure: {e}"})

    return Response(gen(), mimetype='text/event-stream')

# ========================= BOOT quando rodar localmente com python app.py =========================
if __name__ == '__main__':
    # garante diretórios (se ainda não existirem)
    try:
        os.makedirs(DATA_DIR, exist_ok=True)
        os.makedirs(BROWSERS_DIR, exist_ok=True)
    except Exception as e:
        print(f"[MAIN] Falha criando dirs no boot: {e}")

    # cria arquivos default caso não existam (seguro)
    try:
        if not os.path.exists(CONFIG_FILE):
            with open(CONFIG_FILE, "w", encoding="utf-8") as f:
                json.dump({}, f)
        if not os.path.exists(HOSPITALS_FILE):
            with open(HOSPITALS_FILE, "w", encoding="utf-8") as f:
                json.dump([], f)
    except Exception as e:
        print(f"[MAIN] Falha ao criar arquivos default: {e}")

    # recarrega dados (caso criados agora)
    load_data()

    # reagenda se necessário
    reagendar()

    port = int(os.getenv('PORT', 5000))
    app.run(host='0.0.0.0', port=port)
