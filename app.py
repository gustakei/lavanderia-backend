# app.py — VERSÃO CORRIGIDA E OTIMIZADA PARA RENDER
from flask import Flask, request, jsonify, Response, stream_with_context, send_file
from flask_cors import CORS
import os
import json
import base64
import tempfile
from datetime import datetime, timedelta
from urllib.parse import urlparse, parse_qs
from bs4 import BeautifulSoup
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill
from openpyxl.chart import BarChart, Reference
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from playwright.sync_api import sync_playwright
from apscheduler.schedulers.background import BackgroundScheduler
import atexit
import logging

# Configurar logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

app = Flask(__name__)
CORS(app)

# ========================= CONFIGURAÇÃO DE PATHS =========================
# Verificação robusta do disco persistente
DATA_DIR = "/data"
try:
    # Testa se podemos escrever no /data
    test_file = os.path.join(DATA_DIR, "test_write.txt")
    with open(test_file, 'w') as f:
        f.write("test")
    os.remove(test_file)
    logger.info(f"[SUCESSO] Disco persistente montado em {DATA_DIR}")
except (OSError, IOError, PermissionError) as e:
    DATA_DIR = tempfile.mkdtemp(prefix="lavanderia_")
    logger.warning(f"[FALLBACK] /data não acessível → usando {DATA_DIR}. Erro: {e}")

# Caminhos finais
CONFIG_FILE = os.path.join(DATA_DIR, "config.json")
HOSPITALS_FILE = os.path.join(DATA_DIR, "hospitals.json")
RELATORIO_PATH = os.path.join(DATA_DIR, "RelatorioLavanderiaSemanal.xlsx")

# Configuração específica para Render - usar sistema browsers
os.environ["PLAYWRIGHT_BROWSERS_PATH"] = "0"  # Usar sistema browsers

# Variáveis globais
config = {}
hospitals = []

# ========================= CARREGAR DADOS =========================
def load_data():
    global config, hospitals
    try:
        if os.path.exists(CONFIG_FILE):
            with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
                loaded_config = json.load(f)
                config.update(loaded_config)
                logger.info("Configurações carregadas com sucesso")
        
        if os.path.exists(HOSPITALS_FILE):
            with open(HOSPITALS_FILE, 'r', encoding='utf-8') as f:
                loaded_hospitals = json.load(f)
                hospitals.extend(loaded_hospitals)
                logger.info(f"{len(loaded_hospitals)} hospitais carregados")
    except Exception as e:
        logger.error(f"Erro ao carregar dados: {e}")

load_data()

def save_data():
    try:
        os.makedirs(DATA_DIR, exist_ok=True)
        with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
            json.dump(config, f, indent=2, ensure_ascii=False)
        with open(HOSPITALS_FILE, 'w', encoding='utf-8') as f:
            json.dump(hospitals, f, indent=2, ensure_ascii=False)
        logger.info("Dados salvos com sucesso")
    except Exception as e:
        logger.error(f"Erro ao salvar dados: {e}")

# ========================= CÁLCULO SEMANA ANTERIOR =========================
def calcular_semana_anterior():
    hoje = datetime.now()
    dias_desde_segunda = hoje.weekday()
    inicio = hoje - timedelta(days=dias_desde_segunda + 7)
    fim = inicio + timedelta(days=6)
    return inicio, fim

# ========================= EXTRAÇÃO COM PLAYWRIGHT =========================
def extrair_dados_semana_anterior(page, hospital):
    try:
        inicio, fim = calcular_semana_anterior()
        periodo_text = f"{inicio.strftime('%d/%m/%Y')} a {fim.strftime('%d/%m/%Y')}"

        parsed = urlparse(hospital['url'])
        params = parse_qs(parsed.query)
        cliente_id = params.get('cliente', [None])[0]
        if not cliente_id:
            logger.warning(f"Cliente ID não encontrado para {hospital['name']}")
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
                logger.warning(f"Tabela não encontrada para {hospital['name']} - {mes_ano}")
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
                except ValueError: 
                    kg = 0.0
                todos_dados.append({'data': data_str, 'kg': kg})
                total_kg += kg

        logger.info(f"Extraídos {total_kg:.2f} kg para {hospital['name']}")
        return total_kg, todos_dados, periodo_text
    except Exception as e:
        logger.error(f"Erro na extração para {hospital['name']}: {e}")
        return 0.0, [], periodo_text

# ========================= RELATÓRIO EXCEL =========================
def gerar_relatorio(resultados):
    try:
        df = pd.DataFrame([{
            'Hospital': r['hospital'],
            'Período': r['periodo'],
            'Total (Kg)': r['total']
        } for r in resultados])

        df.to_excel(RELATORIO_PATH, index=False)

        wb = openpyxl.load_workbook(RELATORIO_PATH)
        ws = wb.active

        # Estilo cabeçalho
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
        logger.info("Relatório Excel gerado com sucesso")
    except Exception as e:
        logger.error(f"Erro ao gerar relatório: {e}")
        raise

# ========================= ENVIO DE EMAIL =========================
def enviar_email(destinatario):
    try:
        email = os.getenv('EMAIL_SEU')
        senha = os.getenv('EMAIL_SENHA')
        if not email or not senha or not destinatario:
            logger.warning("Credenciais de email ou destinatário não configuradas")
            return

        msg = MIMEMultipart()
        msg['From'] = email
        msg['To'] = destinatario
        msg['Subject'] = f"Relatório Lavanderia - {datetime.now().strftime('%d/%m/%Y')}"

        msg.attach(MIMEText("Segue em anexo o relatório semanal da lavanderia.", 'plain'))

        with open(RELATORIO_PATH, "rb") as f:
            part = MIMEBase('application', 'octet-stream')
            part.set_payload(f.read())
            encoders.encode_base64(part)
            part.add_header('Content-Disposition', f'attachment; filename=RelatorioLavanderia.xlsx')
            msg.attach(part)

        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()
        server.login(email, senha)
        server.sendmail(email, destinatario, msg.as_string())
        server.quit()
        logger.info("Email enviado com sucesso")
    except Exception as e:
        logger.error(f"Erro ao enviar email: {e}")

# ========================= AGENDAMENTO =========================
scheduler = BackgroundScheduler()
scheduler.start()
atexit.register(lambda: scheduler.shutdown(wait=False))

def executar_relatorio_agendado():
    if not hospitals or not config.get('email'):
        logger.warning("Agendamento: Hospitais ou email não configurados")
        return
    
    logger.info("Executando relatório agendado...")
    try:
        with sync_playwright() as p:
            # CORREÇÃO CRÍTICA: Configuração específica para Render
            browser = p.chromium.launch(
                headless=True,
                args=['--no-sandbox', '--disable-dev-shm-usage', '--disable-gpu']
            )
            page = browser.new_page()
            
            try:
                page.goto('https://sistemasaogeraldoservice.com.br/sistema/Login.aspx', timeout=30000)
                page.fill('#txtUsuario', config.get('username', 'Guilherme Duarte'))
                page.fill('#txtSenha', config.get('password', '13072006'))
                page.click('#Button1')
                page.wait_for_load_state('networkidle')

                resultados = []
                for h in hospitals:
                    kg, dados, periodo = extrair_dados_semana_anterior(page, h)
                    resultados.append({'hospital': h['name'], 'periodo': periodo, 'total': kg, 'dados': dados})

                gerar_relatorio(resultados)
                enviar_email(config['email'])
                logger.info("Relatório agendado executado com sucesso")
            except Exception as e:
                logger.error(f"Erro durante execução agendada: {e}")
            finally:
                browser.close()
    except Exception as e:
        logger.error(f"Erro ao inicializar Playwright no agendamento: {e}")

def reagendar():
    try:
        scheduler.remove_all_jobs()
        horario = config.get('schedule', '').strip()
        if not horario:
            logger.info("Nenhum agendamento configurado")
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
                    hour=h, minute=m,
                    id='recorrente'
                )
                logger.info(f"Agendamento recorrente configurado: {dia} às {h:02d}:{m:02d}")
            except Exception as e:
                logger.error(f"Erro ao configurar agendamento recorrente: {e}")
        else:
            try:
                dt = datetime.fromisoformat(horario.replace('Z', '+00:00'))
                if dt > datetime.now():
                    scheduler.add_job(executar_relatorio_agendado, 'date', run_date=dt)
                    logger.info(f"Agendamento único configurado para: {dt}")
            except Exception as e:
                logger.error(f"Erro ao configurar agendamento único: {e}")
    except Exception as e:
        logger.error(f"Erro no reagendamento: {e}")

# ========================= ROTAS =========================
@app.route('/')
def health_check():
    return jsonify({
        'status': 'online', 
        'data_dir': DATA_DIR,
        'hospitals_count': len(hospitals),
        'config_loaded': bool(config)
    })

@app.route('/api/data', methods=['GET'])
def get_data():
    return jsonify({'hospitals': hospitals, 'config': config})

@app.route('/api/config', methods=['POST'])
def update_config():
    global config
    try:
        config = request.json or {}
        save_data()
        reagendar()
        return jsonify({'status': 'ok', 'message': 'Configurações atualizadas'})
    except Exception as e:
        logger.error(f"Erro ao atualizar configurações: {e}")
        return jsonify({'status': 'error', 'message': str(e)}), 500

@app.route('/api/hospitals', methods=['POST'])
def add_hospital():
    global hospitals
    try:
        data = request.json
        if data: 
            hospitals.append(data)
            save_data()
            return jsonify({'status': 'ok', 'message': 'Hospital adicionado'})
        return jsonify({'status': 'error', 'message': 'Dados inválidos'}), 400
    except Exception as e:
        logger.error(f"Erro ao adicionar hospital: {e}")
        return jsonify({'status': 'error', 'message': str(e)}), 500

@app.route('/api/hospitals/<int:i>', methods=['DELETE'])
def remove_hospital(i):
    global hospitals
    try:
        if 0 <= i < len(hospitals):
            hospitals.pop(i)
            save_data()
            return jsonify({'status': 'ok', 'message': 'Hospital removido'})
        return jsonify({'status': 'error', 'message': 'Índice inválido'}), 404
    except Exception as e:
        logger.error(f"Erro ao remover hospital: {e}")
        return jsonify({'status': 'error', 'message': str(e)}), 500

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

        try:
            with sync_playwright() as p:
                # CORREÇÃO CRÍTICA: Configuração específica para Render
                browser = p.chromium.launch(
                    headless=True,
                    args=['--no-sandbox', '--disable-dev-shm-usage', '--disable-gpu']
                )
                page = browser.new_page()
                
                try:
                    page.goto('https://sistemasaogeraldoservice.com.br/sistema/Login.aspx', timeout=30000)
                    page.fill('#txtUsuario', config.get('username', 'Guilherme Duarte'))
                    page.fill('#txtSenha', config.get('password', '13072006'))
                    page.click('#Button1')
                    page.wait_for_load_state('networkidle')

                    resultados = []
                    for idx, h in enumerate(hospitals, 1):
                        yield event({'type': 'progress', 'idx': idx, 'hospital': h['name'], 'status': 'Extraindo...'})
                        kg, dados, periodo = extrair_dados_semana_anterior(page, h)
                        resultados.append({'hospital': h['name'], 'periodo': periodo, 'total': kg, 'dados': dados})
                        yield event({'type': 'progress', 'idx': idx, 'hospital': h['name'], 'status': f'{kg:.2f} kg'})

                    gerar_relatorio(resultados)
                    if config.get('email'):
                        enviar_email(config['email'])

                    # ENVIA O EXCEL PARA O FRONTEND
                    with open(RELATORIO_PATH, 'rb') as f:
                        b64 = base64.b64encode(f.read()).decode()
                        yield event({'type': 'excel', 'data': b64, 'filename': 'RelatorioLavanderia.xlsx'})

                    yield event({'type': 'done', 'results': resultados})
                except Exception as e:
                    logger.error(f"Erro durante execução: {e}")
                    yield event({'type': 'error', 'error': str(e)})
                finally:
                    browser.close()
        except Exception as e:
            logger.error(f"Erro ao inicializar Playwright: {e}")
            yield event({'type': 'error', 'error': f'Falha ao inicializar navegador: {str(e)}'})

    return Response(generate(), mimetype='text/event-stream')

if __name__ == '__main__':
    logger.info("Iniciando aplicação Flask...")
    reagendar()
    port = int(os.getenv('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=False)
