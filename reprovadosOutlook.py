import pandas as pd
import os
import smtplib
from email.message import EmailMessage
from datetime import datetime

# ===== CONFIGURAÇÕES =====
CAMINHO_PLANILHA = r'C:\Users\b621329\OneDrive - IBERDROLA S.A\pymail\pyMail\reprovados.xlsx'
REGISTRO_ALERTAS = 'registro_alertas.txt'
EMAIL_REMETENTE = 'tauanlaranjeiras@hotmail.com'
SENHA = 'dewzbxpzpcmrcglg'
EMAIL_DESTINO = 'tauan.santos@neoenergia.com'
MODO_TESTE = False  # Altere para False para enviar e-mails de verdade

# ===== FUNÇÕES =====

def log(msg):
    print(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] {msg}")

def carregar_registros_enviados():
    if os.path.exists(REGISTRO_ALERTAS):
        with open(REGISTRO_ALERTAS, 'r') as f:
            return set(l.strip() for l in f.readlines())
    return set()

def salvar_registros_enviados(registros):
    with open(REGISTRO_ALERTAS, 'w') as f:
        for registro in registros:
            f.write(f"{registro}\n")

def enviar_email(destinatario, assunto, corpo):
    log(f"Preparando envio para {destinatario}")
    if MODO_TESTE:
        log(f"[MODO TESTE] Simulando envio: {assunto}")
        return

    msg = EmailMessage()
    msg['Subject'] = assunto
    msg['From'] = EMAIL_REMETENTE
    msg['To'] = destinatario
    msg.set_content(corpo)

    try:
        log("Conectando ao servidor SMTP...")
        with smtplib.SMTP('smtp.office365.com', 587) as smtp:
            smtp.ehlo()
            log("Iniciando STARTTLS...")
            smtp.starttls()
            smtp.ehlo()
            log("Realizando login...")
            smtp.login(EMAIL_REMETENTE, SENHA)
            log("Enviando mensagem...")
            smtp.send_message(msg)
            log(f"E-mail enviado para {destinatario}")
    except Exception as e:
        log(f"Erro ao enviar e-mail: {e}")


# ===== PROGRAMA PRINCIPAL =====

log("Carregando planilha...")
df = pd.read_excel(CAMINHO_PLANILHA)
df.columns = df.columns.str.lower()
log("Planilha carregada com sucesso.")
log(f"Colunas: {df.columns.tolist()}")

reprovados = df[df['status'] == 'Reprovado'].copy()
reprovados['data_alerta'] = pd.to_datetime(reprovados['data_alerta'])
reprovados['identificador'] = reprovados['regiao'] + '_' + reprovados['data_alerta'].dt.strftime('%Y-%m-%d_%H-%M-%S')

enviados = carregar_registros_enviados()
novos_alertas = set(reprovados['identificador']) - enviados

log(f"Novos alertas encontrados: {len(novos_alertas)}")

for ident in novos_alertas:
    regiao, data_alerta_str = ident.split('_', 1)
    data_alerta_legivel = data_alerta_str.replace('_', ' ')
    assunto = f"[Alerta] Região {regiao} com reprovados em {data_alerta_legivel}"
    corpo = (
        f"Olá,\n\n"
        f"A região {regiao} possui colaborador(es) com status REPROVADO em {data_alerta_legivel}.\n"
        f"Favor verificar a situação.\n\n"
        f"Atenciosamente,\nSistema de Monitoramento"
    )
    enviar_email(EMAIL_DESTINO, assunto, corpo)

salvar_registros_enviados(enviados.union(novos_alertas))
log("Execução finalizada.")
