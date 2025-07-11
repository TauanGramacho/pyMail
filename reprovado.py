import pandas as pd
import os
import smtplib
from email.message import EmailMessage

# ===== CONFIGURAÇÕES =====
CAMINHO_PLANILHA = r'C:\pyMail\reprovados.xlsx'  # caminho completo para o arquivo
REGISTRO_ALERTAS = 'registro_alertas.txt'
EMAIL_REMETENTE = 'tauanlaranjeiras@gmail.com'  # seu Gmail
SENHA_APP = 'imta flju vbne xukm'          # senha de aplicativo do Gmail
EMAIL_DESTINO = 'tauanlaranjeiras@gmail.com'  # para todos

# ===== FUNÇÕES =====

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
    msg = EmailMessage()
    msg['Subject'] = assunto
    msg['From'] = EMAIL_REMETENTE
    msg['To'] = destinatario
    msg.set_content(corpo)

    with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
        smtp.login(EMAIL_REMETENTE, SENHA_APP)
        smtp.send_message(msg)
        print(f"E-mail enviado para {destinatario}")

# ===== PROGRAMA PRINCIPAL =====

df = pd.read_excel(CAMINHO_PLANILHA)

# Normaliza o nome das colunas para minúsculas para evitar problemas
df.columns = df.columns.str.lower()

print(df.head())
print("Colunas da planilha:", df.columns.tolist())

# Filtra os reprovados
reprovados = df[df['status'] == 'Reprovado'].copy()

# Garante que a coluna data_alerta está em datetime
reprovados['data_alerta'] = pd.to_datetime(reprovados['data_alerta'])

# Cria identificador único: regiao + data_alerta formatada
reprovados['identificador'] = reprovados['regiao'] + '_' + reprovados['data_alerta'].dt.strftime('%Y-%m-%d_%H-%M-%S')

# Carrega os registros já enviados
enviados = carregar_registros_enviados()

# Identifica os novos alertas que ainda não foram enviados
novos_alertas = set(reprovados['identificador']) - enviados

for ident in novos_alertas:
    regiao, data_alerta_str = ident.split('_', 1)
    # Troca '_' por ' ' para data legível
    data_alerta_legivel = data_alerta_str.replace('_', ' ')
    assunto = f"[Alerta] Região {regiao} com reprovados em {data_alerta_legivel}"
    corpo = (
        f"Olá,\n\n"
        f"A região {regiao} possui colaborador(es) com status REPROVADO em {data_alerta_legivel}.\n"
        f"Favor verificar a situação.\n\n"
        f"Atenciosamente,\nSistema de Monitoramento"
    )
    enviar_email(EMAIL_DESTINO, assunto, corpo)

# Atualiza o arquivo com todos os alertas enviados, incluindo os novos
salvar_registros_enviados(enviados.union(novos_alertas))
