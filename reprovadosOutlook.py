import pandas as pd
import os
import smtplib
from email.message import EmailMessage

# ===== CONFIGURAÇÕES =====
CAMINHO_PLANILHA = r'C:\pyMail\reprovados.xlsx'
REGISTRO_ALERTAS = 'registro_alertas.txt'
EMAIL_REMETENTE = 'seu_email@outlook.com'  # seu Outlook
SENHA = 'sua_senha_ou_token_aqui'          # senha ou token de app Outlook
EMAIL_DESTINO = 'destino@example.com'      # para quem vai enviar

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

    # SMTP Outlook - usa STARTTLS na porta 587
    with smtplib.SMTP('smtp.office365.com', 587) as smtp:
        smtp.ehlo()
        smtp.starttls()
        smtp.ehlo()
        smtp.login(EMAIL_REMETENTE, SENHA)
        smtp.send_message(msg)
        print(f"E-mail enviado para {destinatario}")

# ===== PROGRAMA PRINCIPAL =====

df = pd.read_excel(CAMINHO_PLANILHA)

df.columns = df.columns.str.lower()

print(df.head())
print("Colunas da planilha:", df.columns.tolist())

reprovados = df[df['status'] == 'Reprovado'].copy()

reprovados['data_alerta'] = pd.to_datetime(reprovados['data_alerta'])

reprovados['identificador'] = reprovados['regiao'] + '_' + reprovados['data_alerta'].dt.strftime('%Y-%m-%d_%H-%M-%S')

enviados = carregar_registros_enviados()

novos_alertas = set(reprovados['identificador']) - enviados

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

# Alterar EMAIL_REMETENTE para seu e-mail Outlook real.

# Colocar a senha ou token no campo SENHA.

# Ajustar EMAIL_DESTINO para o destinatário correto.

# Se sua conta tiver autenticação multifator (MFA) ativada, recomendo criar uma senha de app no Outlook para esse script funcionar. 