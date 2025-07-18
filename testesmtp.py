import smtplib

print("Testando conexão SMTP...")
with smtplib.SMTP('smtp.office365.com', 587, timeout=10) as smtp:
    smtp.ehlo()
    smtp.starttls()
    smtp.ehlo()
print("Conexão bem-sucedida!")
