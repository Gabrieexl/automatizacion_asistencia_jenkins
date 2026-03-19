import os
import smtplib
from email.message import EmailMessage

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
OUTPUT_DIR = os.path.join(BASE_DIR, 'output')
PDF_FILE = os.path.join(OUTPUT_DIR, 'reporte_asistencia.pdf')
EXCEL_FILE = os.path.join(OUTPUT_DIR, 'resumen_asistencia.xlsx')

SMTP_SERVER = 'smtp.gmail.com'
SMTP_PORT = 587
EMAIL_USER = os.environ.get('EMAIL_USER')
EMAIL_PASS = os.environ.get('EMAIL_PASS')
EMAIL_TO = os.environ.get('EMAIL_TO')

print('=== ETAPA SEND MAIL ===')

if not EMAIL_USER or not EMAIL_PASS or not EMAIL_TO:
    raise ValueError('Faltan variables de entorno EMAIL_USER, EMAIL_PASS o EMAIL_TO en Jenkins.')

if not os.path.exists(PDF_FILE):
    raise FileNotFoundError('No se encontró el PDF para adjuntar.')
if not os.path.exists(EXCEL_FILE):
    raise FileNotFoundError('No se encontró el Excel resumen para adjuntar.')

msg = EmailMessage()
msg['Subject'] = 'Reporte automático de asistencia - Jenkins'
msg['From'] = EMAIL_USER
msg['To'] = EMAIL_TO
msg.set_content(
    'Hola,\n\n'
    'Se adjunta el reporte automático de asistencia generado por Jenkins.\n\n'
    'Adjuntos:\n'
    '- reporte_asistencia.pdf\n'
    '- resumen_asistencia.xlsx\n\n'
    'Saludos.'
)

with open(PDF_FILE, 'rb') as f:
    msg.add_attachment(f.read(), maintype='application', subtype='pdf', filename='reporte_asistencia.pdf')

with open(EXCEL_FILE, 'rb') as f:
    msg.add_attachment(f.read(), maintype='application', subtype='vnd.openxmlformats-officedocument.spreadsheetml.sheet', filename='resumen_asistencia.xlsx')

with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
    server.starttls()
    server.login(EMAIL_USER, EMAIL_PASS)
    server.send_message(msg)

print('Correo enviado correctamente.')