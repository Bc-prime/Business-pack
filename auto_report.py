import smtplib
import schedule
import time
import os
import pandas as pd
from dotenv import load_dotenv
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from datetime import datetime

# Cargar variables de entorno
load_dotenv()

# ============================================
# CONFIGURACIÓN — El cliente solo edita esto
# ============================================
EMAIL_REMITENTE = os.getenv("EMAIL_REMITENTE")
CONTRASEÑA_APP = os.getenv("CONTRASEÑA_APP")
EMAIL_CLIENTE = os.getenv("EMAIL_CLIENTE")
NOMBRE_NEGOCIO = os.getenv("NOMBRE_NEGOCIO")
HORA_ENVIO = "08:00"
STOCK_MINIMO = 10
# ============================================

def generar_reporte():
    df = pd.read_excel("inventario.xlsx")
    stock_bajo = df[df["Stock"] <= STOCK_MINIMO].copy().sort_values("Stock")
    stock_ok = df[df["Stock"] > STOCK_MINIMO].copy()
    return df, stock_bajo, stock_ok

def generar_html(df, stock_bajo, stock_ok):
    fecha = datetime.now().strftime("%B %d, %Y")
    total = len(df)
    alertas = len(stock_bajo)
    ok = len(stock_ok)

    # Filas de productos con stock bajo
    filas_bajo = ""
    for _, row in stock_bajo.iterrows():
        color = "#FEE2E2" if row["Stock"] <= 3 else "#FED7AA"
        emoji = "🔴" if row["Stock"] <= 3 else "🟠"
        filas_bajo += f"""
        <tr style="background:{color}">
            <td style="padding:10px 14px;font-weight:600">{emoji} {row['Product']}</td>
            <td style="padding:10px 14px;text-align:center">{row['Category']}</td>
            <td style="padding:10px 14px;text-align:center;font-weight:700;color:#DC2626">{row['Stock']} units</td>
            <td style="padding:10px 14px;text-align:center">
                <span style="background:#DC2626;color:white;padding:3px 10px;border-radius:20px;font-size:11px;font-weight:700">REORDER NOW</span>
            </td>
        </tr>"""

    # Filas de productos OK
    filas_ok = ""
    for _, row in stock_ok.iterrows():
        filas_ok += f"""
        <tr style="background:#F8FAFC">
            <td style="padding:10px 14px;font-weight:600">✅ {row['Product']}</td>
            <td style="padding:10px 14px;text-align:center">{row['Category']}</td>
            <td style="padding:10px 14px;text-align:center;font-weight:700;color:#16A34A">{row['Stock']} units</td>
            <td style="padding:10px 14px;text-align:center">
                <span style="background:#16A34A;color:white;padding:3px 10px;border-radius:20px;font-size:11px;font-weight:700">OK</span>
            </td>
        </tr>"""

    html = f"""
<!DOCTYPE html>
<html>
<head><meta charset="UTF-8"></head>
<body style="margin:0;padding:0;background:#F1F5F9;font-family:Arial,sans-serif">

    <!-- WRAPPER -->
    <table width="100%" cellpadding="0" cellspacing="0" style="background:#F1F5F9;padding:30px 0">
    <tr><td align="center">
    <table width="600" cellpadding="0" cellspacing="0" style="background:white;border-radius:12px;overflow:hidden;box-shadow:0 4px 20px rgba(0,0,0,0.1)">

        <!-- HEADER -->
        <tr>
            <td style="background:linear-gradient(135deg,#1E293B 0%,#334155 100%);padding:30px 35px">
                <table width="100%">
                    <tr>
                        <td>
                            <p style="margin:0;color:#93C5FD;font-size:12px;font-weight:600;letter-spacing:2px;text-transform:uppercase">Automated Report</p>
                            <h1 style="margin:6px 0 0;color:white;font-size:22px">{NOMBRE_NEGOCIO}</h1>
                        </td>
                        <td align="right">
                            <p style="margin:0;color:#64748B;font-size:12px">{fecha}</p>
                            <p style="margin:4px 0 0;color:#93C5FD;font-size:13px;font-weight:600">📦 Inventory Alert</p>
                        </td>
                    </tr>
                </table>
            </td>
        </tr>

        <!-- KPI CARDS -->
        <tr>
            <td style="padding:25px 35px 10px">
                <table width="100%" cellspacing="0" cellpadding="0">
                    <tr>
                        <td width="32%" style="padding:5px">
                            <div style="background:#F8FAFC;border-radius:10px;padding:18px;text-align:center;border-top:4px solid #3B82F6">
                                <p style="margin:0;color:#64748B;font-size:11px;font-weight:600;text-transform:uppercase;letter-spacing:1px">Total Products</p>
                                <p style="margin:8px 0 0;color:#1E293B;font-size:28px;font-weight:700">{total}</p>
                            </div>
                        </td>
                        <td width="32%" style="padding:5px">
                            <div style="background:#F0FDF4;border-radius:10px;padding:18px;text-align:center;border-top:4px solid #16A34A">
                                <p style="margin:0;color:#16A34A;font-size:11px;font-weight:600;text-transform:uppercase;letter-spacing:1px">Stock OK</p>
                                <p style="margin:8px 0 0;color:#16A34A;font-size:28px;font-weight:700">{ok}</p>
                            </div>
                        </td>
                        <td width="32%" style="padding:5px">
                            <div style="background:#FFF1F2;border-radius:10px;padding:18px;text-align:center;border-top:4px solid #DC2626">
                                <p style="margin:0;color:#DC2626;font-size:11px;font-weight:600;text-transform:uppercase;letter-spacing:1px">Low Stock</p>
                                <p style="margin:8px 0 0;color:#DC2626;font-size:28px;font-weight:700">{alertas}</p>
                            </div>
                        </td>
                    </tr>
                </table>
            </td>
        </tr>

        <!-- ALERTA -->
        {"" if alertas == 0 else f'''
        <tr>
            <td style="padding:10px 35px">
                <div style="background:#FEF2F2;border:1px solid #FECACA;border-radius:8px;padding:14px 18px">
                    <p style="margin:0;color:#DC2626;font-weight:700;font-size:14px">🚨 {alertas} product(s) need to be reordered immediately!</p>
                </div>
            </td>
        </tr>'''}

        <!-- TABLA PRODUCTOS -->
        <tr>
            <td style="padding:20px 35px">
                <p style="margin:0 0 12px;color:#1E293B;font-size:15px;font-weight:700">Inventory Status</p>
                <table width="100%" cellpadding="0" cellspacing="0" style="border-radius:8px;overflow:hidden;border:1px solid #E2E8F0">
                    <tr style="background:#1E293B">
                        <th style="padding:10px 14px;color:white;font-size:12px;text-align:left">Product</th>
                        <th style="padding:10px 14px;color:white;font-size:12px;text-align:center">Category</th>
                        <th style="padding:10px 14px;color:white;font-size:12px;text-align:center">Stock</th>
                        <th style="padding:10px 14px;color:white;font-size:12px;text-align:center">Status</th>
                    </tr>
                    {filas_bajo}
                    {filas_ok}
                </table>
            </td>
        </tr>

        <!-- FOOTER -->
        <tr>
            <td style="background:#F8FAFC;padding:20px 35px;border-top:1px solid #E2E8F0">
                <p style="margin:0;color:#64748B;font-size:12px;text-align:center">
                    This report was generated automatically by <strong>{NOMBRE_NEGOCIO} Automation System</strong><br>
                    Powered by Python 🐍
                </p>
            </td>
        </tr>

    </table>
    </td></tr>
    </table>

</body>
</html>"""
    return html

def enviar_email():
    print("⏰ Generando y enviando reporte...")
    df, stock_bajo, stock_ok = generar_reporte()
    html = generar_html(df, stock_bajo, stock_ok)

    msg = MIMEMultipart("alternative")
    msg["From"] = EMAIL_REMITENTE
    msg["To"] = EMAIL_CLIENTE
    msg["Subject"] = f"📦 Daily Inventory Report — {NOMBRE_NEGOCIO}"
    msg.attach(MIMEText(html, "html"))

    # Adjuntar Excel
    with open("inventario.xlsx", "rb") as f:
        adjunto = MIMEBase("application", "octet-stream")
        adjunto.set_payload(f.read())
        encoders.encode_base64(adjunto)
        adjunto.add_header("Content-Disposition", "attachment", filename="inventario.xlsx")
        msg.attach(adjunto)

    try:
        servidor = smtplib.SMTP("smtp.gmail.com", 587)
        servidor.starttls()
        servidor.login(EMAIL_REMITENTE, CONTRASEÑA_APP)
        servidor.sendmail(EMAIL_REMITENTE, EMAIL_CLIENTE, msg.as_string())
        servidor.quit()
        print(f"✅ Reporte enviado correctamente a {EMAIL_CLIENTE}")
    except Exception as e:
        print(f"❌ Error al enviar: {e}")

# Programar y enviar prueba
print(f"🤖 Sistema iniciado — Enviando todos los días a las {HORA_ENVIO}")
schedule.every().day.at(HORA_ENVIO).do(enviar_email)
print("📧 Enviando reporte de prueba ahora...")
enviar_email()

while True:
    schedule.run_pending()
    time.sleep(60)