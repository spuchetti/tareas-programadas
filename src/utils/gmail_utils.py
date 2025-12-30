"""
Funciones para enviar emails con Gmail API
"""

import os
import json
import base64
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
from google.oauth2.credentials import Credentials as GmailCreds
from googleapiclient.discovery import build as gmail_build
from datetime import datetime
from zoneinfo import ZoneInfo
from utils.common_utils import nombre_mes, obtener_mes_anterior


def enviar_email_html_con_adjuntos(asunto, html, lista_adjuntos=None):
    """
    Envía un email HTML con adjuntos usando Gmail API
    """
    if lista_adjuntos is None:
        lista_adjuntos = []
    
    token_json = os.getenv("GMAIL_TOKEN")
    mail_to = os.getenv("SMTP_TO")

    if not token_json or not mail_to:
        print("⚠ NO se enviará email: faltan credenciales.")
        return

    try:
        creds = GmailCreds.from_authorized_user_info(json.loads(token_json))
        servicio = gmail_build("gmail", "v1", credentials=creds)

        msg = MIMEMultipart()
        msg["To"] = mail_to
        msg["From"] = "me"
        msg["Subject"] = asunto

        msg.attach(MIMEText(html, "html", "utf-8"))

        for ruta in lista_adjuntos:
            with open(ruta, "rb") as f:
                part = MIMEBase("application", "octet-stream")
                part.set_payload(f.read())
                encoders.encode_base64(part)
                part.add_header("Content-Disposition", f'attachment; filename="{os.path.basename(ruta)}"')
                msg.attach(part)

        encoded = base64.urlsafe_b64encode(msg.as_bytes()).decode("utf-8")
        servicio.users().messages().send(userId="me", body={"raw": encoded}).execute()

        print("📧 Email enviado correctamente.")

    except Exception as e:
        print(f"❌ Error enviando email: {e}")


def generar_html_resumen_fv(periodo, procesados, detectados, lista, fecha):
    """Genera HTML para el bot de fondo voluntario"""
    if lista:
        lista_html = "\n".join(f"<li>✔ {os.path.splitext(item)[0]}</li>" for item in lista)
    else:
        lista_html = "<li>No se encontraron archivos</li>"

    html = f"""<!DOCTYPE html>
<html lang="es">
<head>
  <meta charset="utf-8"/>
  <meta name="viewport" content="width=device-width,initial-scale=1"/>
  <title>Resultado</title>
</head>

<body style="font-family: Arial, Helvetica, sans-serif; color:#222; line-height:1.4; padding:18px;">

  <div style="
      background: linear-gradient(90deg,#0a7bdc,#16a085);
      padding: 18px;
      border-radius: 8px;
      color: white;
      margin-bottom: 18px;
    ">
    <h2 style="margin:0;">🟢🔵 OSER - FONDO VOLUNTARIO</h2>
    <div style="opacity:0.9; font-size:14px; margin-top:4px;">Reporte de control automático </div>
  </div>

  <p><strong>Periodo:</strong> {periodo}</p>
  <p>
    <strong>Archivos procesados:</strong> {procesados}<br/>
    <strong>Total detectados:</strong> {detectados}
  </p>

  <hr style="margin:18px 0;">

  <p><strong>Reparticiones:</strong></p>
  <ul>
    {lista_html}
  </ul>

  <hr style="margin:18px 0;">

  <p style="font-size:0.9em; color:#555;">
    Generado: {fecha}
  </p>

<div style="text-align:right; margin-top:25px;">
    <img src="https://raw.githubusercontent.com/spuchetti/tareas-programadas/main/assets/robot.jpg"
         width="140"
         style="opacity:0.55; display:inline-block;"/>
  </div>

</body>
</html>"""

    return html

def formatear_numero(numero):
    """
    Formatea un número con punto como separador de miles.
    
    Ejemplos:
    - 27034 → "27.034"
    - 1000 → "1.000"
    - 500 → "500"
    - 1234567 → "1.234.567"
    """
    try:
        # Convertir a entero si es posible
        if isinstance(numero, (int, float)):
            numero_int = int(numero)
        else:
            # Si es string, intentar convertir
            numero_int = int(float(numero))
        
        # Formatear con puntos como separadores de miles
        # Primero convertir a string y luego invertir para agregar puntos cada 3 dígitos
        numero_str = str(numero_int)
        
        # Manejar números negativos
        if numero_int < 0:
            signo = "-"
            numero_str = numero_str[1:]  # Quitar el signo negativo
        else:
            signo = ""
        
        # Agregar puntos cada 3 dígitos desde el final
        partes = []
        while len(numero_str) > 3:
            partes.append(numero_str[-3:])
            numero_str = numero_str[:-3]
        partes.append(numero_str)
        
        # Unir las partes con puntos y agregar el signo si es necesario
        resultado = signo + ".".join(reversed(partes))
        return resultado
        
    except (ValueError, TypeError):
        # Si no se puede formatear, devolver el número como string
        return str(numero)

def generar_html_resumen_unificador(periodo, fecha, sheets_creado=False, total_filas_combinadas=0):
    """Genera HTML simplificado para el unificador mensual"""
    
    # Calcular cantidad de agentes
    cantidad_agentes = total_filas_combinadas
    cantidad_agentes_formateada = formatear_numero(cantidad_agentes)
    
    html = f"""
<!DOCTYPE html>
<html>
<head>
    <meta charset="utf-8">
    <title>Resultado</title>
    <style>
        .header-box {{
            background: linear-gradient(90deg, #0a7bdc, #16a085);
            padding: 20px;
            border-radius: 8px;
            color: white;
            margin-bottom: 20px;
        }}
    </style>
</head>
<body style="font-family: Arial, Helvetica, sans-serif; color:#222; line-height:1.4; padding:18px;">

     <div style="
      background: linear-gradient(90deg,#0a7bdc,#16a085);
      padding: 18px;
      border-radius: 8px;
      color: white;
      margin-bottom: 18px;
    ">
    <h2 style="margin:0;">🟢🔵 OSER - UNIFICADO MENSUAL</h2>
    <div style="opacity:0.9; font-size:14px; margin-top:4px;">Reporte unificado automático</div>
  </div>
    
    
    <p><strong>Periodo:</strong> {periodo}</p>
    <p><strong>Agentes procesados:</strong> {cantidad_agentes_formateada:}</p>
    
    <hr style="margin: 18px 0;">
    
    <p style="font-size:0.9em; color:#555;">
    Generado: {fecha}
  </p>

<div style="text-align:right; margin-top:25px;">
    <img src="https://raw.githubusercontent.com/spuchetti/tareas-programadas/main/assets/robot.jpg"
         width="140"
         style="opacity:0.55; display:inline-block;"/>
  </div>
</body>
</html>
    """
    return html