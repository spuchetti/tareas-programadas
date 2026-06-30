"""
Funciones para enviar emails con Gmail API y generar HTML de resumen
para todos los bots del proyecto.
"""

import os
import smtplib
import ssl
import csv
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders

# =============================================================================
# GENERADORES DE HTML - NUEVO ESTILO (monitoreo_bot)
# =============================================================================

def generar_html_resumen_monitoreo(
    reparticion,
    nombre_archivo,
    total_cambios,
    total_eliminados,
    total_nuevos,
    total_modificados,
    cambios_por_periodo,
    adjuntos_info,
    fecha
):
    """
    Genera HTML con el mismo estilo que el Apps Script de monitoreo.
    
    Args:
        reparticion: Nombre de la repartición
        nombre_archivo: Nombre del archivo procesado
        total_cambios: Total de cambios
        total_eliminados: Total de eliminados
        total_nuevos: Total de nuevos
        total_modificados: Total de modificados
        cambios_por_periodo: Lista de dicts con cambios por período
            Cada dict: {periodo, eliminados, nuevos, modificados, complementarias}
        adjuntos_info: Lista de dicts con info de adjuntos
            Cada dict: {es_xlsx, nombre, descripcion}
        fecha: Fecha de generación
    """
    
    # ── FILAS DE PERÍODOS ──────────────────────────────────────────────────────
    filas_periodos = ""
    for p in cambios_por_periodo:
        badges = ""
        if p.get("eliminados", 0) > 0:
            badges += f'<span style="font-size:11px;font-weight:500;padding:3px 10px;border-radius:4px;white-space:nowrap;background:#fff1f2;color:#9f1239;border:0.5px solid #fecdd3;margin-left:6px;">{p["eliminados"]} eliminado{"s" if p["eliminados"] != 1 else ""}</span>'
        if p.get("nuevos", 0) > 0:
            badges += f'<span style="font-size:11px;font-weight:500;padding:3px 10px;border-radius:4px;white-space:nowrap;background:#f0fdf4;color:#14532d;border:0.5px solid #bbf7d0;margin-left:6px;">{p["nuevos"]} nuevo{"s" if p["nuevos"] != 1 else ""}</span>'
        if p.get("modificados", 0) > 0:
            badges += f'<span style="font-size:11px;font-weight:500;padding:3px 10px;border-radius:4px;white-space:nowrap;background:#eff6ff;color:#1e3a8a;border:0.5px solid #bfdbfe;margin-left:6px;">{p["modificados"]} modificado{"s" if p["modificados"] != 1 else ""}</span>'
        if p.get("complementarias", False):
            badges += f'<span style="font-size:11px;font-weight:500;padding:3px 10px;border-radius:4px;white-space:nowrap;background:#f5f3ff;color:#4c1d95;border:0.5px solid #ddd6fe;margin-left:6px;">con complementarias</span>'
        
        filas_periodos += f"""
        <tr>
            <td style="padding:12px 0;border-bottom:0.5px solid #e5e7eb;font-size:14px;font-weight:500;color:#111827;vertical-align:middle;">{p["periodo"]}</td>
            <td style="padding:12px 0;border-bottom:0.5px solid #e5e7eb;text-align:right;vertical-align:middle;">{badges}</td>
        </tr>"""
    
    # ── FILAS DE ADJUNTOS ──────────────────────────────────────────────────────
    filas_adjuntos = ""
    for adj in adjuntos_info:
        icono = "XLSX" if adj.get("es_xlsx", False) else "CSV"
        bg_icono = "#dcfce7" if adj.get("es_xlsx", False) else "#dbeafe"
        color_icono = "#14532d" if adj.get("es_xlsx", False) else "#1e3a8a"
        filas_adjuntos += f"""
        <tr>
            <td style="width:36px;padding:6px 10px 6px 0;vertical-align:middle;">
                <div style="width:32px;height:32px;border-radius:6px;background:{bg_icono};color:{color_icono};font-size:11px;font-weight:600;text-align:center;line-height:32px;">{icono}</div>
            </td>
            <td style="padding:6px 0;vertical-align:middle;">
                <div style="font-size:13px;font-weight:500;color:#111827;">{adj["nombre"]}</div>
                <div style="font-size:11px;color:#9ca3af;margin-top:2px;">{adj["descripcion"]}</div>
            </td>
        </tr>"""
    
    seccion_adjuntos = ""
    if adjuntos_info:
        seccion_adjuntos = f"""
        <div style="background:#f8fafc;border:0.5px solid #e5e7eb;border-radius:8px;padding:16px;margin-top:24px;">
            <div style="font-size:11px;font-weight:500;letter-spacing:.06em;text-transform:uppercase;color:#9ca3af;margin-bottom:12px;">Archivos adjuntos</div>
            <table style="width:100%;border-collapse:collapse;">{filas_adjuntos}</table>
        </div>"""
    
    # ── HTML COMPLETO ──────────────────────────────────────────────────────────
    html = f"""
<div style="font-family:'Segoe UI',Arial,sans-serif;font-size:14px;color:#111827;max-width:640px;">

    <!-- HEADER -->
    <div style="background:#0d2a5e;border-radius:12px 12px 0 0;padding:28px 32px 24px;border-bottom:3px solid #1a4fa0;">
        <div style="font-size:11px;font-weight:500;letter-spacing:.1em;color:#5b8ad4;text-transform:uppercase;margin-bottom:10px;">
            <span style="color:#3aaa35;">●</span> Monitoreo de liquidaciones
        </div>
        <div style="font-size:21px;font-weight:500;color:#f0f6ff;line-height:1.3;margin-bottom:4px;">
            Cambios detectados - {reparticion}
        </div>
        <div style="font-size:13px;color:#5b8ad4;">{nombre_archivo.replace('.xlsx', '').replace('-', ' - ')}</div>
    </div>

    <!-- STRIP MÉTRICAS -->
    <table style="width:100%;border-collapse:collapse;background:#102d6b;border-bottom:1px solid #163580;">
        <tr>
            <td style="padding:14px 20px;text-align:center;border-right:0.5px solid #1a3d7a;">
                <div style="font-size:10px;letter-spacing:.08em;color:#5b8ad4;text-transform:uppercase;margin-bottom:3px;">Total cambios</div>
                <div style="font-size:22px;font-weight:500;color:#f0f6ff;">{total_cambios}</div>
            </td>
            <td style="padding:14px 20px;text-align:center;border-right:0.5px solid #1a3d7a;">
                <div style="font-size:10px;letter-spacing:.08em;color:#5b8ad4;text-transform:uppercase;margin-bottom:3px;">Eliminados</div>
                <div style="font-size:22px;font-weight:500;color:#f87171;">{total_eliminados}</div>
            </td>
            <td style="padding:14px 20px;text-align:center;border-right:0.5px solid #1a3d7a;">
                <div style="font-size:10px;letter-spacing:.08em;color:#5b8ad4;text-transform:uppercase;margin-bottom:3px;">Nuevos</div>
                <div style="font-size:22px;font-weight:500;color:#4ade80;">{total_nuevos}</div>
            </td>
            <td style="padding:14px 20px;text-align:center;">
                <div style="font-size:10px;letter-spacing:.08em;color:#5b8ad4;text-transform:uppercase;margin-bottom:3px;">Modificados</div>
                <div style="font-size:22px;font-weight:500;color:#60a5fa;">{total_modificados}</div>
            </td>
        </tr>
    </table>

    <!-- BODY -->
    <div style="background:#ffffff;border:0.5px solid #e5e7eb;border-top:none;padding:28px 32px;border-radius:0 0 12px 12px;">

        <div style="font-size:11px;font-weight:500;letter-spacing:.08em;color:#9ca3af;text-transform:uppercase;margin-bottom:12px;">Detalle por periodo</div>

        <table style="width:100%;border-collapse:collapse;">
            <thead>
                <tr>
                    <th style="font-size:11px;letter-spacing:.06em;text-transform:uppercase;color:#9ca3af;padding:0 0 8px;text-align:left;border-bottom:0.5px solid #e5e7eb;font-weight:500;">Periodo</th>
                    <th style="font-size:11px;letter-spacing:.06em;text-transform:uppercase;color:#9ca3af;padding:0 0 8px;text-align:right;border-bottom:0.5px solid #e5e7eb;font-weight:500;">Registros</th>
                </tr>
            </thead>
            <tbody>{filas_periodos}</tbody>
        </table>

        {seccion_adjuntos}

        <!-- FOOTER -->
        <table style="width:100%;border-collapse:collapse;margin-top:24px;border-top:0.5px solid #e5e7eb;">
            <tr>
                <td style="padding-top:14px;font-size:11px;color:#9ca3af;">
                    Generado automáticamente por <strong>APORTES</strong> · Monitoreo de Liquidaciones
                </td>
            </tr>
        </table>

    </div>

</div>"""
    
    return html


# =============================================================================
# ENVÍO DE EMAILS
# =============================================================================

def enviar_email_html_con_adjuntos(asunto, html, lista_adjuntos=None, env_destinatario="SMTP_TO"):
    """
    Envía email usando SMTP con App Password de Gmail.
    """
    if lista_adjuntos is None:
        lista_adjuntos = []
    
    mail_to = os.getenv(env_destinatario)
    mail_from = os.getenv("SMTP_FROM")
    smtp_password = os.getenv("SMTP_PASSWORD")
    
    if not mail_to:
        print(f"⚠️ {env_destinatario} no configurado. No se enviará email.")
        return
    
    if not mail_from:
        mail_from = mail_to
    
    if not smtp_password:
        print("❌ SMTP_PASSWORD no configurado. No se enviará email.")
        return
    
    destinatarios = [dest.strip() for dest in mail_to.split(',')]
    
    print(f"📧 Configurando email:")
    print(f"   De: {mail_from}")
    print(f"   Para: {', '.join(destinatarios)}")
    print(f"   Adjuntos: {len(lista_adjuntos)} archivo(s)")

    try:
        msg = MIMEMultipart()
        msg["From"] = mail_from
        msg["To"] = mail_to
        msg["Subject"] = asunto
        msg.attach(MIMEText(html, "html", "utf-8"))

        for ruta_adjunto in lista_adjuntos:
            if os.path.exists(ruta_adjunto):
                nombre_archivo = os.path.basename(ruta_adjunto)
                print(f"   📎 Adjuntando: {nombre_archivo}")
                with open(ruta_adjunto, "rb") as archivo:
                    part = MIMEBase("application", "octet-stream")
                    part.set_payload(archivo.read())
                    encoders.encode_base64(part)
                    part.add_header("Content-Disposition", f'attachment; filename="{nombre_archivo}"')
                    msg.attach(part)
            else:
                print(f"   ⚠ Archivo no encontrado: {ruta_adjunto}")

        contexto = ssl.create_default_context()
        print("🔗 Conectando a SMTP Gmail...")
        with smtplib.SMTP_SSL("smtp.gmail.com", 465, context=contexto) as servidor:
            servidor.login(mail_from, smtp_password)
            servidor.send_message(msg, from_addr=mail_from, to_addrs=destinatarios)
        print("✅ Email enviado exitosamente")

    except Exception as e:
        print(f"❌ Error enviando email: {e}")
        import traceback
        traceback.print_exc()


# =============================================================================
# FUNCIONES LEGACY (mantenidas para compatibilidad con otros bots)
# =============================================================================

def generar_html_resumen_fv(periodo, procesados, reparticiones_detectadas, total_agentes, lista, fecha):
    """Genera HTML para el bot de fondo voluntario (legacy)."""
    if lista:
        lista_html = "\n".join(f"<li>{os.path.splitext(item)[0]}</li>" for item in lista)
    else:
        lista_html = "<li>No se encontraron reparticiones con casos</li>"

    html = f"""<!DOCTYPE html>
<html lang="es">
<head>
  <meta charset="utf-8"/>
  <meta name="viewport" content="width=device-width,initial-scale=1"/>
  <title>Resultado Fondo Voluntario</title>
</head>

<body style="font-family: Arial, Helvetica, sans-serif; color:#222; line-height:1.4; padding:18px;">

  <div style="background: linear-gradient(90deg,#0a7bdc,#16a085);padding:18px;border-radius:8px;color:white;margin-bottom:18px;">
    <h2 style="margin:0;">🟢🔵 OSER - FONDO VOLUNTARIO</h2>
    <div style="opacity:0.9; font-size:14px; margin-top:4px;">Reporte de control automático</div>
  </div>

  <p><strong>Período:</strong> {periodo}</p>
  <p>
    <strong>Archivos procesados:</strong> {procesados}<br/>
    <strong>Reparticiones detectadas:</strong> {reparticiones_detectadas}<br/>
    <strong>Total de agentes detectados:</strong> {total_agentes}
  </p>

  <hr style="margin:20px 0;border:none;border-top:2px solid #dee2e6;">

  <p><strong>Reparticiones:</strong></p>
  <ul>{lista_html}</ul>

  <hr style="margin:18px 0;">
  <p style="font-size:0.9em;color:#555;">Generado: {fecha}</p>

  <div style="text-align:right;margin-top:25px;">
    <img src="https://raw.githubusercontent.com/spuchetti/tareas-programadas/main/assets/robot.jpg" width="140" style="opacity:0.55;display:inline-block;"/>
  </div>

</body>
</html>"""
    return html


def generar_html_resumen_unificador(periodos, fecha, cantidades_por_periodo, anio_actual,
                                   sumatorias_por_periodo=None, aportantes_por_periodo=None,
                                   total_dnis_unicos_por_periodo=None):
    """Genera HTML para el unificador mensual (legacy)."""
    periodos_html = ""
    for periodo in periodos:
        if periodo in ["1° sac", "2° sac", "1º sac", "2º sac"]:
            periodo_mostrar = periodo.upper().replace("º", "°")
        else:
            from utils.common_utils import nombre_mes
            periodo_mostrar = nombre_mes(periodo)
        cantidad = cantidades_por_periodo.get(periodo, 0)
        periodos_html += f"""
        <div style="margin-bottom:16px;border:1px solid #dcdfe3;border-radius:6px;padding:12px;">
            <h4 style="margin:0 0 8px 0;">{periodo_mostrar}/{anio_actual}</h4>
            <div>Registros: {cantidad}</div>
        </div>"""

    html = f"""<!DOCTYPE html>
<html lang="es">
<head><meta charset="utf-8"><title>Resultado Unificador Mensual</title></head>
<body style="font-family:Arial,sans-serif;padding:18px;">
    <div style="background:linear-gradient(90deg,#0a7bdc,#16a085);padding:18px;border-radius:8px;color:white;margin-bottom:22px;">
        <h2 style="margin:0;">🟢🔵 OSER - UNIFICADO MENSUAL</h2>
        <div style="opacity:0.9;font-size:14px;margin-top:4px;">Reporte unificado automático</div>
    </div>
    {periodos_html}
    <hr style="margin:18px 0;">
    <p style="font-size:0.9em;color:#555;">Generado: {fecha}</p>
    <div style="text-align:right;margin-top:25px;">
        <img src="https://raw.githubusercontent.com/spuchetti/tareas-programadas/main/assets/robot.jpg" width="140" style="opacity:0.55;"/>
    </div>
</body>
</html>"""
    return html


def generar_html_resumen_anual(anio, archivos_procesados, resumen_por_mes, fecha):
    """Genera HTML para el reporte anual (legacy)."""
    total_registros = sum(v["registros"] for v in resumen_por_mes.values())
    meses_con_casos = sum(1 for v in resumen_por_mes.values() if v["registros"] > 0)

    filas_meses = ""
    for mes_num in sorted(resumen_por_mes.keys()):
        r = resumen_por_mes[mes_num]
        color_fila = "#f0f9f4" if r["registros"] > 0 else "#ffffff"
        badge_color = "#16a085" if r["registros"] > 0 else "#9ca3af"
        filas_meses += f"""
        <tr style="border-bottom:1px solid #e5e7eb;background:{color_fila};">
            <td style="padding:9px 14px;font-size:13px;font-weight:600;color:#374151;">{r['nombre']}</td>
            <td style="padding:9px 14px;font-size:13px;text-align:center;color:#374151;">{r['archivos']}</td>
            <td style="padding:9px 14px;font-size:13px;text-align:center;">
                <span style="background:{badge_color};color:white;padding:2px 10px;border-radius:12px;font-size:12px;font-weight:700;">{r['registros']}</span>
            </td>
        </tr>"""

    html = f"""<!DOCTYPE html>
<html lang="es">
<head><meta charset="utf-8"><title>Reporte Anual Fondo Voluntario {anio}</title></head>
<body style="font-family:Arial,sans-serif;padding:18px;">
    <div style="background:linear-gradient(90deg,#0a7bdc,#16a085);padding:18px;border-radius:8px;color:white;margin-bottom:22px;">
        <h2 style="margin:0;">🟢🔵 OSER - FONDO VOLUNTARIO</h2>
        <div style="opacity:0.9;font-size:14px;margin-top:4px;">Reporte anual {anio}</div>
    </div>
    <table style="width:100%;border-collapse:collapse;margin-bottom:24px;">
        <tr>
            <td style="padding:10px 16px;background:#f3f4f6;border-radius:6px;text-align:center;width:33%;">
                <div style="font-size:11px;color:#6b7280;text-transform:uppercase;letter-spacing:.5px;">Archivos procesados</div>
                <div style="font-size:26px;font-weight:700;color:#1f4e79;">{archivos_procesados}</div>
            </td>
            <td style="width:2%;"></td>
            <td style="padding:10px 16px;background:#f3f4f6;border-radius:6px;text-align:center;width:33%;">
                <div style="font-size:11px;color:#6b7280;text-transform:uppercase;letter-spacing:.5px;">Meses con casos</div>
                <div style="font-size:26px;font-weight:700;color:#1f4e79;">{meses_con_casos} / 12</div>
            </td>
            <td style="width:2%;"></td>
            <td style="padding:10px 16px;background:#eef7f4;border-radius:6px;text-align:center;width:33%;">
                <div style="font-size:11px;color:#6b7280;text-transform:uppercase;letter-spacing:.5px;">Total agentes detectados</div>
                <div style="font-size:26px;font-weight:700;color:#16a085;">{total_registros}</div>
            </td>
        </tr>
    </table>
    <div style="border:1px solid #e5e7eb;border-radius:8px;overflow:hidden;margin-bottom:24px;">
        <div style="background:#1f4e79;padding:12px 16px;">
            <span style="font-size:15px;font-weight:700;color:white;">📆 Detalle por mes</span>
        </div>
        <table style="width:100%;border-collapse:collapse;">
            <thead>
                <tr style="background:#f3f4f6;">
                    <th style="padding:8px 14px;text-align:left;font-size:12px;color:#374151;font-weight:600;">Mes</th>
                    <th style="padding:8px 14px;text-align:center;font-size:12px;color:#374151;font-weight:600;">Reparticiones con casos</th>
                    <th style="padding:8px 14px;text-align:center;font-size:12px;color:#374151;font-weight:600;">Registros detectados</th>
                </tr>
            </thead>
            <tbody>{filas_meses}</tbody>
            <tfoot>
                <tr style="background:#eef7f4;border-top:2px solid #16a085;">
                    <td style="padding:10px 14px;font-weight:700;font-size:13px;">TOTAL</td>
                    <td style="text-align:center;"></td>
                    <td style="padding:10px 14px;text-align:center;font-weight:700;font-size:15px;color:#16a085;">{total_registros}</td>
                </tr>
            </tfoot>
        </table>
    </div>
    <hr style="margin:18px 0;border:none;border-top:1px solid #e5e7eb;">
    <p style="font-size:0.85em;color:#888;">Generado: {fecha}</p>
    <div style="text-align:right;margin-top:20px;">
        <img src="https://raw.githubusercontent.com/spuchetti/tareas-programadas/main/assets/robot.jpg" width="140" style="opacity:0.55;"/>
    </div>
</body>
</html>"""
    return html


# =============================================================================
# UTILIDADES ADICIONALES (legacy)
# =============================================================================

def formatear_numero(numero):
    """Formatea un número con punto como separador de miles."""
    try:
        numero_int = int(numero)
        numero_str = str(abs(numero_int))
        partes = []
        while len(numero_str) > 3:
            partes.append(numero_str[-3:])
            numero_str = numero_str[:-3]
        partes.append(numero_str)
        return ("-" if numero_int < 0 else "") + ".".join(reversed(partes))
    except (ValueError, TypeError):
        return str(numero)


def formatear_dinero(monto):
    """Formatea un monto monetario con punto como separador de miles y 2 decimales."""
    try:
        monto_float = float(monto)
        parte_entera = int(abs(monto_float))
        parte_decimal = round((abs(monto_float) - parte_entera) * 100)
        parte_entera_str = str(parte_entera)
        partes = []
        while len(parte_entera_str) > 3:
            partes.append(parte_entera_str[-3:])
            parte_entera_str = parte_entera_str[:-3]
        partes.append(parte_entera_str)
        resultado = ("-" if monto_float < 0 else "") + ".".join(reversed(partes)) + f",{parte_decimal:02d}"
        return f"$ {resultado}"
    except (ValueError, TypeError):
        return f"$ {monto}"


def calcular_sumatorias_csv(ruta_csv):
    """Calcula las sumatorias de las columnas relevantes de un archivo CSV."""
    try:
        sumatorias = {
            'creditos_asistenciales': 0.0,
            'fondo_voluntario': 0.0,
            'personal': 0.0,
            'adherente': 0.0,
            'patronal': 0.0,
            'total': 0.0
        }
        with open(ruta_csv, 'r', encoding='utf-8') as f:
            reader = csv.reader(f, delimiter='|')
            next(reader, None)
            for fila in reader:
                if len(fila) < 24:
                    continue
                try:
                    def safe_float(val):
                        if not val or val == '':
                            return 0.0
                        try:
                            return float(str(val).strip())
                        except ValueError:
                            return 0.0
                    sumatorias['personal'] += safe_float(fila[8]) + safe_float(fila[16])
                    sumatorias['adherente'] += (
                        safe_float(fila[9]) + safe_float(fila[11]) + safe_float(fila[12]) +
                        safe_float(fila[17]) + safe_float(fila[19]) + safe_float(fila[20])
                    )
                    sumatorias['fondo_voluntario'] += safe_float(fila[10]) + safe_float(fila[18])
                    sumatorias['creditos_asistenciales'] += safe_float(fila[13]) + safe_float(fila[21])
                    sumatorias['patronal'] += safe_float(fila[22]) + safe_float(fila[23])
                except (ValueError, IndexError):
                    continue
        sumatorias['total'] = sum(sumatorias.values())
        return sumatorias
    except Exception as e:
        print(f"❌ Error calculando sumatorias: {e}")
        return {k: 0.0 for k in ['creditos_asistenciales', 'fondo_voluntario', 'personal', 'adherente', 'patronal', 'total']}


def obtener_reparticiones_unicas_csv(ruta_csv):
    """Obtiene la lista de códigos únicos de la columna 25 de un archivo CSV."""
    try:
        codigos = set()
        with open(ruta_csv, 'r', encoding='utf-8') as f:
            reader = csv.reader(f, delimiter='|')
            next(reader, None)
            for fila in reader:
                if len(fila) >= 25:
                    codigo = fila[24].strip()
                    if codigo and codigo != "":
                        codigos.add(codigo)
        return sorted(list(codigos)), len(codigos)
    except Exception as e:
        print(f"❌ Error obteniendo códigos únicos: {e}")
        return [], 0