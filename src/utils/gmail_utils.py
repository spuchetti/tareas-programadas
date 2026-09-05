"""
Funciones para enviar emails con Gmail API
"""

import os
import smtplib
import ssl
import csv
import re
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
from utils.common_utils import nombre_mes


# =============================================================================
# GENERADORES DE HTML - NUEVO ESTILO (monitoreo_bot)
# =============================================================================

def _nombre_archivo_sin_anio(nombre_archivo):
    """
    '015-Ministerio de Salud-2026.xlsx' -> '015 - Ministerio de Salud'
    Saca la extensión y el año final (si lo tiene), y separa el resto con ' - '.

    Además del caso simple ('...-2026.xlsx'), contempla nombres con sufijo
    de copia de Google Drive ('...-2026 (1).xlsx', '...-2026 (2).xlsx'),
    que antes hacían que la regex no matcheara (por no estar el año pegado
    al final del string) y el año se colaba en el subtítulo del mail.
    """
    sin_ext = re.sub(r"\.(xlsx|xlsm|xls)$", "", nombre_archivo.strip(), flags=re.IGNORECASE)
    sin_anio = re.sub(r"[\s\-]*(19|20)\d{2}(\s*\(\d+\))?\s*$", "", sin_ext).strip()
    return sin_anio.replace("-", " - ")


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
        <div style="font-size:13px;color:#5b8ad4;">{_nombre_archivo_sin_anio(nombre_archivo)}</div>
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
    """
    Genera HTML para el bot de fondo voluntario
    
    Args:
        periodo: Período procesado
        procesados: Total de archivos procesados
        reparticiones_detectadas: Cantidad de reparticiones con casos
        total_agentes: Total de agentes detectados
        lista: Lista de nombres de reparticiones
        fecha: Fecha de generación
    """
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

  <div style="
      background: linear-gradient(90deg,#0a7bdc,#16a085);
      padding: 18px;
      border-radius: 8px;
      color: white;
      margin-bottom: 18px;
    ">
    <h2 style="margin:0;">🟢🔵 OSER - FONDO VOLUNTARIO</h2>
    <div style="opacity:0.9; font-size:14px; margin-top:4px;">Reporte de control automático</div>
  </div>

  <p><strong>Período:</strong> {periodo}</p>
  <p>
    <strong>Archivos procesados:</strong> {procesados}<br/>
    <strong>Reparticiones detectadas:</strong> {reparticiones_detectadas}<br/>
    <strong>Total de agentes detectados:</strong> {total_agentes}
  </p>

  <hr style="margin: 20px 0; border: none; border-top: 2px solid #dee2e6;">

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


def formatear_dinero(monto):
    """
    Formatea un monto monetario con punto como separador de miles y 2 decimales.
    
    Ejemplos:
    - 1234.56 → "$ 1.234,56"
    - 1000 → "$ 1.000,00"
    - 1234567.89 → "$ 1.234.567,89"
    """
    try:
        # Convertir a float
        if isinstance(monto, str):
            # Si es string, limpiar y convertir
            monto_str = monto.replace(',', '.')
            monto_float = float(monto_str)
        else:
            monto_float = float(monto)
        
        # Separar parte entera y decimal
        parte_entera = int(monto_float)
        parte_decimal = round((monto_float - parte_entera) * 100)
        
        # Formatear parte entera con puntos
        parte_entera_str = str(abs(parte_entera))
        partes = []
        while len(parte_entera_str) > 3:
            partes.append(parte_entera_str[-3:])
            parte_entera_str = parte_entera_str[:-3]
        partes.append(parte_entera_str)
        
        parte_entera_formateada = ".".join(reversed(partes))
        
        # Agregar signo negativo si corresponde
        if monto_float < 0:
            parte_entera_formateada = f"-{parte_entera_formateada}"
        
        # Formatear parte decimal con 2 dígitos
        parte_decimal_formateada = f"{parte_decimal:02d}"
        
        # Combinar
        return f"$ {parte_entera_formateada},{parte_decimal_formateada}"
        
    except (ValueError, TypeError):
        # Si no se puede formatear, devolver el monto original
        return f"$ {monto}"


def calcular_sumatorias_csv(ruta_csv):
    """
    Calcula las sumatorias de las columnas relevantes de un archivo CSV.
    
    CORRECCIONES:
    - Personal: I (9-aporte personal) + Q (17-reajs aporte pers) ✓
    - Adherente: J (10-adherente sec) + L (12-hijo menor de 35) + M (13-menor a cargo) + 
                R (18-reaj adherente sec) + T (20-reajuste hijo menor) + U (21-reajuste menor a cargo)
    - Fondo Voluntario: K (11-fondo v) + S (19-reajuste fv) ✓
    - Créditos Asistenciales: N (14-cred asist) + V (22-reajuste cred asistencial) ✓
    - Patronal: W (23-aporte patronal) + X (24-reajuste aporte patronal) ✓
    """
    try:
        import csv
        
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
            
            # Saltar encabezado
            next(reader, None)
            
            fila_numero = 1
            for fila in reader:
                fila_numero += 1
                
                # Verificar que la fila tenga al menos 24 columnas para cálculos
                if len(fila) < 24:
                    print(f"⚠ Fila {fila_numero} tiene solo {len(fila)} columnas")
                    continue
                
                try:
                    # Función para convertir valores
                    def safe_float(val):
                        if not val or val == '' or str(val).strip() == '':
                            return 0.0
                        try:
                            return float(str(val).strip())
                        except ValueError:
                            return 0.0
                    
                    # OBTENER VALORES CON CORRECCIONES
                    # Columna I (9-aporte personal)
                    aporte_personal = safe_float(fila[8])      
                    
                    # Columna J (10-adherente sec)
                    adherente_sec = safe_float(fila[9])        
                    
                    # Columna K (11-fondo v)
                    fondo_v = safe_float(fila[10])            
                    
                    # Columna L (12-hijo menor de 35)
                    hijo_menor_35 = safe_float(fila[11])
                    
                    # Columna M (13-menor a cargo)
                    menor_cargo = safe_float(fila[12])
                    
                    # Columna N (14-cred asist)
                    cred_asist = safe_float(fila[13])         
                    
                    # Columna Q (17-reajs aporte pers)
                    reaj_aporte_pers = safe_float(fila[16])   
                    
                    # Columna R (18-reaj adherente sec)
                    reaj_adherente_sec = safe_float(fila[17]) 
                    
                    # Columna S (19-reajuste fv)
                    reajuste_fv = safe_float(fila[18])        
                    
                    # Columna T (20-reajuste hijo menor)
                    reajuste_hijo_menor = safe_float(fila[19])
                    
                    # Columna U (21-reajuste menor a cargo)
                    reajuste_menor_cargo = safe_float(fila[20])
                    
                    # Columna V (22-reajuste cred asistencial)
                    reaj_cred_asist = safe_float(fila[21])    
                    
                    # Columna W (23-aporte patronal)
                    aporte_patronal = safe_float(fila[22])    
                    
                    # Columna X (24-reajuste aporte patronal)
                    reaj_aporte_patronal = safe_float(fila[23]) 
                    
                    # CALCULAR SUMAS POR CONCEPTO - FÓRMULAS CORREGIDAS
                    # Personal = columna I + columna Q
                    sum_personal_fila = aporte_personal + reaj_aporte_pers
                    sumatorias['personal'] += sum_personal_fila
                    
                    # Adherente = columna J + columna L + columna M + columna R + columna T + columna U
                    sum_adherente_fila = (
                        adherente_sec + 
                        hijo_menor_35 + 
                        menor_cargo + 
                        reaj_adherente_sec + 
                        reajuste_hijo_menor + 
                        reajuste_menor_cargo
                    )
                    sumatorias['adherente'] += sum_adherente_fila
                    
                    # Fondo Voluntario = columna K + columna S
                    sum_fv_fila = fondo_v + reajuste_fv
                    sumatorias['fondo_voluntario'] += sum_fv_fila
                    
                    # Créditos Asistenciales = columna N + columna V
                    sum_creditos_fila = cred_asist + reaj_cred_asist
                    sumatorias['creditos_asistenciales'] += sum_creditos_fila
                    
                    # Patronal = columna W + columna X
                    sum_patronal_fila = aporte_patronal + reaj_aporte_patronal
                    sumatorias['patronal'] += sum_patronal_fila
                    
                except (ValueError, IndexError) as e:
                    print(f"⚠ Error en fila {fila_numero}: {e}")
                    print(f"  Fila: {fila}")
                    continue
        
        # Calcular total general
        sumatorias['total'] = (
            sumatorias['personal'] + 
            sumatorias['adherente'] + 
            sumatorias['fondo_voluntario'] + 
            sumatorias['creditos_asistenciales'] + 
            sumatorias['patronal']
        )
        
        return sumatorias
        
    except Exception as e:
        print(f"❌ Error calculando sumatorias de {ruta_csv}: {e}")
        import traceback
        traceback.print_exc()
        return {
            'creditos_asistenciales': 0.0,
            'fondo_voluntario': 0.0,
            'personal': 0.0,
            'adherente': 0.0,
            'patronal': 0.0,
            'total': 0.0
        }


def obtener_reparticiones_unicas_csv(ruta_csv):
    """
    Obtiene la lista de códigos únicos de la columna 25 (codigo) de un archivo CSV.
    AHORA USA LA COLUMNA 25 EN VEZ DE LA COLUMNA 8.
    
    Args:
        ruta_csv: Ruta al archivo CSV
        
    Returns:
        list: Lista de códigos únicos (ordenada alfabéticamente)
        int: Cantidad de códigos únicos
    """
    try:
        codigos = set()
        
        with open(ruta_csv, 'r', encoding='utf-8') as f:
            reader = csv.reader(f, delimiter='|')
            
            # Saltar encabezado
            next(reader, None)
            
            fila_numero = 0
            for fila in reader:
                fila_numero += 1
                # Verificar que la fila tenga al menos 25 columnas
                if len(fila) >= 25:
                    codigo = fila[24].strip()  # Columna 25 (índice 24)
                    if codigo and codigo != "" and codigo != "SIN_CODIGO":
                        codigos.add(codigo)
                    elif codigo and codigo != "":
                        codigos.add(codigo)  # Incluir SIN_CODIGO también
                else:
                    print(f"⚠ Fila {fila_numero} tiene solo {len(fila)} columnas, no se puede extraer código")
        
        # Ordenar alfabéticamente
        lista_codigos = sorted(list(codigos))
        
        print(f"📊 Códigos únicos encontrados en columna 25: {len(lista_codigos)}")
        if lista_codigos:
            print(f"   Ejemplos: {lista_codigos[:5]}")
        
        return lista_codigos, len(lista_codigos)
        
    except Exception as e:
        print(f"❌ Error obteniendo códigos únicos de {ruta_csv}: {e}")
        return [], 0


def generar_html_resumen_unificador(periodos, fecha, cantidades_por_periodo, anio_actual, sumatorias_por_periodo=None, aportantes_por_periodo=None, total_dnis_unicos_por_periodo=None):
    """
    Genera HTML para el unificador mensual con el formato especificado.
    AHORA LA SECCIÓN DE REPARTICIONES PROCESADAS USA LA COLUMNA 25 (CÓDIGO).
    
    Args:
        periodos: Lista de períodos (ej: ["06", "1° sac"])
        fecha: Fecha de generación
        cantidades_por_periodo: Diccionario {periodo: cantidad}
        anio_actual: Año actual (int)
        sumatorias_por_periodo: Diccionario con sumatorias por período {periodo: {tipo_entidad: sumatorias}}
        aportantes_por_periodo: Diccionario con datos de aportantes por período {periodo: [(codigo, nombre, cantidad), ...]}
    """
    
    periodos_html = ""
    
    # Primero, listar todos los archivos CSV disponibles
    archivos_csv_disponibles = {}
    if os.path.exists("generados"):
        for csv_file in os.listdir("generados"):
            if csv_file.endswith(".csv"):
                nombre_sin_extension = os.path.splitext(csv_file)[0]
                archivos_csv_disponibles[nombre_sin_extension] = os.path.join("generados", csv_file)
    
    print(f"📁 Archivos CSV disponibles: {list(archivos_csv_disponibles.keys())}")
    
    for i, periodo in enumerate(periodos):
        # Obtener nombre del período para mostrar
        if periodo in ["1° sac", "2° sac", "1º sac", "2º sac"]:
            periodo_mostrar = periodo.upper().replace("º", "°")
            periodo_buscar = periodo.upper().replace("º", "°").replace(" ", "").replace("°", "")
        else:
            # Para meses normales
            periodo_mostrar = nombre_mes(periodo)
            periodo_buscar = nombre_mes(periodo)
        
        print(f"🔍 Buscando CSV para período: {periodo} (mostrar: {periodo_mostrar}, buscar: {periodo_buscar})")
        
        # Obtener cantidad de registros para este período
        cantidad_registros = cantidades_por_periodo.get(periodo, 0)
        cantidad_registros_formateada = formatear_numero(cantidad_registros)
        
        # Buscar el archivo CSV correspondiente a este período
        archivo_csv = None
        
        # Intentar diferentes patrones de búsqueda
        patrones_busqueda = [
            f"Unificado_{periodo_buscar}{anio_actual}",  # Unificado_Junio2025
        ]
        
        if periodo in ["1° sac", "1º sac"]:
            patrones_busqueda.extend([
                f"Unificado_1SAC{anio_actual}", # Unificado_1SAC2025
            ])
        elif periodo in ["2° sac", "2º sac"]:
            patrones_busqueda.extend([
                f"Unificado_2SAC{anio_actual}", # Unificado_2SAC2025
            ])
        
        for patron in patrones_busqueda:
            if patron in archivos_csv_disponibles:
                archivo_csv = archivos_csv_disponibles[patron]
                print(f"   ✅ Encontrado: {patron} -> {os.path.basename(archivo_csv)}")
                break
        
        if not archivo_csv:
            # Si no se encuentra con patrones exactos, buscar por coincidencia parcial
            for nombre_archivo, ruta_archivo in archivos_csv_disponibles.items():
                nombre_archivo_lower = nombre_archivo.lower()
                periodo_buscar_lower = periodo_buscar.lower()
                
                # Verificar si el nombre del archivo contiene el período
                if periodo_buscar_lower in nombre_archivo_lower:
                    archivo_csv = ruta_archivo
                    print(f"   ⚠ Encontrado por coincidencia parcial: {nombre_archivo}")
                    break
        
        # Inicializar variables
        codigos_unicos = []
        cantidad_codigos = 0
        
        # Calcular sumatorias y códigos únicos si se encontró el archivo
        sumatorias_periodo = {
            'creditos_asistenciales': 0.0,
            'fondo_voluntario': 0.0,
            'personal': 0.0,
            'adherente': 0.0,
            'patronal': 0.0,
            'total': 0.0
        }
        
        if archivo_csv and os.path.exists(archivo_csv):
            print(f"   📊 Calculando sumatorias de: {os.path.basename(archivo_csv)}")
            sumatorias_periodo = calcular_sumatorias_csv(archivo_csv)
            
            # Obtener códigos únicos de la columna 25
            codigos_unicos, cantidad_codigos = obtener_reparticiones_unicas_csv(archivo_csv)
            print(f"   🏢 Códigos únicos encontrados (columna 25): {cantidad_codigos}")
        else:
            print(f"   ⚠ No se encontró archivo CSV para período {periodo_mostrar}")
        
        # Mostrar sumatorias para debug
        print(f"   💰 Sumatorias para {periodo_mostrar}:")
        print(f"     - Créditos Asistenciales: {sumatorias_periodo['creditos_asistenciales']}")
        print(f"     - Fondo Voluntario: {sumatorias_periodo['fondo_voluntario']}")
        print(f"     - Personal: {sumatorias_periodo['personal']}")
        print(f"     - Adherente: {sumatorias_periodo['adherente']}")
        print(f"     - Patronal: {sumatorias_periodo['patronal']}")
        print(f"     - TOTAL: {sumatorias_periodo['total']}")
        
        # Generar HTML para este período (TOTALES DEL PERÍODO)
        periodos_html += f'''
    <!-- PERÍODO {periodo_mostrar} {anio_actual} -->
    <div style="margin-bottom:22px;">

        <div style="
            border:1px solid #dcdfe3;
            border-radius:6px;
            padding:10px 12px;
        ">

            <!-- REPARTICIONES PROCESADAS (USA CÓDIGOS DE COLUMNA 25) -->
            <div style="
                font-size:12px;
                color:#374151;
                margin-bottom:8px;
            ">
                <span style="font-weight:600;">Reparticiones procesadas:</span>
                <span style="
                    background:#e5e7eb;
                    color:#1f2937;
                    font-weight:600;
                    padding:2px 8px;
                    border-radius:12px;
                    margin-left:2px;
                ">{cantidad_codigos}</span>
            </div>

            <!-- SEPARADOR DE PUNTITOS -->
            <div style="border-top:1px dashed #d0d7de; margin:0 0 11px 0;"></div>

            <!-- PERÍODO -->
            <h4 style="
                margin:0;
                font-size:14px;
            ">
                <span style="font-weight:600;">Periodo:</span>
                <span style="color:#1F395E; font-weight:600;">{periodo_mostrar}/{anio_actual}</span>
            </h4>

            <!-- REGISTROS -->
            <div style="
                font-size:12px;
                color:#6b7280;
                margin:2px 0 6px 0;
            ">
                Registros procesados: <strong style="font-weight:500;">{cantidad_registros_formateada}</strong>
            </div>

            <!-- SEPARADOR -->
            <div style="border-top:1px solid #d0d7de; margin:6px 0 8px 0;"></div>

            <table style="width:100%; border-collapse:collapse; font-size:13px;">
                <tr>
                    <td style="padding:4px 0;">Créditos Asistenciales</td>
                    <td style="text-align:right;">{formatear_dinero(sumatorias_periodo['creditos_asistenciales'])}</td>
                </tr>
                <tr>
                    <td style="padding:4px 0;">Fondo Voluntario</td>
                    <td style="text-align:right;">{formatear_dinero(sumatorias_periodo['fondo_voluntario'])}</td>
                </tr>
                <tr>
                    <td style="padding:4px 0;">Personal</td>
                    <td style="text-align:right;">{formatear_dinero(sumatorias_periodo['personal'])}</td>
                </tr>
                <tr>
                    <td style="padding:4px 0;">Adherente</td>
                    <td style="text-align:right;">{formatear_dinero(sumatorias_periodo['adherente'])}</td>
                </tr>
                <tr>
                    <td style="padding:4px 0;">Patronal</td>
                    <td style="text-align:right;">{formatear_dinero(sumatorias_periodo['patronal'])}</td>
                </tr>

                <tr>
                    <td colspan="2">
                        <div style="border-top:2px solid #bfc7cf; margin:8px 0;"></div>
                    </td>
                </tr>

                <tr style="background:#eef7f4;">
                    <td style="font-weight:700;">TOTAL</td>
                    <td style="
                        text-align:right;
                        font-weight:700;
                        font-size:16px;
                        color:#0f766e;
                    ">
                        {formatear_dinero(sumatorias_periodo['total'])}
                    </td>
                </tr>
            </table>

        </div>
    </div>
        '''
        
        # Agregar desglose por tipo de entidad
        if sumatorias_por_periodo and periodo in sumatorias_por_periodo:
            sumatorias_tipo_periodo = sumatorias_por_periodo[periodo]
            
            # Orden de los tipos de reparticiones
            tipos_ordenados = ['Municipios', 'Comunas', 'Entes Descentralizados', 'Cajas Municipales', 'Escuela', 'Pasantias']
            
            # Verificar si hay algún tipo de entidad con datos
            tiene_datos = False
            for tipo_entidad in tipos_ordenados:
                if tipo_entidad in sumatorias_tipo_periodo and sumatorias_tipo_periodo[tipo_entidad]['total'] > 0:
                    tiene_datos = True
                    break
            
            # Solo mostrar si hay datos
            if tiene_datos:
                for tipo_entidad in tipos_ordenados:
                    if tipo_entidad in sumatorias_tipo_periodo and sumatorias_tipo_periodo[tipo_entidad]['total'] > 0:
                        sumatorias = sumatorias_tipo_periodo[tipo_entidad]
                        
                        # Determinar si es el último elemento
                        es_ultimo = (tipo_entidad == tipos_ordenados[-1])
                        
                        periodos_html += f'''
    <!-- DESGLOSE {periodo_mostrar} - {tipo_entidad.upper()} -->
    <div style="margin:10px 0 {"20px" if es_ultimo else "10px"} 0; padding:15px;">
        <h4 style="margin:0 0 10px 0; color:#333; font-size:14px;">{tipo_entidad}</h4>
        
        <table style="width:50%; border-collapse:collapse; font-size:13px;">
            <tr>
                <td style="padding:4px 0;">Créditos Asistenciales</td>
                <td style="text-align:right; font-weight:500;">{formatear_dinero(sumatorias['creditos_asistenciales'])}</td>
            </tr>
            <tr>
                <td style="padding:4px 0;">Fondo Voluntario</td>
                <td style="text-align:right; font-weight:500;">{formatear_dinero(sumatorias['fondo_voluntario'])}</td>
            </tr>
            <tr>
                <td style="padding:4px 0;">Personal</td>
                <td style="text-align:right; font-weight:500;">{formatear_dinero(sumatorias['personal'])}</td>
            </tr>
            <tr>
                <td style="padding:4px 0;">Adherente</td>
                <td style="text-align:right; font-weight:500;">{formatear_dinero(sumatorias['adherente'])}</td>
            </tr>
            <tr>
                <td style="padding:4px 0;">Patronal</td>
                <td style="text-align:right; font-weight:500;">{formatear_dinero(sumatorias['patronal'])}</td>
            </tr>
            <tr>
                <td colspan="2">
                    <div style="border-top:1px solid #ccc; margin:8px 0;"></div>
                </td>
            </tr>
            <tr style="background:#e6f7f3;">
                <td style="padding:6px 0; font-weight:700;">TOTAL</td>
                <td style="padding:6px 0; text-align:right; font-weight:700; color:#0f766e; font-size:14px;">
                    {formatear_dinero(sumatorias['total'])}
                </td>
            </tr>
        </table>
    </div>
                        '''

        # Tabla de aportantes
        if aportantes_por_periodo:
            TOP_N = 10
            datos = aportantes_por_periodo.get(periodo, [])
            if datos:
                datos_ordenados = sorted(datos, key=lambda x: x[2], reverse=True)

                if periodo in ["1° sac", "2° sac", "1º sac", "2º sac"]:
                    periodo_label = periodo.upper().replace("º", "°")
                else:
                    periodo_label = nombre_mes(periodo)

                total_aportantes = sum(cantidad for _, _, cantidad in datos_ordenados)
                # Si hay un conteo de DNIs únicos globales para este período, usarlo como total
                if total_dnis_unicos_por_periodo and periodo in total_dnis_unicos_por_periodo:
                    total_aportantes_mostrar = total_dnis_unicos_por_periodo[periodo]
                else:
                    total_aportantes_mostrar = total_aportantes
                total_reparticiones = len(datos_ordenados)
                top = datos_ordenados[:TOP_N]
                restantes = total_reparticiones - TOP_N

                filas_top = ""
                for codigo, reparticion, cantidad in top:
                    filas_top += f"""
                <tr style="border-bottom:1px solid #e5e7eb;">
                    <td style="padding:8px 16px; font-size:13px; text-align:center; color:#374151;">{codigo}</td>
                    <td style="padding:8px 16px; font-size:13px; text-align:center; color:#374151;">{reparticion}</td>
                    <td style="padding:8px 16px; font-size:13px; text-align:center; color:#374151;">{formatear_numero(cantidad)}</td>
                </tr>"""

                periodos_html += f"""
    <!-- TOP APORTANTES - {periodo_label} -->
    <div style="margin:22px 0 10px 0; border:1px solid #e5e7eb; border-radius:8px; overflow:hidden;">

        <!-- TÍTULO -->
        <div style="background:#e5e7eb; padding:12px 16px; text-align:center; border-bottom:1px solid #e5e7eb;">
            <span style="font-size:15px; font-weight:700; color:#111827;">Aportantes</span>
        </div>

        <!-- SUBTÍTULO -->
        <div style="padding:8px 12px; font-size:12px; color:#374151; border-bottom:1px dashed #d0d7de; text-align:center;">
            Total: <strong>{formatear_numero(total_aportantes)}</strong> aportantes
            &nbsp;·&nbsp; Mostrando Top {min(TOP_N, total_reparticiones)}
        </div>

        <!-- TABLA -->
        <table style="width:100%; border-collapse:collapse;">
            <thead>
                <tr style="background:#f3f4f6">
                    <th style="padding:6px 8px;text-align:center;font-size:12px;color:#374151;font-weight:600">Código</th>
                    <th style="padding:6px 8px;text-align:center;font-size:12px;color:#374151;font-weight:600">Repartición</th>
                    <th style="padding:6px 8px;text-align:center;font-size:12px;color:#374151;font-weight:600">Cantidad</th>
                </tr>
            </thead>
            <tbody>
                {filas_top}
            </tbody>
            <tfoot>
                <tr style="background:#f9fafb; border-top:1px solid #e5e7eb;">
                    <td colspan="3" style="padding:10px 16px; font-size:12px; color:#9ca3af; text-align:center;">
                        +{formatear_numero(restantes)} reparticiones más &nbsp;·&nbsp; Ver detalle completo en el CSV adjunto
                    </td>
                </tr>
            </tfoot>
        </table>

    </div>"""

    html = f"""<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="utf-8">
    <title>Resultado Unificador Mensual</title>
    <style>
        body {{
            font-family: Arial, Helvetica, sans-serif;
            color: #222;
            line-height: 1.35;
            padding: 18px;
        }}
    </style>
</head>

<body>

    <!-- CABECERA -->
    <div style="
        background: linear-gradient(90deg,#0a7bdc,#16a085);
        padding: 18px;
        border-radius: 8px;
        color: white;
        margin-bottom: 22px;
    ">
        <h2 style="margin:0;">🟢🔵 OSER - UNIFICADO MENSUAL</h2>
        <div style="opacity:0.9; font-size:14px; margin-top:4px;">
            Reporte unificado automático
        </div>
    </div>

{periodos_html}
    <hr style="margin:18px 0;">

    <p style="font-size:0.9em; color:#555;">
        Generado: {fecha}
    </p>

    <div style="text-align:right; margin-top:25px;">
        <img
            src="https://raw.githubusercontent.com/spuchetti/tareas-programadas/main/assets/robot.jpg"
            width="140"
            style="opacity:0.55;"
        />
    </div>

</body>
</html>"""
    
    return html


def generar_html_resumen_anual(anio, archivos_procesados, resumen_por_mes, fecha):
    """
    Genera HTML para el reporte anual de fondo voluntario.

    Args:
        anio: Año procesado (int)
        archivos_procesados: Total de archivos de Drive procesados (int)
        resumen_por_mes: dict { "01": {"nombre": "Enero", "registros": N, "archivos": M}, ... }
        fecha: Fecha de generación (str)
    """
    total_registros = sum(v["registros"] for v in resumen_por_mes.values())
    meses_con_casos = sum(1 for v in resumen_por_mes.values() if v["registros"] > 0)

    filas_meses = ""
    # NOTA: se recorre en el orden de inserción del dict (no sorted()), porque
    # reporte_anual_bot.py ya arma resumen_por_mes en el orden cronológico real
    # (con "1° sac" intercalado después de junio y "2° sac" después de
    # noviembre). sorted() ordena como texto y empuja ambos SAC al final de
    # la tabla, lo cual no refleja el orden real del año.
    for mes_num in resumen_por_mes.keys():
        r = resumen_por_mes[mes_num]
        registros = r["registros"]
        archivos  = r["archivos"]
        color_fila  = "#f0f9f4" if registros > 0 else "#ffffff"
        badge_color = "#16a085" if registros > 0 else "#9ca3af"
        filas_meses += f"""
                <tr style="border-bottom:1px solid #e5e7eb; background:{color_fila};">
                    <td style="padding:9px 14px; font-size:13px; font-weight:600; color:#374151;">{r['nombre']}</td>
                    <td style="padding:9px 14px; font-size:13px; text-align:center; color:#374151;">{archivos}</td>
                    <td style="padding:9px 14px; font-size:13px; text-align:center;">
                        <span style="
                            background:{badge_color};
                            color:white;
                            padding:2px 10px;
                            border-radius:12px;
                            font-size:12px;
                            font-weight:700;
                        ">{registros}</span>
                    </td>
                </tr>"""

    html = f"""<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="utf-8">
    <title>Reporte Anual Fondo Voluntario {anio}</title>
    <style>
        body {{
            font-family: Arial, Helvetica, sans-serif;
            color: #222;
            line-height: 1.4;
            padding: 18px;
        }}
    </style>
</head>

<body>

    <!-- CABECERA -->
    <div style="
        background: linear-gradient(90deg,#0a7bdc,#16a085);
        padding: 18px;
        border-radius: 8px;
        color: white;
        margin-bottom: 22px;
    ">
        <h2 style="margin:0;">🟢🔵 OSER - FONDO VOLUNTARIO</h2>
        <div style="opacity:0.9; font-size:14px; margin-top:4px;">Reporte anual {anio}</div>
    </div>

    <!-- RESUMEN GENERAL -->
    <table style="width:100%; border-collapse:collapse; margin-bottom:24px;">
        <tr>
            <td style="padding:10px 16px; background:#f3f4f6; border-radius:6px; text-align:center; width:33%;">
                <div style="font-size:11px; color:#6b7280; text-transform:uppercase; letter-spacing:.5px;">Archivos procesados</div>
                <div style="font-size:26px; font-weight:700; color:#1f4e79;">{archivos_procesados}</div>
            </td>
            <td style="width:2%;"></td>
            <td style="padding:10px 16px; background:#f3f4f6; border-radius:6px; text-align:center; width:33%;">
                <div style="font-size:11px; color:#6b7280; text-transform:uppercase; letter-spacing:.5px;">Meses con casos</div>
                <div style="font-size:26px; font-weight:700; color:#1f4e79;">{meses_con_casos} / 12</div>
            </td>
            <td style="width:2%;"></td>
            <td style="padding:10px 16px; background:#eef7f4; border-radius:6px; text-align:center; width:33%;">
                <div style="font-size:11px; color:#6b7280; text-transform:uppercase; letter-spacing:.5px;">Total agentes detectados</div>
                <div style="font-size:26px; font-weight:700; color:#16a085;">{total_registros}</div>
            </td>
        </tr>
    </table>

    <!-- TABLA POR MES -->
    <div style="border:1px solid #e5e7eb; border-radius:8px; overflow:hidden; margin-bottom:24px;">

        <div style="background:#1f4e79; padding:12px 16px;">
            <span style="font-size:15px; font-weight:700; color:white;">📆 Detalle por mes</span>
        </div>

        <table style="width:100%; border-collapse:collapse;">
            <thead>
                <tr style="background:#f3f4f6;">
                    <th style="padding:8px 14px; text-align:left;   font-size:12px; color:#374151; font-weight:600;">Mes</th>
                    <th style="padding:8px 14px; text-align:center; font-size:12px; color:#374151; font-weight:600;">Reparticiones con casos</th>
                    <th style="padding:8px 14px; text-align:center; font-size:12px; color:#374151; font-weight:600;">Registros detectados</th>
                </tr>
            </thead>
            <tbody>
                {filas_meses}
            </tbody>
            <tfoot>
                <tr style="background:#eef7f4; border-top:2px solid #16a085;">
                    <td style="padding:10px 14px; font-weight:700; font-size:13px;">TOTAL</td>
                    <td style="text-align:center;"></td>
                    <td style="padding:10px 14px; text-align:center; font-weight:700; font-size:15px; color:#16a085;">{total_registros}</td>
                </tr>
            </tfoot>
        </table>
    </div>

    <p style="font-size:13px; color:#555;">
        📎 El Excel adjunto contiene el detalle completo con una hoja por mes.
    </p>

    <hr style="margin:18px 0; border:none; border-top:1px solid #e5e7eb;">

    <p style="font-size:0.85em; color:#888;">Generado: {fecha}</p>

    <div style="text-align:right; margin-top:20px;">
        <img
            src="https://raw.githubusercontent.com/spuchetti/tareas-programadas/main/assets/robot.jpg"
            width="140"
            style="opacity:0.55;"
        />
    </div>

</body>
</html>"""

    return html
