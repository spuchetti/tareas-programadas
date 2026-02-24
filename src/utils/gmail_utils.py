"""
Funciones para enviar emails con Gmail API
"""

import os
import smtplib
import ssl
import csv
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
from utils.common_utils import nombre_mes


def enviar_email_html_con_adjuntos(asunto, html, lista_adjuntos=None, env_destinatario="SMTP_TO"):
    """
    Envía email usando SMTP con App Password de Gmail
    Ahora acepta múltiples destinatarios separados por comas
    
    Args:
        asunto: Asunto del email
        html: Contenido HTML del email
        lista_adjuntos: Lista de rutas de archivos a adjuntar
        env_destinatario: Nombre de la variable de entorno para destinatarios
    """
    if lista_adjuntos is None:
        lista_adjuntos = []
    
    # Obtener configuraciones desde variables de entorno
    mail_to = os.getenv(env_destinatario)  # destinatario
    mail_from = os.getenv("SMTP_FROM")
    smtp_password = os.getenv("SMTP_PASSWORD")
    
    # Validar configuraciones
    if not mail_to:
        print(f"⚠️ {env_destinatario} no configurado. No se enviará email.")
        return
    
    if not mail_from:
        print(f"⚠️ SMTP_FROM no configurado. Usando {env_destinatario} como remitente.")
        mail_from = mail_to
    
    if not smtp_password:
        print("❌ SMTP_PASSWORD no configurado. No se enviará email.")
        print("💡 Verifica que hayas agregado SMTP_PASSWORD en GitHub Secrets")
        return
    
    # Procesa múltiples destinatarios (separados por comas)
    destinatarios = [dest.strip() for dest in mail_to.split(',')]
    
    print(f"📧 Configurando email:")
    print(f"   Variable usada: {env_destinatario}")
    print(f"   De: {mail_from}")
    print(f"   Para: {', '.join(destinatarios)}")
    print(f"   Adjuntos: {len(lista_adjuntos)} archivo(s)")

    try:
        # 1. Crear mensaje MIME
        msg = MIMEMultipart()
        msg["From"] = mail_from
        msg["To"] = mail_to  # Mantener formato original para el header
        msg["Subject"] = asunto

        # 2. Agregar cuerpo HTML
        msg.attach(MIMEText(html, "html", "utf-8"))

        # 3. Agregar archivos adjuntos
        for ruta_adjunto in lista_adjuntos:
            if os.path.exists(ruta_adjunto):
                nombre_archivo = os.path.basename(ruta_adjunto)
                print(f"   📎 Adjuntando: {nombre_archivo}")
                
                with open(ruta_adjunto, "rb") as archivo:
                    part = MIMEBase("application", "octet-stream")
                    part.set_payload(archivo.read())
                    encoders.encode_base64(part)
                    part.add_header(
                        "Content-Disposition",
                        f'attachment; filename="{nombre_archivo}"'
                    )
                    msg.attach(part)
            else:
                print(f"   ⚠ Archivo no encontrado: {ruta_adjunto}")

        # 4. Configurar contexto SSL
        contexto = ssl.create_default_context()

        # 5. Conectar y enviar vía SMTP
        print("🔗 Conectando a SMTP Gmail (smtp.gmail.com:465)...")
        
        with smtplib.SMTP_SSL("smtp.gmail.com", 465, context=contexto) as servidor:
            # Autenticar con App Password
            print("🔑 Autenticando con App Password...")
            servidor.login(mail_from, smtp_password)
            
            # Enviar email a todos los destinatarios
            print("📤 Enviando mensaje...")
            servidor.send_message(msg, from_addr=mail_from, to_addrs=destinatarios)
        
        print("✅ Email enviado exitosamente vía SMTP")
        print(f"   Asunto: {asunto}")
        for destinatario in destinatarios:
            print(f"   Destinatario: {destinatario}")

    except smtplib.SMTPAuthenticationError as auth_error:
        print("❌ ERROR DE AUTENTICACIÓN SMTP")
        print(f"   Código: {auth_error.smtp_code}")
        print(f"   Mensaje: {auth_error.smtp_error}")
        print("\n💡 SOLUCIÓN:")
        print("   1. Verifica que SMTP_PASSWORD sea el App Password de 16 caracteres")
        print("   2. NO uses tu contraseña normal de Gmail")
        print("   3. El App Password debe verse así: 'abcd efgh ijkl mnop'")
        print("   4. Asegúrate de que la verificación en 2 pasos esté ACTIVADA")
        
    except Exception as e:
        print(f"❌ Error enviando email SMTP: {type(e).__name__}: {e}")
        import traceback
        traceback.print_exc()


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


def generar_html_resumen_unificador(periodos, fecha, cantidades_por_periodo, anio_actual, sumatorias_por_periodo=None):
    """
    Genera HTML para el unificador mensual con el formato especificado.
    AHORA LA SECCIÓN DE REPARTICIONES PROCESADAS USA LA COLUMNA 25 (CÓDIGO).
    
    Args:
        periodos: Lista de períodos (ej: ["06", "1° sac"])
        fecha: Fecha de generación
        cantidades_por_periodo: Diccionario {periodo: cantidad}
        anio_actual: Año actual (int)
        sumatorias_por_periodo: Diccionario con sumatorias por período {periodo: {tipo_entidad: sumatorias}}
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
                    <td>Créditos Asistenciales</td>
                    <td style="text-align:right;">{formatear_dinero(sumatorias_periodo['creditos_asistenciales'])}</td>
                </tr>
                <tr>
                    <td>Fondo Voluntario</td>
                    <td style="text-align:right;">{formatear_dinero(sumatorias_periodo['fondo_voluntario'])}</td>
                </tr>
                <tr>
                    <td>Personal</td>
                    <td style="text-align:right;">{formatear_dinero(sumatorias_periodo['personal'])}</td>
                </tr>
                <tr>
                    <td>Adherente</td>
                    <td style="text-align:right;">{formatear_dinero(sumatorias_periodo['adherente'])}</td>
                </tr>
                <tr>
                    <td>Patronal</td>
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
        
        # Agregar desglose por tipo de entidad si hay datos
        if sumatorias_por_periodo and periodo in sumatorias_por_periodo:
            sumatorias_tipo_periodo = sumatorias_por_periodo[periodo]
            
            # Orden de los tipos que queremos mostrar
            tipos_ordenados = ['Municipios', 'Comunas', 'Entes Descentralizados', 'Cajas Municipales', 'Escuela']
            
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
                        es_ultimo = (tipo_entidad == 'Escuela')
                        
                        periodos_html += f'''
    <!-- DESGLOSE {periodo_mostrar} - {tipo_entidad.upper()} -->
    <div style="margin:10px 0 {"20px" if es_ultimo else "10px"} 0; padding:15px;">
        <h4 style="margin:0 0 10px 0; color:#333; font-size:14px;">{tipo_entidad}</h4>
        
        <table style="width:50%; border-collapse:collapse; font-size:13px;">
            <tr>
                <td>Créditos Asistenciales</td>
                <td style="text-align:right; font-weight:500;">{formatear_dinero(sumatorias['creditos_asistenciales'])}</td>
            </tr>
            <tr>
                <td>Fondo Voluntario</td>
                <td style="text-align:right; font-weight:500;">{formatear_dinero(sumatorias['fondo_voluntario'])}</td>
            </tr>
            <tr>
                <td>Personal</td>
                <td style="text-align:right; font-weight:500;">{formatear_dinero(sumatorias['personal'])}</td>
            </tr>
            <tr>
                <td>Adherente</td>
                <td style="text-align:right; font-weight:500;">{formatear_dinero(sumatorias['adherente'])}</td>
            </tr>
            <tr>
                <td>Patronal</td>
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