"""
Funciones comunes para ambos bots
"""

import time
import os
import socket
import ssl
from http.client import IncompleteRead
from datetime import datetime
from zoneinfo import ZoneInfo

from googleapiclient.errors import HttpError


def obtener_zona_horaria():
    """Retorna la zona horaria de Argentina"""
    return ZoneInfo("America/Argentina/Buenos_Aires")


def nombre_mes(numero):
    """
    Convierte número de mes a nombre.

    NOTA: se normaliza 'º' (ordinal, U+00BA) a '°' (grado, U+00B0) antes de
    buscar en el diccionario. Los distintos bots generan el token de SAC
    con signos distintos (ej. unificador_mensual_bot.py arma "1º sac" con
    ordinal), así que sin esta normalización una de las dos variantes no
    encuentra la clave y devuelve "???" — pasaba concretamente con "1º sac"
    generando nombres de archivo tipo "Unificado_???2025.csv".
    """
    meses = {
        "01": "Enero", "02": "Febrero", "03": "Marzo", "04": "Abril",
        "05": "Mayo", "06": "Junio", "1° sac": "1SAC", "07": "Julio",
        "08": "Agosto", "09": "Septiembre", "10": "Octubre", "11": "Noviembre",
        "12": "Diciembre", "2° sac": "2SAC"
    }
    clave = str(numero).replace("º", "°")
    return meses.get(clave, "???")


def obtener_mes_anterior():
    """Obtiene el número del mes anterior (formato MM)"""
    ahora = datetime.now(obtener_zona_horaria())
    mes_anterior = ahora.month - 1 or 12
    return f"{mes_anterior:02d}"

def obtener_anio(mes_a_procesar):
    """Obtiene el año actual, validando si el mes a procesar es Diciembre incluido(en ese caso tiene que tomar el año anterior al actual)"""
    ahora = datetime.now(obtener_zona_horaria())
    anio_actual = ahora.year

    mes_normalizado = str(mes_a_procesar).replace("º", "°")
    if mes_normalizado == "12" or mes_normalizado == "2° sac":
        return anio_actual - 1

    return anio_actual

def registrar_inicio(nombre_proceso):
    """Registra el inicio del proceso"""
    ahora = datetime.now(obtener_zona_horaria())
    print("=" * 60)
    print(f"🚀 INICIO - {nombre_proceso}")
    print(f"📅 Fecha y hora: {ahora.strftime('%Y-%m-%d %H:%M:%S')}")
    print("=" * 60 + "\n")
    return ahora


def registrar_resumen(inicio, archivos_procesados=0, archivos_encontrados=0, 
                     filas_procesadas=0, errores=None):
    """Muestra resumen del proceso"""
    duracion = time.time() - inicio
    
    print("\n" + "=" * 60)
    print("📊 RESUMEN FINAL")
    print("=" * 60)
    
    if archivos_procesados > 0:
        print(f"📁 Archivos encontrados: {archivos_encontrados}")
        print(f"✅ Archivos procesados: {archivos_procesados}")
    
    if filas_procesadas > 0:
        print(f"📊 Filas procesadas: {filas_procesadas}")
    
    print(f"⏱ Tiempo total: {duracion:.2f} segundos")
    
    if errores:
        print(f"❌ Errores: {len(errores)}")
        for error in errores[:5]:
            print(f"  ⚠ {error}")
    
    print("=" * 60)


def crear_directorio_salida():
    """Crea directorio para archivos generados"""
    os.makedirs("generados", exist_ok=True)
    return "generados"


def ejecutar_con_reintentos_sheets(request, descripcion, intentos_max=4, espera_base=20):
    """
    Ejecuta un request de la Sheets API reintentando ante:
      - Límite de cuota (429 / RATE_LIMIT_EXCEEDED — "Write requests per
        minute per user", 60/min). La cuota es "por minuto", así que la
        espera empieza en ~20s y se va duplicando.
      - HttpError transitorio de servidor (500/503 — "The service is
        currently unavailable", visto en logs reales en
        obtener_titulos_hojas_snapshot y leer_hojas_snapshot_batch).
        Antes esto NO se reintentaba: solo se detectaba cuota buscando
        los strings "429"/"RATE_LIMIT_EXCEEDED"/"Quota exceeded" dentro
        de str(e), y un 503 no matchea ninguno de esos ni es una excepción
        de red/SSL de bajo nivel, así que se relanzaba directo sin
        reintentar ni una vez. Mismo criterio que ya usa
        request_drive_con_reintentos en drive_utils.py para 403/500/503.
      - Errores transitorios de red/SSL (conexión cortada, timeout, etc.)
        que pueden pasar en cualquier llamada HTTP, incluido el refresh de
        credenciales — antes de esto se relanzaban sin reintentar ninguna
        vez, mismo problema que tenía request_drive_con_reintentos en
        drive_utils.py hasta que apareció en un log real. Acá la espera es
        fija (espera_base), no exponencial, porque no es un problema de
        cuota que empeora con la insistencia.

    Vive acá (no en registro_utils.py, donde nació) para que también lo
    use monitoreo_utils.py sin generar un import circular entre ambos:
    common_utils.py no depende de ningún otro módulo del proyecto.

    Cualquier escritura a la Sheets API (values().update, values().clear,
    spreadsheets().batchUpdate, etc.) debería pasar por acá en vez de
    llamar .execute() directo — un archivo con muchas hojas puede disparar
    varias escrituras seguidas y superar la cuota por minuto fácilmente.
    """
    excepciones_red = (ssl.SSLError, ConnectionError, TimeoutError, socket.timeout, IncompleteRead, OSError)
    for intento in range(intentos_max):
        try:
            return request.execute()
        except excepciones_red as e:
            if intento < intentos_max - 1:
                print(f"  ⏳ Error de red/SSL, reintentando '{descripcion}' en {espera_base}s "
                      f"({intento + 1}/{intentos_max}): {e}")
                time.sleep(espera_base)
                continue
            raise
        except HttpError as e:
            status = e.resp.status if getattr(e, "resp", None) is not None else None
            es_quota = status == 429 or "RATE_LIMIT_EXCEEDED" in str(e) or "Quota exceeded" in str(e)
            es_transitorio_servidor = status in (500, 503)
            if intento < intentos_max - 1 and (es_quota or es_transitorio_servidor):
                if es_quota:
                    espera = espera_base * (2 ** intento)
                    print(f"  ⏳ Límite de cuota de Sheets API, reintentando '{descripcion}' en {espera}s "
                          f"({intento + 1}/{intentos_max})...")
                else:
                    espera = espera_base
                    print(f"  ⏳ Error {status} de Sheets API (servicio no disponible), reintentando "
                          f"'{descripcion}' en {espera}s ({intento + 1}/{intentos_max})...")
                time.sleep(espera)
                continue
            raise
        except Exception as e:
            es_quota = "429" in str(e) or "RATE_LIMIT_EXCEEDED" in str(e) or "Quota exceeded" in str(e)
            if es_quota and intento < intentos_max - 1:
                espera = espera_base * (2 ** intento)
                print(f"  ⏳ Límite de cuota de Sheets API, reintentando '{descripcion}' en {espera}s "
                      f"({intento + 1}/{intentos_max})...")
                time.sleep(espera)
                continue
            raise