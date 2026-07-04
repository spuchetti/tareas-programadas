"""
Funciones comunes para ambos bots
"""

import time
import os
from datetime import datetime
from zoneinfo import ZoneInfo


def obtener_zona_horaria():
    """Retorna la zona horaria de Argentina"""
    return ZoneInfo("America/Argentina/Buenos_Aires")


def _normalizar_periodo(s):
    """
    Normaliza variantes de escritura de los períodos especiales (SAC).
    Unifica todo al símbolo de grado '°' (U+00B0), que es el que se usa
    en el resto del proyecto (monitoreo_utils.py, reporte_anual_bot.py).
    Antes había una mezcla de '°' (grado) y 'º' (ordinal masculino) que
    hacía que nombre_mes()/obtener_anio() no reconocieran "2º sac".
    """
    return str(s).strip().replace("º", "°")


def nombre_mes(numero):
    """Convierte número de mes (o período SAC) a nombre"""
    meses = {
        "01": "Enero", "02": "Febrero", "03": "Marzo", "04": "Abril",
        "05": "Mayo", "06": "Junio", "1° sac": "1SAC", "07": "Julio",
        "08": "Agosto", "09": "Septiembre", "10": "Octubre", "11": "Noviembre",
        "12": "Diciembre", "2° sac": "2SAC"
    }
    return meses.get(_normalizar_periodo(numero), "???")


def obtener_mes_anterior():
    """Obtiene el número del mes anterior (formato MM)"""
    ahora = datetime.now(obtener_zona_horaria())
    mes_anterior = ahora.month - 1 or 12
    return f"{mes_anterior:02d}"


def obtener_anio(mes_a_procesar, anio_override=None):
    """
    Obtiene el año a procesar.

    - Si se pasa `anio_override` (string u int no vacío), se usa ese valor
      directamente — esto es lo que permite que el input "año" de los
      workflows de GitHub Actions tenga efecto real.
    - Si no se pasa override, calcula el año automáticamente a partir de
      la fecha actual, restando 1 si el período es Diciembre o el 2° SAC
      (que en la práctica se liquida/corresponde al año anterior).
    """
    if anio_override not in (None, ""):
        return int(anio_override)

    ahora = datetime.now(obtener_zona_horaria())
    anio_actual = ahora.year

    if _normalizar_periodo(mes_a_procesar) in ("12", "2° sac"):
        return anio_actual - 1

    return anio_actual


def leer_override_env(nombre_var):
    """
    Lee una variable de entorno usada como override manual desde
    workflow_dispatch (ej: MES_OVERRIDE, ANIO_OVERRIDE) y la devuelve
    "limpia". Si no está seteada o viene vacía, devuelve "" (para que el
    caller pueda decidir el fallback automático en lugar de crashear).
    """
    return os.getenv(nombre_var, "").strip()

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
