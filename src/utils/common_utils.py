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


def nombre_mes(numero):
    """Convierte número de mes a nombre"""
    meses = {
        "01": "Enero", "02": "Febrero", "03": "Marzo", "04": "Abril",
        "05": "Mayo", "06": "Junio", "1° sac": "1SAC", "07": "Julio",
        "08": "Agosto", "09": "Septiembre", "10": "Octubre", "11": "Noviembre",
        "12": "Diciembre", "2º sac": "2SAC"
    }
    return meses.get(str(numero), "???")


def obtener_mes_anterior():
    """Obtiene el número del mes anterior (formato MM)"""
    ahora = datetime.now(obtener_zona_horaria())
    mes_anterior = ahora.month - 1 or 12
    return f"{mes_anterior:02d}"

def obtener_anio(mes_a_procesar):
    """Obtiene el año actual, validando si el mes a procesar es Diciembre incluido(en ese caso tiene que tomar el año anterior al actual)"""
    ahora = datetime.now(obtener_zona_horaria())
    anio_actual = ahora.year

    if mes_a_procesar == "12" or mes_a_procesar == "2° sac":
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
