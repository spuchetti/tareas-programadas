"""
Unificador Mensual - Actualiza CSV consolidado en Drive agregando datos del mes actual
"""

import sys
import os
import io
import csv
import time
from datetime import datetime

# Agregar esta línea para que Python encuentre los módulos utils
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

# Luego el resto de imports
from utils.common_utils import (
    registrar_inicio, registrar_resumen, 
    nombre_mes, obtener_mes_anterior, obtener_anio, obtener_zona_horaria
)
from utils.drive_utils import (
    inicializar_drive, obtener_archivos, descargar_archivo,
    guardar_csv_localmente
)
from utils.excel_utils import extraer_datos_excel
from utils.gmail_utils import (
    enviar_email_html_con_adjuntos, 
    generar_html_resumen_unificador
)

# Configuración
MES_ACTUAL = obtener_mes_anterior()  # Mes que estamos procesando


def obtener_nombre_csv():
    """Obtiene el nombre del archivo CSV basado en el mes y año actual"""
    mes_actual_formateado = MES_ACTUAL
    anio_actual = obtener_anio(mes_actual_formateado)
    nombre_mes_actual = nombre_mes(mes_actual_formateado)
    
    # Formato: Unificado_MesAño.csv (ej: Unificado_Noviembre2025.csv)
    nombre_csv = f"Unificado_{nombre_mes_actual}{anio_actual}.csv"
    return nombre_csv


def extraer_y_preparar_datos_mes(drive, archivos_excel):
    """Extrae datos de archivos Excel y los prepara para el CSV"""
    mes_procesando = MES_ACTUAL

    datos_mes = []
    errores = []
    archivos_procesados = 0
    filas_totales = 0
    
    for archivo in archivos_excel:
        print(f"\n📄 Procesando: {archivo['name']}")
        
        try:
            # Descargar archivo
            fh = descargar_archivo(drive, archivo)
            if not fh:
                errores.append(f"No se pudo descargar: {archivo['name']}")
                continue
            
            # Extraer datos del Excel (solo los 24 campos A-X)
            datos_excel = extraer_datos_excel(fh, archivo['name'])
            
            if datos_excel:
                # Agregar directamente los datos del Excel a datos_mes
                # Cada fila_excel ya contiene los 24 campos
                datos_mes.extend(datos_excel)
                
                filas_agregadas = len(datos_excel)
                filas_totales += filas_agregadas
                archivos_procesados += 1
                print(f"   ✅ {filas_agregadas} filas extraídas (24 campos cada una)")
            else:
                print(f"   ⚠ Sin datos en la hoja del mes {mes_procesando}")
                errores.append(f"Sin datos en hoja {mes_procesando}: {archivo['name']}")
                
        except Exception as e:
            error_msg = f"Error procesando {archivo['name']}: {str(e)}"
            print(f"   ❌ {error_msg}")
            errores.append(error_msg)
    
    return datos_mes, archivos_procesados, filas_totales, errores


def combinar_con_existente(csv_existente, datos_nuevos):
    """
    Combina CSV existente con nuevos datos.
    """
    if not csv_existente:
        # Si no hay CSV existente, crear uno nuevo con encabezados de 24 campos
        encabezados = [
            "1-cuil",
            "2-dni",
            "3-tipo doc",
            "4-nombre y apellido",
            "5-cod liq",
            "6-sit revista",
            "7-estado del afil",
            "8-reparticion",
            "9-aporte personal",
            "10-adherente sec",
            "11-fondo v",
            "12-hijo menor de 35",
            "13-menor a cargo",
            "14-cred asist",
            "15-sueldo sin desc",
            "16-sueldo con desc",
            "17-reajs aporte pers",
            "18-reaj adherente sec",
            "19-reajuste fv",
            "20-reajuste hijo menor",
            "21-reajuste menor a cargo",
            "22-reajuste cred asistencial",
            "23-aporte patronal",
            "24-reajuste aporte patronal"
        ]
        return encabezados, datos_nuevos
    
    # Leer CSV existente
    reader = csv.reader(io.StringIO(csv_existente))
    filas_existentes = list(reader)
    
    if not filas_existentes:
        return [], datos_nuevos
    
    # Separar encabezados y datos
    encabezados = filas_existentes[0]
    datos_existentes = filas_existentes[1:]
    
    # Combinar todos los datos existentes con nuevos datos
    datos_combinados = datos_existentes + datos_nuevos
    
    return encabezados, datos_combinados


def ejecutar_principal():
    """Función principal del unificador mensual"""
    inicio = time.time()
    mes_actual = MES_ACTUAL
    anio_actual = obtener_anio(mes_actual)
    nombre_mes_actual = nombre_mes(mes_actual)
    
    periodo_legible = f"{nombre_mes_actual}/{anio_actual}"
    periodo_legible_upper = periodo_legible.upper()

    # Obtener nombre del CSV
    NOMBRE_CSV = obtener_nombre_csv()

    ahora = registrar_inicio(f"UNIFICADOR MENSUAL - {nombre_mes_actual} {anio_actual} ({mes_actual})")
    
    # 1. Inicializar Drive (solo para leer archivos)
    drive = inicializar_drive()
    if not drive:
        print("❌ No se pudo inicializar Drive")
        return
    
    # 2. Obtener archivos Excel
    print("📁 Buscando archivos Excel en Drive...")
    archivos = obtener_archivos(drive)
    
    # Filtrar solo Excel
    archivos_excel = [
        a for a in archivos 
        if a["name"].lower().endswith(('.xlsx', '.xlsm')) 
        or a["mimeType"] == "application/vnd.google-apps.spreadsheet"
    ]
    
    if not archivos_excel:
        print("❌ No se encontraron archivos Excel")
        return
    
    print(f"✅ Encontrados {len(archivos_excel)} archivos Excel\n")
    print(f"📅 Procesando datos del mes: {mes_actual} ({nombre_mes_actual})\n")
    
    # 3. Extraer datos del mes actual de todos los archivos
    datos_mes_actual, archivos_procesados, filas_totales, errores = extraer_y_preparar_datos_mes(
        drive, archivos_excel
    )
    
    if not datos_mes_actual:
        print("⚠️ No se extrajeron datos del mes actual")
        if not errores:
            errores.append("No se pudieron extraer datos de ningún archivo")
        return
    
    # 4. Como cada mes genera archivo nuevo, no buscamos CSV existente
    print(f"\n📥 Cada mes genera archivo nuevo con nombre diferente, no se buscará archivo previo")
    csv_existente = None
    
    # 5. Combinar datos existentes con nuevos datos del mes
    print(f"🔄 Preparando datos del mes {mes_actual}...")
    encabezados, datos_combinados = combinar_con_existente(csv_existente, datos_mes_actual)
    
    # Asegurar que tenemos encabezados
    if not encabezados:
        encabezados = [
            "1-cuil",
            "2-dni",
            "3-tipo doc",
            "4-nombre y apellido",
            "5-cod liq",
            "6-sit revista",
            "7-estado del afil",
            "8-reparticion",
            "9-aporte personal",
            "10-adherente sec",
            "11-fondo v",
            "12-hijo menor de 35",
            "13-menor a cargo",
            "14-cred asist",
            "15-sueldo sin desc",
            "16-sueldo con desc",
            "17-reajs aporte pers",
            "18-reaj adherente sec",
            "19-reajuste fv",
            "20-reajuste hijo menor",
            "21-reajuste menor a cargo",
            "22-reajuste cred asistencial",
            "23-aporte patronal",
            "24-reajuste aporte patronal"
        ]
    
    # Agregar encabezados a los datos
    datos_finales = [encabezados] + datos_combinados
    
    lineas_totales = len(datos_combinados)
    
    print(f"📊 Datos del mes: {lineas_totales} filas totales")
    print(f"📈 Archivo: {NOMBRE_CSV}")
    
    # 6. GUARDAR CSV LOCALMENTE (única salida)
    print(f"\n💾 Guardando CSV localmente...")
    ruta_csv_local = guardar_csv_localmente(datos_finales, NOMBRE_CSV)
    
    if ruta_csv_local:
        print(f"✅ CSV guardado localmente en: {ruta_csv_local}")
        print(f"   📊 Total de datos: {lineas_totales} filas (24 columnas cada una)")

    else:
        errores.append("No se pudo guardar CSV localmente")
        print("❌ Error crítico: No se pudo guardar el CSV")
        return
    
    # 7. Mostrar resumen
    registrar_resumen(
        inicio, 
        archivos_procesados=archivos_procesados,
        archivos_encontrados=len(archivos_excel),
        filas_procesadas=filas_totales,
        errores=errores
    )
    
    # 8. Enviar email de resumen CON ARCHIVO ADJUNTO
    print("\n📧 Preparando email de resumen con archivo adjunto...")
    asunto = f"🟢🔵 OSER - UNIFICADO MENSUAL AUTOMÁTICO | PERIODO: {periodo_legible_upper}"
    
    html = generar_html_resumen_unificador(
        periodo_legible,
        ahora.strftime("%d-%m-%Y %H:%M:%S"),
        False,  # sheets_creado siempre False
        lineas_totales
    )
    
    # Adjuntar el CSV completo al email
    adjuntos = []
    if ruta_csv_local and os.path.exists(ruta_csv_local):
        # Verificar tamaño del archivo
        file_size = os.path.getsize(ruta_csv_local) / (1024 * 1024)  # MB
        print(f"📎 Tamaño del archivo: {file_size:.2f} MB")
        
        if file_size < 25:  # Gmail permite hasta 25MB
            adjuntos.append(ruta_csv_local)
            print(f"📎 Adjuntando archivo: {ruta_csv_local}")
        else:
            print("⚠️ Archivo demasiado grande para adjuntar al email (>25MB)")
            errores.append(f"Archivo demasiado grande para email: {file_size:.2f} MB")
    else:
        print("⚠️ No se pudo adjuntar archivo al email")
        errores.append("No se pudo adjuntar archivo al email")
    
    enviar_email_html_con_adjuntos(asunto, html, adjuntos)
    
    print("\n" + "=" * 60)
    print("✅ PROCESO COMPLETADO!")
    print("=" * 60)
    print(f"📁 Archivo generado: {NOMBRE_CSV}")
    print(f"📊 Total de filas: {lineas_totales}")
    print(f"📅 Mes procesado: {mes_actual}/{anio_actual}")
    print(f"💾 Guardado en: {ruta_csv_local}")
    print(f"📧 Email enviado con {'archivo adjunto' if adjuntos else 'sin adjunto'}")
    print("=" * 60)


if __name__ == "__main__":
    ejecutar_principal()