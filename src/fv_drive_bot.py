"""
Bot de Fondo Voluntario - Busca archivos con diferencias
"""

import sys
import os
import time
import concurrent.futures
from datetime import datetime
from io import BytesIO

# Agregar esta línea para que Python encuentre los módulos utils
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

# Luego los imports normales
from utils.common_utils import (
    registrar_inicio, registrar_resumen, 
    nombre_mes, obtener_mes_anterior, obtener_anio, crear_directorio_salida
)
from utils.drive_utils import inicializar_drive, obtener_archivos, descargar_archivo
from utils.excel_utils import sanitizar_libro_remover_filtros
from utils.gmail_utils import enviar_email_html_con_adjuntos, generar_html_resumen_fv

import openpyxl
from copy import copy

# Configuración
TEXTO_OBJETIVO = "REVISAR"
COLUMNA_OBJETIVO = 33
FILA_INICIO = 4
MAXIMO_HILOS = 5


def buscar_en_hoja(fh, nombre, hoja):
    """Busca texto objetivo en una hoja específica"""
    try:
        fh.seek(0)
        wb = openpyxl.load_workbook(fh, data_only=True, read_only=True)

        if hoja not in wb.sheetnames:
            wb.close()
            return False

        ws = wb[hoja]

        for row in ws.iter_rows(min_row=FILA_INICIO, values_only=True):
            if row is None:
                break

            if len(row) < COLUMNA_OBJETIVO:
                continue

            val = row[COLUMNA_OBJETIVO - 1]

            if val is None:
                break

            if TEXTO_OBJETIVO.upper() in str(val).upper():
                wb.close()
                return True

        wb.close()
        return False

    except Exception as e:
        print(f"❌ Error leyendo {nombre}: {e}")
        return False


def generar_archivo_solo_hojas(fh, nombre_archivo, hoja_periodo):
    """Genera un archivo Excel solo con las hojas necesarias"""
    try:
        fh.seek(0)
        libro_original = None
        intento_sanitizar = False

        for intento in range(2):
            try:
                libro_original = openpyxl.load_workbook(
                    fh,
                    data_only=False,
                    keep_vba=True,
                    keep_links=False
                )
                break
            except Exception:
                if not intento_sanitizar:
                    fh_sanitizado = sanitizar_libro_remover_filtros(fh)
                    intento_sanitizar = True
                    if fh_sanitizado is None:
                        break
                    else:
                        fh = fh_sanitizado
                        continue
                else:
                    libro_original = None
                    break

        if libro_original is None:
            return None

        hojas_a_copiar = [hoja_periodo]
        if "MINIMOS" in libro_original.sheetnames:
            hojas_a_copiar.append("MINIMOS")

        nuevo_wb = openpyxl.Workbook()
        if nuevo_wb.active and nuevo_wb.active.title == "Sheet":
            nuevo_wb.remove(nuevo_wb.active)

        for hoja_nombre in hojas_a_copiar:
            origen = libro_original[hoja_nombre]
            nueva = nuevo_wb.create_sheet(title=hoja_nombre)

            for row in origen.iter_rows():
                for cell in row:
                    nueva_cell = nueva.cell(row=cell.row, column=cell.column)

                    if cell.data_type == 'f' and cell.value:
                        nueva_cell.value = f"={cell.value}"
                    else:
                        nueva_cell.value = cell.value

                    try:
                        nueva_cell.font = copy(cell.font)
                        nueva_cell.fill = copy(cell.fill)
                        nueva_cell.border = copy(cell.border)
                        nueva_cell.alignment = copy(cell.alignment)
                        nueva_cell.number_format = copy(cell.number_format)
                        nueva_cell.protection = copy(cell.protection)
                    except Exception:
                        pass

            nueva.auto_filter = None

        # Guardar archivo
        carpeta = crear_directorio_salida()
        ruta = os.path.join(carpeta, f"{os.path.splitext(nombre_archivo)[0]}.xlsx")
        nuevo_wb.save(ruta)
        nuevo_wb.close()
        libro_original.close()

        print(f"   → Archivo generado: {ruta}")
        return ruta

    except Exception as e:
        print(f"❌ Error generando archivo {nombre_archivo}: {e}")
        return None


def procesar_archivo(archivo, hoja_periodo):
    """Procesa un archivo individual"""
    try:
        servicio_drive = inicializar_drive()
        if not servicio_drive:
            return archivo["name"], None, None

        fh = descargar_archivo(servicio_drive, archivo)
        if not fh:
            return archivo["name"], None, None

        encontrado = buscar_en_hoja(fh, archivo["name"], hoja_periodo)
        fh_copia = BytesIO(fh.getvalue())

        return archivo["name"], encontrado, fh_copia
    except Exception as e:
        print(f"❌ Error procesando {archivo['name']}: {e}")
        return archivo["name"], None, None


def ejecutar_principal():
    """Función principal del bot"""
    inicio = time.time()
    ahora = registrar_inicio("BOT FONDO VOLUNTARIO")

    # 1. Inicializar Drive
    drive = inicializar_drive()
    if not drive:
        return

    # 2. Obtener archivos
    archivos = obtener_archivos(drive)
    if not archivos:
        print("❌ No se encontraron archivos.")
        return

    # 3. Determinar período
    periodo = obtener_mes_anterior()
    anio_actual = obtener_anio(periodo)
    periodo_legible = f"{nombre_mes(periodo)}/{anio_actual}"
    periodo_legible_upper = periodo_legible.upper()

    print(f"📄 Hoja a controlar: {periodo} ({periodo_legible})\n")

    # 4. Procesar archivos en paralelo
    encontrados = []
    adjuntos = []

    with concurrent.futures.ThreadPoolExecutor(max_workers=MAXIMO_HILOS) as pool:
        tareas = [
            pool.submit(procesar_archivo, a, periodo)
            for a in archivos
        ]

        for future in concurrent.futures.as_completed(tareas):
            nombre, ok, fh_copy = future.result()

            if ok:
                print(f"   ✔ {nombre} → REVISAR")
                encontrados.append(nombre)

                if fh_copy:
                    generado = generar_archivo_solo_hojas(fh_copy, nombre, periodo)
                    if generado:
                        adjuntos.append(generado)
            else:
                print(f"   ✖ {nombre}")

    # 5. Enviar email
    html = generar_html_resumen_fv(
        periodo_legible,
        len(archivos),
        len(encontrados),
        encontrados,
        ahora.strftime("%d-%m-%Y %H:%M:%S")
    )

    asunto = f"🟢🔵 OSER - CONTROL AUTOMÁTICO FONDO VOLUNTARIO | PERIODO: {periodo_legible_upper}"
    enviar_email_html_con_adjuntos(asunto, html, adjuntos)

    # 6. Mostrar resumen
    registrar_resumen(inicio, len(archivos), len(encontrados))


if __name__ == "__main__":
    ejecutar_principal()