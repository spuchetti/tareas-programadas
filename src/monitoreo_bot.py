"""
Bot de Monitoreo de Liquidaciones

Compara archivos actuales con sus snapshots y genera reportes de cambios.

======= EJECUCIÓN =========
Correr manualmente desde GitHub Actions → workflow_dispatch
O bien: python src/monitoreo_bot.py

======= CREDENCIALES =========
Corre 100% con la Service Account (GDRIVE_JSON) — no depende de OAuth.
No crea archivos nuevos en Drive:
  - Lee el .xlsx actual de cada repartición con openpyxl (vía descarga
    normal de Drive), igual que fv_drive_bot.py / reporte_anual_bot.py.
  - Lee el snapshot (que ES un Google Sheet, creado por snapshot_bot.py)
    exportándolo como .xlsx y leyéndolo también con openpyxl.
  - Si hay cambios, actualiza el CONTENIDO del snapshot existente vía
    Sheets API (values.clear + values.update) — es edición de un archivo
    que ya existe, no creación, así que la Service Account puede hacerlo.

Si una repartición todavía no tiene snapshot, este bot NO lo crea (eso
requeriría crear un archivo nuevo). Se limita a registrar los agentes en
el padrón y avisar que hace falta correr snapshot_bot.py manualmente para
esa repartición.
"""

import sys
import os
import time
import traceback
from datetime import datetime

sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from utils.common_utils import registrar_inicio, registrar_resumen
from utils.drive_utils import inicializar_drive, obtener_archivos, descargar_archivo
from utils.gmail_utils import enviar_email_html_con_adjuntos, generar_html_resumen_monitoreo
from utils.registro_utils import (
    inicializar_sheets,
    obtener_o_crear_hoja_registro,
    obtener_id_agente,
    flush_registro_pendientes,
)
from utils.monitoreo_utils import (
    CONFIG,
    HOJAS_ORDEN,
    leer_rango,
    leer_hoja_xlsx,
    actualizar_snapshot_hoja,
    comparar_hojas_normal,
    comparar_hojas_caja,
    separar_complementarias_agrupado,
    generar_xlsx_cambios,
    generar_csv_modificados,
    generar_csv_complementarias,
    generar_csv_liquidacion_completa,
    extraer_reparticion,
    extraer_anio_desde_nombre,
    hoja_a_periodo,
    normalizar_nombre,
    normalizar_periodo,
    construir_asunto_monitoreo,
)

# Nombres genéricos que se muestran en el cuerpo del mail (sección "Archivos
# adjuntos"), igual que hacía el Apps Script original: el archivo adjunto de
# verdad lleva el nombre real (repartición + timestamp), pero en el listado
# del mail se muestra el patrón genérico para que sea fácil de leer de un
# vistazo, sin repetir el nombre completo de la repartición 4 veces.
NOMBRE_GENERICO_XLSX_MODIF = "Modificaciones_ReparticionPeriodoAño.xlsx"
NOMBRE_GENERICO_CSV_MODIF = "Modificaciones_ReparticionPeriodoAño.csv"
NOMBRE_GENERICO_CSV_COMPLEMENTARIAS = "Complementarias_ReparticionPeriodoAño.csv"
NOMBRE_GENERICO_CSV_RECTIFICATIVA = "Rectificativa_ReparticionPeriodoAño.csv"


# =============================================================================
# BÚSQUEDA DEL SNAPSHOT
# =============================================================================

def obtener_snapshot_de_archivo(nombre_archivo, carpeta_snapshots_id, drive_svc):
    """Busca si existe un snapshot (Google Sheet) para el archivo. Solo lectura."""
    nombre_snap = f"[SNAP] {nombre_archivo.replace('.xlsx', '')}"
    query = f"'{carpeta_snapshots_id}' in parents and name='{nombre_snap}' and trashed=false"
    result = drive_svc.files().list(
        q=query,
        fields="files(id, name, mimeType)",
        supportsAllDrives=True,
        includeItemsFromAllDrives=True
    ).execute()
    archivos = result.get("files", [])
    return archivos[0] if archivos else None


def obtener_carpeta_snapshots(drive_svc):
    """Busca la carpeta de snapshots. NO la crea (eso es tarea de snapshot_bot)."""
    query = (
        f"'{CONFIG['CARPETA_INTERNA_ID']}' in parents "
        f"and name='{CONFIG['CARPETA_SNAPSHOTS']}' "
        f"and mimeType='application/vnd.google-apps.folder' and trashed=false"
    )
    result = drive_svc.files().list(
        q=query, fields="files(id)",
        supportsAllDrives=True, includeItemsFromAllDrives=True
    ).execute()
    folders = result.get("files", [])
    return folders[0]["id"] if folders else None


# =============================================================================
# PROCESAMIENTO DE ARCHIVO
# =============================================================================

def procesar_archivo(archivo, carpeta_snapshots_id, drive_svc, sheets_svc_registro):
    """
    Procesa un archivo Excel: lo compara con su snapshot (si existe) y
    genera reportes de cambios. Actualiza el snapshot al final (in-place).

    Envuelve todo en try/finally para garantizar que, pase lo que pase, se
    escriban en la Sheets API (en UN solo lote) los agentes acumulados
    durante el procesamiento — así nunca se pierden altas/actualizaciones
    de agentes aunque el archivo falle a mitad de camino.
    """
    estado = {"hoja_registro": None}
    try:
        return _procesar_archivo_impl(archivo, carpeta_snapshots_id, drive_svc, sheets_svc_registro, estado)
    finally:
        flush_registro_pendientes(estado["hoja_registro"])


def _procesar_archivo_impl(archivo, carpeta_snapshots_id, drive_svc, sheets_svc_registro, estado):
    nombre_archivo = archivo["name"]
    es_caja = "caja" in nombre_archivo.lower()
    fila_inicio = CONFIG["FILA_INICIO_CAJA"] if es_caja else CONFIG["FILA_INICIO_DEFAULT"]
    reparticion = extraer_reparticion(nombre_archivo)
    anio = extraer_anio_desde_nombre(nombre_archivo)

    print(f"\n📄 Procesando: {nombre_archivo}")
    print(f"   Tipo: {'Caja' if es_caja else 'Normal'}")
    print(f"   Repartición: {reparticion}")

    # ── 1. Buscar snapshot (sin crearlo si no existe) ───────────────────────
    snapshot = obtener_snapshot_de_archivo(nombre_archivo, carpeta_snapshots_id, drive_svc)

    # ── 2. Descargar el archivo actual (una sola vez) ───────────────────────
    fh_actual = descargar_archivo(drive_svc, archivo)
    if not fh_actual:
        print("   ❌ No se pudo descargar el archivo actual")
        return None

    hoja_registro = obtener_o_crear_hoja_registro(sheets_svc_registro, nombre_archivo) if sheets_svc_registro else None
    estado["hoja_registro"] = hoja_registro

    # ── 3. Si no hay snapshot: solo registrar agentes, sin comparar ─────────
    if not snapshot:
        print("   📸 Sin snapshot todavía — correr snapshot_bot.py para crear el inicial "
              "de esta repartición. Se registran los agentes igualmente.")
        for nombre_hoja in HOJAS_ORDEN:
            datos = leer_rango(leer_hoja_xlsx(fh_actual, nombre_hoja, fila_inicio))
            if hoja_registro:
                for fila in datos:
                    cuil = str(fila[0] if len(fila) > 0 else "").strip()
                    dni = str(fila[CONFIG["COL_DNI"] - CONFIG["COL_INICIO"]] if len(fila) > 1 else "").strip()
                    nombre = str(fila[3] if len(fila) > 3 else "").strip()
                    if cuil or dni or nombre:
                        obtener_id_agente(cuil, dni, nombre, hoja_registro)
        return None

    # ── 4. Descargar el snapshot (exportado como xlsx, una sola vez) ────────
    fh_snapshot = descargar_archivo(drive_svc, snapshot)
    if not fh_snapshot:
        print("   ⚠️  No se pudo descargar/exportar el snapshot — se omite la comparación esta vez")
        return None

    # ── 5. Recorrer las hojas del período y comparar ─────────────────────────
    cambios_por_hoja = []
    todos_los_cambios = []
    datos_por_hoja_actual = {}   # se reutiliza después para reescribir el snapshot

    for nombre_hoja in HOJAS_ORDEN:
        datos_actual = leer_rango(leer_hoja_xlsx(fh_actual, nombre_hoja, fila_inicio))
        datos_snap = leer_rango(leer_hoja_xlsx(fh_snapshot, nombre_hoja, fila_inicio))
        datos_por_hoja_actual[nombre_hoja] = datos_actual

        if not datos_actual and not datos_snap:
            continue

        resultado = (
            comparar_hojas_caja(datos_actual, datos_snap, hoja_registro)
            if es_caja else
            comparar_hojas_normal(datos_actual, datos_snap, hoja_registro)
        )
        cambios = resultado.get("cambios", [])
        mapa_actual = resultado.get("mapa_actual", {})

        if cambios:
            elims = [c for c in cambios if c["tipo"] == "eliminado"]
            nuevos = [c for c in cambios if c["tipo"] == "nuevo"]
            modifs = [c for c in cambios if c["tipo"] == "modificado"]

            tiene_comp = False
            if not es_caja and mapa_actual:
                comps = separar_complementarias_agrupado(mapa_actual)
                tiene_comp = len(comps.get("complementarias", {})) > 0

            cambios_por_hoja.append({
                "periodo": hoja_a_periodo(nombre_hoja, anio),
                "cambios": cambios,
                "mapa_actual": mapa_actual,
                "eliminados": len(elims),
                "nuevos": len(nuevos),
                "modificados": len(modifs),
                "complementarias": tiene_comp,
            })
            todos_los_cambios.extend(cambios)

    # ── 6. Si hay cambios: generar adjuntos y enviar mail ────────────────────
    if cambios_por_hoja:
        total_cambios = len(todos_los_cambios)
        total_eliminados = sum(h["eliminados"] for h in cambios_por_hoja)
        total_nuevos = sum(h["nuevos"] for h in cambios_por_hoja)
        total_modificados = sum(h["modificados"] for h in cambios_por_hoja)

        print(f"\n   📊 Cambios detectados: {total_cambios} "
              f"(elim: {total_eliminados}, nuevos: {total_nuevos}, modif: {total_modificados})")

        periodos_html = [{
            "periodo": h["periodo"], "eliminados": h["eliminados"], "nuevos": h["nuevos"],
            "modificados": h["modificados"], "complementarias": h.get("complementarias", False),
        } for h in cambios_por_hoja]

        os.makedirs("generados", exist_ok=True)
        adjuntos_info, adjuntos_paths = [], []

        # Se genera un set de adjuntos POR PERÍODO (igual que el Apps Script
        # original): así el nombre real de archivo puede incluir el período
        # exacto (ej: Modificaciones_MinisterioDeSaludMayo2026.xlsx) en vez
        # de un timestamp genérico que no dice nada del contenido.
        hay_modif_xlsx = hay_modif_csv = hay_complementarias_csv = hay_rectificativa_csv = False

        for h in cambios_por_hoja:
            periodo = h["periodo"]
            cambios_h = h["cambios"]
            mapa_actual_h = h["mapa_actual"]
            modifs_h = [c for c in cambios_h if c["tipo"] == "modificado"]
            elims_h = [c for c in cambios_h if c["tipo"] == "eliminado"]
            nuevos_h = [c for c in cambios_h if c["tipo"] == "nuevo"]

            sufijo = f"{normalizar_nombre(reparticion)}{normalizar_periodo(periodo)}"

            try:
                nombre_xlsx = f"Modificaciones_{sufijo}.xlsx"
                ruta_xlsx = os.path.join("generados", nombre_xlsx)
                generar_xlsx_cambios(modifs_h, elims_h, nuevos_h, periodo, reparticion, ruta_xlsx)
                if os.path.exists(ruta_xlsx):
                    adjuntos_paths.append(ruta_xlsx)
                    hay_modif_xlsx = True
            except Exception as e:
                print(f"   ⚠️ Error generando XLSX ({periodo}): {e}")

            try:
                if modifs_h or nuevos_h:
                    nombre_csv = f"Modificaciones_{sufijo}.csv"
                    ruta_csv = os.path.join("generados", nombre_csv)
                    generar_csv_modificados(modifs_h, nuevos_h, ruta_csv)
                    if os.path.exists(ruta_csv):
                        adjuntos_paths.append(ruta_csv)
                        hay_modif_csv = True
            except Exception as e:
                print(f"   ⚠️ Error generando CSV modificados ({periodo}): {e}")

            try:
                if not es_caja and mapa_actual_h:
                    comps = separar_complementarias_agrupado(mapa_actual_h)
                    if comps.get("complementarias", {}):
                        nombre_comp = f"Complementarias_{sufijo}.csv"
                        ruta_comp = os.path.join("generados", nombre_comp)
                        if generar_csv_complementarias(comps["complementarias"], ruta_comp) and os.path.exists(ruta_comp):
                            adjuntos_paths.append(ruta_comp)
                            hay_complementarias_csv = True
            except Exception as e:
                print(f"   ⚠️ Error generando CSV complementarias ({periodo}): {e}")

            try:
                if mapa_actual_h:
                    nombre_rect = f"Rectificativa_{sufijo}.csv"
                    ruta_rect = os.path.join("generados", nombre_rect)
                    generar_csv_liquidacion_completa(mapa_actual_h, es_caja, ruta_rect)
                    if os.path.exists(ruta_rect):
                        adjuntos_paths.append(ruta_rect)
                        hay_rectificativa_csv = True
            except Exception as e:
                print(f"   ⚠️ Error generando CSV liquidación completa ({periodo}): {e}")

        # El cuerpo del mail muestra UNA fila genérica por categoría (no una
        # por archivo real), igual que el Apps Script original.
        if hay_modif_xlsx:
            adjuntos_info.append({"es_xlsx": True, "nombre": NOMBRE_GENERICO_XLSX_MODIF,
                                   "descripcion": "Detalles de registros modificados, nuevos y eliminados"})
        if hay_modif_csv:
            adjuntos_info.append({"es_xlsx": False, "nombre": NOMBRE_GENERICO_CSV_MODIF,
                                   "descripcion": "Registros modificados y nuevos"})
        if hay_complementarias_csv:
            adjuntos_info.append({"es_xlsx": False, "nombre": NOMBRE_GENERICO_CSV_COMPLEMENTARIAS,
                                   "descripcion": "Liquidación complementaria"})
        if hay_rectificativa_csv:
            adjuntos_info.append({"es_xlsx": False, "nombre": NOMBRE_GENERICO_CSV_RECTIFICATIVA,
                                   "descripcion": "Liquidación completa - solo ordinarias"})

        fecha_hora = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
        html = generar_html_resumen_monitoreo(
            reparticion=reparticion, nombre_archivo=nombre_archivo,
            total_cambios=total_cambios, total_eliminados=total_eliminados,
            total_nuevos=total_nuevos, total_modificados=total_modificados,
            cambios_por_periodo=periodos_html, adjuntos_info=adjuntos_info, fecha=fecha_hora,
        )
        periodos_lista = [p["periodo"] for p in periodos_html]
        asunto = construir_asunto_monitoreo(reparticion, periodos_lista)

        # Primero el mail. Si falla, no actualizamos el snapshot: la próxima
        # corrida va a volver a detectar los mismos cambios y reintentar.
        print("\n   📧 Enviando email...")
        enviar_email_html_con_adjuntos(asunto, html, adjuntos_paths, "SMTP_TO_MONITOREO")

        _actualizar_snapshot_in_place(sheets_svc_registro, snapshot, datos_por_hoja_actual, fila_inicio)
        return {"cambios": total_cambios, "archivo": nombre_archivo}

    # ── 7. Sin cambios: igual refrescamos el snapshot por las dudas ─────────
    print("   ⏭️  Sin cambios detectados")
    _actualizar_snapshot_in_place(sheets_svc_registro, snapshot, datos_por_hoja_actual, fila_inicio)
    return None


def _actualizar_snapshot_in_place(sheets_svc, snapshot, datos_por_hoja_actual, fila_inicio):
    """
    Reescribe el contenido del snapshot existente con los datos actuales,
    vía Sheets API (values.clear + values.update). No crea ni reemplaza
    el archivo — solo edita su contenido, que la Service Account sí puede
    hacer.

    Se escribe a partir de `fila_inicio` (el mismo usado para leer tanto el
    archivo real como el propio snapshot en leer_hoja_xlsx) para que el
    snapshot no quede "corrido" respecto de cómo se lee en la próxima
    corrida — ver el docstring de actualizar_snapshot_hoja() en
    monitoreo_utils.py para el detalle del bug que esto evita.
    """
    if not sheets_svc:
        print("   ⚠️  Sin Sheets service disponible — no se pudo actualizar el snapshot")
        return
    try:
        for nombre_hoja, datos in datos_por_hoja_actual.items():
            if not datos:
                continue
            actualizar_snapshot_hoja(sheets_svc, snapshot["id"], nombre_hoja, datos, fila_inicio=fila_inicio)
        print("   ✅ Snapshot actualizado (contenido)")
    except Exception as e:
        print(f"   ⚠️  Error actualizando snapshot: {e}")
        traceback.print_exc()


# =============================================================================
# PRINCIPAL
# =============================================================================

def ejecutar_principal():
    """Función principal del bot de monitoreo. Corre 100% con Service Account."""
    inicio = time.time()
    ahora = registrar_inicio("MONITOREO DE LIQUIDACIONES")

    print("🔑 Inicializando Drive (Service Account)...")
    drive_svc = inicializar_drive()
    if not drive_svc:
        print("   ❌ No se pudo inicializar Drive — revisar el secret GDRIVE_JSON")
        return
    print("   ✅ Drive inicializado")

    sheets_svc_registro = inicializar_sheets()
    if not sheets_svc_registro:
        print("   ⚠️ No se pudo inicializar Sheets — se continúa sin registro de agentes "
              "ni actualización de snapshots")

    carpeta_snapshots_id = obtener_carpeta_snapshots(drive_svc)
    if not carpeta_snapshots_id:
        print("   ❌ Carpeta de snapshots no encontrada — correr snapshot_bot.py al menos "
              "una vez para crearla")
        return
    print(f"   📁 Carpeta snapshots: {carpeta_snapshots_id}")

    print(f"\n📂 Buscando archivos en carpeta {CONFIG['CARPETA_ID']}...")
    archivos = obtener_archivos(drive_svc, CONFIG["CARPETA_ID"])
    if not archivos:
        print("   ❌ No se encontraron archivos Excel")
        return
    print(f"   ✅ Archivos Excel encontrados: {len(archivos)}")

    procesados, con_cambios, errores = 0, 0, 0
    errores_lista = []

    for i, archivo in enumerate(archivos, 1):
        print(f"\n{'='*60}")
        print(f"[{i}/{len(archivos)}] Procesando...")
        print(f"{'='*60}")
        try:
            resultado = procesar_archivo(archivo, carpeta_snapshots_id, drive_svc, sheets_svc_registro)
            procesados += 1
            if resultado and resultado.get("cambios", 0) > 0:
                con_cambios += 1
        except Exception as e:
            print(f"   ❌ Error procesando {archivo['name']}: {e}")
            traceback.print_exc()
            errores += 1
            errores_lista.append(archivo["name"])

    duracion = time.time() - inicio
    print(f"\n{'='*60}")
    print("📊 RESUMEN FINAL")
    print(f"{'='*60}")
    print(f"📁 Archivos procesados: {procesados}")
    print(f"📝 Con cambios: {con_cambios}")
    print(f"⏭️  Sin cambios: {procesados - con_cambios - errores}")
    print(f"❌ Errores: {errores}")
    if errores_lista:
        print("   Archivos con error:")
        for e in errores_lista:
            print(f"     ⚠️ {e}")
    print(f"⏱️  Tiempo total: {duracion:.0f}s ({duracion/60:.1f} min)")
    print(f"{'='*60}")

    registrar_resumen(inicio, procesados, len(archivos), 0, errores_lista)


if __name__ == "__main__":
    ejecutar_principal()