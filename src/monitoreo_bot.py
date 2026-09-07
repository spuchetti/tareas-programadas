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

from utils.common_utils import registrar_inicio
from utils.drive_utils import (
    inicializar_drive,
    obtener_archivos,
    descargar_archivo,
    request_drive_con_reintentos,
)
from utils.gmail_utils import enviar_email_html_con_adjuntos, generar_html_resumen_monitoreo
from utils.registro_utils import (
    inicializar_sheets,
    obtener_o_crear_hoja_registro,
    obtener_id_agente,
    flush_registro_pendientes,
    obtener_avisos_cuil_invalido,
)
from utils.monitoreo_utils import (
    CONFIG,
    HOJAS_ORDEN,
    leer_rango,
    leer_hoja_xlsx,
    obtener_titulos_hojas_snapshot,
    leer_hojas_snapshot_batch,
    actualizar_snapshot_hoja,
    comparar_hojas_normal,
    comparar_hojas_caja,
    separar_complementarias,
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
    """
    Busca si existe un snapshot (Google Sheet) para el archivo. Solo lectura.

    Pasa por request_drive_con_reintentos (propagar_error=True) para que:
      - errores transitorios de red/SSL (ej. un SSLEOFError durante el
        refresh del token de la Service Account, visto en logs reales a
        mitad de una corrida larga) se reintenten en vez de tirar abajo
        el procesamiento de este archivo.
      - si los reintentos se agotan, la excepción se relance en vez de
        devolver None en silencio. Un None acá se interpreta más abajo
        como "no existe snapshot todavía" (dispara el flujo de
        solo-registro sin comparar) — confundir eso con "no se pudo ni
        preguntar" haría creer que falta el snapshot cuando en realidad
        existe y solo hubo un corte de red.
    El caller (_procesar_archivo_impl) es quien decide qué hacer con el
    error: lo atrapa y omite la comparación de este archivo, igual que ya
    se hace con la metadata del snapshot (paso 4 más abajo).
    """
    nombre_snap = f"[SNAP] {nombre_archivo.replace('.xlsx', '')}"
    query = f"'{carpeta_snapshots_id}' in parents and name='{nombre_snap}' and trashed=false"
    result = request_drive_con_reintentos(
        drive_svc.files().list(
            q=query,
            fields="files(id, name, mimeType)",
            supportsAllDrives=True,
            includeItemsFromAllDrives=True
        ).execute,
        f"buscar snapshot de '{nombre_archivo}'",
        propagar_error=True,
    )
    archivos = result.get("files", [])
    return archivos[0] if archivos else None


def obtener_carpeta_snapshots(drive_svc):
    """
    Busca la carpeta de snapshots. NO la crea (eso es tarea de snapshot_bot).

    Se llama una sola vez al arrancar el bot (no dentro del loop por
    archivo), pero igual pasa por request_drive_con_reintentos con
    propagar_error=True: si esto falla por un corte de red transitorio no
    reintentado, mejor un error explícito y visible en el job que un
    "carpeta de snapshots no encontrada" engañoso, que llevaría a pensar
    que hace falta correr snapshot_bot.py cuando en realidad la carpeta
    existe.
    """
    query = (
        f"'{CONFIG['CARPETA_SERVICES_ID']}' in parents "
        f"and name='{CONFIG['CARPETA_SNAPSHOTS']}' "
        f"and mimeType='application/vnd.google-apps.folder' and trashed=false"
    )
    result = request_drive_con_reintentos(
        drive_svc.files().list(
            q=query, fields="files(id)",
            supportsAllDrives=True, includeItemsFromAllDrives=True
        ).execute,
        "buscar carpeta de snapshots",
        propagar_error=True,
    )
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
        # Solo consola/log de GitHub Actions — nunca al mail (mismo criterio
        # que avisos_conceptos en monitoreo_utils.py). Se imprime en el
        # finally, no en _procesar_archivo_impl, para que salga pase lo que
        # pase con el resto del procesamiento de este archivo — igual que
        # flush_registro_pendientes.
        for aviso in obtener_avisos_cuil_invalido(estado["hoja_registro"]):
            print(f"   ⚠️  [CUIL inválido] {archivo['name']}: {aviso}")
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

    # ── 1. Descargar el archivo actual (una sola vez) ───────────────────────
    fh_actual = descargar_archivo(drive_svc, archivo)
    if not fh_actual:
        print("   ❌ No se pudo descargar el archivo actual")
        return {"estado": "sin_comparar", "motivo": "descarga", "archivo": nombre_archivo}

    # ── 2. Obtener hoja de registro ANTES de buscar el snapshot ─────────────
    # A propósito en este orden: si la búsqueda del snapshot (paso 3) falla
    # por un corte de red transitorio que agota los reintentos, igual
    # queremos haber podido registrar los agentes de este archivo — antes
    # el orden era al revés y un fallo en la búsqueda del snapshot (que
    # pasa primero) hacía perder también el registro de agentes de toda
    # la repartición en esa corrida.
    hoja_registro = obtener_o_crear_hoja_registro(sheets_svc_registro, nombre_archivo) if sheets_svc_registro else None
    estado["hoja_registro"] = hoja_registro

    # ── 3. Buscar snapshot (sin crearlo si no existe) ───────────────────────
    # obtener_snapshot_de_archivo ya reintenta ante errores transitorios de
    # red/SSL; si aun así falla, relanza (propagar_error=True) en vez de
    # devolver None, porque un None acá es ambiguo con "no existe snapshot
    # todavía" (ver su docstring). Se lo atrapa acá, igual que ya se hace
    # con la metadata del snapshot más abajo, para no perder el resto del
    # procesamiento de este archivo por un problema de red puntual.
    try:
        snapshot = obtener_snapshot_de_archivo(nombre_archivo, carpeta_snapshots_id, drive_svc)
    except Exception as e:
        print(f"   ⚠️  No se pudo buscar el snapshot (error de red) — se omite la comparación "
              f"esta vez, agentes igual registrados: {e}")
        for nombre_hoja in HOJAS_ORDEN:
            datos = leer_rango(leer_hoja_xlsx(fh_actual, nombre_hoja, fila_inicio))
            if hoja_registro:
                for fila in datos:
                    cuil = str(fila[0] if len(fila) > 0 else "").strip()
                    dni = str(fila[CONFIG["COL_DNI"] - CONFIG["COL_INICIO"]] if len(fila) > 1 else "").strip()
                    nombre = str(fila[3] if len(fila) > 3 else "").strip()
                    if cuil or dni or nombre:
                        obtener_id_agente(cuil, dni, nombre, hoja_registro)
        return {"estado": "sin_comparar", "motivo": "red_snapshot", "archivo": nombre_archivo}

    # ── 4. Si no hay snapshot (genuinamente): solo registrar agentes ────────
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
        return {"estado": "sin_comparar", "motivo": "sin_snapshot", "archivo": nombre_archivo}

    # ── 5. Resolver las pestañas del snapshot vía Sheets API ────────────────
    # Antes acá se exportaba el snapshot completo a .xlsx
    # (descargar_archivo -> drive.files().export) y se leía con openpyxl,
    # igual que el archivo actual. Google le pone un límite de tamaño a esa
    # exportación que NO es transitorio: con reparticiones grandes (ej.
    # MUNICIPIO PARANA) siempre rompía con "This file is too large to be
    # exported" y la comparación se saltaba en TODAS las corridas, no solo
    # en una puntual. Ahora se lee directo por Sheets API (ver
    # leer_hojas_snapshot_batch() en monitoreo_utils.py), que no tiene ese
    # límite.
    if not sheets_svc_registro:
        print("   ⚠️  Sin Sheets service disponible — se omite la comparación esta vez")
        return {"estado": "sin_comparar", "motivo": "sin_sheets_service", "archivo": nombre_archivo}
    try:
        titulos_hojas = obtener_titulos_hojas_snapshot(sheets_svc_registro, snapshot["id"])
    except Exception as e:
        print(f"   ⚠️  No se pudo leer la metadata del snapshot — se omite la comparación esta vez: {e}")
        return {"estado": "sin_comparar", "motivo": "metadata_snapshot", "archivo": nombre_archivo}

    # Las ~14 hojas se leen en UNA sola llamada batchGet (ver
    # leer_hojas_snapshot_batch en monitoreo_utils.py) en vez de 14
    # values().get() sueltos — evita agotar la cuota de lectura de la
    # Sheets API (60 req/min/usuario) a mitad de un archivo.
    #
    # Si aun así falla tras los reintentos, se propaga y se omite la
    # comparación de TODO el archivo esta corrida, SIN tocar el snapshot
    # (return antes de _actualizar_snapshot_in_place). A propósito no se
    # degrada a "comparar solo las hojas que sí se pudieron leer": hacerlo
    # sobreescribiría igual el snapshot de las hojas no leídas con los datos
    # actuales al final de la corrida (_actualizar_snapshot_in_place escribe
    # todo datos_por_hoja_actual sin condicionar por hoja), absorbiendo en
    # silencio cualquier cambio real no comparado. Saltar el archivo entero
    # es más simple y no tiene ese riesgo: la próxima corrida vuelve a
    # comparar contra el mismo snapshot bueno.
    try:
        datos_snap_por_hoja = leer_hojas_snapshot_batch(
            sheets_svc_registro, snapshot["id"], HOJAS_ORDEN, fila_inicio, titulos_hojas
        )
    except Exception as e:
        print(f"   ⚠️  No se pudo leer el snapshot vía Sheets API (cuota/red) — "
              f"se omite la comparación esta vez: {e}")
        return {"estado": "sin_comparar", "motivo": "lectura_snapshot", "archivo": nombre_archivo}

    # ── 6. Recorrer las hojas del período y comparar ─────────────────────────
    cambios_por_hoja = []
    todos_los_cambios = []
    datos_por_hoja_actual = {}   # se reutiliza después para reescribir el snapshot

    # Se pone en True si ALGUNA hoja tuvo una diferencia real contra el
    # snapshot, aunque esa diferencia no se termine reportando por mail
    # (ver "primera carga del período" más abajo). Sirve para decidir si
    # hay que reescribir el snapshot incluso cuando no se manda mail — si
    # no, la próxima corrida seguiría comparando contra un snapshot vacío
    # y cualquier modificación real posterior se seguiría viendo como
    # "nuevo" en vez de "modificado" (ver comentario de "primera carga").
    hubo_diferencias = False

    for nombre_hoja in HOJAS_ORDEN:
        datos_actual = leer_rango(leer_hoja_xlsx(fh_actual, nombre_hoja, fila_inicio))
        datos_snap = leer_rango(datos_snap_por_hoja.get(nombre_hoja, []))
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

        # Solo consola/log de GitHub Actions — a propósito NO se suma a
        # cambios_por_hoja/todos_los_cambios ni a ningún dato que después
        # se use para armar el HTML del mail (ver generar_html_resumen_monitoreo
        # más abajo). Avisa cuando un agente con 2+ liquidaciones tiene algún
        # valor de "sit. revista" que no está en CONCEPTOS_SIT_REVISTA
        # (utils/monitoreo_utils.py) — el caso donde una variante nueva sin
        # mapear puede romper el agrupamiento ordinaria/complementaria.
        for aviso in resultado.get("avisos_conceptos", []):
            print(f"   ⚠️  [sit. revista sin mapear] {hoja_a_periodo(nombre_hoja, anio)}: {aviso}")

        if not cambios:
            continue

        hubo_diferencias = True

        # Primera carga del período: la hoja no tenía NINGÚN registro en el
        # snapshot (datos_snap vacío) y ahora el archivo trae datos por
        # primera vez -> comparar_hojas_normal/caja marca cada fila como
        # "nuevo" porque ningún aid está en mapa_snap. Eso es la llegada
        # normal de la liquidación del período, no una anomalía: no
        # corresponde alertarla como si fueran altas sospechosas. Se
        # distingue de un agente puntual que se agrega más tarde a una hoja
        # que YA tenía otros registros (eso sigue reportándose como
        # "nuevo" normalmente, porque ahí datos_snap no está vacío).
        if not datos_snap:
            print(f"   ⏭️  {hoja_a_periodo(nombre_hoja, anio)}: primera carga del período "
                  f"({len(cambios)} registro(s)) — no se reporta, solo se actualiza el snapshot")
            continue

        elims = [c for c in cambios if c["tipo"] == "eliminado"]
        nuevos = [c for c in cambios if c["tipo"] == "nuevo"]
        modifs = [c for c in cambios if c["tipo"] == "modificado"]

        tiene_comp = False
        if mapa_actual:
            comps = separar_complementarias(mapa_actual, es_caja)
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

    # ── 7. Si hay cambios: generar adjuntos y enviar mail ────────────────────
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
                if mapa_actual_h:
                    comps = separar_complementarias(mapa_actual_h, es_caja)
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
        return {"estado": "con_cambios", "cambios": total_cambios, "archivo": nombre_archivo}

    # ── 7b. Solo hubo "primera carga de período" (sin nada reportable) ──────
    # hubo_diferencias=True pero cambios_por_hoja quedó vacío: todas las
    # diferencias detectadas fueron hojas que pasaron de "sin ningún
    # registro" a "con datos" por primera vez (ver el "continue" de más
    # arriba). No corresponde mail, pero SÍ hay que persistir el snapshot
    # con el estado actual — si no, la próxima corrida seguiría comparando
    # esa hoja contra un snapshot vacío, y una modificación real posterior
    # se seguiría viendo (incorrectamente) como alta nueva en vez de como
    # modificación.
    if hubo_diferencias:
        print("   📸 Solo hubo primera carga de período (sin cambios reportables) — "
              "se actualiza el snapshot sin enviar mail")
        _actualizar_snapshot_in_place(sheets_svc_registro, snapshot, datos_por_hoja_actual, fila_inicio)
        return {"estado": "sin_cambios", "archivo": nombre_archivo}

    # ── 8. Sin cambios: NO tocamos el snapshot ──────────────────────────────
    # Antes esto igual llamaba a _actualizar_snapshot_in_place() "por las
    # dudas". Pero si comparar_hojas() no encontró diferencias, el contenido
    # del snapshot ya es idéntico al del archivo actual — reescribirlo es un
    # no-op en términos de contenido. El costo NO es un no-op: cada llamada
    # recorre hasta 14 pestañas (01..12, 1°sac, 2°sac) y hace clear+update
    # por cada una, es decir hasta 28 escrituras a Sheets POR ARCHIVO. Con
    # una corrida de 246 archivos y demanda de cuota "60 write requests per
    # minute per user", eso agota la cuota rápido — es justo lo que se vio
    # en logs reales: reintentos de 20/40/80s en 'limpiar pestaña' seguidos
    # de un timeout, en un archivo que ni siquiera tenía cambios que
    # justificaran la escritura. Sacar este refresh innecesario resuelve el
    # problema de raíz en vez de solo reintentar más ante la cuota agotada.
    print("   ⏭️  Sin cambios detectados — snapshot no necesita actualizarse")
    return {"estado": "sin_cambios", "archivo": nombre_archivo}


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

    try:
        carpeta_snapshots_id = obtener_carpeta_snapshots(drive_svc)
    except Exception as e:
        print(f"   ❌ No se pudo buscar la carpeta de snapshots (error de red persistente "
              f"tras varios reintentos): {e}")
        return
    if not carpeta_snapshots_id:
        print("   ❌ Carpeta de snapshots no encontrada — correr snapshot_bot.py al menos "
              "una vez para crearla")
        return
    print(f"   📁 Carpeta snapshots: {carpeta_snapshots_id}")

    print(f"\n📂 Buscando archivos en carpeta {CONFIG['CARPETA_REPARTICIONES_ID']}...")
    archivos = obtener_archivos(drive_svc, CONFIG["CARPETA_REPARTICIONES_ID"])
    if not archivos:
        print("   ❌ No se encontraron archivos Excel")
        return
    print(f"   ✅ Archivos Excel encontrados: {len(archivos)}")

    procesados, con_cambios, errores = 0, 0, 0
    errores_lista = []
    # Archivos que "procesaron" sin excepción pero cuya comparación se
    # omitió por algún motivo (descarga fallida, error de red buscando el
    # snapshot, sin Sheets service, etc. — ver _procesar_archivo_impl).
    # Antes esto quedaba contado en silencio dentro de "Sin cambios",
    # porque tanto "comparé y no había cambios" como "no pude comparar"
    # devolvían None. Se separa explícitamente para que un día con varias
    # descargas fallidas no se vea como una corrida 100% en verde.
    sin_comparar_lista = []

    for i, archivo in enumerate(archivos, 1):
        print(f"\n{'='*60}")
        print(f"[{i}/{len(archivos)}] Procesando...")
        print(f"{'='*60}")
        try:
            resultado = procesar_archivo(archivo, carpeta_snapshots_id, drive_svc, sheets_svc_registro)
            procesados += 1
            estado = resultado.get("estado") if resultado else None
            if estado == "con_cambios":
                con_cambios += 1
            elif estado == "sin_comparar":
                sin_comparar_lista.append({
                    "archivo": archivo["name"],
                    "motivo": resultado.get("motivo", "desconocido"),
                })
        except Exception as e:
            print(f"   ❌ Error procesando {archivo['name']}: {e}")
            traceback.print_exc()
            errores += 1
            errores_lista.append(archivo["name"])

    sin_comparar = len(sin_comparar_lista)
    sin_cambios = procesados - con_cambios - errores - sin_comparar

    duracion = time.time() - inicio
    print(f"\n{'='*60}")
    print("📊 RESUMEN FINAL")
    print(f"{'='*60}")
    print(f"📁 Archivos procesados: {procesados}")
    print(f"📝 Con cambios: {con_cambios}")
    print(f"⏭️  Sin cambios: {sin_cambios}")
    print(f"🚫 Sin comparar (no se pudo comparar esta vez): {sin_comparar}")
    if sin_comparar_lista:
        print("   Archivos sin comparar:")
        for item in sin_comparar_lista:
            print(f"     ⚠️ {item['archivo']} (motivo: {item['motivo']})")
    print(f"❌ Errores: {errores}")
    if errores_lista:
        print("   Archivos con error:")
        for e in errores_lista:
            print(f"     ⚠️ {e}")
    print(f"⏱️  Tiempo total: {duracion:.0f}s ({duracion/60:.1f} min)")
    print(f"{'='*60}")


if __name__ == "__main__":
    ejecutar_principal()