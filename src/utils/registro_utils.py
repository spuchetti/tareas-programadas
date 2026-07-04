"""
Registro de agentes en Google Sheets.

Mantiene el mismo esquema que el Apps Script:
  - Una o más planillas  _registro_agentes_N  en CARPETA_INTERNA_ID
  - Una hoja por repartición (nombre del archivo sin .xlsx)
  - Columnas: ID | CUIL | DNI | NOMBRE | FECHA_ALTA | ULTIMA_VEZ
  - Cache en memoria por (spreadsheet_id, nombre_hoja) para evitar llamadas redundantes

El ID de cada planilla de registro se persiste en un archivo local
  /tmp/monitoreo_registro_ids.json
(En GitHub Actions el runner es efímero, así que la primera ejecución
del día lo descubre desde Drive si el archivo no existe.)

--------------------------------------------------------------------------
IMPORTANTE — separación Service Account / OAuth
--------------------------------------------------------------------------
No tenemos Unidad Compartida, así que la Service Account (GDRIVE_JSON) NO
puede crear archivos nuevos en Drive — solo puede leer, y editar contenido
de archivos que ya existen (agregar una hoja/tab dentro de una planilla
existente sí es "editar contenido", no "crear archivo", así que la SA
puede hacerlo sin problema).

Por eso este módulo separa dos responsabilidades:

  · obtener_o_crear_hoja_registro()  → uso normal, con SA. Busca una
    planilla _registro_agentes_N con lugar y le agrega la hoja (tab) que
    falte. NUNCA crea una planilla nueva. Si no encuentra ninguna con
    lugar, devuelve None y loguea que hace falta correr snapshot_bot.

  · verificar_y_ampliar_capacidad() → SOLO debe llamarse desde
    snapshot_bot.py, pasándole servicios armados con OAuth. Es la única
    función de este módulo que puede crear una planilla _registro_agentes_N
    nueva (bootstrap inicial, o por adelantado al acercarse al límite de
    150 hojas).
"""

import json
import os
import re
import time
import traceback
import unicodedata
from datetime import datetime
from zoneinfo import ZoneInfo

from utils.auth_utils import (
    obtener_sheets_service_sa,
    obtener_drive_service_sa,
    obtener_email_service_account,
    CredencialesFaltantesError,
)

# ---------------------------------------------------------------------------
# Config
# ---------------------------------------------------------------------------

CARPETA_INTERNA_ID  = "1XJj3pMySybGeK7cW5-PRFPf1q5w2Dch5"
NOMBRE_REGISTRO     = "_registro_agentes"
MAX_HOJAS_POR_PLANILLA = 150
UMBRAL_AMPLIACION   = 140  # crear la siguiente planilla al llegar acá, no esperar al límite
IDS_CACHE_PATH      = "/tmp/monitoreo_registro_ids.json"

TZ_AR = ZoneInfo("America/Argentina/Buenos_Aires")

# Cache en memoria: { "spreadsheet_id__nombre_hoja": { porCuil, porDni, porNombre, ultimoId } }
_cache_registro: dict = {}

# ---------------------------------------------------------------------------
# Inicialización del servicio (Service Account — uso normal, sin crear archivos)
# ---------------------------------------------------------------------------

def inicializar_sheets():
    """Sheets service con la Service Account. Para uso normal (lectura/append)."""
    try:
        return obtener_sheets_service_sa()
    except CredencialesFaltantesError as e:
        print(f"❌ {e}")
        return None
    except Exception as e:
        print(f"❌ Error iniciando Sheets: {e}")
        traceback.print_exc()
        return None


def _inicializar_drive_registro():
    """Drive service con la Service Account. Para uso normal (búsqueda/lectura)."""
    return obtener_drive_service_sa()


# ---------------------------------------------------------------------------
# Persistencia de IDs de planillas de registro
# ---------------------------------------------------------------------------

def _cargar_ids_guardados():
    if os.path.exists(IDS_CACHE_PATH):
        try:
            with open(IDS_CACHE_PATH, "r") as f:
                return json.load(f)
        except Exception:
            pass
    return []


def _guardar_ids(ids):
    with open(IDS_CACHE_PATH, "w") as f:
        json.dump(ids, f)


def _descubrir_planillas_desde_drive(drive):
    """Busca planillas _registro_agentes_N en CARPETA_INTERNA_ID."""
    q = (
        f"'{CARPETA_INTERNA_ID}' in parents and trashed=false "
        f"and name contains '{NOMBRE_REGISTRO}' "
        f"and mimeType='application/vnd.google-apps.spreadsheet'"
    )
    res = drive.files().list(
        q=q, pageSize=50, fields="files(id, name)",
        supportsAllDrives=True, includeItemsFromAllDrives=True,
    ).execute()
    return [a["id"] for a in res.get("files", [])]


def _ids_actualizados(sheets_svc, drive_svc=None):
    """Carga el listado de IDs de planillas, descubriéndolo desde Drive si hace falta."""
    ids = _cargar_ids_guardados()
    if not ids:
        drive = drive_svc or _inicializar_drive_registro()
        ids = _descubrir_planillas_desde_drive(drive)
        if ids:
            _guardar_ids(ids)
    return ids


def _cantidad_hojas(sheets_svc, spreadsheet_id):
    try:
        meta = sheets_svc.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
        return len(meta.get("sheets", []))
    except Exception:
        return None


# ---------------------------------------------------------------------------
# Búsqueda de planilla con lugar (SOLO LECTURA — apta para Service Account)
# ---------------------------------------------------------------------------

def _buscar_planilla_con_espacio(sheets_svc, ids):
    """
    Devuelve el spreadsheet_id de la primera planilla con lugar (< 150
    hojas), o None si no hay ninguna. NO crea nada — es seguro llamarla
    con la Service Account.
    """
    for sid in ids:
        n_hojas = _cantidad_hojas(sheets_svc, sid)
        if n_hojas is not None and n_hojas < MAX_HOJAS_POR_PLANILLA:
            return sid
    return None


# ---------------------------------------------------------------------------
# Obtener hoja de registro para un archivo (Service Account — sin crear planillas)
# ---------------------------------------------------------------------------

def obtener_o_crear_hoja_registro(sheets_svc, nombre_archivo):
    """
    Devuelve un dict con toda la info necesaria para trabajar con la hoja:
      { "spreadsheet_id": str, "nombre_hoja": str, "sheets_svc": obj }

    Puede agregar una hoja (tab) nueva dentro de una planilla YA EXISTENTE
    (eso es edición de contenido, la SA lo puede hacer). Lo que NO hace es
    crear una planilla _registro_agentes_N nueva — eso es responsabilidad
    exclusiva de verificar_y_ampliar_capacidad() (OAuth, vía snapshot_bot).

    Si no hay ninguna planilla con lugar disponible, devuelve None y
    loguea un mensaje claro en vez de intentar crear un archivo (lo que
    fallaría con la Service Account).
    """
    nombre_hoja = nombre_archivo.replace(".xlsx", "").replace(".XLSX", "")[:31]
    ids = _ids_actualizados(sheets_svc)

    if not ids:
        print("⚠️  Todavía no existe ninguna planilla _registro_agentes — "
              "correr snapshot_bot para crearla.")
        return None

    # 1. ¿La hoja ya existe en alguna planilla?
    for sid in ids:
        try:
            meta = sheets_svc.spreadsheets().get(spreadsheetId=sid).execute()
            hojas = [s["properties"]["title"] for s in meta.get("sheets", [])]
            if nombre_hoja in hojas:
                return {"spreadsheet_id": sid, "nombre_hoja": nombre_hoja, "sheets_svc": sheets_svc}
        except Exception:
            pass

    # 2. No existe todavía: buscar una planilla con lugar (solo lectura)
    sid = _buscar_planilla_con_espacio(sheets_svc, ids)
    if sid is None:
        print("⚠️  Capacidad de registro agotada en todas las planillas "
              "existentes — correr snapshot_bot para ampliar capacidad. "
              "Se continúa sin registro numérico para este archivo "
              "(fallback a DNI/CUIL como identificador).")
        return None

    # 3. Agregar la hoja (tab) dentro de la planilla existente — no crea archivo
    sheets_svc.spreadsheets().batchUpdate(
        spreadsheetId=sid,
        body={"requests": [{"addSheet": {"properties": {"title": nombre_hoja}}}]},
    ).execute()

    # Agregar encabezados
    sheets_svc.spreadsheets().values().update(
        spreadsheetId=sid,
        range=f"'{nombre_hoja}'!A1:F1",
        valueInputOption="RAW",
        body={"values": [["ID", "CUIL", "DNI", "NOMBRE", "FECHA_ALTA", "ULTIMA_VEZ"]]},
    ).execute()

    print(f"  → Hoja de registro creada: {nombre_hoja}")
    return {"spreadsheet_id": sid, "nombre_hoja": nombre_hoja, "sheets_svc": sheets_svc}


# ---------------------------------------------------------------------------
# Ampliación de capacidad (OAuth — SOLO desde snapshot_bot.py)
# ---------------------------------------------------------------------------

def verificar_y_ampliar_capacidad(oauth_sheets_svc, oauth_drive_svc):
    """
    Garantiza que exista al menos una planilla _registro_agentes_N con
    lugar disponible, creándola por adelantado si hace falta.

    SOLO debe llamarse desde snapshot_bot.py, pasándole servicios armados
    con OAuth (services.aportes.oser@gmail.com) — es la única credencial
    del proyecto habilitada para crear archivos nuevos en Drive.

    Se llama al principio de cada corrida manual de snapshot_bot, antes de
    procesar snapshots, así el resto de los bots (que usan Service Account)
    siempre encuentran una planilla con lugar y nunca necesitan crear una.
    """
    ids = _ids_actualizados(oauth_sheets_svc, drive_svc=oauth_drive_svc)

    capacidades = [(sid, _cantidad_hojas(oauth_sheets_svc, sid)) for sid in ids]
    capacidades = [(sid, n) for sid, n in capacidades if n is not None]

    hay_espacio = any(n < UMBRAL_AMPLIACION for _, n in capacidades)

    if ids and hay_espacio:
        detalle = ", ".join(f"{sid[:8]}…({n}/{MAX_HOJAS_POR_PLANILLA})" for sid, n in capacidades)
        print(f"✓ Capacidad de registro OK — {detalle}")
        return

    if not ids:
        print("📋 No existe ninguna planilla _registro_agentes — creando la primera (bootstrap)...")
    else:
        detalle = ", ".join(f"{sid[:8]}…({n}/{MAX_HOJAS_POR_PLANILLA})" for sid, n in capacidades)
        print(f"📋 Todas las planillas están cerca del límite ({detalle}) — creando una nueva por adelantado...")

    numero = len(ids) + 1
    nueva = oauth_sheets_svc.spreadsheets().create(body={
        "properties": {"title": f"{NOMBRE_REGISTRO}_{numero}"},
        "sheets": [{"properties": {"title": "_info"}}],
    }).execute()
    sid = nueva["spreadsheetId"]

    # Mover a CARPETA_INTERNA_ID (que ya está compartida con la Service
    # Account como Editor, así que el archivo hereda ese acceso).
    try:
        file_meta = oauth_drive_svc.files().get(fileId=sid, fields="parents").execute()
        padres_actuales = ",".join(file_meta.get("parents", []))
        oauth_drive_svc.files().update(
            fileId=sid,
            addParents=CARPETA_INTERNA_ID,
            removeParents=padres_actuales,
            supportsAllDrives=True,
            fields="id, parents",
        ).execute()
    except Exception as e:
        print(f"  ⚠️  No se pudo mover la planilla a la carpeta interna: {e}")

    # Refuerzo explícito: compartir directamente con la Service Account por
    # si la herencia de permisos de la carpeta no alcanzara (por ejemplo,
    # si la carpeta usa permisos "solo esta carpeta" en vez de heredar a
    # los archivos que se agregan después). No debería hacer falta, pero
    # es barato y evita un fallo silencioso.
    sa_email = obtener_email_service_account()
    if sa_email:
        try:
            oauth_drive_svc.permissions().create(
                fileId=sid,
                body={"type": "user", "role": "writer", "emailAddress": sa_email},
                fields="id",
                supportsAllDrives=True,
            ).execute()
            print(f"  ✓ Planilla compartida explícitamente con la Service Account ({sa_email})")
        except Exception as e:
            print(f"  ⚠️  No se pudo compartir explícitamente con la Service Account: {e}")

    ids.append(sid)
    _guardar_ids(ids)
    print(f"  ✓ Planilla de registro #{numero} creada por adelantado ({sid})")
    return sid


# ---------------------------------------------------------------------------
# Normalización
# ---------------------------------------------------------------------------

def _normalizar(s):
    t = str(s or "").upper()
    t = unicodedata.normalize("NFD", t)
    t = "".join(c for c in t if unicodedata.category(c) != "Mn")
    return " ".join(t.split())


def _limpiar_num(s):
    return re.sub(r"[^0-9]", "", str(s or ""))


# ---------------------------------------------------------------------------
# Cargar caché de una hoja de registro
# ---------------------------------------------------------------------------

def _cargar_cache(hoja_info):
    sid   = hoja_info["spreadsheet_id"]
    nhoja = hoja_info["nombre_hoja"]
    svc   = hoja_info["sheets_svc"]
    clave = f"{sid}__{nhoja}"

    if clave in _cache_registro:
        return _cache_registro[clave]

    cache = {"porCuil": {}, "porDni": {}, "porNombre": {}, "ultimoId": 0, "filas": []}
    try:
        res = svc.spreadsheets().values().get(
            spreadsheetId=sid, range=f"'{nhoja}'!A:F",
        ).execute()
        filas = res.get("values", [])
        for i, fila in enumerate(filas[1:], start=2):   # fila 1 = encabezados
            if len(fila) < 4:
                continue
            aid, cuil, dni, nombre = (fila + ["", "", "", ""])[:4]
            aid_num = int(aid) if str(aid).isdigit() else 0
            if aid_num > cache["ultimoId"]:
                cache["ultimoId"] = aid_num
            entrada = {"id": aid_num, "fila_sheet": i}
            cuil_l  = _limpiar_num(cuil)
            dni_l   = _limpiar_num(dni)
            nombre_n = _normalizar(nombre)
            if cuil_l:   cache["porCuil"][cuil_l]     = entrada
            if dni_l:    cache["porDni"][dni_l]        = entrada
            if nombre_n: cache["porNombre"][nombre_n]  = entrada
            cache["filas"].append({"id": aid_num, "cuil": cuil_l, "dni": dni_l, "nombre": nombre_n})
    except Exception as e:
        print(f"  ⚠️  Error cargando caché de registro: {e}")

    _cache_registro[clave] = cache
    return cache


# ---------------------------------------------------------------------------
# Obtener o crear ID de agente
# ---------------------------------------------------------------------------
#
# IMPORTANTE — por qué esto NO llama a la API en cada agente:
# Un archivo puede tener cientos de agentes. Si cada uno dispara su propia
# llamada a la Sheets API (un append o un update), se choca rápido contra
# el límite de 60 escrituras por minuto por usuario de la API ("Quota
# exceeded... Write requests per minute"). Por eso acá solo se actualiza el
# caché en memoria (instantáneo, sin red) y se ENCOLAN los cambios; recién
# se escriben todos juntos en una sola llamada `append` (agentes nuevos) y
# una sola `batchUpdate` (agentes existentes) cuando se llama a
# flush_registro_pendientes() — una vez por archivo procesado, no una vez
# por agente.

_pending_nuevos: dict = {}          # clave -> [[id, cuil, dni, nombre, alta, ultima_vez], ...]
_pending_actualizaciones: dict = {}  # clave -> [(fila_sheet, [cuil, dni, nombre, "", ultima_vez]), ...]


def obtener_id_agente(cuil, dni, nombre, hoja_info):
    """
    Busca el agente en el caché/hoja. Si no existe, lo agrega al caché y a
    la cola de escritura pendiente (NO llama a la API acá — ver
    flush_registro_pendientes). Devuelve el ID entero del agente.
    hoja_info puede ser None (en ese caso devuelve dni o cuil como clave de texto).
    """
    if hoja_info is None:
        return dni or cuil or nombre

    cache   = _cargar_cache(hoja_info)
    sid     = hoja_info["spreadsheet_id"]
    nhoja   = hoja_info["nombre_hoja"]
    clave_c = f"{sid}__{nhoja}"

    cuil_l   = _limpiar_num(cuil)
    dni_l    = _limpiar_num(dni)
    nombre_n = _normalizar(nombre)

    # Buscar en caché
    entrada = (
        cache["porCuil"].get(cuil_l)
        or cache["porDni"].get(dni_l)
        or (cache["porNombre"].get(nombre_n) if nombre_n and len(nombre_n) > 3 else None)
    )

    ahora = datetime.now(TZ_AR).strftime("%d/%m/%Y %H:%M")

    if entrada:
        # Encolar actualización de ULTIMA_VEZ (se escribe recién en el flush)
        _pending_actualizaciones.setdefault(clave_c, []).append(
            (entrada["fila_sheet"], [cuil or "", dni or "", nombre or "", "", ahora])
        )
        # Actualizar caché local
        if cuil_l:   cache["porCuil"][cuil_l]     = entrada
        if dni_l:    cache["porDni"][dni_l]        = entrada
        if nombre_n: cache["porNombre"][nombre_n]  = entrada
        return entrada["id"]

    # Nuevo agente
    nuevo_id           = cache["ultimoId"] + 1
    cache["ultimoId"]  = nuevo_id
    nueva_fila_num     = len(cache["filas"]) + 2   # +2: 1 encabezado + 1-based

    _pending_nuevos.setdefault(clave_c, []).append(
        [nuevo_id, cuil or "", dni or "", nombre or "", ahora, ahora]
    )

    nueva_entrada = {"id": nuevo_id, "fila_sheet": nueva_fila_num}
    if cuil_l:   cache["porCuil"][cuil_l]     = nueva_entrada
    if dni_l:    cache["porDni"][dni_l]        = nueva_entrada
    if nombre_n: cache["porNombre"][nombre_n]  = nueva_entrada
    cache["filas"].append({"id": nuevo_id, "cuil": cuil_l, "dni": dni_l, "nombre": nombre_n})
    _cache_registro[clave_c] = cache
    return nuevo_id


def flush_registro_pendientes(hoja_info):
    """
    Escribe en la Sheets API, en lote, todos los agentes nuevos y las
    actualizaciones de "última vez" acumulados para esta hoja_info desde
    la última vez que se llamó a esta función.

    Hay que llamarla una vez por archivo procesado (no por agente) —
    normalmente al final de procesar_archivo(), incluso si hubo un error a
    mitad de camino, para no perder lo que ya se acumuló.
    """
    if hoja_info is None:
        return

    sid   = hoja_info["spreadsheet_id"]
    nhoja = hoja_info["nombre_hoja"]
    svc   = hoja_info["sheets_svc"]
    clave = f"{sid}__{nhoja}"

    nuevos = _pending_nuevos.pop(clave, [])
    actualizaciones = _pending_actualizaciones.pop(clave, [])

    if nuevos:
        try:
            _ejecutar_con_reintentos(
                svc.spreadsheets().values().append(
                    spreadsheetId=sid,
                    range=f"'{nhoja}'!A:F",
                    valueInputOption="RAW",
                    insertDataOption="INSERT_ROWS",
                    body={"values": nuevos},
                ),
                f"agregar {len(nuevos)} agente(s) nuevo(s) en {nhoja}",
            )
        except Exception as e:
            print(f"  ⚠️  Error guardando {len(nuevos)} agente(s) nuevo(s) en bloque: {e}")

    if actualizaciones:
        try:
            data = [
                {"range": f"'{nhoja}'!B{fila}:F{fila}", "values": [valores]}
                for fila, valores in actualizaciones
            ]
            _ejecutar_con_reintentos(
                svc.spreadsheets().values().batchUpdate(
                    spreadsheetId=sid,
                    body={"valueInputOption": "RAW", "data": data},
                ),
                f"actualizar {len(actualizaciones)} agente(s) existente(s) en {nhoja}",
            )
        except Exception as e:
            print(f"  ⚠️  Error actualizando {len(actualizaciones)} agente(s) existente(s) en bloque: {e}")


def _ejecutar_con_reintentos(request, descripcion, intentos_max=4, espera_base=20):
    """
    Ejecuta un request de la Sheets API reintentando con backoff si choca
    con el límite de cuota (429 / RATE_LIMIT_EXCEEDED). La cuota es "por
    minuto", así que la espera empieza en ~20s y se va duplicando.
    """
    for intento in range(intentos_max):
        try:
            return request.execute()
        except Exception as e:
            es_quota = "429" in str(e) or "RATE_LIMIT_EXCEEDED" in str(e) or "Quota exceeded" in str(e)
            if es_quota and intento < intentos_max - 1:
                espera = espera_base * (2 ** intento)
                print(f"  ⏳ Límite de cuota de Sheets API, reintentando '{descripcion}' en {espera}s "
                      f"({intento + 1}/{intentos_max})...")
                time.sleep(espera)
                continue
            raise