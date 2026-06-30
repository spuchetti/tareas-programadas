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
"""

import json
import os
import re
import time
import traceback
import unicodedata
from datetime import datetime
from zoneinfo import ZoneInfo

from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from google.oauth2.service_account import Credentials

# ---------------------------------------------------------------------------
# Config
# ---------------------------------------------------------------------------

CARPETA_INTERNA_ID  = "1XJj3pMySybGeK7cW5-PRFPf1q5w2Dch5"
NOMBRE_REGISTRO     = "_registro_agentes"
MAX_HOJAS_POR_PLANILLA = 150
IDS_CACHE_PATH      = "/tmp/monitoreo_registro_ids.json"

TZ_AR = ZoneInfo("America/Argentina/Buenos_Aires")

# Cache en memoria: { "spreadsheet_id__nombre_hoja": { porCuil, porDni, porNombre, ultimoId } }
_cache_registro: dict = {}

# ---------------------------------------------------------------------------
# Inicialización del servicio
# ---------------------------------------------------------------------------

def inicializar_sheets():
    try:
        cfg   = json.loads(os.getenv("GDRIVE_JSON"))
        creds = Credentials.from_service_account_info(
            cfg,
            scopes=[
                "https://www.googleapis.com/auth/drive",
                "https://www.googleapis.com/auth/spreadsheets",
            ],
        )
        return build("sheets", "v4", credentials=creds, cache_discovery=False)
    except Exception as e:
        print(f"❌ Error iniciando Sheets: {e}")
        traceback.print_exc()
        return None


def _inicializar_drive_registro():
    cfg   = json.loads(os.getenv("GDRIVE_JSON"))
    creds = Credentials.from_service_account_info(
        cfg,
        scopes=[
            "https://www.googleapis.com/auth/drive",
            "https://www.googleapis.com/auth/spreadsheets",
        ],
    )
    return build("drive", "v3", credentials=creds, cache_discovery=False)


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


# ---------------------------------------------------------------------------
# Obtener o crear planilla con capacidad
# ---------------------------------------------------------------------------

def _obtener_o_crear_planilla(sheets_svc):
    """Devuelve (spreadsheet_id, sheets_svc, drive_svc) de la planilla con capacidad."""
    drive = _inicializar_drive_registro()
    ids   = _cargar_ids_guardados()

    # Descubrir desde Drive si el caché local está vacío
    if not ids:
        ids = _descubrir_planillas_desde_drive(drive)
        if ids:
            _guardar_ids(ids)

    # Buscar planilla con capacidad
    for sid in ids:
        try:
            meta = sheets_svc.spreadsheets().get(spreadsheetId=sid).execute()
            if len(meta.get("sheets", [])) < MAX_HOJAS_POR_PLANILLA:
                return sid
        except Exception:
            pass

    # Crear nueva planilla
    numero = len(ids) + 1
    nueva  = sheets_svc.spreadsheets().create(body={
        "properties": {"title": f"{NOMBRE_REGISTRO}_{numero}"},
        "sheets": [{"properties": {"title": "_info"}}],
    }).execute()
    sid = nueva["spreadsheetId"]

    # Mover a CARPETA_INTERNA_ID
    try:
        file_meta = drive.files().get(fileId=sid, fields="parents").execute()
        padres_actuales = ",".join(file_meta.get("parents", []))
        drive.files().update(
            fileId=sid,
            addParents=CARPETA_INTERNA_ID,
            removeParents=padres_actuales,
            supportsAllDrives=True,
            fields="id, parents",
        ).execute()
    except Exception as e:
        print(f"  ⚠️  No se pudo mover planilla a carpeta interna: {e}")

    ids.append(sid)
    _guardar_ids(ids)
    print(f"  ✓ Planilla de registro #{numero} creada ({sid})")
    return sid


# ---------------------------------------------------------------------------
# Obtener o crear hoja de registro para un archivo
# ---------------------------------------------------------------------------

def obtener_o_crear_hoja_registro(sheets_svc, nombre_archivo):
    """
    Devuelve un dict con toda la info necesaria para trabajar con la hoja:
      { "spreadsheet_id": str, "nombre_hoja": str, "sheets_svc": obj }
    """
    nombre_hoja = nombre_archivo.replace(".xlsx", "").replace(".XLSX", "")[:31]
    ids         = _cargar_ids_guardados()

    if not ids:
        drive = _inicializar_drive_registro()
        ids   = _descubrir_planillas_desde_drive(drive)
        if ids:
            _guardar_ids(ids)

    # Buscar si la hoja ya existe en alguna planilla
    for sid in ids:
        try:
            meta   = sheets_svc.spreadsheets().get(spreadsheetId=sid).execute()
            hojas  = [s["properties"]["title"] for s in meta.get("sheets", [])]
            if nombre_hoja in hojas:
                return {"spreadsheet_id": sid, "nombre_hoja": nombre_hoja, "sheets_svc": sheets_svc}
        except Exception:
            pass

    # Crear hoja nueva
    sid = _obtener_o_crear_planilla(sheets_svc)
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

def obtener_id_agente(cuil, dni, nombre, hoja_info):
    """
    Busca el agente en el caché/hoja.  Si no existe, lo crea.
    Devuelve el ID entero del agente.
    hoja_info puede ser None (en ese caso devuelve dni o cuil como clave de texto).
    """
    if hoja_info is None:
        return dni or cuil or nombre

    cache   = _cargar_cache(hoja_info)
    sid     = hoja_info["spreadsheet_id"]
    nhoja   = hoja_info["nombre_hoja"]
    svc     = hoja_info["sheets_svc"]
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
        # Actualizar ULTIMA_VEZ en Sheets (fire-and-forget; ignoramos errores)
        try:
            fila_num = entrada["fila_sheet"]
            svc.spreadsheets().values().update(
                spreadsheetId=sid,
                range=f"'{nhoja}'!B{fila_num}:F{fila_num}",
                valueInputOption="RAW",
                body={"values": [[cuil or "", dni or "", nombre or "", "", ahora]]},
            ).execute()
        except Exception:
            pass
        # Actualizar caché local
        if cuil_l:   cache["porCuil"][cuil_l]     = entrada
        if dni_l:    cache["porDni"][dni_l]        = entrada
        if nombre_n: cache["porNombre"][nombre_n]  = entrada
        return entrada["id"]

    # Nuevo agente
    nuevo_id           = cache["ultimoId"] + 1
    cache["ultimoId"]  = nuevo_id
    nueva_fila_num     = len(cache["filas"]) + 2   # +2: 1 encabezado + 1-based

    try:
        svc.spreadsheets().values().append(
            spreadsheetId=sid,
            range=f"'{nhoja}'!A:F",
            valueInputOption="RAW",
            insertDataOption="INSERT_ROWS",
            body={"values": [[nuevo_id, cuil or "", dni or "", nombre or "", ahora, ahora]]},
        ).execute()
    except Exception as e:
        print(f"  ⚠️  Error guardando agente nuevo: {e}")

    nueva_entrada = {"id": nuevo_id, "fila_sheet": nueva_fila_num}
    if cuil_l:   cache["porCuil"][cuil_l]     = nueva_entrada
    if dni_l:    cache["porDni"][dni_l]        = nueva_entrada
    if nombre_n: cache["porNombre"][nombre_n]  = nueva_entrada
    cache["filas"].append({"id": nuevo_id, "cuil": cuil_l, "dni": dni_l, "nombre": nombre_n})
    _cache_registro[clave_c] = cache
    return nuevo_id