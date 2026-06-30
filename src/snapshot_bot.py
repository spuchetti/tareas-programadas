"""
Bot de construcción de Snapshots para Monitoreo de Liquidaciones

Descarga cada .xlsx de la carpeta de reparticiones y lo sube como
Google Sheets en la carpeta de snapshots. Los SNAPs ya existentes
se saltean automáticamente.

======= EJECUCIÓN =======
Correr manualmente desde GitHub Actions → workflow_dispatch
O bien: python src/snapshot_bot.py
"""

import io
import json
import os
import sys
import time
import traceback

from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload, MediaIoBaseDownload
from googleapiclient.errors import HttpError
from google.oauth2.service_account import Credentials

sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from utils.common_utils import registrar_inicio, registrar_resumen

# ---------------------------------------------------------------------------
# Configuración
# ---------------------------------------------------------------------------

CARPETA_XLSX_ID    = "1_Xb2jrtr3Sjwi8-2nhT2k53KZ6CLE5hJ"   # Reparticiones
CARPETA_INTERNA_ID = "1XJj3pMySybGeK7cW5-PRFPf1q5w2Dch5"   # Carpeta interna OSER
SNAP_FOLDER_NAME   = "_snapshots_liquidaciones"
SNAP_PREFIX        = "[SNAP] "

INTENTOS_MAX    = 3
ESPERA_REINTENTO = 6   # segundos entre reintentos de subida
PAUSA_ENTRE_ARCH = 2   # segundos entre archivos (evita rate limit)


# ---------------------------------------------------------------------------
# Drive helpers
# ---------------------------------------------------------------------------

def inicializar_drive_con_scopes():
    """Inicializa Drive con scopes completos (lectura + escritura)."""
    try:
        cfg   = json.loads(os.getenv("GDRIVE_JSON"))
        creds = Credentials.from_service_account_info(
            cfg,
            scopes=[
                "https://www.googleapis.com/auth/drive",
                "https://www.googleapis.com/auth/spreadsheets",
            ],
        )
        return build("drive", "v3", credentials=creds, cache_discovery=False)
    except Exception as e:
        print(f"❌ Error iniciando Drive: {e}")
        traceback.print_exc()
        return None


def listar_archivos(drive, carpeta_id, solo_xlsx=True):
    """Devuelve todos los archivos de una carpeta (con paginación)."""
    mime_xlsx = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

    if solo_xlsx:
        q = (
            f"'{carpeta_id}' in parents and trashed=false and ("
            f"name contains '.xlsx' or mimeType='{mime_xlsx}')"
        )
    else:
        q = f"'{carpeta_id}' in parents and trashed=false"

    archivos, page_token = [], None
    while True:
        res = drive.files().list(
            q=q,
            pageSize=200,
            fields="nextPageToken, files(id, name, mimeType)",
            supportsAllDrives=True,
            includeItemsFromAllDrives=True,
            pageToken=page_token,
        ).execute()
        archivos.extend(res.get("files", []))
        page_token = res.get("nextPageToken")
        if not page_token:
            break
    return archivos


def obtener_o_crear_carpeta_snaps(drive):
    """Busca la carpeta de snapshots dentro de CARPETA_INTERNA_ID; la crea si no existe."""
    res = drive.files().list(
        q=(
            f"'{CARPETA_INTERNA_ID}' in parents "
            f"and name='{SNAP_FOLDER_NAME}' "
            f"and mimeType='application/vnd.google-apps.folder' "
            f"and trashed=false"
        ),
        fields="files(id)",
        supportsAllDrives=True,
        includeItemsFromAllDrives=True,
    ).execute()

    archivos = res.get("files", [])
    if archivos:
        folder_id = archivos[0]["id"]
        print(f"📁 Carpeta snapshots encontrada: {folder_id}")
        return folder_id

    nueva = drive.files().create(
        body={
            "name": SNAP_FOLDER_NAME,
            "mimeType": "application/vnd.google-apps.folder",
            "parents": [CARPETA_INTERNA_ID],
        },
        fields="id",
        supportsAllDrives=True,
    ).execute()
    print(f"📁 Carpeta snapshots creada: {nueva['id']}")
    return nueva["id"]


def listar_snaps_existentes(drive, snap_folder_id):
    """Devuelve un set con los nombres de SNAPs ya creados."""
    archivos = listar_archivos(drive, snap_folder_id, solo_xlsx=False)
    return {a["name"] for a in archivos}


def descargar_bytes(drive, file_id, mime_type):
    """Descarga un archivo y devuelve BytesIO."""
    try:
        if mime_type == "application/vnd.google-apps.spreadsheet":
            req = drive.files().export_media(
                fileId=file_id,
                mimeType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        else:
            req = drive.files().get_media(fileId=file_id)

        fh = io.BytesIO()
        downloader = MediaIoBaseDownload(fh, req)
        done = False
        while not done:
            _, done = downloader.next_chunk()
        fh.seek(0)
        return fh
    except Exception as e:
        print(f"   ❌ Error descargando: {e}")
        return None


def subir_como_gsheet(drive, fh, nombre_snap, snap_folder_id):
    """
    Sube el BytesIO como Google Sheets (conversión automática de Drive).
    Reintenta hasta INTENTOS_MAX veces ante errores de rate limit.
    """
    fh.seek(0)
    media = MediaIoBaseUpload(
        fh,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        resumable=True,
    )
    for intento in range(INTENTOS_MAX):
        try:
            resultado = drive.files().create(
                body={
                    "name": nombre_snap,
                    "mimeType": "application/vnd.google-apps.spreadsheet",
                    "parents": [snap_folder_id],
                },
                media_body=media,
                fields="id, name",
                supportsAllDrives=True,
            ).execute()
            return resultado["id"]
        except HttpError as e:
            if e.resp.status in (403, 429, 500, 503):
                espera = ESPERA_REINTENTO * (intento + 1)
                print(f"   ⏳ Error {e.resp.status}, reintento {intento+1}/{INTENTOS_MAX} en {espera}s...")
                time.sleep(espera)
            else:
                raise
    return None


# ---------------------------------------------------------------------------
# Principal
# ---------------------------------------------------------------------------

def ejecutar_principal():
    inicio = time.time()
    ahora  = registrar_inicio("BOT SNAPSHOT BUILDER - Monitoreo de Liquidaciones")

    drive = inicializar_drive_con_scopes()
    if not drive:
        return

    # 1. Carpeta de snapshots
    snap_folder_id = obtener_o_crear_carpeta_snaps(drive)

    # 2. SNAPs ya existentes
    snaps_existentes = listar_snaps_existentes(drive, snap_folder_id)
    print(f"✅ SNAPs existentes: {len(snaps_existentes)}")

    # 3. Archivos .xlsx a procesar
    archivos = listar_archivos(drive, CARPETA_XLSX_ID, solo_xlsx=True)
    print(f"📊 Archivos .xlsx encontrados: {len(archivos)}\n")

    procesados, saltados, errores = 0, 0, 0
    lista_errores = []

    for i, archivo in enumerate(archivos, 1):
        nombre_base = archivo["name"].replace(".xlsx", "").replace(".XLSX", "")
        nombre_snap = f"{SNAP_PREFIX}{nombre_base}"

        print(f"[{i}/{len(archivos)}] {archivo['name']}")

        # Saltear si ya existe
        if nombre_snap in snaps_existentes:
            print(f"   ⏭️  SNAP ya existe, saltando.")
            saltados += 1
            continue

        # Descargar
        fh = descargar_bytes(drive, archivo["id"], archivo["mimeType"])
        if not fh:
            print(f"   ❌ No se pudo descargar.")
            errores += 1
            lista_errores.append(archivo["name"])
            continue

        # Subir como Google Sheets
        snap_id = subir_como_gsheet(drive, fh, nombre_snap, snap_folder_id)
        if snap_id:
            print(f"   ✅ SNAP creado ({snap_id})")
            snaps_existentes.add(nombre_snap)
            procesados += 1
        else:
            print(f"   ❌ Falló la subida tras {INTENTOS_MAX} intentos.")
            errores += 1
            lista_errores.append(archivo["name"])

        # Pausa entre archivos para no saturar la API
        if i < len(archivos):
            time.sleep(PAUSA_ENTRE_ARCH)

    # Resumen consola
    duracion = time.time() - inicio
    print(f"\n{'='*60}")
    print(f"✅ Creados:      {procesados}")
    print(f"⏭️  Ya existían:  {saltados}")
    print(f"❌ Errores:      {errores}")
    print(f"⏱️  Tiempo:       {duracion:.0f}s ({duracion/60:.1f} min)")
    print(f"{'='*60}")

    if lista_errores:
        print("\nArchivos con error:")
        for e in lista_errores:
            print(f"  ⚠️  {e}")

    # Log de resumen
    registrar_resumen(inicio, procesados, len(archivos))
    print(f"\n📝 Resumen registrado: {procesados} SNAPs creados, {errores} errores, {saltados} saltados")


if __name__ == "__main__":
    ejecutar_principal()