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

# Forzar flush de prints para ver logs en tiempo real en GitHub Actions
sys.stdout.reconfigure(line_buffering=True)

from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload, MediaIoBaseDownload
from googleapiclient.errors import HttpError
from google.oauth2.credentials import Credentials
from google.auth.transport.requests import Request

sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from utils.common_utils import registrar_inicio, registrar_resumen

# ---------------------------------------------------------------------------
# Configuración
# ---------------------------------------------------------------------------

CARPETA_XLSX_ID    = "1_Xb2jrtr3Sjwi8-2nhT2k53KZ6CLE5hJ"   # Reparticiones
CARPETA_INTERNA_ID = "1XJj3pMySybGeK7cW5-PRFPf1q5w2Dch5"   # Carpeta interna OSER
SNAP_FOLDER_NAME   = "_snapshots_liquidaciones"
SNAP_PREFIX        = "[SNAP] "

INTENTOS_MAX       = 3
ESPERA_REINTENTO   = 6   # segundos entre reintentos de subida
PAUSA_ENTRE_ARCH   = 2   # segundos entre archivos (evita rate limit)

# ⬇️ MODO PRODUCCIÓN: procesar TODOS los archivos
MODO_PRUEBA          = False   # <--- CAMBIADO A False
MAX_ARCHIVOS_PRUEBA  = 3       # Ya no se usa

# ---------------------------------------------------------------------------
# Drive helpers
# ---------------------------------------------------------------------------

def inicializar_drive_con_scopes():
    """
    Inicializa el servicio de Google Drive usando OAuth 2.0 con refresh token.
    
    Returns:
        googleapiclient.discovery.Resource: Servicio de Drive autenticado
    """
    try:
        token_data = json.loads(os.getenv("OAUTH_REFRESH_TOKEN"))
        
        creds = Credentials(
            token=token_data.get("token"),
            refresh_token=token_data["refresh_token"],
            token_uri=token_data["token_uri"],
            client_id=token_data["client_id"],
            client_secret=token_data["client_secret"],
            scopes=token_data["scopes"]
        )
        
        if creds.expired:
            print("🔄 Refrescando token...", flush=True)
            creds.refresh(Request())
            print("✅ Token refrescado", flush=True)
        
        return build("drive", "v3", credentials=creds, cache_discovery=False)
    
    except Exception as e:
        print(f"❌ Error iniciando Drive: {e}", flush=True)
        traceback.print_exc()
        return None


def listar_archivos(drive, carpeta_id, solo_xlsx=True):
    """
    Lista todos los archivos de una carpeta de Drive con paginación.
    
    Args:
        drive: Servicio de Drive autenticado
        carpeta_id: ID de la carpeta a listar
        solo_xlsx: Si True, solo devuelve archivos .xlsx
    
    Returns:
        list: Lista de archivos con id, name y mimeType
    """
    mime_xlsx = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    
    if solo_xlsx:
        q = (
            f"'{carpeta_id}' in parents and trashed=false and ("
            f"name contains '.xlsx' or mimeType='{mime_xlsx}')"
        )
    else:
        q = f"'{carpeta_id}' in parents and trashed=false"
    
    archivos = []
    page_token = None
    
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
    """
    Busca la carpeta de snapshots dentro de CARPETA_INTERNA_ID.
    Si no existe, la crea.
    
    Returns:
        str: ID de la carpeta de snapshots
    """
    print(f"🔍 Buscando carpeta '{SNAP_FOLDER_NAME}' en {CARPETA_INTERNA_ID}...", flush=True)
    
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
        print(f"📁 Carpeta snapshots encontrada: {folder_id}", flush=True)
        return folder_id
    
    print(f"📁 Carpeta '{SNAP_FOLDER_NAME}' no encontrada, creando...", flush=True)
    
    nueva = drive.files().create(
        body={
            "name": SNAP_FOLDER_NAME,
            "mimeType": "application/vnd.google-apps.folder",
            "parents": [CARPETA_INTERNA_ID],
        },
        fields="id",
        supportsAllDrives=True,
    ).execute()
    
    print(f"📁 Carpeta snapshots creada: {nueva['id']}", flush=True)
    return nueva["id"]


def listar_snaps_existentes(drive, snap_folder_id):
    """
    Devuelve un set con los nombres de SNAPs ya creados.
    
    Returns:
        set: Conjunto de nombres de archivos SNAP existentes
    """
    archivos = listar_archivos(drive, snap_folder_id, solo_xlsx=False)
    return {a["name"] for a in archivos}


def descargar_bytes(drive, file_id, mime_type):
    """
    Descarga un archivo de Drive y devuelve su contenido como BytesIO.
    
    Args:
        drive: Servicio de Drive autenticado
        file_id: ID del archivo a descargar
        mime_type: MIME type del archivo
    
    Returns:
        io.BytesIO: Contenido del archivo o None si falla
    """
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
        print(f"   ❌ Error descargando: {e}", flush=True)
        return None


def subir_como_gsheet(drive, fh, nombre_snap, snap_folder_id):
    """
    Sube un archivo .xlsx como Google Sheets utilizando el enfoque de dos pasos:
    1. Crear el archivo vacío con el mimeType de Google Sheets
    2. Actualizar su contenido con el archivo .xlsx
    
    Este enfoque es más robusto y evita errores de conversión.
    
    Args:
        drive: Servicio de Drive autenticado
        fh: BytesIO con el contenido del archivo .xlsx
        nombre_snap: Nombre que tendrá el archivo en Drive
        snap_folder_id: ID de la carpeta donde se guardará
    
    Returns:
        str: ID del archivo creado o None si falla
    """
    fh.seek(0)
    
    for intento in range(INTENTOS_MAX):
        try:
            # PASO 1: Crear el archivo vacío con mimeType de Google Sheets
            file_metadata = {
                "name": nombre_snap,
                "mimeType": "application/vnd.google-apps.spreadsheet",
                "parents": [snap_folder_id]
            }
            
            print(f"   📄 Creando archivo vacío...", flush=True)
            file = drive.files().create(
                body=file_metadata,
                fields="id",
                supportsAllDrives=True
            ).execute()
            
            file_id = file.get("id")
            print(f"   📄 Archivo creado (ID: {file_id})", flush=True)
            
            # PASO 2: Subir el contenido usando el método update
            fh.seek(0)
            media = MediaIoBaseUpload(
                fh,
                mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                resumable=True
            )
            
            print(f"   ⬆️  Subiendo contenido...", flush=True)
            updated_file = drive.files().update(
                fileId=file_id,
                media_body=media,
                fields="id",
                supportsAllDrives=True
            ).execute()
            
            print(f"   ✅ Contenido subido", flush=True)
            return updated_file.get("id")
        
        except HttpError as e:
            print(f"   ❌ Error {e.resp.status}: {e._get_reason()}", flush=True)
            
            if e.resp.status in (403, 429, 500, 503):
                espera = ESPERA_REINTENTO * (intento + 1)
                print(f"   ⏳ Reintento {intento+1}/{INTENTOS_MAX} en {espera}s...", flush=True)
                time.sleep(espera)
            else:
                # Si es un error 400, no tiene sentido reintentar
                print(f"   📝 Detalle: {e.content}", flush=True)
                return None
    
    return None


# ---------------------------------------------------------------------------
# Principal
# ---------------------------------------------------------------------------

def ejecutar_principal():
    """
    Función principal del bot.
    """
    print("🚀 INICIANDO SNAPSHOT BUILDER", flush=True)
    inicio = time.time()
    ahora = registrar_inicio("BOT SNAPSHOT BUILDER - Monitoreo de Liquidaciones")
    
    print("🔑 Inicializando Drive con OAuth...", flush=True)
    drive = inicializar_drive_con_scopes()
    
    if not drive:
        print("❌ No se pudo inicializar Drive", flush=True)
        return
    
    print("✅ Drive inicializado correctamente", flush=True)
    
    # 1. Obtener o crear carpeta de snapshots
    print("📁 Obteniendo carpeta de snapshots...", flush=True)
    snap_folder_id = obtener_o_crear_carpeta_snaps(drive)
    
    # 2. Listar SNAPs existentes
    print("📋 Listando SNAPs existentes...", flush=True)
    snaps_existentes = listar_snaps_existentes(drive, snap_folder_id)
    print(f"✅ SNAPs existentes: {len(snaps_existentes)}", flush=True)
    
    # 3. Listar archivos .xlsx a procesar
    print(f"📂 Listando archivos en carpeta {CARPETA_XLSX_ID}...", flush=True)
    archivos = listar_archivos(drive, CARPETA_XLSX_ID, solo_xlsx=True)
    print(f"📊 Archivos .xlsx encontrados: {len(archivos)}", flush=True)
    
    # ⬇️ MODO PRODUCCIÓN: sin límite de archivos
    if MODO_PRUEBA and len(archivos) > MAX_ARCHIVOS_PRUEBA:
        archivos = archivos[:MAX_ARCHIVOS_PRUEBA]
        print(f"⚠️  MODO PRUEBA: procesando solo {len(archivos)} archivos", flush=True)
    else:
        print(f"🚀 MODO PRODUCCIÓN: procesando {len(archivos)} archivos", flush=True)
    
    print()
    
    # 4. Procesar cada archivo
    procesados, saltados, errores = 0, 0, 0
    lista_errores = []
    
    for i, archivo in enumerate(archivos, 1):
        nombre_base = archivo["name"].replace(".xlsx", "").replace(".XLSX", "")
        nombre_snap = f"{SNAP_PREFIX}{nombre_base}"
        
        print(f"[{i}/{len(archivos)}] {archivo['name']}", flush=True)
        
        # Saltear si ya existe
        if nombre_snap in snaps_existentes:
            print(f"   ⏭️  SNAP ya existe, saltando.", flush=True)
            saltados += 1
            continue
        
        # Descargar
        print(f"   ⬇️  Descargando...", flush=True)
        fh = descargar_bytes(drive, archivo["id"], archivo["mimeType"])
        
        if not fh:
            print(f"   ❌ No se pudo descargar.", flush=True)
            errores += 1
            lista_errores.append(archivo["name"])
            continue
        
        tamanio_kb = fh.getbuffer().nbytes / 1024
        print(f"   ✅ Descargado ({tamanio_kb:.1f} KB)", flush=True)
        
        # Subir como Google Sheets
        print(f"   ⬆️  Subiendo como Google Sheets...", flush=True)
        snap_id = subir_como_gsheet(drive, fh, nombre_snap, snap_folder_id)
        
        if snap_id:
            print(f"   ✅ SNAP creado ({snap_id})", flush=True)
            snaps_existentes.add(nombre_snap)
            procesados += 1
        else:
            print(f"   ❌ Falló la subida tras {INTENTOS_MAX} intentos.", flush=True)
            errores += 1
            lista_errores.append(archivo["name"])
        
        # Pausa entre archivos para no saturar la API
        if i < len(archivos):
            print(f"   ⏳ Esperando {PAUSA_ENTRE_ARCH}s...", flush=True)
            time.sleep(PAUSA_ENTRE_ARCH)
    
    # 5. Resumen
    duracion = time.time() - inicio
    
    print(f"\n{'='*60}", flush=True)
    print(f"✅ Creados:      {procesados}", flush=True)
    print(f"⏭️  Ya existían:  {saltados}", flush=True)
    print(f"❌ Errores:      {errores}", flush=True)
    print(f"⏱️  Tiempo:       {duracion:.0f}s ({duracion/60:.1f} min)", flush=True)
    print(f"{'='*60}", flush=True)
    
    if lista_errores:
        print("\nArchivos con error:", flush=True)
        for e in lista_errores:
            print(f"  ⚠️  {e}", flush=True)
    
    # 6. Log de resumen
    registrar_resumen(inicio, procesados, len(archivos))
    print(f"\n📝 Resumen registrado: {procesados} SNAPs creados, {errores} errores, {saltados} saltados", flush=True)
    print("🏁 SNAPSHOT BUILDER FINALIZADO", flush=True)


if __name__ == "__main__":
    ejecutar_principal()