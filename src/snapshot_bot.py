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
import os
import ssl
import sys
import time
import traceback
from http.client import IncompleteRead

# Forzar flush de prints para ver logs en tiempo real en GitHub Actions
sys.stdout.reconfigure(line_buffering=True)

from googleapiclient.http import MediaIoBaseUpload, MediaIoBaseDownload
from googleapiclient.errors import HttpError

sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from utils.common_utils import registrar_inicio, registrar_resumen
from utils.auth_utils import (
    obtener_drive_service_oauth,
    obtener_sheets_service_oauth,
    obtener_drive_service_sa,
    obtener_sheets_service_sa,
    CredencialesFaltantesError,
)
from utils.registro_utils import verificar_y_ampliar_capacidad
from utils.drive_utils import request_drive_con_reintentos
from utils.config_drive import FOLDER_REPARTICIONES_ID, FOLDER_SERVICES_ID

# ---------------------------------------------------------------------------
# Configuración
# ---------------------------------------------------------------------------
#
# snapshot_bot es el ÚNICO proceso del proyecto que usa OAuth
# (services.aportes.oser@gmail.com) en vez de la Service Account, porque es
# el único que necesita CREAR archivos nuevos en Drive (snapshots y, cuando
# hace falta, nuevas planillas de registro). Se ejecuta solo manualmente
# (workflow_dispatch), cuando se sabe que hay reparticiones nuevas para
# crear — no tiene trigger automático.
#
# FOLDER_REPARTICIONES_ID y FOLDER_SERVICES_ID viven en utils/config_drive.py
# (única fuente de verdad, compartida con monitoreo_utils.py y
# registro_utils.py). No redefinir acá.

SNAP_FOLDER_NAME   = "_snapshots_liquidaciones"
SNAP_PREFIX        = "[SNAP] "

# Errores de red/transporte que NO llegan como HttpError porque la
# conexión se corta antes de recibir una respuesta HTTP (ej. el socket se
# cuelga esperando datos y el sistema operativo lo mata por timeout). Sin
# esto, un TimeoutError como el que tiró abajo la corrida completa el
# 27/07 se propagaba sin reintento y mataba todo el proceso, dejando sin
# procesar el resto de las reparticiones pendientes.
ERRORES_RED_REINTENTABLES = (
    TimeoutError,       # incluye socket.timeout (alias desde Python 3.10)
    ConnectionError,    # ConnectionReset/Aborted/Refused, BrokenPipe
    ssl.SSLError,
    IncompleteRead,
    OSError,            # red de última instancia (ej. "Network is unreachable")
)

INTENTOS_MAX       = 3
ESPERA_REINTENTO   = 6   # segundos entre reintentos de subida
PAUSA_ENTRE_ARCH   = 2   # segundos entre archivos (evita rate limit)

# Modo producción: procesar TODOS los archivos
MODO_PRUEBA          = False
MAX_ARCHIVOS_PRUEBA  = 3

# ---------------------------------------------------------------------------
# Drive helpers
# ---------------------------------------------------------------------------

def inicializar_drive_con_scopes():
    """Inicializa el servicio de Google Drive usando OAuth 2.0 con refresh token."""
    try:
        return obtener_drive_service_oauth()
    except CredencialesFaltantesError as e:
        print(f"❌ {e}", flush=True)
        return None
    except Exception as e:
        print(f"❌ Error iniciando Drive: {e}", flush=True)
        traceback.print_exc()
        return None


def listar_archivos(drive, carpeta_id, solo_xlsx=True):
    """Lista todos los archivos de una carpeta de Drive con paginación."""
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
        res = request_drive_con_reintentos(
            drive.files().list(
                q=q,
                pageSize=200,
                fields="nextPageToken, files(id, name, mimeType)",
                supportsAllDrives=True,
                includeItemsFromAllDrives=True,
                pageToken=page_token,
            ).execute,
            f"listar archivos en {carpeta_id} (paginación)",
        )
        if not res:
            print("   ⚠️  Error listando archivos tras reintentos — se corta la paginación "
                  "con lo obtenido hasta ahora", flush=True)
            break
        archivos.extend(res.get("files", []))
        page_token = res.get("nextPageToken")
        if not page_token:
            break
    return archivos


def obtener_o_crear_carpeta_snaps(drive):
    """
    Busca la carpeta de snapshots dentro de FOLDER_SERVICES_ID. Si no existe, la crea.

    Las dos llamadas a Drive pasan por request_drive_con_reintentos con
    propagar_error=True: esto corre una sola vez al arrancar el bot y su
    resultado (snap_folder_id) se usa durante toda la corrida. Si el
    request de búsqueda fallara por un corte de red y se tratara como
    "None -> no existe", terminaría CREANDO una carpeta duplicada en vez de
    reusar la existente — peor que frenar con un error claro y reintentar
    la corrida.
    """
    print(f"🔍 Buscando carpeta '{SNAP_FOLDER_NAME}' en {FOLDER_SERVICES_ID}...", flush=True)
    res = request_drive_con_reintentos(
        drive.files().list(
            q=(
                f"'{FOLDER_SERVICES_ID}' in parents "
                f"and name='{SNAP_FOLDER_NAME}' "
                f"and mimeType='application/vnd.google-apps.folder' "
                f"and trashed=false"
            ),
            fields="files(id)",
            supportsAllDrives=True,
            includeItemsFromAllDrives=True,
        ).execute,
        f"buscar carpeta '{SNAP_FOLDER_NAME}'",
        propagar_error=True,
    )
    archivos = res.get("files", [])
    if archivos:
        folder_id = archivos[0]["id"]
        print(f"📁 Carpeta snapshots encontrada: {folder_id}", flush=True)
        return folder_id
    print(f"📁 Carpeta '{SNAP_FOLDER_NAME}' no encontrada, creando...", flush=True)
    nueva = request_drive_con_reintentos(
        drive.files().create(
            body={
                "name": SNAP_FOLDER_NAME,
                "mimeType": "application/vnd.google-apps.folder",
                "parents": [FOLDER_SERVICES_ID],
            },
            fields="id",
            supportsAllDrives=True,
        ).execute,
        f"crear carpeta '{SNAP_FOLDER_NAME}'",
        propagar_error=True,
    )
    print(f"📁 Carpeta snapshots creada: {nueva['id']}", flush=True)
    return nueva["id"]


def listar_snaps_existentes(drive, snap_folder_id):
    """Devuelve un set con los nombres de SNAPs ya creados."""
    archivos = listar_archivos(drive, snap_folder_id, solo_xlsx=False)
    return {a["name"] for a in archivos}


def descargar_bytes(drive, file_id, mime_type):
    """Descarga un archivo de Drive y devuelve su contenido como BytesIO."""
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
    Sube un archivo .xlsx como Google Sheets usando enfoque de dos pasos
    (crear vacío + subir contenido).

    Si un intento falla DESPUÉS de crear el archivo vacío (ej. timeout
    subiendo el contenido), ese archivo queda huérfano en Drive: vacío o
    con contenido parcial, con el mismo nombre que el SNAP final. El
    siguiente intento crea uno nuevo en vez de reusar el huérfano (Drive
    permite nombres duplicados), así que sin este cuidado terminan
    quedando varias copias del mismo SNAP. Acá se registra el ID de cada
    intento fallido y se lo borra apenas el intento final tiene éxito (o
    al agotar los reintentos, para no dejar basura ni en el caso de
    fallo total).
    """
    fh.seek(0)
    archivos_huerfanos = []

    def _borrar_huerfanos():
        for huerfano_id in archivos_huerfanos:
            try:
                request_drive_con_reintentos(
                    drive.files().delete(fileId=huerfano_id, supportsAllDrives=True).execute,
                    f"borrar archivo huérfano {huerfano_id}",
                    propagar_error=True,
                )
                print(f"   🗑️  Archivo huérfano de un intento anterior borrado ({huerfano_id})", flush=True)
            except Exception as e:
                print(f"   ⚠️  No se pudo borrar archivo huérfano {huerfano_id}: {e}", flush=True)

    for intento in range(INTENTOS_MAX):
        file_id = None
        try:
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
            archivos_huerfanos.append(file_id)  # huérfano hasta que se confirme el contenido
            print(f"   📄 Archivo creado (ID: {file_id})", flush=True)
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
            archivos_huerfanos.remove(file_id)  # este es el definitivo, no un huérfano
            _borrar_huerfanos()
            return updated_file.get("id")
        except HttpError as e:
            print(f"   ❌ Error {e.resp.status}: {e._get_reason()}", flush=True)
            if e.resp.status in (403, 429, 500, 503):
                espera = ESPERA_REINTENTO * (intento + 1)
                print(f"   ⏳ Reintento {intento+1}/{INTENTOS_MAX} en {espera}s...", flush=True)
                time.sleep(espera)
            else:
                print(f"   📝 Detalle: {e.content}", flush=True)
                _borrar_huerfanos()
                return None
        except ERRORES_RED_REINTENTABLES as e:
            # Ej: TimeoutError leyendo la respuesta del upload — no llega a
            # ser un HttpError porque nunca hubo respuesta HTTP. Se trata
            # igual que un error transitorio de servidor: reintentar.
            espera = ESPERA_REINTENTO * (intento + 1)
            print(f"   ❌ Error de red ({type(e).__name__}: {e})", flush=True)
            print(f"   ⏳ Reintento {intento+1}/{INTENTOS_MAX} en {espera}s...", flush=True)
            time.sleep(espera)
    _borrar_huerfanos()
    return None


# ---------------------------------------------------------------------------
# Principal
# ---------------------------------------------------------------------------

def ejecutar_principal():
    print("🚀 INICIANDO SNAPSHOT BUILDER", flush=True)
    inicio = time.time()
    ahora = registrar_inicio("BOT SNAPSHOT BUILDER - Monitoreo de Liquidaciones")
    
    print("🔑 Inicializando Drive con OAuth...", flush=True)
    drive = inicializar_drive_con_scopes()
    if not drive:
        print("❌ No se pudo inicializar Drive", flush=True)
        return
    print("✅ Drive inicializado correctamente", flush=True)

    print(f"📂 Listando archivos en carpeta {FOLDER_REPARTICIONES_ID}...", flush=True)
    archivos = listar_archivos(drive, FOLDER_REPARTICIONES_ID, solo_xlsx=True)
    print(f"📊 Archivos .xlsx encontrados: {len(archivos)}", flush=True)

    # ── Ampliar capacidad del registro de agentes si hace falta ─────────────
    # Único paso del proyecto que puede crear planillas _registro_agentes_N
    # nuevas (bootstrap, o las que hagan falta). Se le pasa la cantidad total
    # de reparticiones para que la capacidad libre resultante alcance para
    # el peor caso de la corrida de monitoreo_bot que sigue (que TODAS
    # necesiten hoja nueva), no solo para un colchón fijo — ver docstring de
    # verificar_y_ampliar_capacidad() para el detalle de por qué. El resto
    # de los bots corre con Service Account y solo espera encontrar lugar.
    #
    # Se le pasan AMBOS pares de credenciales: la lectura de las planillas
    # ya existentes (descubrirlas + contar hojas) se hace con la Service
    # Account, que tiene acceso garantizado a todo lo ya existente; OAuth
    # se usa solo para el acto de crear una planilla nueva si hace falta.
    # Antes se usaba OAuth también para leer capacidad, y si OAuth no tenía
    # acceso de Sheets a alguna planilla vieja, esa planilla se contaba
    # como "sin lugar" y se terminaban creando duplicados de más.
    print("📋 Verificando capacidad del registro de agentes...", flush=True)
    try:
        sheets_sa = obtener_sheets_service_sa()
        drive_sa = obtener_drive_service_sa()
        sheets_oauth = obtener_sheets_service_oauth()
        verificar_y_ampliar_capacidad(
            sheets_sa, drive_sa, sheets_oauth, drive, cantidad_reparticiones=len(archivos)
        )
    except Exception as e:
        print(f"⚠️  No se pudo verificar/ampliar la capacidad del registro: {e}", flush=True)
        traceback.print_exc()

    print("📁 Obteniendo carpeta de snapshots...", flush=True)
    try:
        snap_folder_id = obtener_o_crear_carpeta_snaps(drive)
    except Exception as e:
        print(f"❌ No se pudo obtener/crear la carpeta de snapshots (error de red persistente "
              f"tras varios reintentos): {e}", flush=True)
        return
    
    print("📋 Listando SNAPs existentes...", flush=True)
    snaps_existentes = listar_snaps_existentes(drive, snap_folder_id)
    print(f"✅ SNAPs existentes: {len(snaps_existentes)}", flush=True)
    
    if MODO_PRUEBA and len(archivos) > MAX_ARCHIVOS_PRUEBA:
        archivos = archivos[:MAX_ARCHIVOS_PRUEBA]
        print(f"⚠️  MODO PRUEBA: procesando solo {len(archivos)} archivos", flush=True)
    else:
        print(f"🚀 MODO PRODUCCIÓN: procesando {len(archivos)} archivos", flush=True)
    print()
    
    procesados, saltados, errores = 0, 0, 0
    lista_errores = []
    
    for i, archivo in enumerate(archivos, 1):
        nombre_base = archivo["name"].replace(".xlsx", "").replace(".XLSX", "")
        nombre_snap = f"{SNAP_PREFIX}{nombre_base}"
        print(f"[{i}/{len(archivos)}] {archivo['name']}", flush=True)

        try:
            if nombre_snap in snaps_existentes:
                print(f"   ⏭️  SNAP ya existe, saltando.", flush=True)
                saltados += 1
                continue

            print(f"   ⬇️  Descargando...", flush=True)
            fh = descargar_bytes(drive, archivo["id"], archivo["mimeType"])
            if not fh:
                print(f"   ❌ No se pudo descargar.", flush=True)
                errores += 1
                lista_errores.append(archivo["name"])
                continue
            tamanio_kb = fh.getbuffer().nbytes / 1024
            print(f"   ✅ Descargado ({tamanio_kb:.1f} KB)", flush=True)

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
        except Exception as e:
            # Cualquier error no previsto en ESTE archivo puntual (ej. un
            # TimeoutError que agotó los INTENTOS_MAX reintentos, o algo
            # totalmente nuevo) no debe tirar abajo el resto del batch.
            # Antes esto no estaba protegido y un solo archivo con
            # problemas de red mataba la corrida entera, dejando sin
            # procesar todas las reparticiones restantes.
            print(f"   ❌ Error inesperado procesando este archivo: {e}", flush=True)
            traceback.print_exc()
            errores += 1
            lista_errores.append(archivo["name"])

        if i < len(archivos):
            print(f"   ⏳ Esperando {PAUSA_ENTRE_ARCH}s...", flush=True)
            time.sleep(PAUSA_ENTRE_ARCH)
    
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
    registrar_resumen(inicio, procesados, len(archivos))
    print(f"\n📝 Resumen registrado: {procesados} SNAPs creados, {errores} errores, {saltados} saltados", flush=True)
    print("🏁 SNAPSHOT BUILDER FINALIZADO", flush=True)


if __name__ == "__main__":
    ejecutar_principal()