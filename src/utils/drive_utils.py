"""
Funciones para interactuar con Google Drive
"""

import io
import json
import os
import socket
import ssl
import time
import traceback
from http.client import IncompleteRead
from googleapiclient.http import MediaIoBaseDownload
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from google.oauth2.service_account import Credentials

# Excepciones transitorias de red/SSL que NO son HttpError: pasan durante el
# refresh del token (JWT grant contra el endpoint de OAuth de Google) o en
# cualquier llamada HTTP de bajo nivel antes de que llegue a convertirse en
# una respuesta HTTP. Ej. real visto en logs: ssl.SSLEOFError durante
# credentials.refresh() en medio de una corrida larga (monitoreo_bot.py).
# Antes estas excepciones caían en el "except Exception: return None"
# genérico de más abajo — es decir, ni se reintentaban ni se relanzaban,
# se devolvía None en silencio como si la request hubiese respondido vacía.
# Para funciones como obtener_snapshot_de_archivo() eso es peor que un
# crash: None se interpreta como "no existe snapshot todavía", lo cual es
# incorrecto cuando en realidad hubo un corte de red.
#
# IncompleteRead y OSError se agregan por el mismo motivo que ya llevó a
# incluirlos en ERRORES_RED_REINTENTABLES de snapshot_bot.py: un socket que
# se cuelga a mitad de una respuesta, o "Network is unreachable", tampoco
# llegan como HttpError. Se mantienen ambos conjuntos alineados a propósito.
EXCEPCIONES_RED_TRANSITORIAS = (
    ssl.SSLError, ConnectionError, TimeoutError, socket.timeout,
    IncompleteRead, OSError,
)

# Configuración común
INTENTOS_MAX = 3
ESPERA_REINTENTO = 5
PAGINA_TAMANIO = 200  # Máximo por página

# Mismos scopes que usa auth_utils.py (snapshot_bot / registro_utils) para la
# Service Account. Sin scopes explícitos, Credentials.from_service_account_info
# arma credenciales sin permisos y las llamadas a la API de Drive/Sheets fallan.
SCOPES_DRIVE_SHEETS = [
    "https://www.googleapis.com/auth/drive",
    "https://www.googleapis.com/auth/spreadsheets",
]


def inicializar_drive():
    """Inicializa el servicio de Google Drive"""
    try:
        cfg = json.loads(os.getenv("GDRIVE_JSON"))
        creds = Credentials.from_service_account_info(cfg, scopes=SCOPES_DRIVE_SHEETS)
        servicio = build("drive", "v3", credentials=creds, cache_discovery=False)
        return servicio
    except Exception as e:
        print(f"❌ Error iniciando Drive: {e}")
        traceback.print_exc()
        return None


def request_drive_con_reintentos(funcion, descripcion, propagar_error=False):
    """
    Ejecuta una función de Drive con reintentos.

    Reintenta ante:
      - HttpError con status 403/500/503 (como antes).
      - Excepciones de red/SSL transitorias (EXCEPCIONES_RED_TRANSITORIAS)
        que pueden ocurrir en cualquier punto de la llamada HTTP, incluido
        el refresh del token de credenciales — no llegan a ser un HttpError
        porque no hay respuesta HTTP de por medio.

    Cualquier otro caso (HttpError con otro status, u otra excepción no
    contemplada) no se reintenta.

    Por default (propagar_error=False, comportamiento histórico) cualquier
    fallo que agota los reintentos devuelve None — así queda para callers
    como obtener_archivos(), donde un None se puede tratar como "esta
    página no trajo nada" sin romper la paginación.

    Con propagar_error=True se relanza la excepción en vez de devolver
    None cuando se agotan los reintentos. Usar esto cuando un None
    "silencioso" sería ambiguo con un resultado legítimamente vacío (ej.
    "no existe el snapshot" vs. "no se pudo ni preguntar si existe" — ver
    obtener_snapshot_de_archivo() en monitoreo_bot.py).
    """
    for intento in range(INTENTOS_MAX):
        try:
            return funcion()
        except HttpError as e:
            if e.resp.status in [403, 500, 503]:
                print(f"⏳ Error {descripcion}, reintento {intento+1}/{INTENTOS_MAX}")
                time.sleep(ESPERA_REINTENTO)
                continue
            if propagar_error:
                raise
            return None
        except EXCEPCIONES_RED_TRANSITORIAS as e:
            if intento < INTENTOS_MAX - 1:
                print(f"⏳ Error de red/SSL {descripcion}, reintento {intento+1}/{INTENTOS_MAX}: {e}")
                time.sleep(ESPERA_REINTENTO)
                continue
            print(f"❌ Error de red/SSL persistente {descripcion} tras {INTENTOS_MAX} intentos: {e}")
            if propagar_error:
                raise
            return None
        except Exception:
            if propagar_error:
                raise
            return None
    return None


def obtener_archivos(servicio_drive, folder_id):
    """
    Obtiene TODOS los archivos de una carpeta de Drive con paginación
    completa. folder_id es obligatorio a propósito: antes tenía un default
    hardcodeado acá adentro que quedó desactualizado sin que nadie lo
    notara (ver config_drive.py). El ID vigente vive en config_drive.py.
    """
    query = f"'{folder_id}' in parents and trashed=false"
    all_files = []
    page_token = None
    
    print(f"📁 Buscando archivos en carpeta: {folder_id}")
    
    while True:
        try:
            # Preparar la solicitud con paginación
            request = servicio_drive.files().list(
                q=query,
                pageSize=PAGINA_TAMANIO,
                fields="nextPageToken, files(id, name, mimeType)",
                supportsAllDrives=True,
                includeItemsFromAllDrives=True,
                pageToken=page_token
            )
            
            # Ejecutar con reintentos
            res = request_drive_con_reintentos(
                request.execute,
                "listar archivos (paginación)"
            )
            
            if not res:
                print("❌ Error en paginación, retornando archivos obtenidos hasta ahora")
                break
            
            # Agregar archivos de esta página
            files_in_page = res.get("files", [])
            all_files.extend(files_in_page)
            
            print(f"📄 Página procesada: {len(files_in_page)} archivos (Total: {len(all_files)})")
            
            # Verificar si hay más páginas
            page_token = res.get("nextPageToken")
            if not page_token:
                print(f"✅ Paginación completa. Total archivos: {len(all_files)}")
                break
                
        except Exception as e:
            print(f"❌ Error en paginación: {e}")
            break
    
    # Filtrar solo archivos Excel (como en el bot original)
    archivos_validos = []
    for a in all_files:
        nombre = a["name"].lower()
        
        es_excel = (
            nombre.endswith(".xlsx") or
            nombre.endswith(".xlsm") or
            nombre.endswith(".xls") or
            a["mimeType"] == "application/vnd.google-apps.spreadsheet"
        )
        
        if es_excel:
            archivos_validos.append(a)
    
    print(f"📊 Archivos Excel válidos: {len(archivos_validos)} de {len(all_files)} totales")
    return archivos_validos


def descargar_archivo(servicio_drive, archivo):
    """
    Descarga un archivo de Drive, con reintentos ante errores de red/SSL
    transitorios (ver EXCEPCIONES_RED_TRANSITORIAS).

    Antes esta era la única función del módulo que no pasaba por
    request_drive_con_reintentos (que se usa para llamadas .execute()
    simples): tenía su propio try/except Exception genérico que atrapaba
    también errores de red/SSL (ej. "EOF occurred in violation of
    protocol") y devolvía None directo, sin ningún reintento — mismo tipo
    de problema que ya se había resuelto en request_drive_con_reintentos y
    en ejecutar_con_reintentos_sheets (common_utils.py), pero que quedó
    afuera acá porque la descarga no usa .execute() sino un loop propio de
    MediaIoBaseDownload.

    Si un chunk falla a mitad de la descarga, se reintenta la descarga
    COMPLETA desde cero (no el chunk suelto): un BytesIO parcialmente
    lleno de un intento anterior quedaría corrupto si se mezclan chunks
    de intentos distintos.
    """
    file_id = archivo["id"]
    mime = archivo["mimeType"]

    for intento in range(INTENTOS_MAX):
        try:
            if mime == "application/vnd.google-apps.spreadsheet":
                req = servicio_drive.files().export_media(
                    fileId=file_id,
                    mimeType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            else:
                req = servicio_drive.files().get_media(fileId=file_id)

            fh = io.BytesIO()
            downloader = MediaIoBaseDownload(fh, req)
            terminado = False

            while not terminado:
                _, terminado = downloader.next_chunk()

            fh.seek(0)
            return fh

        except EXCEPCIONES_RED_TRANSITORIAS as e:
            if intento < INTENTOS_MAX - 1:
                print(f"⏳ Error de red/SSL descargando {archivo['name']}, "
                      f"reintento {intento+1}/{INTENTOS_MAX}: {e}")
                time.sleep(ESPERA_REINTENTO)
                continue
            print(f"❌ Error de red/SSL persistente descargando {archivo['name']} "
                  f"tras {INTENTOS_MAX} intentos: {e}")
            return None
        except Exception as e:
            print(f"❌ Error descargando {archivo['name']}: {e}")
            traceback.print_exc()
            return None

    return None


def guardar_csv_localmente(datos, nombre_archivo="UNIFICADO_MENSUAL.csv"):
    """Guarda CSV localmente usando | como delimitador con codificación UTF-8"""
    try:
        import csv
        
        # Crear directorio si no existe
        os.makedirs("generados", exist_ok=True)
        ruta = os.path.join("generados", nombre_archivo)
        
        # IMPORTANTE: Usar encoding='utf-8' y newline=''
        with open(ruta, 'w', encoding='utf-8', newline='') as f:
            # Usar delimiter='|' y quoting=csv.QUOTE_MINIMAL
            writer = csv.writer(f, delimiter='|', quoting=csv.QUOTE_MINIMAL)
            writer.writerows(datos)
        
        print(f"💾 CSV guardado localmente con delimitador '|': {ruta} ({len(datos)} filas)")
        print(f"   🔤 Codificación: UTF-8")
        
        # Verificar el formato
        with open(ruta, 'r', encoding='utf-8') as f:
            lineas = f.readlines()
            if lineas:
                print(f"   📝 Formato: {len(lineas[0].split('|'))} columnas separadas por '|'")
                if len(lineas) > 1:
                    # Mostrar primeros 100 caracteres de la segunda línea
                    muestra = lineas[1][:100]
                    print(f"   📊 Ejemplo primera fila de datos: {muestra}...")
                    
                    # Verificar si hay caracteres no ASCII (tildes deberían estar)
                    non_ascii = sum(1 for c in muestra if ord(c) > 127)
                    if non_ascii > 0:
                        print(f"   ✅ Se detectaron {non_ascii} caracteres con tildes/acentos")
        
        return ruta
        
    except Exception as e:
        print(f"❌ Error guardando CSV local: {e}")
        return None
