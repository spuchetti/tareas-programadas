"""
Funciones para interactuar con Google Drive
"""

import io
import json
import os
import time
import traceback
from googleapiclient.http import MediaIoBaseUpload, MediaIoBaseDownload
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from google.oauth2.service_account import Credentials

# Configuración común
FOLDER_ID_REPARTICIONES = "1_Xb2jrtr3Sjwi8-2nhT2k53KZ6CLE5hJ"
INTENTOS_MAX = 3
ESPERA_REINTENTO = 5


def inicializar_drive():
    """Inicializa el servicio de Google Drive"""
    try:
        cfg = json.loads(os.getenv("GDRIVE_JSON"))
        creds = Credentials.from_service_account_info(cfg)
        servicio = build("drive", "v3", credentials=creds, cache_discovery=False)
        return servicio
    except Exception as e:
        print(f"❌ Error iniciando Drive: {e}")
        traceback.print_exc()
        return None


def request_drive_con_reintentos(funcion, descripcion):
    """Ejecuta una función de Drive con reintentos"""
    for intento in range(INTENTOS_MAX):
        try:
            return funcion()
        except HttpError as e:
            if e.resp.status in [403, 500, 503]:
                print(f"⏳ Error {descripcion}, reintento {intento+1}/{INTENTOS_MAX}")
                time.sleep(ESPERA_REINTENTO)
                continue
            return None
        except Exception:
            return None
    return None


def obtener_archivos(servicio_drive, folder_id=None):
    """Obtiene archivos de una carpeta de Drive"""
    if folder_id is None:
        folder_id = FOLDER_ID_REPARTICIONES
    
    query = f"'{folder_id}' in parents and trashed=false"
    
    res = request_drive_con_reintentos(
        lambda: servicio_drive.files().list(
            q=query,
            pageSize=200,
            fields="files(id,name,mimeType)",
            supportsAllDrives=True,
            includeItemsFromAllDrives=True,
        ).execute(),
        "listar archivos"
    )
    
    if not res:
        return []
    
    archivos = res.get("files", [])
    
    # Filtrar solo archivos Excel (como en el bot original)
    archivos_validos = []
    for a in archivos:
        nombre = a["name"].lower()
        
        if (
            nombre.endswith(".xlsx") or
            nombre.endswith(".xlsm") or
            a["mimeType"] == "application/vnd.google-apps.spreadsheet"
        ):
            archivos_validos.append(a)
    
    return archivos_validos


def descargar_archivo(servicio_drive, archivo):
    """Descarga un archivo de Drive"""
    file_id = archivo["id"]
    mime = archivo["mimeType"]
    
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
        
    except Exception as e:
        print(f"❌ Error descargando {archivo['name']}: {e}")
        traceback.print_exc()
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