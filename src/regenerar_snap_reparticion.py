"""
Regenera el snapshot de UNA repartición y limpia su hoja en el registro
de agentes.

Por qué hace falta (caso real Municipio Gualeguaychú Agosto 2026):
  El SNAP y/o el registro quedaron con CUILs truncados (~5 cifras
  significativas, ej. 20229754328 → 20230000000). Eso hizo que el
  monitoreo reportara miles de falsos "1-cuil modificado" y, cuando el
  truncado todavía pasaba el checksum AFIP, fusionara agentes distintos
  bajo un mismo id.

snapshot_bot.py saltea SNAPs que ya existen, así que regenerar implica:
  1) borrar el SNAP viejo
  2) volver a crearlo desde el .xlsx actual de reparticiones
  3) vaciar la hoja de esa repartición en _registro_agentes_N (se vuelve
     a llenar sola en la próxima corrida de monitoreo)

======= EJECUCIÓN =======
  python src/regenerar_snap_reparticion.py "16-MUNICIPIO GUALEGUAYCHU-2026.xlsx"

O vía GitHub Actions → workflow "Regenerar snap repartición".
"""

import os
import sys
import time
import traceback

sys.stdout.reconfigure(line_buffering=True)

_SRC = os.path.dirname(os.path.abspath(__file__))
_ROOT = os.path.dirname(_SRC)
sys.path.append(_ROOT)
sys.path.append(_SRC)

from utils.auth_utils import (
    obtener_drive_service_oauth,
    obtener_drive_service_sa,
    obtener_sheets_service_sa,
    CredencialesFaltantesError,
)
from utils.config_drive import FOLDER_REPARTICIONES_ID_SNAPSHOT
from utils.common_utils import ejecutar_con_reintentos_sheets
from utils.drive_utils import request_drive_con_reintentos
from utils.registro_utils import _descubrir_planillas_desde_drive

from snapshot_bot import (
    SNAP_PREFIX,
    PAUSA_ENTRE_ARCH,
    obtener_o_crear_carpeta_snaps,
    listar_archivos,
    descargar_bytes,
    subir_como_gsheet,
)


def _normalizar_nombre_xlsx(nombre):
    nombre = (nombre or "").strip().strip('"').strip("'")
    if not nombre:
        return ""
    if not nombre.lower().endswith(".xlsx"):
        nombre = f"{nombre}.xlsx"
    return nombre


def _borrar_snaps(drive, snap_folder_id, nombre_snap):
    """Borra todos los archivos con ese nombre exacto en la carpeta de SNAPs."""
    res = request_drive_con_reintentos(
        drive.files().list(
            q=(
                f"'{snap_folder_id}' in parents "
                f"and name='{nombre_snap}' "
                f"and trashed=false"
            ),
            fields="files(id, name)",
            supportsAllDrives=True,
            includeItemsFromAllDrives=True,
        ).execute,
        f"buscar SNAP '{nombre_snap}' para borrar",
        propagar_error=True,
    )
    archivos = res.get("files", [])
    if not archivos:
        print(f"   ℹ️  No había SNAP previo llamado '{nombre_snap}'", flush=True)
        return 0

    for a in archivos:
        request_drive_con_reintentos(
            drive.files().delete(fileId=a["id"], supportsAllDrives=True).execute,
            f"borrar SNAP {a['id']}",
            propagar_error=True,
        )
        print(f"   🗑️  SNAP borrado: {a['name']} ({a['id']})", flush=True)
    return len(archivos)


def _buscar_xlsx_reparticion(drive, nombre_xlsx):
    archivos = listar_archivos(drive, FOLDER_REPARTICIONES_ID_SNAPSHOT, solo_xlsx=True)
    objetivo = nombre_xlsx.lower()
    base = objetivo.replace(".xlsx", "")
    for a in archivos:
        n = a["name"].lower()
        if n == objetivo or n.replace(".xlsx", "") == base:
            return a
    return None


def limpiar_registro_reparticion(sheets_sa, drive_sa, nombre_xlsx):
    """
    Busca la hoja de la repartición en las planillas _registro_agentes_N
    y deja solo la fila de encabezados (A1:F1).
    """
    nombre_hoja = nombre_xlsx.replace(".xlsx", "").replace(".XLSX", "")[:31]
    print(f"📋 Buscando hoja de registro '{nombre_hoja}'...", flush=True)

    ids = _descubrir_planillas_desde_drive(drive_sa)
    limpiadas = 0
    for sid in ids:
        try:
            meta = ejecutar_con_reintentos_sheets(
                sheets_sa.spreadsheets().get(
                    spreadsheetId=sid, fields="sheets.properties.title"
                ),
                f"leer hojas de registro {sid}",
            )
            titulos = [s["properties"]["title"] for s in meta.get("sheets", [])]
        except Exception as e:
            print(f"   ⚠️  No se pudo leer planilla {sid}: {e}", flush=True)
            continue

        if nombre_hoja not in titulos:
            continue

        ejecutar_con_reintentos_sheets(
            sheets_sa.spreadsheets().values().clear(
                spreadsheetId=sid,
                range=f"'{nombre_hoja}'!A2:F",
                body={},
            ),
            f"limpiar datos de '{nombre_hoja}' en {sid}",
        )
        ejecutar_con_reintentos_sheets(
            sheets_sa.spreadsheets().values().update(
                spreadsheetId=sid,
                range=f"'{nombre_hoja}'!A1:F1",
                valueInputOption="RAW",
                body={"values": [["ID", "CUIL", "DNI", "NOMBRE", "FECHA_ALTA", "ULTIMA_VEZ"]]},
            ),
            f"reescribir encabezados de '{nombre_hoja}' en {sid}",
        )
        print(f"   ✅ Registro limpiado: '{nombre_hoja}' en {sid}", flush=True)
        limpiadas += 1

    if limpiadas == 0:
        print(
            f"   ℹ️  No se encontró hoja '{nombre_hoja}' en el registro "
            f"(se creará sola en el próximo monitoreo)",
            flush=True,
        )
    return limpiadas


def regenerar(nombre_xlsx):
    nombre_xlsx = _normalizar_nombre_xlsx(nombre_xlsx)
    if not nombre_xlsx:
        print("❌ Falta el nombre del archivo .xlsx de la repartición", flush=True)
        return 2

    nombre_base = nombre_xlsx[:-5] if nombre_xlsx.lower().endswith(".xlsx") else nombre_xlsx
    nombre_snap = f"{SNAP_PREFIX}{nombre_base}"

    print("=" * 60, flush=True)
    print("🔄 REGENERAR SNAP + LIMPIAR REGISTRO", flush=True)
    print(f"   Archivo: {nombre_xlsx}", flush=True)
    print(f"   SNAP:    {nombre_snap}", flush=True)
    print("=" * 60, flush=True)

    try:
        drive_oauth = obtener_drive_service_oauth()
    except CredencialesFaltantesError as e:
        print(f"❌ {e}", flush=True)
        return 1

    print("\n[1/3] Borrando SNAP previo...", flush=True)
    snap_folder_id = obtener_o_crear_carpeta_snaps(drive_oauth)
    _borrar_snaps(drive_oauth, snap_folder_id, nombre_snap)

    print("\n[2/3] Recreando SNAP desde el .xlsx de reparticiones...", flush=True)
    archivo = _buscar_xlsx_reparticion(drive_oauth, nombre_xlsx)
    if not archivo:
        print(
            f"❌ No se encontró '{nombre_xlsx}' en la carpeta de reparticiones "
            f"({FOLDER_REPARTICIONES_ID_SNAPSHOT})",
            flush=True,
        )
        return 1

    print(f"   ⬇️  Descargando {archivo['name']} ({archivo['id']})...", flush=True)
    fh = descargar_bytes(drive_oauth, archivo["id"], archivo["mimeType"])
    if not fh:
        print("❌ No se pudo descargar el .xlsx", flush=True)
        return 1

    print(f"   ⬆️  Subiendo como Google Sheets '{nombre_snap}'...", flush=True)
    snap_id = subir_como_gsheet(drive_oauth, fh, nombre_snap, snap_folder_id)
    if not snap_id:
        print("❌ Falló la creación del SNAP", flush=True)
        return 1
    print(f"   ✅ SNAP nuevo: {snap_id}", flush=True)

    print("\n[3/3] Limpiando hoja en _registro_agentes...", flush=True)
    try:
        sheets_sa = obtener_sheets_service_sa()
        drive_sa = obtener_drive_service_sa()
        limpiar_registro_reparticion(sheets_sa, drive_sa, nombre_xlsx)
    except CredencialesFaltantesError as e:
        print(f"   ⚠️  Sin Service Account ({e}) — SNAP ok, registro NO limpiado", flush=True)
    except Exception as e:
        print(f"   ⚠️  Error limpiando registro: {e}", flush=True)
        traceback.print_exc()

    print(
        "\n✅ Listo. La próxima corrida de monitoreo va a comparar contra el SNAP "
        "recién creado y va a volver a cargar el registro con CUIL buenos.",
        flush=True,
    )
    time.sleep(PAUSA_ENTRE_ARCH)
    return 0


def main():
    if len(sys.argv) < 2:
        print(
            'Uso: python src/regenerar_snap_reparticion.py '
            '"16-MUNICIPIO GUALEGUAYCHU-2026.xlsx"',
            flush=True,
        )
        return 2
    try:
        return regenerar(sys.argv[1])
    except Exception as e:
        print(f"❌ Error inesperado: {e}", flush=True)
        traceback.print_exc()
        return 1


if __name__ == "__main__":
    sys.exit(main())
