"""
Regenera los snapshots de TODAS las reparticiones.

Borra cada SNAP existente y lo vuelve a crear desde el .xlsx actual en
la carpeta de reparticiones. NO toca _registro_agentes (el matching ya
tolera CUIL truncados vía DNI; vaciar el registro global perdería
FECHA_ALTA / historial de IDs sin necesidad).

Pensado para dejar una base limpia después de corrupción masiva de CUILs
en los SNAPs (caso Gualeguaychú Agosto 2026 y posibles hermanas).

======= EJECUCIÓN =======
  python src/regenerar_todos_snaps.py

O vía GitHub Actions → workflow "Regenerar todos los snaps".
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

from utils.auth_utils import obtener_drive_service_oauth, CredencialesFaltantesError
from utils.config_drive import FOLDER_REPARTICIONES_ID_SNAPSHOT

from snapshot_bot import (
    PAUSA_ENTRE_ARCH,
    obtener_o_crear_carpeta_snaps,
    listar_archivos,
)
from regenerar_snap_reparticion import regenerar_snap_archivo


def ejecutar():
    print("=" * 60, flush=True)
    print("🔄 REGENERAR TODOS LOS SNAPS (sin tocar registro)", flush=True)
    print("=" * 60, flush=True)

    try:
        drive = obtener_drive_service_oauth()
    except CredencialesFaltantesError as e:
        print(f"❌ {e}", flush=True)
        return 1

    print(f"📂 Listando .xlsx en {FOLDER_REPARTICIONES_ID_SNAPSHOT}...", flush=True)
    archivos = listar_archivos(drive, FOLDER_REPARTICIONES_ID_SNAPSHOT, solo_xlsx=True)
    # Evitar procesar dos veces el mismo nombre si Drive devolviera variantes
    vistos = set()
    unicos = []
    for a in archivos:
        key = a["name"].lower()
        if key in vistos:
            continue
        vistos.add(key)
        unicos.append(a)
    archivos = unicos
    print(f"📊 Reparticiones a regenerar: {len(archivos)}", flush=True)

    snap_folder_id = obtener_o_crear_carpeta_snaps(drive)

    ok, errores = 0, 0
    fallidos = []
    inicio = time.time()

    for i, archivo in enumerate(archivos, 1):
        print(f"\n[{i}/{len(archivos)}] {archivo['name']}", flush=True)
        try:
            snap_id = regenerar_snap_archivo(drive, snap_folder_id, archivo)
            if snap_id:
                ok += 1
            else:
                errores += 1
                fallidos.append(archivo["name"])
        except Exception as e:
            print(f"   ❌ Error inesperado: {e}", flush=True)
            traceback.print_exc()
            errores += 1
            fallidos.append(archivo["name"])

        if i < len(archivos):
            time.sleep(PAUSA_ENTRE_ARCH)

    duracion = time.time() - inicio
    print(f"\n{'=' * 60}", flush=True)
    print(f"✅ Regenerados: {ok}", flush=True)
    print(f"❌ Errores:     {errores}", flush=True)
    print(f"⏱️  Tiempo:      {duracion:.0f}s ({duracion / 60:.1f} min)", flush=True)
    print(f"{'=' * 60}", flush=True)
    if fallidos:
        print("\nArchivos con error:", flush=True)
        for n in fallidos:
            print(f"  ⚠️  {n}", flush=True)
        return 1

    print(
        "\n✅ Listo. La próxima corrida de monitoreo compara contra SNAPs frescos.",
        flush=True,
    )
    return 0


if __name__ == "__main__":
    try:
        sys.exit(ejecutar())
    except Exception as e:
        print(f"❌ Error fatal: {e}", flush=True)
        traceback.print_exc()
        sys.exit(1)
