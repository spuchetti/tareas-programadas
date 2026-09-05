"""
Diagnóstico: universo real de valores de la columna "06-sit. revista"
en archivos "Caja".

Por qué existe:
  _normalizar_situacion_revista() en monitoreo_utils.py hoy agrupa por
  prefijo ("empieza con JUB" -> JUBILADO, "empieza con PEN" -> PENSIONADO).
  Eso corre riesgo de falsos positivos (ej. "PENDIENTE" empieza con "PEN")
  y de falsos negativos si se lo reemplaza sin más por un match de texto
  exacto (variantes de escritura de un mismo concepto, ej. "JUB." vs
  "JUBILADO", no colapsarían).

  La solución correcta es un diccionario explícito de variantes conocidas
  -> concepto canónico. Este script junta la materia prima para armarlo:
  todos los valores RAW (tal como están escritos) que aparecen hoy en la
  columna, con su frecuencia y un ejemplo de dónde aparece cada uno.

  Además, para cada valor calcula la CLAVE NORMALIZADA (el mismo
  preprocesamiento que hace _normalizar_situacion_revista antes de buscar
  en CONCEPTOS_SIT_REVISTA: sin tildes, sin puntos, sin espacios de más,
  en mayúsculas) e indica si ya está mapeada o no. Así no hace falta
  transformar cada valor a mano — la sección final del reporte ya deja
  las claves de los valores NO mapeados listas para copiar/pegar dentro
  de CONCEPTOS_SIT_REVISTA (monitoreo_utils.py), quedando solo definir a
  mano el concepto canónico de cada una.

Uso:
  python src/diagnostico_sit_revista.py

Requiere solo GDRIVE_JSON (Service Account) — es de solo lectura, no
escribe ni modifica nada en Drive.
"""

import sys
import os
from collections import defaultdict

sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from utils.drive_utils import inicializar_drive, obtener_archivos, descargar_archivo
from utils.config_drive import FOLDER_REPARTICIONES_ID
from utils.monitoreo_utils import leer_hoja_xlsx, CONFIG, HOJAS_ORDEN, COL_SIT_REVISTA, CONCEPTOS_SIT_REVISTA
from utils.excel_utils import normalizar_texto


def _calcular_clave_normalizada(val_str):
    """
    Reproduce EXACTAMENTE el preprocesamiento de
    _normalizar_situacion_revista() (monitoreo_utils.py) para que la clave
    calculada acá sea la misma que el código va a usar para buscar en
    CONCEPTOS_SIT_REVISTA en tiempo de ejecución. Si algún día cambia esa
    función, hay que reflejar el mismo cambio acá.
    """
    return normalizar_texto(val_str).replace(".", "").strip().upper()


def ejecutar():
    print("🔑 Inicializando Drive (Service Account)...")
    drive = inicializar_drive()
    if not drive:
        print("❌ No se pudo inicializar Drive — revisar el secret GDRIVE_JSON")
        return

    archivos = obtener_archivos(drive, FOLDER_REPARTICIONES_ID)
    archivos_caja = [a for a in archivos if "caja" in a["name"].lower()]
    print(f"📁 Archivos 'Caja' encontrados: {len(archivos_caja)} de {len(archivos)} totales")

    col_dni = CONFIG["COL_DNI"] - CONFIG["COL_INICIO"]
    conteo = defaultdict(int)
    ejemplo = {}
    archivos_por_valor = defaultdict(set)

    for i, archivo in enumerate(archivos_caja, 1):
        print(f"  [{i}/{len(archivos_caja)}] {archivo['name']}")
        fh = descargar_archivo(drive, archivo)
        if not fh:
            print("     ⚠️  No se pudo descargar, se salta")
            continue

        for hoja in HOJAS_ORDEN:
            try:
                filas = leer_hoja_xlsx(fh, hoja, CONFIG["FILA_INICIO_CAJA"])
            except Exception as e:
                print(f"     ⚠️  Error leyendo hoja '{hoja}': {e}")
                continue

            for fila in filas:
                val = fila[COL_SIT_REVISTA] if len(fila) > COL_SIT_REVISTA else None
                val_str = str(val).strip() if val not in (None, "") else "(vacío)"
                conteo[val_str] += 1
                archivos_por_valor[val_str].add(archivo["name"])
                if val_str not in ejemplo:
                    dni = fila[col_dni] if len(fila) > col_dni else ""
                    ejemplo[val_str] = f"{archivo['name']} / hoja {hoja} / DNI {dni}"

    print(f"\n{'='*80}")
    print(f"📊 VALORES DISTINTOS ENCONTRADOS EN '06-sit. revista': {len(conteo)}")
    print(f"{'='*80}")
    print(f"{'Ocurrencias':>11}  {'En N archivos':>13}  {'Valor (raw)':<35}  {'Clave normalizada':<25}  {'¿Mapeado?':<9}  Ejemplo")
    print("-" * 100)

    # Agrupa por clave normalizada los valores que todavía no están en
    # CONCEPTOS_SIT_REVISTA — varios valores "raw" distintos (ej. "Jubilado"
    # y "JUBILADO", o "Jubil." y "JUB.") pueden colapsar a la MISMA clave
    # normalizada, así que se consolidan acá para no listar la misma clave
    # más de una vez en la sección final.
    sin_mapear = defaultdict(lambda: {"ocurrencias": 0, "archivos": set(), "variantes_raw": set()})

    for val, cant in sorted(conteo.items(), key=lambda x: -x[1]):
        n_archivos = len(archivos_por_valor[val])

        if val == "(vacío)":
            clave_norm = "—"
            mapeado = "—"
        else:
            clave_norm = _calcular_clave_normalizada(val)
            mapeado = "sí" if clave_norm in CONCEPTOS_SIT_REVISTA else "NO"
            if clave_norm not in CONCEPTOS_SIT_REVISTA:
                acumulado = sin_mapear[clave_norm]
                acumulado["ocurrencias"] += cant
                acumulado["archivos"] |= archivos_por_valor[val]
                acumulado["variantes_raw"].add(val)

        print(f"{cant:>11}  {n_archivos:>13}  {val!r:<35}  {clave_norm:<25}  {mapeado:<9}  {ejemplo[val]}")

    print("-" * 100)
    print(f"\nCopiá/pegá esta tabla para armar el diccionario explícito de "
          f"variantes -> concepto canónico.")

    print(f"\n{'='*80}")
    print(f"🧩 CLAVES SIN MAPEAR — listas para pegar en CONCEPTOS_SIT_REVISTA "
          f"(monitoreo_utils.py): {len(sin_mapear)}")
    print(f"{'='*80}")

    if not sin_mapear:
        print("✅ No hay valores sin mapear — todos los valores encontrados ya "
              "tienen una entrada en CONCEPTOS_SIT_REVISTA.")
    else:
        print("Reemplazá el '???' por el concepto canónico que corresponda en cada "
              "línea (ej. \"JUBILADO\", \"PENSIONADO\", \"ACTIVO\", o dejá la clave "
              "afuera del diccionario si en realidad es un concepto propio que NO "
              "debe agruparse con otro):\n")
        for clave_norm, info in sorted(sin_mapear.items(), key=lambda x: -x[1]["ocurrencias"]):
            variantes = ", ".join(sorted(f"{v!r}" for v in info["variantes_raw"]))
            print(f'    "{clave_norm}": "???",'
                  f'  # {info["ocurrencias"]} ocurrencia(s) en {len(info["archivos"])} archivo(s)'
                  f' — variante(s) raw: {variantes}')

    print()


if __name__ == "__main__":
    ejecutar()