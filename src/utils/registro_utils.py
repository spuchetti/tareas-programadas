"""
Registro de agentes en Google Sheets.

Mantiene el mismo esquema que el Apps Script:
  - Una o más planillas  _registro_agentes_N  en FOLDER_SERVICES_ID
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
    snapshot_bot.py, pasándole servicios de AMBAS credenciales (SA y
    OAuth). Es la única función de este módulo que puede crear una
    planilla _registro_agentes_N nueva (bootstrap inicial, o las que
    hagan falta para que la capacidad libre total alcance para el peor
    caso de la corrida de monitoreo_bot que sigue: que TODAS las
    reparticiones necesiten una hoja nueva). La LECTURA de las planillas
    ya existentes (descubrirlas + contar hojas) se hace con la SA, que
    tiene acceso garantizado a todo lo existente; OAuth se usa solo para
    el acto de crear el archivo nuevo en sí.
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
from utils.config_drive import FOLDER_SERVICES_ID
from utils.common_utils import ejecutar_con_reintentos_sheets
from utils.drive_utils import request_drive_con_reintentos

# ---------------------------------------------------------------------------
# Config
# ---------------------------------------------------------------------------

NOMBRE_REGISTRO     = "_registro_agentes"
MAX_HOJAS_POR_PLANILLA = 150
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
    """
    Busca planillas _registro_agentes_N en FOLDER_SERVICES_ID.

    propagar_error=True a propósito: un None/lista vacía acá se interpreta
    más arriba como "todavía no existe ninguna planilla de registro", así
    que un fallo de red no debe devolver eso en silencio — hay que
    distinguir "no hay planillas" de "no se pudo ni preguntar".

    El resultado se ordena por el número real del sufijo "_N" (no por el
    orden en que Drive devuelve los resultados, que no está garantizado
    y de hecho varía entre corridas). Sin esto, _buscar_planilla_con_espacio
    -- que recorre `ids` en orden y usa la primera con lugar -- podía
    arrancar a llenar _registro_agentes_2 estando _registro_agentes_1
    completamente vacía, si Drive devolvía el orden "al revés" esa vez.
    Pasa sobre todo en GitHub Actions: el runner es efímero, así que el
    cache de /tmp casi nunca sobrevive entre corridas y esta función se
    termina llamando de nuevo con más frecuencia de la que parece.

    Ordenar por nombre como STRING (orderBy="name" en la query de Drive)
    no alcanza: con 10+ planillas, "_registro_agentes_10" ordena antes que
    "_registro_agentes_2" como texto. Por eso se extrae el número y se
    ordena numéricamente acá, en Python.
    """
    q = (
        f"'{FOLDER_SERVICES_ID}' in parents and trashed=false "
        f"and name contains '{NOMBRE_REGISTRO}' "
        f"and mimeType='application/vnd.google-apps.spreadsheet'"
    )
    res = request_drive_con_reintentos(
        drive.files().list(
            q=q, pageSize=50, fields="files(id, name)",
            supportsAllDrives=True, includeItemsFromAllDrives=True,
        ).execute,
        "descubrir planillas _registro_agentes",
        propagar_error=True,
    )
    archivos = res.get("files", [])

    def _numero_sufijo(nombre):
        m = re.search(r"_(\d+)$", nombre)
        return int(m.group(1)) if m else float("inf")  # sin sufijo numérico -> al final

    archivos.sort(key=lambda a: _numero_sufijo(a.get("name", "")))
    return [a["id"] for a in archivos]


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
        meta = ejecutar_con_reintentos_sheets(
            sheets_svc.spreadsheets().get(spreadsheetId=spreadsheet_id),
            f"leer cantidad de hojas de {spreadsheet_id}",
        )
        return len(meta.get("sheets", []))
    except Exception as e:
        print(f"  ⚠️  No se pudo leer metadata de {spreadsheet_id}: {e}")
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
    exclusiva de verificar_y_ampliar_capacidad() (vía snapshot_bot).

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
            meta = ejecutar_con_reintentos_sheets(
                sheets_svc.spreadsheets().get(spreadsheetId=sid),
                f"leer metadata de {sid} (buscar hoja existente)",
            )
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
    ejecutar_con_reintentos_sheets(
        sheets_svc.spreadsheets().batchUpdate(
            spreadsheetId=sid,
            body={"requests": [{"addSheet": {"properties": {"title": nombre_hoja}}}]},
        ),
        f"agregar hoja '{nombre_hoja}' en {sid}",
    )

    # Agregar encabezados
    ejecutar_con_reintentos_sheets(
        sheets_svc.spreadsheets().values().update(
            spreadsheetId=sid,
            range=f"'{nombre_hoja}'!A1:F1",
            valueInputOption="RAW",
            body={"values": [["ID", "CUIL", "DNI", "NOMBRE", "FECHA_ALTA", "ULTIMA_VEZ"]]},
        ),
        f"escribir encabezados de '{nombre_hoja}' en {sid}",
    )

    print(f"  → Hoja de registro creada: {nombre_hoja}")
    return {"spreadsheet_id": sid, "nombre_hoja": nombre_hoja, "sheets_svc": sheets_svc}


# ---------------------------------------------------------------------------
# Ampliación de capacidad (OAuth — SOLO desde snapshot_bot.py)
# ---------------------------------------------------------------------------

def _crear_planilla_registro(oauth_sheets_svc, oauth_drive_svc, numero):
    """
    Crea UNA planilla _registro_agentes_N nueva, la mueve a FOLDER_SERVICES_ID
    y la comparte explícitamente con la Service Account. Función auxiliar de
    verificar_y_ampliar_capacidad(); no la usa nada más.
    """
    nueva = ejecutar_con_reintentos_sheets(
        oauth_sheets_svc.spreadsheets().create(body={
            "properties": {"title": f"{NOMBRE_REGISTRO}_{numero}"},
            "sheets": [{"properties": {"title": "_info"}}],
        }),
        f"crear planilla {NOMBRE_REGISTRO}_{numero}",
    )
    sid = nueva["spreadsheetId"]

    # Mover a FOLDER_SERVICES_ID (que ya está compartida con la Service
    # Account como Editor, así que el archivo hereda ese acceso).
    try:
        file_meta = request_drive_con_reintentos(
            oauth_drive_svc.files().get(fileId=sid, fields="parents").execute,
            f"leer parents de planilla nueva {sid}",
            propagar_error=True,
        )
        padres_actuales = ",".join(file_meta.get("parents", []))
        request_drive_con_reintentos(
            oauth_drive_svc.files().update(
                fileId=sid,
                addParents=FOLDER_SERVICES_ID,
                removeParents=padres_actuales,
                supportsAllDrives=True,
                fields="id, parents",
            ).execute,
            f"mover planilla nueva {sid} a carpeta interna",
            propagar_error=True,
        )
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
            request_drive_con_reintentos(
                oauth_drive_svc.permissions().create(
                    fileId=sid,
                    body={"type": "user", "role": "writer", "emailAddress": sa_email},
                    fields="id",
                    supportsAllDrives=True,
                ).execute,
                f"compartir planilla nueva {sid} con Service Account",
                propagar_error=True,
            )
            print(f"  ✓ Planilla compartida explícitamente con la Service Account ({sa_email})")
        except Exception as e:
            print(f"  ⚠️  No se pudo compartir explícitamente con la Service Account: {e}")

    print(f"  ✓ Planilla de registro #{numero} creada ({sid})")
    return sid


def verificar_y_ampliar_capacidad(
    sa_sheets_svc, sa_drive_svc, oauth_sheets_svc, oauth_drive_svc, cantidad_reparticiones=0
):
    """
    Garantiza que la capacidad libre TOTAL entre todas las planillas
    _registro_agentes_N alcance para el peor caso de la corrida de
    monitoreo_bot que sigue: que TODAS las `cantidad_reparticiones` que se
    van a procesar necesiten una hoja nueva. Si no alcanza, crea tantas
    planillas nuevas como hagan falta (no solo una).

    SOLO debe llamarse desde snapshot_bot.py, pasándole AMBOS pares de
    servicios:
      - sa_sheets_svc / sa_drive_svc     → Service Account
      - oauth_sheets_svc / oauth_drive_svc → OAuth (services.aportes.oser@gmail.com)

    Por qué se necesitan los dos, y no solo OAuth como antes
    ----------------------------------------------------------------
    Todo lo que es LECTURA de las planillas ya existentes (descubrirlas en
    Drive, contar cuántas hojas tiene cada una) se hace con la Service
    Account, que — según la arquitectura documentada en auth_utils.py — es
    la credencial con acceso garantizado a todos los archivos ya
    existentes del proyecto. OAuth se reserva ÚNICAMENTE para el acto de
    CREAR una planilla nueva (_crear_planilla_registro), que es la única
    operación que la SA no puede hacer sin Unidad Compartida.

    Antes esta función usaba oauth_sheets_svc/oauth_drive_svc también para
    leer la capacidad de las planillas existentes. Si OAuth no tenía
    acceso de Sheets a alguna planilla vieja (por ejemplo, una planilla
    creada o compartida antes de que existiera el paso de "compartir
    explícitamente con la SA" en _crear_planilla_registro), _cantidad_hojas
    fallaba en silencio (except Exception: return None) y esa planilla
    simplemente se EXCLUÍA del cálculo de capacidad_libre — como si no
    existiera. Resultado real observado: con _registro_agentes_1 y _2 ya
    existentes y con lugar de sobra, la función igual creó _3 y _4, porque
    para OAuth la capacidad libre calculada daba 0.

    Por qué "capacidad libre total vs. cantidad_reparticiones" y no un
    colchón fijo (versión anterior: crear una planilla nueva si alguna
    existente tenía < 140/150 hojas): un colchón fijo de 10 hojas alcanza
    para el ritmo normal ("aparecen un par de reparticiones nuevas por
    corrida"), pero no para un pico dentro de una misma corrida — p.ej. la
    primera vez que corre este mecanismo sobre el batch completo, donde
    monitoreo_bot puede necesitar decenas de hojas nuevas de una. Eso fue
    justamente lo que pasó: con ~250 reparticiones y una sola planilla con
    lugar para 149, la capacidad se agotó a mitad de la corrida de
    monitoreo_bot (repartición #150), sin forma de ampliarla en ese momento
    porque crear planillas es exclusivo de OAuth/snapshot_bot, que ya había
    terminado de correr.

    Con ~250 reparticiones y 149 hojas útiles por planilla, en régimen
    permanente hacen falta al menos 2 planillas simultáneamente — esto NO
    es un pico puntual a absorber con un colchón, es la capacidad de
    régimen. Por eso la cuenta se hace contra el total de reparticiones,
    no contra un umbral fijo.
    """
    MARGEN_SEGURIDAD = 20  # margen extra sobre la demanda conocida, por si acaso

    ids = _ids_actualizados(sa_sheets_svc, drive_svc=sa_drive_svc)

    capacidades = [(sid, _cantidad_hojas(sa_sheets_svc, sid)) for sid in ids]
    fallidas = [sid for sid, n in capacidades if n is None]
    if fallidas:
        # Antes: se descartaban en silencio y se creaban planillas nuevas
        # de más para "compensar" una capacidad que en realidad sí existía.
        # Ahora: se corta acá. Es preferible que este paso quede en rojo
        # (y se revise a mano el permiso/estado de esas planillas) a que
        # se sigan generando _registro_agentes_N duplicados cada vez que
        # corre snapshot_bot.
        raise RuntimeError(
            f"No se pudo leer la cantidad de hojas de {len(fallidas)} "
            f"planilla(s) de registro, incluso con la Service Account "
            f"(que debería tener acceso garantizado): {fallidas}. Se "
            f"aborta sin crear planillas nuevas para no generar "
            f"duplicados — revisar permisos/estado de esas planillas."
        )
    capacidad_libre = sum(max(0, MAX_HOJAS_POR_PLANILLA - n) for _, n in capacidades)

    objetivo = cantidad_reparticiones + MARGEN_SEGURIDAD

    if ids and capacidad_libre >= objetivo:
        detalle = ", ".join(f"{sid[:8]}…({n}/{MAX_HOJAS_POR_PLANILLA})" for sid, n in capacidades)
        print(f"✓ Capacidad de registro OK — libre {capacidad_libre} (objetivo {objetivo}) — {detalle}")
        return

    if not ids:
        print("📋 No existe ninguna planilla _registro_agentes — creando la primera (bootstrap)...")
    else:
        detalle = ", ".join(f"{sid[:8]}…({n}/{MAX_HOJAS_POR_PLANILLA})" for sid, n in capacidades)
        print(f"📋 Capacidad libre insuficiente (libre {capacidad_libre}, objetivo {objetivo}) "
              f"— {detalle} — creando planilla(s) nueva(s)...")

    numero = len(ids)
    creadas = []
    while capacidad_libre < objetivo:
        numero += 1
        sid = _crear_planilla_registro(oauth_sheets_svc, oauth_drive_svc, numero)
        ids.append(sid)
        creadas.append(sid)
        capacidad_libre += MAX_HOJAS_POR_PLANILLA - 1  # -1 por la hoja "_info" con la que nace

    _guardar_ids(ids)
    print(f"  ✓ {len(creadas)} planilla(s) nueva(s) creada(s) — capacidad libre ahora: {capacidad_libre}")
    return creadas


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


_PREFIJOS_CUIL_VALIDOS = {"20", "23", "24", "25", "26", "27", "30", "33", "34"}
_MULT_CUIL = [5, 4, 3, 2, 7, 6, 5, 4, 3, 2]


def _cuil_es_confiable(cuil_l):
    """
    Valida un CUIL ya limpio (solo dígitos, ver _limpiar_num) contra el
    algoritmo de dígito verificador de AFIP (módulo 11), no solo la forma.

    Un CUIL puede tener longitud (11) y prefijo correctos y aun así ser
    inválido — caso real: '27548000000' pasaba longitud/prefijo pero el
    dígito verificador no coincidía (esperado 4, venía 0), y terminó
    fusionando en obtener_id_agente() a dos agentes distintos (mismo CUIL
    mal cargado, DNIs correctamente distintos) bajo el mismo id — se
    reportaron como "complementarias" sin serlo.

    NOTA: AFIP tiene una excepción cuando el resto da 10 (ligada al
    prefijo 23, personas físicas de tipo ambiguo) que acá NO se maneja de
    forma especial — un CUIL real que caiga en ese caso raro se
    rechazaría como "no confiable" y el matching caería a DNI, que sigue
    siendo el fallback seguro (ver obtener_id_agente). Preferible a
    aceptar checksums que no cierran.
    """
    if len(cuil_l) != 11 or cuil_l[:2] not in _PREFIJOS_CUIL_VALIDOS:
        return False
    if cuil_l == cuil_l[0] * 11:
        return False

    total = sum(int(d) * m for d, m in zip(cuil_l[:10], _MULT_CUIL))
    resto = total % 11
    verificador_esperado = 0 if resto == 0 else (11 - resto)
    if verificador_esperado == 11:
        verificador_esperado = 0

    return verificador_esperado == int(cuil_l[10])


# clave_c -> [aviso, ...]. Se acumulan en obtener_id_agente() cada vez que
# un CUIL se descarta como clave de matching por no pasar _cuil_es_confiable,
# y se vacían por archivo procesado vía obtener_avisos_cuil_invalido() —
# mismo patrón que _pending_nuevos/_pending_actualizaciones más abajo.
_avisos_cuil_invalido: dict = {}


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
        res = ejecutar_con_reintentos_sheets(
            svc.spreadsheets().values().get(
                spreadsheetId=sid, range=f"'{nhoja}'!A:F",
            ),
            f"leer caché de registro de '{nhoja}' en {sid}",
        )
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
    if cuil_l and not _cuil_es_confiable(cuil_l):
        _avisos_cuil_invalido.setdefault(clave_c, []).append(
            f"CUIL '{cuil}' inválido (checksum no coincide) — DNI {dni} ({nombre}), "
            f"se ignora como clave de matching"
        )
        cuil_l = ""
    dni_l    = _limpiar_num(dni)
    nombre_n = _normalizar(nombre)

    # Buscar en caché.
    # El fallback por nombre SOLO debe usarse cuando esta fila no trae ni
    # CUIL ni DNI utilizables (el único caso real que lo justifica: una
    # persona cuyo identificador numérico no está disponible en este
    # archivo puntual). Si la fila SÍ trae CUIL o DNI y simplemente no
    # están todavía en caché (agente nuevo), NO hay que caer al nombre:
    # dos personas distintas con el mismo nombre y apellido normalizado
    # (caso real: "RICLE MARIA INES" con DNI 13183215 y con DNI 25287600)
    # terminaban heredando el mismo id, y el reporte de monitoreo las
    # mostraba agrupadas como si fueran una sola persona.
    entrada = cache["porCuil"].get(cuil_l) or cache["porDni"].get(dni_l)
    if entrada is None and not cuil_l and not dni_l:
        entrada = cache["porNombre"].get(nombre_n) if nombre_n and len(nombre_n) > 3 else None

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


def obtener_avisos_cuil_invalido(hoja_info):
    """
    Devuelve y vacía los avisos de CUIL inválido acumulados para esta
    hoja_info desde la última llamada (ver _avisos_cuil_invalido y
    _cuil_es_confiable). Solo para log — nunca se manda al mail.

    Igual que flush_registro_pendientes, hay que llamarla una vez por
    archivo procesado, en el finally de procesar_archivo(), para que se
    imprima pase lo que pase con el resto del procesamiento.
    """
    if hoja_info is None:
        return []
    clave = f"{hoja_info['spreadsheet_id']}__{hoja_info['nombre_hoja']}"
    return _avisos_cuil_invalido.pop(clave, [])


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
    Alias local: la implementación real vive en common_utils.py
    (ejecutar_con_reintentos_sheets), compartida con monitoreo_utils.py.
    Se mantiene este wrapper para no tocar los call-sites existentes.
    """
    return ejecutar_con_reintentos_sheets(request, descripcion, intentos_max, espera_base)