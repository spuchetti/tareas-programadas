"""
Utilidades del bot de monitoreo de liquidaciones.

Contiene:
  - CONFIG / constantes (espejo del Apps Script)
  - Lógica de comparación (normal y caja)
  - Generadores de adjuntos: XLSX de cambios y CSVs
  - Helpers de nombres/periodos
"""

import os
import re
from datetime import datetime

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side


# ---------------------------------------------------------------------------
# Configuración (espejo del CONFIG del Apps Script)
# ---------------------------------------------------------------------------

CONFIG = {
    "CARPETA_ID": "1_Xb2jrtr3Sjwi8-2nhT2k53KZ6CLE5hJ",
    "CARPETA_INTERNA_ID": "1XJj3pMySybGeK7cW5-PRFPf1q5w2Dch5",
    "CARPETA_SNAPSHOTS": "_snapshots_liquidaciones",
    "NOMBRE_REGISTRO": "_registro_agentes",

    "COL_INICIO": 1,
    "COL_FIN": 24,
    "COL_DNI": 2,

    "FILA_INICIO_DEFAULT": 4,
    "FILA_INICIO_CAJA": 5,
}

HOJAS_ORDEN = [
    "01", "02", "03", "04", "05", "06", "1° sac",
    "07", "08", "09", "10", "11", "12", "2° sac",
]

# Columnas numéricas (0-based dentro del rango A..X): cols 9-24 → offset 8-23
COLS_NUMERICAS = set(range(8, 24))

NOMBRES_COLUMNAS = [
    "cuil", "dni", "tipo doc", "nombre y apellido", "cod. liq.",
    "sit. revista", "estado afil.", "reparticion", "aporte personal",
    "adherente sec.", "fondo vol.", "hijo menor de 35", "menor a cargo",
    "cred. asist.", "sueldo sin desc.", "sueldo con desc.",
    "reaj. aporte pers.", "reaj. adh. sec.", "reaj. fv",
    "reaj. hijo menor", "reaj. menor cargo", "reaj. cred. asist.",
    "aporte patronal", "reaj. ap. patronal",
]

ENCABEZADOS_EXCEL = [
    "1-cuil", "2-dni", "3-tipo doc", "4-nombre y apellido", "5-cod. liq.",
    "6-sit. revista", "7-estado afil.", "8-reparticion", "9-aporte personal",
    "10-adherente sec.", "11-fondo vol.", "12-hijo menor de 35", "13-menor a cargo",
    "14-cred. asist.", "15-sueldo sin desc.", "16-sueldo con desc.",
    "17-reaj. aporte pers.", "18-reaj. adh. sec.", "19-reaj. fv",
    "20-reaj. hijo menor", "21-reaj. menor cargo", "22-reaj. cred. asist.",
    "23-aporte patronal", "24-reaj. ap. patronal",
]

# Columna de sueldo sin descuentos (0-based) para separar complementarias
COL_SUELDO_SIN_DESC = 14


# ---------------------------------------------------------------------------
# Helpers numéricos / texto
# ---------------------------------------------------------------------------

def parse_numero(val):
    if val is None or val == "":
        return 0.0
    if isinstance(val, (int, float)):
        return 0.0 if (isinstance(val, float) and (val != val)) else float(val)
    s = str(val).strip()
    if not s:
        return 0.0
    if "," in s:
        s = s.replace(".", "").replace(",", ".")
    try:
        return float(s)
    except ValueError:
        return 0.0


def formatear_importe(val):
    return f"{parse_numero(val):.2f}"


def normalizar_texto(s):
    import unicodedata
    t = str(s or "").upper()
    t = unicodedata.normalize("NFD", t)
    t = "".join(c for c in t if unicodedata.category(c) != "Mn")
    return " ".join(t.split())


def limpiar_numero(s):
    return re.sub(r"[^0-9]", "", str(s or ""))


def normalizar_cuil(val, col_idx):
    """Quita guiones de CUIL (col 0) y puntos/guiones de DNI (col 1)."""
    if col_idx == 0:
        return val.replace("-", "")
    if col_idx == 1:
        return re.sub(r"[\.\-\s]", "", val)
    return val


# ---------------------------------------------------------------------------
# Lectura de rango (recibe lista de filas ya leída)
# ---------------------------------------------------------------------------

def leer_rango(datos):
    """
    Filtra filas donde el DNI (col offset 1) esté vacío, sea '-' o '0'.
    Recibe lista de listas tal como vienen de openpyxl values_only.
    """
    col_dni = CONFIG["COL_DNI"] - CONFIG["COL_INICIO"]   # 0-based
    return [
        fila for fila in datos
        if str(fila[col_dni] if col_dni < len(fila) else "").strip() not in ("", "-", "0")
    ]


# ---------------------------------------------------------------------------
# Indexado por ID de agente
# ---------------------------------------------------------------------------

def _indexar(filas, hoja_registro, get_id_fn):
    col_dni = CONFIG["COL_DNI"] - CONFIG["COL_INICIO"]
    mapa = {}
    for fila in filas:
        cuil = str(fila[0] if len(fila) > 0 else "").strip()
        dni = str(fila[col_dni] if len(fila) > col_dni else "").strip()
        nombre = str(fila[3] if len(fila) > 3 else "").strip()
        if not cuil and not dni and not nombre:
            continue
        aid = get_id_fn(cuil, dni, nombre, hoja_registro) if hoja_registro else (dni or cuil)
        mapa.setdefault(aid, []).append(fila)
    return mapa


# ---------------------------------------------------------------------------
# Comparación modo normal
# ---------------------------------------------------------------------------

def comparar_hojas_normal(datos_actual, datos_snap, hoja_registro):
    from utils.registro_utils import obtener_id_agente
    cambios = []
    cant_cols = CONFIG["COL_FIN"] - CONFIG["COL_INICIO"] + 1
    col_dni = CONFIG["COL_DNI"] - CONFIG["COL_INICIO"]

    mapa_snap = _indexar(datos_snap, hoja_registro, obtener_id_agente)
    mapa_act = _indexar(datos_actual, hoja_registro, obtener_id_agente)

    # Eliminados completamente
    for aid, filas in mapa_snap.items():
        if aid not in mapa_act:
            f = filas[0]
            dni = str(f[col_dni] if len(f) > col_dni else "").strip()
            nombre = f[3] if len(f) > 3 else "(sin nombre)"
            cambios.append({"tipo": "eliminado", "dni": dni, "nombre": nombre, "fila": f})

    for aid, filas_act in mapa_act.items():
        fref = filas_act[0]
        dni = str(fref[col_dni] if len(fref) > col_dni else "").strip()
        nombre = fref[3] if len(fref) > 3 else "(sin nombre)"

        if aid not in mapa_snap:
            for f in filas_act:
                cambios.append({"tipo": "nuevo", "dni": dni, "nombre": nombre, "fila": f})
            continue

        filas_sn = mapa_snap[aid]

        # Filas eliminadas (había más antes)
        for i in range(len(filas_act), len(filas_sn)):
            cambios.append({"tipo": "eliminado", "dni": dni, "nombre": nombre, "fila": filas_sn[i]})

        # Filas nuevas (hay más ahora)
        for i in range(len(filas_sn), len(filas_act)):
            cambios.append({"tipo": "nuevo", "dni": dni, "nombre": nombre, "fila": filas_act[i]})

        # Comparar fila a fila
        for i in range(min(len(filas_act), len(filas_sn))):
            fa, fs = filas_act[i], filas_sn[i]
            for c in range(cant_cols):
                va = normalizar_cuil(str(fa[c] if len(fa) > c else "").strip(), c)
                vs = normalizar_cuil(str(fs[c] if len(fs) > c else "").strip(), c)
                if va != vs:
                    cambios.append({
                        "tipo": "modificado",
                        "id": aid,
                        "dni": dni,
                        "nombre": nombre,
                        "columna": NOMBRES_COLUMNAS[c] if c < len(NOMBRES_COLUMNAS) else f"col{c+1}",
                        "anterior": vs or "(vacío)",
                        "actual": va or "(vacío)",
                        "es_no_numerico": c not in COLS_NUMERICAS,
                        "fila": fa,
                    })

    return {"cambios": cambios, "mapa_actual": mapa_act}


# ---------------------------------------------------------------------------
# Comparación modo caja
# ---------------------------------------------------------------------------

def comparar_hojas_caja(datos_actual, datos_snap, hoja_registro):
    from utils.registro_utils import obtener_id_agente
    cambios = []
    cant_cols = CONFIG["COL_FIN"] - CONFIG["COL_INICIO"] + 1
    col_dni = CONFIG["COL_DNI"] - CONFIG["COL_INICIO"]

    grupos_act = _indexar(datos_actual, hoja_registro, obtener_id_agente)
    grupos_snap = _indexar(datos_snap, hoja_registro, obtener_id_agente)
    todos_ids = set(list(grupos_act.keys()) + list(grupos_snap.keys()))

    for aid in todos_ids:
        fa_list = grupos_act.get(aid, [])
        fs_list = grupos_snap.get(aid, [])
        fref = (fa_list or fs_list)[0]
        dni = str(fref[col_dni] if len(fref) > col_dni else "").strip()
        nombre = fref[3] if len(fref) > 3 else "(sin nombre)"

        for i in range(len(fa_list), len(fs_list)):
            cambios.append({
                "tipo": "eliminado",
                "dni": dni,
                "nombre": nombre,
                "registro": i + 1,
                "fila": fs_list[i] if i < len(fs_list) else None
            })

        for i in range(len(fs_list), len(fa_list)):
            cambios.append({
                "tipo": "nuevo",
                "dni": dni,
                "nombre": nombre,
                "registro": i + 1,
                "fila": fa_list[i] if i < len(fa_list) else None
            })

        for i in range(min(len(fa_list), len(fs_list))):
            fa, fs = fa_list[i], fs_list[i]
            for c in range(cant_cols):
                va = normalizar_cuil(str(fa[c] if len(fa) > c else "").strip(), c)
                vs = normalizar_cuil(str(fs[c] if len(fs) > c else "").strip(), c)
                if va != vs:
                    cambios.append({
                        "tipo": "modificado",
                        "id": aid,
                        "dni": dni,
                        "nombre": nombre,
                        "registro": i + 1,
                        "columna": NOMBRES_COLUMNAS[c] if c < len(NOMBRES_COLUMNAS) else f"col{c+1}",
                        "anterior": vs or "(vacío)",
                        "actual": va or "(vacío)",
                        "es_no_numerico": c not in COLS_NUMERICAS,
                        "fila": fa,
                    })

    return {"cambios": cambios, "mapa_actual": grupos_act}


# ---------------------------------------------------------------------------
# Separar ordinarias / complementarias
# ---------------------------------------------------------------------------

def separar_complementarias_agrupado(mapa_agrupado):
    """
    mapa_agrupado: { id: [fila, ...] }
    Retorna: { "ordinarias": {id: fila}, "complementarias": {id: [fila, ...]} }
    """
    ordinarias = {}
    complementarias = {}

    for aid, filas in mapa_agrupado.items():
        if not filas:
            continue
        if len(filas) == 1:
            ordinarias[aid] = filas[0]
            continue
        idx_max = max(range(len(filas)), key=lambda i: parse_numero(filas[i][COL_SUELDO_SIN_DESC] if len(filas[i]) > COL_SUELDO_SIN_DESC else 0))
        ordinarias[aid] = filas[idx_max]
        comps = [f for i, f in enumerate(filas) if i != idx_max]
        if comps:
            complementarias[aid] = comps

    return {"ordinarias": ordinarias, "complementarias": complementarias}


# ---------------------------------------------------------------------------
# Generadores de CSV
# ---------------------------------------------------------------------------

def _escribir_csv(lineas, ruta):
    with open(ruta, "w", encoding="utf-8", newline="") as f:
        f.write("\r\n".join(lineas))


def _fila_a_csv(fila, cant_cols):
    cols = []
    for i in range(cant_cols):
        val = fila[i] if i < len(fila) else None
        s = str(val or "").strip()
        if COLS_NUMERICAS and i in COLS_NUMERICAS:
            cols.append(formatear_importe(val))
        else:
            cols.append(s.replace("|", " ").replace("\n", " "))
    return "|".join(cols)


def generar_csv_modificados(modifs, nuevos, ruta):
    cant_cols = CONFIG["COL_FIN"] - CONFIG["COL_INICIO"] + 1
    vistas = set()
    lineas = []

    for c in modifs:
        clave = f"{c.get('id','')}_{c.get('fila','')}"
        if clave not in vistas and c.get("fila"):
            vistas.add(clave)
            lineas.append(_fila_a_csv(c["fila"], cant_cols))

    for c in nuevos:
        if c.get("fila"):
            lineas.append(_fila_a_csv(c["fila"], cant_cols))

    _escribir_csv(lineas, ruta)


def generar_csv_complementarias(complementarias, ruta):
    """Devuelve True si se escribió algo, False si no había complementarias."""
    cant_cols = CONFIG["COL_FIN"] - CONFIG["COL_INICIO"] + 1
    if not complementarias:
        return False

    lineas = []
    for filas in complementarias.values():
        for f in filas:
            lineas.append(_fila_a_csv(f, cant_cols))

    _escribir_csv(lineas, ruta)
    return True


def generar_csv_liquidacion_completa(mapa_agrupado, es_caja, ruta):
    """Solo ordinarias en modo normal; todas las filas en modo caja."""
    cant_cols = CONFIG["COL_FIN"] - CONFIG["COL_INICIO"] + 1
    lineas = []

    if es_caja:
        for filas in mapa_agrupado.values():
            for f in (filas if isinstance(filas[0], list) else [filas]):
                lineas.append(_fila_a_csv(f, cant_cols))
    else:
        resultado = separar_complementarias_agrupado(mapa_agrupado)
        for f in resultado["ordinarias"].values():
            lineas.append(_fila_a_csv(f, cant_cols))

    _escribir_csv(lineas, ruta)


# ---------------------------------------------------------------------------
# Generador de XLSX de cambios (con openpyxl)
# ---------------------------------------------------------------------------

# Estilos
_THIN = Side(style="thin", color="D1D5DB")
_MED = Side(style="medium", color="6B7280")
_BRD_N = Border(left=_THIN, right=_THIN, top=_THIN, bottom=_THIN)
_BRD_B = Border(left=_THIN, right=_THIN, top=_THIN, bottom=_MED)

_FILL_HEADER = PatternFill("solid", fgColor="074F69")
_FILL_BLUE = PatternFill("solid", fgColor="EFF6FF")
_FILL_WHITE = PatternFill("solid", fgColor="FFFFFF")
_FILL_RED = PatternFill("solid", fgColor="C83C2D")
_FILL_GREEN = PatternFill("solid", fgColor="275317")

_FONT_HDR = Font(name="Calibri", size=11, bold=True, color="8ED973")
_FONT_NORMAL = Font(name="Calibri", size=11)
_FONT_RED = Font(name="Calibri", size=11, color="B91C1C")
_FONT_GREEN = Font(name="Calibri", size=11, bold=True, color="15803D")
_FONT_WHITE = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
_FONT_RED_S = Font(name="Calibri", size=11, bold=True, color="C83C2D")
_FONT_GRN_S = Font(name="Calibri", size=11, bold=True, color="275317")
_FONT_BLUE_S = Font(name="Calibri", size=11, bold=True, color="215C98")

_ALIGN_CTR = Alignment(horizontal="center", vertical="center")
_ALIGN_LEFT = Alignment(horizontal="left", vertical="center")
_NUM_FMT = "#,##0.00"

_ANCHOS = [14, 30, 22, 22, 14, 14, 14, 35] + [16] * 16


def _celda(ws, row, col, valor, font=None, fill=None, border=None, align=None, num_fmt=None):
    c = ws.cell(row=row, column=col, value=valor)
    if font:
        c.font = font
    if fill:
        c.fill = fill
    if border:
        c.border = border
    if align:
        c.alignment = align
    if num_fmt:
        c.number_format = num_fmt
    return c


def generar_xlsx_cambios(modifs, elims, nuevos, periodo, reparticion, ruta_salida):
    wb = openpyxl.Workbook()
    ws = wb.active
    nombre_hoja = re.sub(r'[\\/:*?\[\]]', '-', str(periodo)).replace("°", "")[:31].strip()
    ws.title = nombre_hoja

    for i, ancho in enumerate(_ANCHOS, 1):
        ws.column_dimensions[ws.cell(1, i).column_letter].width = ancho

    row = 1

    # ── REGISTROS MODIFICADOS ──
    if modifs:
        _celda(ws, row, 1, "REGISTROS MODIFICADOS", font=_FONT_BLUE_S)
        row += 1

        hdrs = ["DNI", "Nombre y Apellido", "Campo modificado", "Valor anterior", "Valor nuevo"]
        for c_idx, h in enumerate(hdrs, 1):
            _celda(ws, row, c_idx, h, font=_FONT_HDR, fill=_FILL_HEADER, border=_BRD_N, align=_ALIGN_CTR)
        row += 1

        # Agrupar por DNI
        grupos = {}
        orden_dni = []
        for c in modifs:
            d = c["dni"]
            if d not in grupos:
                grupos[d] = {"nombre": c["nombre"], "cambios": []}
                orden_dni.append(d)
            grupos[d]["cambios"].append(c)

        for gi, dni in enumerate(orden_dni):
            g = grupos[dni]
            fill = _FILL_BLUE if gi % 2 == 0 else _FILL_WHITE
            cc = g["cambios"]
            for i, c in enumerate(cc):
                es_ult = i == len(cc) - 1
                brd = _BRD_B if es_ult else _BRD_N
                campo = c["columna"]
                idx_c = NOMBRES_COLUMNAS.index(campo) if campo in NOMBRES_COLUMNAS else -1
                campo_label = f"{idx_c+1}-{campo}" if idx_c >= 0 else campo

                _celda(ws, row, 1, dni if i == 0 else "", font=_FONT_NORMAL, fill=fill, border=brd, align=_ALIGN_LEFT)
                _celda(ws, row, 2, g["nombre"] if i == 0 else "", font=_FONT_NORMAL, fill=fill, border=brd, align=_ALIGN_LEFT)
                _celda(ws, row, 3, campo_label, font=_FONT_NORMAL, fill=fill, border=brd, align=_ALIGN_LEFT)

                if c["es_no_numerico"]:
                    _celda(ws, row, 4, str(c["anterior"]), font=_FONT_RED, fill=fill, border=brd, align=_ALIGN_LEFT)
                    _celda(ws, row, 5, str(c["actual"]), font=_FONT_GREEN, fill=fill, border=brd, align=_ALIGN_LEFT)
                else:
                    _celda(ws, row, 4, parse_numero(c["anterior"]), font=_FONT_RED, fill=fill, border=brd, align=_ALIGN_LEFT, num_fmt=_NUM_FMT)
                    _celda(ws, row, 5, parse_numero(c["actual"]), font=_FONT_GREEN, fill=fill, border=brd, align=_ALIGN_LEFT, num_fmt=_NUM_FMT)
                row += 1
        row += 1

    # ── Helper para secciones de filas completas ──
    def _seccion(titulo, lista, font_titulo, fill_fila, font_fila):
        nonlocal row
        _celda(ws, row, 1, titulo, font=font_titulo)
        row += 1
        for ci, enc in enumerate(ENCABEZADOS_EXCEL, 1):
            _celda(ws, row, ci, enc, font=_FONT_HDR, fill=_FILL_HEADER, border=_BRD_N, align=_ALIGN_CTR)
        row += 1
        for i, c in enumerate(lista):
            brd = _BRD_B if i == len(lista) - 1 else _BRD_N
            fila = c.get("fila", [])
            cant = CONFIG["COL_FIN"] - CONFIG["COL_INICIO"] + 1
            for ci in range(cant):
                val = fila[ci] if ci < len(fila) else None
                if ci in COLS_NUMERICAS:
                    _celda(ws, row, ci+1, parse_numero(val), font=font_fila, fill=fill_fila, border=brd, align=_ALIGN_CTR, num_fmt=_NUM_FMT)
                else:
                    _celda(ws, row, ci+1, str(val or ""), font=font_fila, fill=fill_fila, border=brd, align=_ALIGN_LEFT)
            row += 1
        row += 1

    if elims:
        _seccion("REGISTROS ELIMINADOS", elims, _FONT_RED_S, _FILL_RED, _FONT_WHITE)

    if nuevos:
        _seccion("REGISTROS NUEVOS", nuevos, _FONT_GRN_S, _FILL_GREEN, _FONT_WHITE)

    wb.save(ruta_salida)


# ---------------------------------------------------------------------------
# Helpers de nombres / periodos
# ---------------------------------------------------------------------------

_MESES = {
    "01": "Enero", "02": "Febrero", "03": "Marzo", "04": "Abril",
    "05": "Mayo", "06": "Junio", "07": "Julio", "08": "Agosto",
    "09": "Septiembre", "10": "Octubre", "11": "Noviembre", "12": "Diciembre",
}


def hoja_a_periodo(hoja, anio):
    hl = hoja.strip().lower().replace("º", "°")
    if hl in ("1° sac", "1°sac"):
        return f"1° SAC/{anio}"
    if hl in ("2° sac", "2°sac"):
        return f"2° SAC/{anio}"
    h = hoja.strip()
    if h in _MESES:
        return f"{_MESES[h]}/{anio}"
    return f"{h}/{anio}"


def extraer_reparticion(nombre_archivo):
    sin_ext = nombre_archivo.replace(".xlsx", "").replace(".XLSX", "")
    partes = sin_ext.split("-")
    if len(partes) >= 3:
        ultimo = partes[-1].strip()
        fin = len(partes) - 1 if re.match(r"^\d{4}$", ultimo) else len(partes)
        return "-".join(partes[1:fin]).strip()
    return sin_ext


def extraer_anio_desde_nombre(nombre):
    m = re.search(r"(19\d{2}|20\d{2})", nombre)
    return int(m.group(1)) if m else datetime.now().year


def normalizar_nombre(r):
    return "".join(w.capitalize() for w in re.split(r"[\s\-]+", r.lower()))


def normalizar_periodo(p):
    return re.sub(r"[^a-zA-Z0-9]", "", p.replace("°", "").replace("º", ""))