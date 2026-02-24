"""
Funciones para trabajar con archivos Excel
"""


import zipfile
import tempfile
import shutil
import xml.etree.ElementTree as ET
from datetime import datetime
from zoneinfo import ZoneInfo
from io import BytesIO
import os


def obtener_hoja_mes_anterior():
    """Obtiene el nombre de la hoja del mes anterior (MM)"""
    ahora = datetime.now(ZoneInfo("America/Argentina/Buenos_Aires"))
    mes_anterior = ahora.month - 1 or 12
    return f"{mes_anterior:02d}"


def eliminar_tildes_latin(texto):
    """
    Elimina tildes y caracteres especiales de un texto.
    Convierte caracteres acentuados a su equivalente sin tilde.
    
    Ejemplos:
    - "María González" → "Maria Gonzalez"
    - "José Pérez" → "Jose Perez"
    - "Ñandú" → "Nandu"
    - "Über" → "Uber"
    """
    
    # caracteres problemáticos
    texto = texto.replace('Ã¡', 'a').replace('Ã¡', 'a')
    texto = texto.replace('Ã©', 'e').replace('Ã‰', 'E')
    texto = texto.replace('Ã­', 'i').replace('Ã', 'I')
    texto = texto.replace('Ã³', 'o').replace('Ã“', 'O')
    texto = texto.replace('Ãº', 'u').replace('Ãš', 'U')
    texto = texto.replace('Ã±', 'ñ').replace('Ã‘', 'Ñ')
    texto = texto.replace('Âº', '°')

    # Diccionario de reemplazos para tildes y caracteres especiales
    reemplazos_tildes = {
        # Vocales minúsculas con tilde
        'á': 'a', 'é': 'e', 'í': 'i', 'ó': 'o', 'ú': 'u',
        'à': 'a', 'è': 'e', 'ì': 'i', 'ò': 'o', 'ù': 'u',
        'ä': 'a', 'ë': 'e', 'ï': 'i', 'ö': 'o', 'ü': 'u',
        'â': 'a', 'ê': 'e', 'î': 'i', 'ô': 'o', 'û': 'u',
        
        # Vocales mayúsculas con tilde
        'Á': 'A', 'É': 'E', 'Í': 'I', 'Ó': 'O', 'Ú': 'U',
        'À': 'A', 'È': 'E', 'Ì': 'I', 'Ò': 'O', 'Ù': 'U',
        'Ä': 'A', 'Ë': 'E', 'Ï': 'I', 'Ö': 'O', 'Ü': 'U',
        'Â': 'A', 'Ê': 'E', 'Î': 'I', 'Ô': 'O', 'Û': 'U',
    }
    
    # Aplicar reemplazos
    resultado = ""
    for caracter in texto:
        if caracter in reemplazos_tildes:
            resultado += reemplazos_tildes[caracter]
        else:
            resultado += caracter
    
    # Asegurar que el texto sea compatible con CSV (| como separador)
    # Reemplazar el separador de CSV si aparece en el texto
    resultado = resultado.replace('|', '-')
    
    return resultado


def normalizar_texto(texto, eliminar_tildes_param=True):
    """
    Normaliza texto para CSV.
    
    Args:
        texto: Texto a normalizar
        eliminar_tildes_param: Si True, elimina tildes. Si False, las mantiene.
                               Por defecto es True para columnas D y H.
    """
    if texto is None:
        return ""
    
    # Convertir a string si no lo es
    if not isinstance(texto, str):
        texto = str(texto)
    
    # Limpiar espacios
    texto = texto.strip()
    
    # Si está vacío después de limpiar
    if not texto:
        return ""
    
    # SOLO si se solicita eliminar tildes
    if eliminar_tildes_param:
        return eliminar_tildes_latin(texto)
    else:
        # Mantener tildes, solo reemplazar separador
        return texto.replace('|', '-')



def sanitizar_libro_remover_filtros(fh):
    """
    Intenta reparar un archivo .xlsx o .xlsm eliminando nodos <autoFilter>
    y otros nodos problemáticos. Devuelve BytesIO con el archivo reempaquetado
    o None si no pudo repararlo.
    """
    try:
        tmpdir = tempfile.mkdtemp()
        ruta_entrada = os.path.join(tmpdir, "input_file")
        with open(ruta_entrada, "wb") as f:
            f.write(fh.getvalue())

        ruta_extraida = os.path.join(tmpdir, "extracted")
        os.makedirs(ruta_extraida, exist_ok=True)

        with zipfile.ZipFile(ruta_entrada, "r") as zin:
            zin.extractall(ruta_extraida)

        ns = {"main": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
        worksheets_dir = os.path.join(ruta_extraida, "xl", "worksheets")

        if os.path.isdir(worksheets_dir):
            for nombre_arch in os.listdir(worksheets_dir):
                if not nombre_arch.lower().endswith(".xml"):
                    continue
                ruta = os.path.join(worksheets_dir, nombre_arch)
                try:
                    tree = ET.parse(ruta)
                    root = tree.getroot()

                    modificado = False
                    for auto in root.findall("main:autoFilter", ns):
                        root.remove(auto)
                        modificado = True

                    for ext in root.findall("main:extLst", ns):
                        root.remove(ext)
                        modificado = True

                    for fcol in root.findall(".//main:filterColumn", ns):
                        for filt in fcol.findall("main:filters", ns):
                            fcol.remove(filt)
                            modificado = True

                    if modificado:
                        tree.write(ruta, encoding="utf-8", xml_declaration=True)
                except ET.ParseError:
                    # Si no se puede parsear una hoja, seguimos con las demás
                    pass

        ruta_salida = os.path.join(tmpdir, "output_file.xlsx")
        with zipfile.ZipFile(ruta_salida, "w", zipfile.ZIP_DEFLATED) as zout:
            for root_dir, dirs, files in os.walk(ruta_extraida):
                for file in files:
                    fullpath = os.path.join(root_dir, file)
                    arcname = os.path.relpath(fullpath, ruta_extraida)
                    zout.write(fullpath, arcname)

        with open(ruta_salida, "rb") as f:
            datos = f.read()

        shutil.rmtree(tmpdir)
        return BytesIO(datos)

    except Exception as e:
        try:
            shutil.rmtree(tmpdir)
        except Exception:
            pass
        print(f"❌ sanitizar_libro_remover_filtros error: {e}")
        return None
