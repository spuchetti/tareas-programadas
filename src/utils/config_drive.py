"""
Fuente única de verdad para los IDs de carpetas de Drive del proyecto.

Antes estos IDs estaban hardcodeados por separado en drive_utils.py,
snapshot_bot.py, registro_utils.py y monitoreo_utils.py. Al no haber un
solo lugar de edición, terminaron desincronizados: drive_utils.py quedó
apuntando a un CARPETA_XLSX_ID viejo que ya se había reemplazado en los
demás archivos. No causó un incidente porque el único caller de
obtener_archivos() pasa el ID explícito, pero era una trampa lista para
activarse con el próximo cambio.

Regla: cualquier módulo que necesite un ID de carpeta lo importa desde
acá. No se redefine ni se copia el valor en otro archivo.
"""

# Carpeta con todas las reparticiones (.xlsx) a monitorear
FOLDER_REPARTICIONES_ID = "1_Xb2jrtr3Sjwi8-2nhT2k53KZ6CLE5hJ"

# Carpeta interna: contiene _snapshots_liquidaciones y las planillas
# _registro_agentes_N
FOLDER_SERVICES_ID = "1XJj3pMySybGeK7cW5-PRFPf1q5w2Dch5"