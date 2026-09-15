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

---------------------------------------------------------------------
Carpeta de reparticiones: UNA constante por bot, no una compartida
---------------------------------------------------------------------
Todos los bots leían la misma carpeta de reparticiones a través de un
único FOLDER_REPARTICIONES_ID. Eso significaba que cambiar el ID para
un bot (ej. apuntar monitoreo a una carpeta distinta) cambiaba el ID
para TODOS los demás bots también -- no había forma de tocar uno sin
afectar al resto.

Ahora cada bot tiene su propia constante, todas inicializadas al mismo
valor vigente. Cambiar la carpeta que usa un bot puntual es editar SOLO
su constante; los demás quedan intactos.

CASO ESPECIAL -- snapshot_builder + monitoreo:
snapshot_builder (snapshot_bot.py) puede correr de dos formas:
  a) Suelto, vía workflow_dispatch -- en ese caso usa
     FOLDER_REPARTICIONES_ID_SNAPSHOT (su propio ID, independiente).
  b) Como paso previo DENTRO de monitoreo.yml (workflow_call), para que
     las reparticiones nuevas tengan snapshot antes de que
     monitoreo_bot.py las compare EN LA MISMA CORRIDA. En este caso
     tiene que mirar la MISMA carpeta que monitoreo
     (FOLDER_REPARTICIONES_ID_MONITOREO) -- si mirara una carpeta
     distinta, una repartición nueva en la carpeta de monitoreo nunca
     tendría snapshot creado, y monitoreo_bot.py la saltearía en
     silencio corrida tras corrida.
  Esto se resuelve con el flag USAR_CARPETA_MONITOREO (ver snapshot_bot.py
  y snapshot_builder.yml) -- NO duplicando el ID en el YAML, que rompería
  la fuente única de verdad de este archivo.

CASO ESPECIAL -- diagnostico_sit_revista.py:
No tiene constante propia: importa FOLDER_REPARTICIONES_ID_MONITOREO
directamente y a propósito, porque el diagnóstico solo tiene sentido si
releva los mismos archivos que el monitoreo está comparando. A diferencia
del caso de snapshot_builder, acá no hace falta ningún flag porque
diagnostico_sit_revista.yml corre suelto (no lo llama otro workflow), así
que basta con importar la constante correcta directo.
"""

# Carpeta interna: contiene _snapshots_liquidaciones y las planillas
# _registro_agentes_N
FOLDER_SERVICES_ID = "1XJj3pMySybGeK7cW5-PRFPf1q5w2Dch5"

# --- Carpeta de reparticiones (.xlsx), una constante por bot ---------------
# Todas arrancan con el mismo valor vigente; editar cada una por separado
# según haga falta.
FOLDER_REPARTICIONES_ID_MONITOREO = "1_Xb2jrtr3Sjwi8-2nhT2k53KZ6CLE5hJ"
FOLDER_REPARTICIONES_ID_SNAPSHOT = "1_Xb2jrtr3Sjwi8-2nhT2k53KZ6CLE5hJ"
FOLDER_REPARTICIONES_ID_UNIFICADOR_MENSUAL = "1_Xb2jrtr3Sjwi8-2nhT2k53KZ6CLE5hJ"
FOLDER_REPARTICIONES_ID_REPORTE_ANUAL = "1_Xb2jrtr3Sjwi8-2nhT2k53KZ6CLE5hJ"
FOLDER_REPARTICIONES_ID_FV = "1_Xb2jrtr3Sjwi8-2nhT2k53KZ6CLE5hJ"

# diagnostico_sit_revista.py NO tiene su propia constante: usa
# FOLDER_REPARTICIONES_ID_MONITOREO a propósito, porque el diagnóstico solo
# tiene sentido si mira los mismos archivos que está comparando el
# monitoreo (ver el docstring de diagnostico_sit_revista.py).