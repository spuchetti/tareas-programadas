#!/usr/bin/env bash
#
# Backup diario: Google Drive (div.aportes.oser@gmail.com) -> Mega (div.aportes.backups@gmail.com)
#
# Requiere:
#   - rclone instalado y ~/.config/rclone/rclone.conf con los remotes:
#       gdrive-div    (tipo drive, OAuth de la cuenta div, scope drive.readonly)
#       mega-backups  (tipo mega, cuenta free de div.aportes.backups@gmail.com)
#   - Variable de entorno COMPUTERS_DRIVE_IDS (opcional) con el formato:
#       "Nombre1:folderId1,Nombre2:folderId2,..."
#     Si no se define, solo se backupea "Mi unidad".
#
# Qué hace:
#   - Deja en Mega un espejo ACTUALIZADO de Mi Unidad + cada Computer.
#   - PROTECCIÓN ANTI-RANSOMWARE/VIRUS (importante):
#       a) Tanto lo que se BORRA como lo que se SOBREESCRIBE (cambió de
#          contenido) queda versionado en "_papelera/<fecha>" antes de
#          tocarse. Si un virus corrompe/encripta archivos en las PCs
#          sincronizadas, la versión de AYER queda a salvo en la papelera
#          en vez de perderse en el momento en que el backup corre.
#       b) Antes de aplicar cambios, se cuenta cuántos archivos cambiaron
#          de contenido respecto al backup anterior. Si el porcentaje supera
#          UMBRAL_ALERTA_PCT (cambio masivo y repentino = sospechoso de
#          ransomware/corrupción), el backup de esa carpeta se ABORTA sin
#          tocar nada y el job de GitHub Actions termina en rojo. Hay que
#          revisar manualmente antes de forzar el backup de nuevo.
#     Esta papelera con historial NO se limpia sola y va a crecer más rápido
#     que si solo guardara borrados; conviene podarla periódicamente (ver
#     nota al final) para no acercarse al límite de 20GB del plan free.
#   - Filtra SOLO xlsx, xls, csv, docx, pdf. Los Google Sheets/Docs nativos
#     se exportan automáticamente a xlsx/docx antes de filtrar.
#
#   - sync_uno() distingue TRES resultados posibles, no dos:
#       0 = éxito
#       1 = abortado por % de cambios sospechoso (posible ransomware)
#       2 = falló por error técnico (login/red/timeout) tras reintentos
#     Esto importa porque `rclone check`/`rclone sync` contra Mega pueden
#     fallar con "login with previous auth keys failed: unexpected end of
#     JSON input" por un simple rate-limit transitorio del login de Mega
#     (visto en logs reales al pegarle a 6 carpetas seguidas) -- eso NO es
#     lo mismo que detectar cambios masivos sospechosos, y antes ambos
#     casos devolvían 1 y se reportaban con el mismo mensaje de "alerta de
#     ransomware", lo cual era engañoso y generaba fatiga de alertas.
#     Solo los destinos con resultado 0 (éxito real) se podan al final;
#     un destino que falló técnicamente NO se toca en la poda, para no
#     reducir la ventana de recuperación de algo que ni siquiera se
#     actualizó en esta corrida.

set -euo pipefail

GDRIVE_REMOTE="gdrive-div:"
MEGA_REMOTE="mega-div-backups:"
DEST_BASE="Backups_OSER"
PAPELERA_BASE="${DEST_BASE}/_papelera"
# Formato DDMMYYYY para el nombre de carpeta dentro de _papelera (más
# legible para revisar a simple vista en Mega). IMPORTANTE: no sirve para
# ordenar/comparar como texto (ver podar_papelera más abajo, que reordena a
# YYYYMMDD internamente antes de comparar fechas).
FECHA="$(date +%d%m%Y)"

EXPORT_FORMATS="xlsx,docx"

# % de archivos que, si cambian de contenido en una sola pasada, se
# considera sospechoso y aborta el backup de esa carpeta sin aplicar nada.
UMBRAL_ALERTA_PCT=80

# Días de retención para la papelera (_papelera/<fecha>). Cada carpeta de
# fecha más vieja que esto se borra al final de la corrida (ver
# podar_papelera). 7 días (una semana cumplida) alcanza para notar un
# problema y recuperar la versión limpia, y deja bastante margen respecto
# al límite de 20GB del plan free de Mega (con retención de 14 días la
# papelera venía ocupando ~2.6GB en uso real).
RETENCION_PAPELERA_DIAS=7

# Reintentos para comandos rclone (check/sync) ante fallos técnicos como
# "login with previous auth keys failed" (rate-limit transitorio de Mega),
# timeouts o cortes de red. rclone no distingue esto en el exit code, así
# que simplemente se reintenta cualquier fallo del comando como tal.
INTENTOS_RCLONE=3
ESPERA_REINTENTO_RCLONE=45

# Pausa entre destinos para no encadenar logins contra Mega demasiado
# rápido (cada destino hace mínimo 2 logins: uno para check, otro para
# sync). Ayuda a evitar el rate-limit que causó el "unexpected end of
# JSON input" en varios destinos seguidos.
PAUSA_ENTRE_DESTINOS=10

# Flags separados: uno para alerta real de ransomware (% sospechoso de
# cambios) y otro para errores técnicos (login/red) que agotaron los
# reintentos. Antes había un solo HUBO_ALERTA que mezclaba ambos casos
# bajo el mismo mensaje de "cambios masivos sospechosos", lo cual era
# falso cuando en realidad era un problema de conectividad con Mega.
HUBO_ALERTA_RANSOMWARE=0
HUBO_ERROR_TECNICO=0

# Extensiones a conservar. Los Google Sheets/Docs quedan como .xlsx/.docx
# gracias a --drive-export-formats, así que ya entran en estos patrones.
FILTROS=(
  --include "*.xlsx"
  --include "*.xls"
  --include "*.csv"
  --include "*.docx"
  --include "*.pdf"
)

# ---------------------------------------------------------------------------
# ejecutar_rclone_con_reintentos <descripcion> <comando rclone completo...>
#
# Ejecuta un comando rclone (check o sync) reintentando ante cualquier
# fallo (login, red, timeout, etc — rclone no distingue el motivo en el
# exit code, así que se reintenta el fallo tal cual). Espera fija entre
# intentos (no exponencial): no es un problema de cuota que empeora con
# la insistencia, sino un hiccup transitorio de sesión/red.
#
# IMPORTANTE: 'rclone check' sale con código != 0 CADA VEZ que encuentra
# diferencias entre origen y destino (ej. "sizes differ") -- eso es su
# comportamiento normal y esperado, no un error. El script original tenía
# un "|| true" en el check justamente por esto. Si acá tratáramos
# cualquier exit code != 0 como "fallo técnico a reintentar", terminamos
# reintentando 3 veces contra las MISMAS diferencias de contenido (que no
# cambian entre intentos) y abortando el destino entero sin sincronizar
# nada -- pasó en un run real: 45 archivos con "sizes differ" (7% del
# total, muy por debajo del umbral de alerta) hicieron que se descartara
# todo el backup de MiUnidad como si fuera un fallo de login.
#
# Por eso NO alcanza con mirar el exit code: hay que inspeccionar el
# stderr y solo reintentar cuando aparece una firma de error técnico real
# (login/red), no cuando el motivo es simplemente que check encontró
# diferencias de contenido.
PATRON_ERROR_TECNICO_RCLONE='Failed to create file system|login with previous auth keys failed|unexpected end of JSON input|couldn.?t connect|no such host|Temporary failure in name resolution|context deadline exceeded|i/o timeout|connection refused|TLS handshake timeout|network is unreachable'

# Devuelve:
#   0 -> el comando terminó con exit 0
#   1 -> el comando terminó con exit != 0 pero SIN firma de error técnico
#        (ej. rclone check reportando diferencias de contenido -- no es
#        un fallo, es información que el caller debe seguir procesando)
#   2 -> se agotaron los reintentos tras detectar fallos técnicos reales
ejecutar_rclone_con_reintentos() {
  local descripcion="$1"
  shift

  local intento tmp_stderr rc
  for ((intento = 1; intento <= INTENTOS_RCLONE; intento++)); do
    tmp_stderr="$(mktemp)"

    if "$@" 2>"$tmp_stderr"; then
      cat "$tmp_stderr" >&2
      rm -f "$tmp_stderr"
      return 0
    fi
    rc=$?

    cat "$tmp_stderr" >&2

    if ! grep -qE "$PATRON_ERROR_TECNICO_RCLONE" "$tmp_stderr"; then
      # Exit code != 0 pero sin firma de error técnico: no es un fallo
      # transitorio de login/red, así que no tiene sentido reintentar
      # (el resultado no va a cambiar). Se devuelve tal cual para que el
      # caller decida qué hacer con esa salida.
      rm -f "$tmp_stderr"
      return 1
    fi

    rm -f "$tmp_stderr"

    if [[ "$intento" -lt "$INTENTOS_RCLONE" ]]; then
      echo "⏳ Fallo técnico en '${descripcion}' (intento ${intento}/${INTENTOS_RCLONE}), reintentando en ${ESPERA_REINTENTO_RCLONE}s..."
      sleep "$ESPERA_REINTENTO_RCLONE"
    fi
  done

  echo "❌ '${descripcion}' falló tras ${INTENTOS_RCLONE} intentos (error técnico: login/red)."
  return 2
}

# ---------------------------------------------------------------------------
# sync_uno <root_folder_id_o_vacio> <subcarpeta_destino_en_mega>
#
# 1) Compara origen vs. lo que ya hay en Mega y clasifica: nuevos, iguales,
#    cambiados de contenido, borrados.
# 2) Si el % de "cambiados" es sospechosamente alto -> ABORTA sin tocar nada.
# 3) Si está todo dentro de lo normal -> corre el sync real, con
#    --backup-dir, así CUALQUIER archivo que se borre o se sobreescriba
#    queda versionado en _papelera/<fecha> antes de perderse.
#
# Devuelve:
#   0 -> éxito
#   1 -> abortado por el umbral de cambios sospechosos (posible ransomware)
#   2 -> falló por error técnico (login/red) tras agotar reintentos
# En los tres casos el caller sigue con las demás carpetas en vez de
# cortar todo el script.
# ---------------------------------------------------------------------------
sync_uno() {
  local root_id="$1"
  local destino="$2"

  local extra_args=()
  if [[ -n "$root_id" ]]; then
    extra_args+=(--drive-root-folder-id "$root_id")
  fi

  local dest_path="${MEGA_REMOTE}${DEST_BASE}/${destino}"
  local papelera_path="${MEGA_REMOTE}${PAPELERA_BASE}/${destino}/${FECHA}"
  local tmp_check
  tmp_check="$(mktemp)"

  echo ""
  echo "=================================================================="
  echo "=== Backup: ${destino}"
  echo "=================================================================="

  # rclone check compara origen (source) contra lo que ya hay en Mega y
  # clasifica cada archivo con un código:
  #   =  idéntico          +  nuevo (solo en origen)
  #   *  cambió de contenido   -  ya no está en origen (se borraría)
  # --one-way evita que marque como "faltante" cosas que Mega tiene de más
  # por otro motivo; acá solo nos interesa el punto de vista del origen.
  #
  # Con reintentos: antes un fallo de login/red acá se tragaba con
  # "|| true" y seguía como si no hubiese pasado nada, calculando el %
  # de cambios sobre un tmp_check vacío o parcial (0/0 = 0%, "todo
  # normal" cuando en realidad no se pudo ni comparar).
  #
  # OJO: 'rclone check' devuelve exit != 0 cada vez que encuentra
  # diferencias de contenido -- eso es NORMAL, no un fallo. Por eso acá
  # solo se aborta el destino si el wrapper devuelve 2 (error técnico
  # real de login/red tras reintentos). Si devuelve 1 (diferencias
  # encontradas, sin firma de error técnico) se sigue procesando: es
  # justo el resultado que --combined guardó en tmp_check para contar el
  # % de cambios más abajo.
  # OJO: llamar a esto como sentencia suelta y capturar "$?" en la línea
  # siguiente NO es seguro bajo 'set -e' (activo al principio del script):
  # si la función devuelve != 0, errexit corta el script ACÁ MISMO, antes
  # de que la línea "local rc_check=\$?" llegue a ejecutarse -- pasó en un
  # run real: el check encontró diferencias normales (rc=1, válido y
  # esperado), pero el script murió ahí mismo sin seguir procesando ni
  # los demás destinos. Por eso se usa "|| rc_check=\$?": al ser parte de
  # una lista con "||" (y no el último comando de la lista), está exento
  # de errexit.
  local rc_check=0
  ejecutar_rclone_con_reintentos "check ${destino}" \
      rclone check "$GDRIVE_REMOTE" "$dest_path" \
      "${extra_args[@]}" \
      --drive-export-formats "$EXPORT_FORMATS" \
      "${FILTROS[@]}" \
      --one-way --combined "$tmp_check" || rc_check=$?

  if [[ "$rc_check" -eq 2 ]]; then
    echo "❌ No se pudo comparar origen/destino de '${destino}' (error técnico)."
    echo "❌ Se omite este destino SIN tocar nada en Mega."
    rm -f "$tmp_check"
    return 2
  fi

  local total cambiados pct
  total=$(grep -c -E '^[=*+]' "$tmp_check" || true)   # todo lo que hay en origen
  cambiados=$(grep -c '^\* ' "$tmp_check" || true)     # cambió de contenido
  pct=0
  if [[ "$total" -gt 0 ]]; then
    pct=$(( cambiados * 100 / total ))
  fi

  echo "🔍 ${cambiados}/${total} archivo(s) cambiaron de contenido (${pct}%)."

  if [[ "$total" -gt 0 && "$pct" -ge "$UMBRAL_ALERTA_PCT" ]]; then
    echo "🚨 ALERTA: ${pct}% de los archivos cambiaron de golpe (umbral: ${UMBRAL_ALERTA_PCT}%)."
    echo "🚨 Esto es sospechoso de ransomware/virus corrompiendo archivos en masa."
    echo "🚨 Se ABORTA el backup de '${destino}' SIN sobreescribir nada. Revisar manualmente"
    echo "🚨 el origen antes de volver a correr el backup."
    rm -f "$tmp_check"
    return 1
  fi

  rm -f "$tmp_check"

  # Sync real: agrega lo nuevo, actualiza lo cambiado, borra lo que ya no
  # está en origen -- pero TODO lo que se borra o se sobreescribe queda
  # versionado en _papelera/<fecha> gracias a --backup-dir.
  #
  # A diferencia de check, 'rclone sync' no tiene un caso "normal" de
  # exit != 0: cualquier fallo acá (técnico o no) es un problema real,
  # así que ambos códigos de retorno del wrapper (1 y 2) se tratan igual.
  # Mismo motivo que en el check de arriba para usar "|| rc_sync=\$?" en
  # vez de una sentencia suelta seguida de "local rc_sync=\$?": bajo
  # set -e, la sentencia suelta corta el script en el momento del fallo.
  local rc_sync=0
  ejecutar_rclone_con_reintentos "sync ${destino}" \
      rclone sync "$GDRIVE_REMOTE" "$dest_path" \
      "${extra_args[@]}" \
      --drive-export-formats "$EXPORT_FORMATS" \
      "${FILTROS[@]}" \
      --backup-dir "$papelera_path" \
      --max-delete 200 \
      --transfers 4 \
      --checkers 8 \
      --fast-list \
      --stats 30s \
      --log-level INFO || rc_sync=$?

  if [[ "$rc_sync" -ne 0 ]]; then
    echo "❌ El sync de '${destino}' falló (código ${rc_sync}: 2=error técnico tras reintentos, 1=otro fallo de rclone)."
    return 2
  fi

  return 0
}

# ---------------------------------------------------------------------------
# podar_papelera <subcarpeta_destino_en_mega>
#
# Borra, dentro de _papelera/<destino>, las carpetas de fecha (DDMMYYYY)
# más viejas que RETENCION_PAPELERA_DIAS. Las carpetas de fecha son las que
# arma sync_uno vía --backup-dir en cada corrida.
#
# Se compara por nombre de carpeta (fecha de backup), no por --min-age de
# rclone, porque --min-age filtraría por la fecha de modificación ORIGINAL
# de cada archivo (la que traía de Drive antes de ser reemplazado/borrado),
# no por la fecha en que cayó a la papelera -- que es lo que nos interesa
# para esta poda.
#
# OJO: el nombre de carpeta está en DDMMYYYY (para que sea legible a simple
# vista), pero eso NO se puede comparar como texto para saber qué es "más
# viejo" (ej. "05012025" ordena antes que "31122024" alfabéticamente, pero
# es una fecha posterior). Por eso cada carpeta se reordena a YYYYMMDD antes
# de compararla contra la fecha límite.
# ---------------------------------------------------------------------------
podar_papelera() {
  local destino="$1"
  local base_path="${MEGA_REMOTE}${PAPELERA_BASE}/${destino}"

  echo "🧹 Podando papelera de '${destino}' (reteniendo ${RETENCION_PAPELERA_DIAS} días)..."

  local fecha_limite
  fecha_limite="$(date -d "-${RETENCION_PAPELERA_DIAS} days" +%Y%m%d)"

  local carpetas
  carpetas="$(rclone lsf --dirs-only "$base_path" 2>/dev/null || true)"

  if [[ -z "$carpetas" ]]; then
    echo "   (sin carpetas de fecha para revisar)"
    return 0
  fi

  local borradas=0
  while IFS= read -r carpeta; do
    carpeta="${carpeta%/}"
    # Solo tocar carpetas con formato de fecha DDMMYYYY (8 dígitos);
    # cualquier otra cosa se ignora por seguridad (no se borra nada que no
    # reconozcamos).
    if [[ "$carpeta" =~ ^[0-9]{8}$ ]]; then
      local dd="${carpeta:0:2}"
      local mm="${carpeta:2:2}"
      local yyyy="${carpeta:4:4}"
      local carpeta_ymd="${yyyy}${mm}${dd}"

      if [[ "$carpeta_ymd" < "$fecha_limite" ]]; then
        echo "   🗑️  Borrando ${destino}/${carpeta} (anterior a $(date -d "-${RETENCION_PAPELERA_DIAS} days" +%d/%m/%Y))"
        if rclone purge "${base_path}/${carpeta}"; then
          borradas=$((borradas + 1))
        else
          echo "   ⚠️  No se pudo borrar ${destino}/${carpeta}"
        fi
      fi
    fi
  done <<< "$carpetas"

  echo "   ✅ ${borradas} carpeta(s) de fecha podada(s) en ${destino}"
}

echo "🚀 Iniciando backup diario Drive (div) -> Mega ($(date))"

# Solo los destinos con sync EXITOSO en esta corrida entran acá, y son los
# únicos que se podan al final. Antes se podaba TODO lo que se había
# intentado backupear, sin importar si el sync realmente corrió -- eso
# hizo que, ante un fallo de login de Mega, se borrara la carpeta de
# papelera del día anterior de un destino que ni siquiera se había
# actualizado en esta corrida, achicando la ventana real de recuperación.
DESTINOS_OK=()

# procesar_destino <root_folder_id_o_vacio> <subcarpeta_destino_en_mega>
# Llama a sync_uno, clasifica el resultado (éxito / ransomware / error
# técnico) y pausa antes del siguiente destino para no encadenar logins
# contra Mega demasiado rápido.
procesar_destino() {
  local root_id="$1"
  local destino="$2"

  # Mismo motivo que en sync_uno: "sync_uno ... ; local rc=\$?" como dos
  # sentencias separadas es inseguro bajo set -e. sync_uno devuelve 1
  # (ransomware) o 2 (error técnico) como resultados NORMALES y
  # esperados, no excepcionales -- si la sentencia fuera suelta, errexit
  # cortaría el script entero en el primer destino que no fuera 100%
  # exitoso, sin llegar a procesar los demás ni a clasificar el motivo.
  local rc=0
  sync_uno "$root_id" "$destino" || rc=$?

  case "$rc" in
    0) DESTINOS_OK+=("$destino") ;;
    1) HUBO_ALERTA_RANSOMWARE=1 ;;
    2) HUBO_ERROR_TECNICO=1 ;;
  esac

  sleep "$PAUSA_ENTRE_DESTINOS"
}

# 1) Mi Unidad (raíz normal del Drive de div)
procesar_destino "" "MiUnidad"

# 2) Computers (hasta 5 discos, definidos vía COMPUTERS_DRIVE_IDS)
if [[ -n "${COMPUTERS_DRIVE_IDS:-}" ]]; then
  IFS=',' read -ra PARES <<< "$COMPUTERS_DRIVE_IDS"
  for par in "${PARES[@]}"; do
    nombre="${par%%:*}"
    folder_id="${par#*:}"
    if [[ -z "$nombre" || -z "$folder_id" || "$nombre" == "$folder_id" ]]; then
      echo "⚠️  Entrada inválida en COMPUTERS_DRIVE_IDS, se salta: '$par'"
      continue
    fi
    procesar_destino "$folder_id" "Computers/${nombre}"
  done
else
  echo "ℹ️  COMPUTERS_DRIVE_IDS no está definida: se omite el backup de Computers."
fi

echo ""
echo "=================================================================="
echo "=== Poda de papelera (retención: ${RETENCION_PAPELERA_DIAS} días)"
echo "=================================================================="
if [[ ${#DESTINOS_OK[@]} -eq 0 ]]; then
  echo "ℹ️  Ningún destino tuvo sync exitoso en esta corrida: se omite la poda por completo."
else
  for destino in "${DESTINOS_OK[@]}"; do
    podar_papelera "$destino"
  done
fi

echo ""
echo "📦 Tamaño actual de la papelera (${PAPELERA_BASE}):"
rclone size "${MEGA_REMOTE}${PAPELERA_BASE}" || echo "⚠️  No se pudo calcular el tamaño de la papelera"

echo ""
if [[ "$HUBO_ALERTA_RANSOMWARE" -eq 1 ]]; then
  echo "🚨 Backup diario TERMINADO CON ALERTA DE RANSOMWARE ($(date))"
  echo "🚨 Al menos una carpeta se abortó por % de cambios sospechoso. Revisar el log arriba"
  echo "🚨 y el origen en Drive antes de forzar el backup de nuevo."
fi
if [[ "$HUBO_ERROR_TECNICO" -eq 1 ]]; then
  echo "⚠️  Backup diario con ERRORES TÉCNICOS ($(date))"
  echo "⚠️  Al menos un destino falló por login/red tras agotar reintentos. Esto NO es indicio"
  echo "⚠️  de ransomware -- revisar conectividad/credenciales de Mega (rclone.conf, rate limits)."
fi
if [[ "$HUBO_ALERTA_RANSOMWARE" -eq 0 && "$HUBO_ERROR_TECNICO" -eq 0 ]]; then
  echo "✅ Backup diario completo, sin alertas ($(date))"
fi

# Exit code distinto según el tipo de problema, para que quien lea el job
# de GitHub Actions (rojo = algo pasó) pueda distinguir en el log si hace
# falta auditar archivos (ransomware, código 1) o solo reintentar/revisar
# conectividad (error técnico, código 2), en vez de asumir siempre lo peor.
if [[ "$HUBO_ALERTA_RANSOMWARE" -eq 1 ]]; then
  exit 1
elif [[ "$HUBO_ERROR_TECNICO" -eq 1 ]]; then
  exit 2
else
  exit 0
fi