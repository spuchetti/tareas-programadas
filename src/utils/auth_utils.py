"""
Autenticación centralizada.

Arquitectura de credenciales del proyecto:
  - SERVICE ACCOUNT (GDRIVE_JSON): usada por TODOS los bots para leer/escribir
    archivos y hojas que ya existen. No puede crear archivos nuevos en Drive
    fuera de una Unidad Compartida (no la tenemos), pero SÍ puede:
      · leer/descargar/exportar archivos existentes
      · editar el contenido de un Google Sheet existente (values.update, etc)
      · agregar una pestaña (hoja) dentro de un Google Sheet existente
      · listar/buscar archivos en carpetas donde tiene acceso

  - OAUTH (OAUTH_REFRESH_TOKEN): usada ÚNICAMENTE por snapshot_bot.py, que es
    el único proceso que necesita CREAR archivos nuevos en Drive (los
    snapshots la primera vez, y las planillas _registro_agentes_N cuando se
    quedan sin lugar). snapshot_bot corre solo de forma manual.

Ningún otro módulo debería leer GDRIVE_JSON u OAUTH_REFRESH_TOKEN
directamente: todos pasan por acá para que haya un solo lugar donde
diagnosticar problemas de credenciales.
"""

import json
import os

from googleapiclient.discovery import build
from google.oauth2.service_account import Credentials as SACredentials
from google.oauth2.credentials import Credentials as OAuthCredentials
from google.auth.transport.requests import Request

SCOPES_DRIVE_SHEETS = [
    "https://www.googleapis.com/auth/drive",
    "https://www.googleapis.com/auth/spreadsheets",
]


class CredencialesFaltantesError(Exception):
    """Se lanza cuando falta un secret/variable de entorno requerido."""
    pass


# ---------------------------------------------------------------------------
# Service Account (uso general — todos los bots)
# ---------------------------------------------------------------------------

_cache_sa_creds = None


def obtener_credenciales_sa(scopes=None):
    """Arma las credenciales de la Service Account a partir de GDRIVE_JSON."""
    global _cache_sa_creds

    raw = os.getenv("GDRIVE_JSON")
    if not raw:
        raise CredencialesFaltantesError(
            "Falta el secret GDRIVE_JSON (credenciales de la Service Account)."
        )

    if _cache_sa_creds is None:
        try:
            cfg = json.loads(raw)
        except json.JSONDecodeError as e:
            raise CredencialesFaltantesError(
                f"GDRIVE_JSON no es un JSON válido: {e}"
            )
        _cache_sa_creds = SACredentials.from_service_account_info(
            cfg, scopes=scopes or SCOPES_DRIVE_SHEETS
        )

    return _cache_sa_creds


def obtener_email_service_account():
    """Devuelve el email de la Service Account (client_email del JSON), o None."""
    raw = os.getenv("GDRIVE_JSON")
    if not raw:
        return None
    try:
        cfg = json.loads(raw)
        return cfg.get("client_email")
    except Exception:
        return None


def obtener_drive_service_sa():
    creds = obtener_credenciales_sa()
    return build("drive", "v3", credentials=creds, cache_discovery=False)


def obtener_sheets_service_sa():
    creds = obtener_credenciales_sa()
    return build("sheets", "v4", credentials=creds, cache_discovery=False)


# ---------------------------------------------------------------------------
# OAuth (uso exclusivo de snapshot_bot.py)
# ---------------------------------------------------------------------------

_cache_oauth_creds = None


def obtener_credenciales_oauth():
    """
    Arma las credenciales OAuth a partir de OAUTH_REFRESH_TOKEN.
    SOLO debe usarse desde snapshot_bot.py — es el único proceso autorizado
    a crear archivos nuevos en Drive.
    """
    global _cache_oauth_creds

    raw = os.getenv("OAUTH_REFRESH_TOKEN")
    if not raw:
        raise CredencialesFaltantesError(
            "Falta el secret OAUTH_REFRESH_TOKEN (credenciales OAuth de "
            "services.aportes.oser@gmail.com)."
        )

    if _cache_oauth_creds is None:
        try:
            token_data = json.loads(raw)
        except json.JSONDecodeError as e:
            raise CredencialesFaltantesError(
                f"OAUTH_REFRESH_TOKEN no es un JSON válido: {e}"
            )
        creds = OAuthCredentials(
            token=token_data.get("token"),
            refresh_token=token_data["refresh_token"],
            token_uri=token_data["token_uri"],
            client_id=token_data["client_id"],
            client_secret=token_data["client_secret"],
            scopes=token_data.get("scopes") or SCOPES_DRIVE_SHEETS,
        )
        if creds.expired or not creds.token:
            creds.refresh(Request())
        _cache_oauth_creds = creds

    return _cache_oauth_creds


def obtener_drive_service_oauth():
    creds = obtener_credenciales_oauth()
    return build("drive", "v3", credentials=creds, cache_discovery=False)


def obtener_sheets_service_oauth():
    creds = obtener_credenciales_oauth()
    return build("sheets", "v4", credentials=creds, cache_discovery=False)
