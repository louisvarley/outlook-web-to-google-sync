"""Google Calendar OAuth authentication.

Uses InstalledAppFlow (auth code + PKCE) with a random localhost redirect.
The deprecated implicit/token flow is intentionally not used.

Tokens are serialised into config["google"]["token"] so they survive process
restarts.  Expired tokens with a refresh_token are refreshed automatically;
the full interactive flow runs only when no valid credentials exist.
"""
from __future__ import annotations

import logging
from typing import Any

from google.auth.transport.requests import Request  # type: ignore
from google.oauth2.credentials import Credentials  # type: ignore
from google_auth_oauthlib.flow import InstalledAppFlow  # type: ignore

logger = logging.getLogger("calendar_sync")

_SCOPES = ["https://www.googleapis.com/auth/calendar"]


def _creds_from_config(token_data: dict) -> Credentials | None:
    if not token_data:
        return None
    try:
        return Credentials.from_authorized_user_info(token_data, _SCOPES)
    except Exception:
        return None


def _creds_to_dict(creds: Credentials) -> dict:
    return {
        "token": creds.token,
        "refresh_token": creds.refresh_token,
        "token_uri": creds.token_uri,
        "client_id": creds.client_id,
        "client_secret": creds.client_secret,
        "scopes": list(creds.scopes) if creds.scopes else _SCOPES,
    }


def get_credentials(g_config: dict[str, Any]) -> Credentials:
    """Return valid Google OAuth credentials.

    Tries the following in order:
      1. Existing valid credentials from config.
      2. Refresh expired credentials using the stored refresh_token.
      3. Full interactive OAuth flow (opens browser).

    Mutates g_config["token"] in-place with updated token data.

    Raises:
        RuntimeError: if the interactive flow fails or is cancelled.
    """
    creds = _creds_from_config(g_config.get("token", {}))

    if creds and creds.valid:
        return creds

    if creds and creds.expired and creds.refresh_token:
        try:
            creds.refresh(Request())
            g_config["token"] = _creds_to_dict(creds)
            logger.debug("Google token refreshed successfully.")
            return creds
        except Exception as exc:
            logger.warning("Google token refresh failed (%s); re-authorising…", exc)

    # Full interactive flow
    client_config = {
        "installed": {
            "client_id": g_config["client_id"],
            "client_secret": g_config["client_secret"],
            "redirect_uris": ["http://localhost"],
            "auth_uri": "https://accounts.google.com/o/oauth2/auth",
            "token_uri": "https://oauth2.googleapis.com/token",
        }
    }
    flow = InstalledAppFlow.from_client_config(client_config, _SCOPES)
    creds = flow.run_local_server(port=0, prompt="consent", access_type="offline")
    g_config["token"] = _creds_to_dict(creds)
    return creds


def clear_token(g_config: dict[str, Any]) -> None:
    """Wipe the stored Google token to force re-authentication."""
    g_config["token"] = {}
