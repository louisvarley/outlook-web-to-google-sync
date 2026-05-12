"""Microsoft 365 OAuth authentication via MSAL.

Uses device code flow for both personal and work/Exchange accounts
(PublicClientApplication).  No redirect URI registration is required.

get_token() tries a silent cache hit first; the interactive device code
prompt is only triggered when the cache is empty or the refresh token
has expired.
"""
from __future__ import annotations

import json
import logging
import webbrowser
from typing import Any

import msal  # type: ignore

logger = logging.getLogger("calendar_sync")

GRAPH_SCOPES = ["https://graph.microsoft.com/Calendars.ReadWrite"]

_AUTHORITY_PERSONAL = "https://login.microsoftonline.com/consumers"
_AUTHORITY_COMMON = "https://login.microsoftonline.com/common"


def _build_authority(tenant_id: str, account_type: str) -> str:
    if account_type == "personal":
        return _AUTHORITY_PERSONAL
    if tenant_id and tenant_id not in ("", "common", "consumers", "organizations"):
        return f"https://login.microsoftonline.com/{tenant_id}"
    return _AUTHORITY_COMMON


def _deserialise_cache(token_cache_data: dict) -> msal.SerializableTokenCache:
    cache = msal.SerializableTokenCache()
    if token_cache_data:
        cache.deserialize(json.dumps(token_cache_data))
    return cache


def get_token(ms_config: dict[str, Any]) -> str:
    """Return a valid Microsoft Graph access token.

    Performs a silent cache lookup first; falls back to device code flow.
    Mutates ms_config["token_cache"] in-place.

    Raises:
        RuntimeError: if authentication fails for any reason.
    """
    account_type = ms_config.get("account_type", "personal")
    client_id = ms_config["client_id"]
    tenant_id = ms_config.get("tenant_id", "")
    authority = _build_authority(tenant_id, account_type)
    cache = _deserialise_cache(ms_config.get("token_cache", {}))

    app = msal.PublicClientApplication(
        client_id=client_id,
        authority=authority,
        token_cache=cache,
    )

    # Silent acquisition
    accounts = app.get_accounts()
    if accounts:
        result = app.acquire_token_silent(GRAPH_SCOPES, account=accounts[0])
        if result and "access_token" in result:
            _persist_cache(app, ms_config)
            return result["access_token"]

    # Interactive acquisition via device code (no redirect URI required)
    result = _device_code_flow(app)

    if "access_token" not in result:
        error = result.get("error_description") or result.get("error", "Unknown error")
        raise RuntimeError(f"Microsoft authentication failed: {error}")

    _persist_cache(app, ms_config)
    return result["access_token"]


def _device_code_flow(app: msal.PublicClientApplication) -> dict:
    """Prompt the user with a device code; no browser redirect needed."""
    from rich.console import Console
    from rich.panel import Panel

    console = Console()
    flow = app.initiate_device_flow(scopes=GRAPH_SCOPES)
    if "user_code" not in flow:
        raise RuntimeError(
            f"Failed to initiate device code flow: {flow.get('error_description', flow)}"
        )

    console.print(
        Panel(
            f"[bold]Microsoft Authentication Required[/bold]\n\n"
            f"1. Visit: [link={flow['verification_uri']}]{flow['verification_uri']}[/link]\n"
            f"2. Enter code: [bold yellow]{flow['user_code']}[/bold yellow]\n\n"
            f"Waiting for you to complete sign-in…",
            title="[cyan]Sign In[/cyan]",
            border_style="cyan",
        )
    )
    webbrowser.open(flow["verification_uri"])
    return app.acquire_token_by_device_flow(flow)





def _persist_cache(app: msal.ClientApplication, ms_config: dict[str, Any]) -> None:
    if app.token_cache.has_state_changed:
        ms_config["token_cache"] = json.loads(app.token_cache.serialize())


def clear_token_cache(ms_config: dict[str, Any]) -> None:
    """Wipe the stored MSAL token cache to force re-authentication."""
    ms_config["token_cache"] = {}
