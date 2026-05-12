"""Microsoft 365 OAuth authentication via MSAL.

Supports two flows:
  - Device code flow   → personal accounts (PublicClientApplication)
  - Auth code + PKCE   → work/Exchange accounts (ConfidentialClientApplication)

The account_type stored in config["microsoft"]["account_type"] ("personal" | "work")
determines which flow is used.  get_token() tries a silent cache hit first; the
interactive flow is only triggered when the cache is empty or the refresh token
has expired.
"""
from __future__ import annotations

import json
import logging
import webbrowser
from typing import Any, Optional

import msal  # type: ignore

logger = logging.getLogger("calendar_sync")

GRAPH_SCOPES = ["Calendars.ReadWrite", "offline_access"]

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

    Performs a silent cache lookup first; falls back to the appropriate
    interactive flow.  Mutates ms_config["token_cache"] in-place.

    Raises:
        RuntimeError: if authentication fails for any reason.
    """
    account_type = ms_config.get("account_type", "personal")
    client_id = ms_config["client_id"]
    tenant_id = ms_config.get("tenant_id", "")
    client_secret = ms_config.get("client_secret", "")
    authority = _build_authority(tenant_id, account_type)
    cache = _deserialise_cache(ms_config.get("token_cache", {}))

    # Choose application type
    use_confidential = account_type == "work" and bool(client_secret)

    if use_confidential:
        app: msal.ClientApplication = msal.ConfidentialClientApplication(
            client_id=client_id,
            client_credential=client_secret,
            authority=authority,
            token_cache=cache,
        )
    else:
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

    # Interactive acquisition
    if use_confidential and isinstance(app, msal.ConfidentialClientApplication):
        result = _auth_code_pkce_flow(app, client_id, authority)
    else:
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


def _auth_code_pkce_flow(
    app: msal.ConfidentialClientApplication,
    client_id: str,
    authority: str,
) -> dict:
    """Auth code + PKCE flow: opens a browser and listens on a random localhost port."""
    import http.server
    import threading
    import urllib.parse
    from rich.console import Console

    console = Console()

    received_code: list[Optional[str]] = [None]
    received_state: list[Optional[str]] = [None]
    server_ready = threading.Event()

    class _CallbackHandler(http.server.BaseHTTPRequestHandler):
        def do_GET(self) -> None:  # noqa: N802
            parsed = urllib.parse.urlparse(self.path)
            params = urllib.parse.parse_qs(parsed.query)
            received_code[0] = params.get("code", [None])[0]
            received_state[0] = params.get("state", [None])[0]
            self.send_response(200)
            self.send_header("Content-Type", "text/html; charset=utf-8")
            self.end_headers()
            self.wfile.write(
                b"<html><body><h2>Authentication complete. You may close this window.</h2></body></html>"
            )

        def log_message(self, format: str, *args: Any) -> None:  # noqa: A002
            pass  # Suppress HTTP server console noise

    httpd = http.server.HTTPServer(("localhost", 0), _CallbackHandler)
    port = httpd.server_address[1]
    redirect_uri = f"http://localhost:{port}"

    server_thread = threading.Thread(target=httpd.handle_request, daemon=True)
    server_thread.start()

    flow = app.initiate_auth_code_flow(scopes=GRAPH_SCOPES, redirect_uri=redirect_uri)
    auth_url = flow["auth_uri"]

    console.print(f"\n[bold cyan]Opening browser for Microsoft authentication…[/bold cyan]")
    console.print(f"If the browser doesn't open automatically, visit:\n  [link={auth_url}]{auth_url}[/link]\n")
    webbrowser.open(auth_url)
    server_thread.join(timeout=120)

    if received_code[0] is None:
        raise RuntimeError("Microsoft authentication timed out or was cancelled.")

    return app.acquire_token_by_auth_code_flow(
        flow,
        {"code": received_code[0], "state": received_state[0]},
    )


def _persist_cache(app: msal.ClientApplication, ms_config: dict[str, Any]) -> None:
    if app.token_cache.has_state_changed:
        ms_config["token_cache"] = json.loads(app.token_cache.serialize())


def clear_token_cache(ms_config: dict[str, Any]) -> None:
    """Wipe the stored MSAL token cache to force re-authentication."""
    ms_config["token_cache"] = {}
