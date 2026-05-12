"""Interactive setup wizard for Microsoft 365 ↔ Google Calendar sync.

Walks the user through:
  1. Microsoft Entra ID app registration (printed guide + credential input)
  2. Microsoft calendar selection
  3. Outlook category selection (from master list + event scan)
  4. Google Cloud Console setup (printed guide + credential input)
  5. Google calendar selection
  6. Google colour filter (optional)
  7. Sync options (lookback/lookahead days, watch interval)
  8. Writes config.json

On re-run the wizard detects existing config and offers to reuse settings.
"""
from __future__ import annotations

import copy
import logging
from datetime import datetime, timedelta, timezone
from typing import Any

from rich.console import Console
from rich.panel import Panel
from rich.prompt import Confirm, Prompt
from rich.table import Table

from src.auth.google import clear_token, get_credentials
from src.auth.microsoft import clear_token_cache, get_token
from src.calendars.google import GoogleCalendarClient
from src.calendars.microsoft import GraphClient
from src.config import DEFAULT_CONFIG, config_exists, load_config, save_config

logger = logging.getLogger("calendar_sync")
console = Console()

GOOGLE_COLOR_MAP: dict[int, str] = {
    1: "Tomato",
    2: "Flamingo",
    3: "Tangerine",
    4: "Banana",
    5: "Sage",
    6: "Basil",
    7: "Peacock",
    8: "Blueberry",
    9: "Lavender",
    10: "Grape",
    11: "Graphite",
}

_COLOR_SWATCHES: dict[int, str] = {
    1: "🔴", 2: "🩷", 3: "🟠", 4: "🟡", 5: "🟢",
    6: "🌲", 7: "🔵", 8: "🔷", 9: "💜", 10: "🍇", 11: "⬛",
}


def run_setup() -> dict[str, Any]:
    """Run the full interactive setup wizard and return the saved config."""
    console.print(
        Panel(
            "[bold cyan]Microsoft 365 ↔ Google Calendar Sync[/bold cyan]\n"
            "Interactive Setup Wizard",
            subtitle="Press Ctrl+C at any time to cancel",
            border_style="cyan",
        )
    )

    if config_exists():
        reuse = Confirm.ask(
            "\nExisting configuration found. Reuse current settings where possible?",
            default=True,
        )
        config = load_config() if reuse else copy.deepcopy(DEFAULT_CONFIG)
    else:
        config = copy.deepcopy(DEFAULT_CONFIG)

    console.print()

    config = _setup_microsoft(config)
    console.print()

    config = _setup_google(config)
    console.print()

    config = _setup_sync_options(config)
    console.print()

    save_config(config)

    console.print(
        Panel(
            "[green bold]Configuration saved.[/green bold]\n\n"
            "• One-time sync:  [bold]python main.py --sync[/bold]\n"
            "• Continuous:     [bold]python main.py --watch[/bold]\n"
            "• Preview only:   [bold]python main.py --dry-run[/bold]",
            border_style="green",
        )
    )
    return config


# ── Microsoft 365 ─────────────────────────────────────────────────────────────

def _setup_microsoft(config: dict) -> dict:
    console.print(Panel("[bold]Step 1 of 3 — Microsoft 365 / Outlook[/bold]", border_style="blue"))
    ms: dict = config.setdefault("microsoft", {})

    show_guide = not ms.get("client_id") or Confirm.ask(
        "Show Entra ID app registration guide?", default=False
    )
    if show_guide:
        _print_entra_guide()

    # Account type
    console.print("\n[bold]Account type:[/bold]")
    console.print("  [cyan]1[/cyan]  Personal Microsoft account (@outlook.com, @hotmail.com, @live.com)")
    console.print("  [cyan]2[/cyan]  Work / Organisation account (Microsoft 365 / Exchange)")
    default_acct = "1" if ms.get("account_type", "personal") == "personal" else "2"
    acct_choice = Prompt.ask("Select", choices=["1", "2"], default=default_acct)
    ms["account_type"] = "personal" if acct_choice == "1" else "work"

    ms["client_id"] = Prompt.ask(
        "Application (Client) ID",
        default=ms.get("client_id", ""),
    ).strip()

    if ms["account_type"] == "work":
        ms["tenant_id"] = Prompt.ask(
            "Directory (Tenant) ID",
            default=ms.get("tenant_id", ""),
        ).strip()
        console.print("[dim]Client Secret is required for work / organisation accounts.[/dim]")
        secret = Prompt.ask(
            "Client Secret",
            default="(keep existing)" if ms.get("client_secret") else "",
            password=True,
        )
        if secret and secret != "(keep existing)":
            ms["client_secret"] = secret
    else:
        ms.setdefault("tenant_id", "consumers")
        ms.setdefault("client_secret", "")

    # Validate credentials with a live auth attempt
    console.print("\nTesting Microsoft authentication…")
    try:
        clear_token_cache(ms)
        get_token(ms)
        console.print("[green]✓ Microsoft authentication successful.[/green]")
    except Exception as exc:
        console.print(f"[red]✗ Authentication failed: {exc}[/red]")
        if not Confirm.ask("Continue setup anyway?", default=False):
            raise SystemExit(1)

    config = _select_microsoft_calendar(config)
    config = _select_microsoft_categories(config)
    return config


def _print_entra_guide() -> None:
    console.print(
        Panel(
            "[bold cyan]Microsoft Entra ID — App Registration[/bold cyan]\n\n"
            "[bold]1.[/bold] Open [link=https://entra.microsoft.com]https://entra.microsoft.com[/link]\n"
            "[bold]2.[/bold] Go to [cyan]App registrations[/cyan] → [cyan]New registration[/cyan]\n"
            "[bold]3.[/bold] Name your app (e.g. [italic]Calendar Sync[/italic])\n"
            "[bold]4.[/bold] Supported account types:\n"
            "   • Personal account → [italic]Personal Microsoft accounts only[/italic]\n"
            "   • Work account     → [italic]Accounts in this organizational directory only[/italic]\n"
            "[bold]5.[/bold] Redirect URI: leave blank — device code flow requires none\n"
            "[bold]6.[/bold] Click [cyan]Register[/cyan]\n"
            "[bold]7.[/bold] [cyan]Authentication[/cyan] → [cyan]Advanced settings[/cyan]\n"
            "   → Enable [cyan]Allow public client flows[/cyan] → Save\n"
            "[bold]8.[/bold] [cyan]API permissions[/cyan] → [cyan]Add a permission[/cyan] → Microsoft Graph\n"
            "   → Delegated → [cyan]Calendars.ReadWrite[/cyan] → Add\n"
            "[bold]9.[/bold] If on a business tenant: [cyan]Grant admin consent[/cyan]\n"
            "[bold]10.[/bold] Note your [yellow]Application (client) ID[/yellow] and [yellow]Directory (tenant) ID[/yellow]",
            title="[cyan]Setup Guide[/cyan]",
            border_style="dim",
            padding=(1, 2),
        )
    )
    Prompt.ask("[dim]Press Enter to continue[/dim]", default="")


def _select_microsoft_calendar(config: dict) -> dict:
    ms = config["microsoft"]
    console.print("\nFetching Outlook calendars…")
    try:
        client = GraphClient(token_getter=lambda: get_token(ms))
        calendars = client.list_calendars()
    except Exception as exc:
        console.print(f"[red]Could not list calendars: {exc}[/red]")
        ms["calendar_id"] = Prompt.ask(
            "Enter calendar ID manually", default=ms.get("calendar_id", "")
        ).strip()
        return config

    table = Table(title="Outlook Calendars")
    table.add_column("#", style="cyan", width=4)
    table.add_column("Name")
    table.add_column("ID", style="dim")
    for i, cal in enumerate(calendars, 1):
        badge = " [green](primary)[/green]" if cal.get("isDefaultCalendar") else ""
        table.add_row(str(i), cal.get("name", "") + badge, cal.get("id", ""))
    console.print(table)

    current_idx = next(
        (i + 1 for i, c in enumerate(calendars) if c["id"] == ms.get("calendar_id")),
        None,
    )
    choice = Prompt.ask(
        "Select calendar number", default=str(current_idx) if current_idx else "1"
    )
    try:
        selected = calendars[int(choice) - 1]
        if selected.get("isDefaultCalendar"):
            console.print(
                "[yellow]⚠ Primary calendar selected — sync will affect all events.[/yellow]"
            )
        ms["calendar_id"] = selected["id"]
        console.print(f"[green]✓ Selected: {selected.get('name', selected['id'])}[/green]")
    except (ValueError, IndexError):
        console.print("[red]Invalid selection; keeping existing calendar ID.[/red]")
    return config


def _select_microsoft_categories(config: dict) -> dict:
    ms = config["microsoft"]
    console.print("\nFetching Outlook categories…")
    try:
        client = GraphClient(token_getter=lambda: get_token(ms))
        master_cats = client.get_master_categories()
        master_names = {
            c.get("displayName", "") for c in master_cats if c.get("displayName")
        }
        now = datetime.now(timezone.utc)
        event_cats = set(
            client.get_event_categories(
                ms["calendar_id"],
                now - timedelta(days=90),
                now + timedelta(days=365),
            )
        )
        all_cats = sorted(master_names | event_cats, key=str.casefold)
    except Exception as exc:
        console.print(f"[red]Could not fetch categories: {exc}[/red]")
        raw = Prompt.ask(
            "Enter category names to sync (comma-separated)",
            default=", ".join(ms.get("sync_categories", [])),
        )
        ms["sync_categories"] = [c.strip() for c in raw.split(",") if c.strip()]
        return config

    if not all_cats:
        console.print(
            "[yellow]No categories found. Leave blank to sync all events.[/yellow]"
        )
        raw = Prompt.ask(
            "Category names to sync (comma-separated, blank = all)",
            default=", ".join(ms.get("sync_categories", [])),
        )
        ms["sync_categories"] = [c.strip() for c in raw.split(",") if c.strip()]
        return config

    table = Table(title="Outlook Categories")
    table.add_column("#", style="cyan", width=4)
    table.add_column("Category Name")
    for i, cat in enumerate(all_cats, 1):
        table.add_row(str(i), cat)
    console.print(table)

    current = ", ".join(
        str(all_cats.index(c) + 1)
        for c in ms.get("sync_categories", [])
        if c in all_cats
    )
    selection = Prompt.ask(
        "Select categories to sync (comma-separated numbers, blank = all)",
        default=current or "",
    )
    selected_cats: list[str] = []
    for part in selection.split(","):
        part = part.strip()
        if part.isdigit():
            idx = int(part) - 1
            if 0 <= idx < len(all_cats):
                selected_cats.append(all_cats[idx])

    ms["sync_categories"] = selected_cats
    display = ", ".join(selected_cats) if selected_cats else "(all events)"
    console.print(f"[green]✓ Sync categories: {display}[/green]")
    return config


# ── Google ────────────────────────────────────────────────────────────────────

def _setup_google(config: dict) -> dict:
    console.print(Panel("[bold]Step 2 of 3 — Google Calendar[/bold]", border_style="blue"))
    g: dict = config.setdefault("google", {})

    show_guide = not g.get("client_id") or Confirm.ask(
        "Show Google Cloud Console setup guide?", default=False
    )
    if show_guide:
        _print_google_guide()

    g["client_id"] = Prompt.ask(
        "Google Client ID", default=g.get("client_id", "")
    ).strip()

    secret = Prompt.ask(
        "Google Client Secret",
        default="(keep existing)" if g.get("client_secret") else "",
        password=True,
    )
    if secret and secret != "(keep existing)":
        g["client_secret"] = secret

    console.print("\nTesting Google authentication (your browser will open)…")
    try:
        clear_token(g)
        get_credentials(g)
        console.print("[green]✓ Google authentication successful.[/green]")
    except Exception as exc:
        console.print(f"[red]✗ Authentication failed: {exc}[/red]")
        if not Confirm.ask("Continue setup anyway?", default=False):
            raise SystemExit(1)

    config = _select_google_calendar(config)
    config = _setup_google_color_filter(config)
    return config


def _print_google_guide() -> None:
    console.print(
        Panel(
            "[bold cyan]Google Cloud Console — Setup Steps[/bold cyan]\n\n"
            "[bold]1.[/bold] Open [link=https://console.cloud.google.com]https://console.cloud.google.com[/link]\n"
            "[bold]2.[/bold] Create or select a project\n"
            "[bold]3.[/bold] [cyan]APIs & Services[/cyan] → [cyan]Enable APIs[/cyan] → enable [cyan]Google Calendar API[/cyan]\n"
            "[bold]4.[/bold] [cyan]APIs & Services[/cyan] → [cyan]Credentials[/cyan]\n"
            "   → [cyan]Create credentials[/cyan] → [cyan]OAuth 2.0 Client ID[/cyan]\n"
            "[bold]5.[/bold] Application type: [cyan]Desktop app[/cyan] → name it → Create\n"
            "[bold]6.[/bold] Note the [yellow]Client ID[/yellow] and [yellow]Client Secret[/yellow]\n"
            "[bold]7.[/bold] [cyan]OAuth consent screen[/cyan] → add your Google email as a [cyan]Test user[/cyan]\n"
            "   (Required while the app is in [italic]Testing[/italic] mode)\n"
            "[bold]8.[/bold] Redirect URI [yellow]http://localhost[/yellow] is handled automatically "
            "by the desktop flow",
            title="[cyan]Setup Guide[/cyan]",
            border_style="dim",
            padding=(1, 2),
        )
    )
    Prompt.ask("[dim]Press Enter to continue[/dim]", default="")


def _select_google_calendar(config: dict) -> dict:
    g = config["google"]
    console.print("\nFetching Google calendars…")
    try:
        creds = get_credentials(g)
        client = GoogleCalendarClient(creds)
        calendars = client.list_calendars()
    except Exception as exc:
        console.print(f"[red]Could not list calendars: {exc}[/red]")
        g["calendar_id"] = Prompt.ask(
            "Enter calendar ID manually", default=g.get("calendar_id", "")
        ).strip()
        return config

    table = Table(title="Google Calendars")
    table.add_column("#", style="cyan", width=4)
    table.add_column("Name")
    table.add_column("ID", style="dim")
    for i, cal in enumerate(calendars, 1):
        name = cal.get("summary", "")
        badge = ""
        if cal.get("primary"):
            badge += " [green](primary)[/green]"
        if name.lower() in ("family", "personal"):
            badge += " [yellow]★[/yellow]"
        table.add_row(str(i), name + badge, cal.get("id", ""))
    console.print(table)

    current_idx = next(
        (i + 1 for i, c in enumerate(calendars) if c["id"] == g.get("calendar_id")),
        None,
    )
    choice = Prompt.ask(
        "Select calendar number", default=str(current_idx) if current_idx else "1"
    )
    try:
        selected = calendars[int(choice) - 1]
        g["calendar_id"] = selected["id"]
        console.print(f"[green]✓ Selected: {selected.get('summary', selected['id'])}[/green]")
    except (ValueError, IndexError):
        console.print("[red]Invalid selection; keeping existing calendar ID.[/red]")
    return config


def _setup_google_color_filter(config: dict) -> dict:
    g = config["google"]
    use_filter = Confirm.ask(
        "\nFilter which Google events sync to Outlook by colour?",
        default=False,
    )
    if not use_filter:
        g["color_filter"] = []
        console.print("[dim]All Google events from the selected calendar will be synced.[/dim]")
        return config

    table = Table(title="Google Calendar Colours")
    table.add_column("#", style="cyan", width=4)
    table.add_column("")
    table.add_column("Colour Name")
    for color_id, name in GOOGLE_COLOR_MAP.items():
        table.add_row(str(color_id), _COLOR_SWATCHES.get(color_id, ""), name)
    table.add_row("—", "", "Default (calendar colour, no colorId)")
    console.print(table)

    current = ", ".join(str(c) for c in g.get("color_filter", []))
    selection = Prompt.ask(
        "Select colour numbers to sync (comma-separated, e.g. 5,7; blank = all)",
        default=current or "",
    )
    selected: list[int] = []
    for part in selection.split(","):
        part = part.strip()
        if part.isdigit():
            n = int(part)
            if n in GOOGLE_COLOR_MAP:
                selected.append(n)

    g["color_filter"] = selected
    if selected:
        names = ", ".join(GOOGLE_COLOR_MAP[c] for c in selected)
        console.print(f"[green]✓ Colour filter: {names}[/green]")
    else:
        console.print("[yellow]No valid colours chosen; all events will be synced.[/yellow]")
    return config


# ── Sync options ──────────────────────────────────────────────────────────────

def _setup_sync_options(config: dict) -> dict:
    console.print(Panel("[bold]Step 3 of 3 — Sync Options[/bold]", border_style="blue"))
    sync: dict = config.setdefault("sync", {})

    sync["lookback_days"] = int(
        Prompt.ask(
            "Lookback window (days in the past to sync)",
            default=str(sync.get("lookback_days", 30)),
        )
    )
    sync["lookahead_days"] = int(
        Prompt.ask(
            "Lookahead window (days in the future to sync)",
            default=str(sync.get("lookahead_days", 365)),
        )
    )
    sync["interval_minutes"] = int(
        Prompt.ask(
            "Watch mode polling interval (minutes)",
            default=str(sync.get("interval_minutes", 15)),
        )
    )

    do_initial = Confirm.ask(
        "\nOn the first --sync run, perform a deeper historical sync?",
        default=True,
    )
    if do_initial:
        sync["initial_lookback_days"] = int(
            Prompt.ask(
                "Initial historical lookback (days)",
                default=str(sync.get("initial_lookback_days", 365)),
            )
        )
    else:
        sync["initial_lookback_days"] = sync["lookback_days"]

    return config
