#!/usr/bin/env python3
"""Microsoft 365 ↔ Google Calendar Sync — CLI entry point.

Usage
-----
  python main.py --setup          # Run the interactive setup wizard
  python main.py --sync           # One-time sync
  python main.py --dry-run        # Preview what would be synced (no changes)
  python main.py --watch          # Continuous sync (polls every N minutes)
  python main.py --watch --interval 5
  python main.py --status         # Show last sync time, pair count, recent errors
"""
from __future__ import annotations

import argparse
import sys
import time
from datetime import datetime
from pathlib import Path

from rich.console import Console
from rich.panel import Panel
from rich.table import Table

# Make the project root importable regardless of working directory
sys.path.insert(0, str(Path(__file__).parent))

from src.config import config_exists, load_config, save_config
from src.utils.logging import get_logger, init_logger

console = Console()


# ── Command handlers ──────────────────────────────────────────────────────────

def cmd_setup(args: argparse.Namespace) -> int:
    from src.setup_wizard import run_setup

    try:
        run_setup()
        return 0
    except SystemExit as exc:
        return int(exc.code) if exc.code is not None else 1
    except KeyboardInterrupt:
        console.print("\n[yellow]Setup cancelled.[/yellow]")
        return 1


def cmd_sync(args: argparse.Namespace) -> int:
    if not config_exists():
        console.print(
            "[red]No configuration found. Run [bold]python main.py --setup[/bold] first.[/red]"
        )
        return 1

    config = load_config()
    init_logger(config["sync"].get("log_dir", "logs"))
    logger = get_logger()

    dry = getattr(args, "dry_run", False)
    if dry:
        console.print("[bold yellow]DRY RUN — no changes will be made.[/bold yellow]\n")

    try:
        from src.sync.engine import run

        summary = run(config, dry_run=dry)
        # Persist any refreshed tokens back to disk
        if not dry:
            save_config(config)
        _print_summary(summary)
        return 0 if summary.errors == 0 else 2
    except Exception as exc:
        logger.exception("Sync failed: %s", exc)
        console.print(f"[red bold]Sync failed:[/red bold] {exc}")
        return 1


def cmd_watch(args: argparse.Namespace) -> int:
    if not config_exists():
        console.print(
            "[red]No configuration found. Run [bold]python main.py --setup[/bold] first.[/red]"
        )
        return 1

    config = load_config()
    interval: int = getattr(args, "interval", None) or config["sync"].get(
        "interval_minutes", 15
    )
    init_logger(config["sync"].get("log_dir", "logs"))
    logger = get_logger()

    console.print(
        Panel(
            f"[bold]Watch mode[/bold] — syncing every [cyan]{interval}[/cyan] minute(s).\n"
            "Press [bold]Ctrl+C[/bold] to stop.",
            border_style="cyan",
        )
    )

    from src.sync.engine import run

    cycle = 0
    try:
        while True:
            cycle += 1
            ts = datetime.now().strftime("%H:%M:%S")
            console.rule(f"[dim]Cycle {cycle} — {ts}[/dim]")
            try:
                summary = run(config, dry_run=False)
                save_config(config)
                _print_summary(summary)
            except Exception as exc:
                logger.error("Cycle %d failed: %s", cycle, exc)
                console.print(f"[red]Cycle {cycle} error:[/red] {exc}")
            console.print(f"[dim]Next sync in {interval} minute(s).[/dim]")
            time.sleep(interval * 60)
    except KeyboardInterrupt:
        console.print("\n[yellow]Watch mode stopped.[/yellow]")
        return 0


def cmd_status(args: argparse.Namespace) -> int:
    if not config_exists():
        console.print("[red]No configuration found.[/red]")
        return 1

    config = load_config()
    db_path = config["sync"].get("state_db_path", "sync_state.db")

    from src.sync.state import SyncStateDB

    try:
        db = SyncStateDB(db_path)
        pair_count = db.pair_count()
        last_sync = db.get_metadata("last_sync_time") or "Never"
    except Exception as exc:
        console.print(f"[red]Could not read sync state: {exc}[/red]")
        return 1

    ms_cfg = config["microsoft"]
    g_cfg = config["google"]

    table = Table(title="Sync Status", show_header=False, box=None)
    table.add_column("Key", style="cyan", min_width=24)
    table.add_column("Value")

    table.add_row("Last sync", last_sync)
    table.add_row("Tracked pairs", str(pair_count))
    table.add_row("Outlook calendar ID", ms_cfg.get("calendar_id") or "[dim]not set[/dim]")
    table.add_row(
        "Outlook sync categories",
        ", ".join(ms_cfg.get("sync_categories", [])) or "[dim]all[/dim]",
    )
    table.add_row("Google calendar ID", g_cfg.get("calendar_id") or "[dim]not set[/dim]")
    color_filter = g_cfg.get("color_filter", [])
    table.add_row(
        "Google colour filter",
        ", ".join(str(c) for c in color_filter) or "[dim]all[/dim]",
    )
    table.add_row(
        "Sync window",
        f"-{config['sync'].get('lookback_days', 30)}d / "
        f"+{config['sync'].get('lookahead_days', 365)}d",
    )

    console.print(table)

    # Tail today's error lines from the log file
    log_dir = Path(config["sync"].get("log_dir", "logs"))
    today_log = log_dir / f"sync-{datetime.now().strftime('%Y-%m-%d')}.log"
    errors: list[str] = []
    if today_log.exists():
        with today_log.open("r", encoding="utf-8") as f:
            for line in f:
                if "ERROR" in line:
                    errors.append(line.rstrip())

    if errors:
        console.print(f"\n[red]Errors today ({len(errors)}):[/red]")
        for line in errors[-5:]:
            console.print(f"  [dim]{line}[/dim]")
    else:
        console.print("\n[green]No errors logged today.[/green]")

    return 0


# ── Output helpers ────────────────────────────────────────────────────────────

def _print_summary(summary) -> None:
    table = Table(title="Sync Summary")
    table.add_column("Action", style="bold")
    table.add_column("Count", justify="right")

    table.add_row("[green]Created[/green]", str(summary.created))
    table.add_row("[yellow]Updated[/yellow]", str(summary.updated))
    table.add_row("[red]Deleted[/red]", str(summary.deleted))
    table.add_row("[dim]Skipped[/dim]", str(summary.skipped))
    if summary.errors:
        table.add_row("[red bold]Errors[/red bold]", str(summary.errors))
    if summary.conflicts:
        table.add_row("[magenta]Conflicts[/magenta]", str(len(summary.conflicts)))

    console.print(table)


# ── Argument parser ───────────────────────────────────────────────────────────

def _build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="calendar-sync",
        description="Bi-directional Microsoft 365 ↔ Google Calendar sync",
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )

    group = parser.add_mutually_exclusive_group(required=True)
    group.add_argument(
        "--setup",
        action="store_true",
        help="Run the interactive setup wizard",
    )
    group.add_argument(
        "--sync",
        action="store_true",
        help="Run a one-time sync",
    )
    group.add_argument(
        "--watch",
        action="store_true",
        help="Run continuously, polling on an interval",
    )
    group.add_argument(
        "--dry-run",
        action="store_true",
        dest="dry_run",
        help="Show what would be synced without making any changes",
    )
    group.add_argument(
        "--status",
        action="store_true",
        help="Show sync status (last run, pair count, recent errors)",
    )

    parser.add_argument(
        "--interval",
        type=int,
        metavar="MINUTES",
        help="Polling interval for --watch mode (overrides config)",
    )

    return parser


def main() -> int:
    parser = _build_parser()
    args = parser.parse_args()

    if args.setup:
        return cmd_setup(args)
    if args.sync:
        return cmd_sync(args)
    if args.watch:
        return cmd_watch(args)
    if args.dry_run:
        return cmd_sync(args)  # cmd_sync reads args.dry_run
    if args.status:
        return cmd_status(args)
    return 0


if __name__ == "__main__":
    sys.exit(main())
