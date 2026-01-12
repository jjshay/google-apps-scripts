#!/usr/bin/env python3
"""
Google Apps Scripts - Demo
Shows sample outputs from the 24+ automation scripts with rich visual output.

Run: python demo.py
"""
from __future__ import annotations

import json
from datetime import datetime

try:
    from rich.console import Console
    from rich.table import Table
    from rich.panel import Panel
    from rich.progress import Progress, SpinnerColumn, TextColumn
    from rich import box
    RICH_AVAILABLE = True
except ImportError:
    RICH_AVAILABLE = False

console = Console() if RICH_AVAILABLE else None


def print_header(text: str) -> None:
    if RICH_AVAILABLE:
        console.print()
        console.rule(f"[bold green]{text}[/bold green]", style="green")
        console.print()
    else:
        print(f"\n{'='*60}\n {text}\n{'='*60}\n")


def show_banner() -> None:
    if RICH_AVAILABLE:
        banner = """
[bold cyan]╔═══════════════════════════════════════════════════════════════════╗
║[/bold cyan] [bold gold1]   ____                   _          _                           [/bold gold1][bold cyan]║
║[/bold cyan] [bold gold1]  / ___| ___   ___   __ _| | ___    / \   _ __  _ __  ___        [/bold gold1][bold cyan]║
║[/bold cyan] [bold gold1] | |  _ / _ \ / _ \ / _` | |/ _ \  / _ \ | '_ \| '_ \/ __|       [/bold gold1][bold cyan]║
║[/bold cyan] [bold gold1] | |_| | (_) | (_) | (_| | |  __/ / ___ \| |_) | |_) \__ \       [/bold gold1][bold cyan]║
║[/bold cyan] [bold gold1]  \____|\___/ \___/ \__, |_|\___/_/   \_\ .__/| .__/|___/       [/bold gold1][bold cyan]║
║[/bold cyan] [bold gold1]                    |___/               |_|   |_|  [bold white]Scripts[/bold white]      [/bold gold1][bold cyan]║
║[/bold cyan]                                                                       [bold cyan]║
║[/bold cyan]             [bold white]24+ Automation Scripts for Sheets & Drive[/bold white]             [bold cyan]║
╚═══════════════════════════════════════════════════════════════════╝[/bold cyan]
"""
        console.print(banner)
    else:
        print("\n" + "="*60 + "\n  GOOGLE APPS SCRIPTS\n" + "="*60)


def demo_3dsellers() -> None:
    print_header("3DSELLERS INVENTORY SYNC")

    sample = {
        "image": "death-nyc-marilyn-001.jpg",
        "artist": "Death NYC",
        "title": "Marilyn Monroe x Chanel",
        "sku": "1_DNYC_Marilyn",
        "price": 225.00,
        "status": "Ready"
    }

    if RICH_AVAILABLE:
        console.print("[dim]Auto-populates 30+ fields from artwork images using Claude Vision AI[/dim]\n")

        table = Table(title="📦 Sample Processed Item", box=box.ROUNDED)
        table.add_column("Field", style="cyan")
        table.add_column("Value", style="gold1")

        for k, v in sample.items():
            table.add_row(k, str(v))
        console.print(table)

        console.print(Panel("""
[bold]What it does:[/bold]
  1. Monitors Google Drive for new images
  2. Sends to Claude Vision API for analysis
  3. Auto-fills inventory sheet (artist, title, price, etc.)
  4. Syncs to 3DSellers for eBay/Etsy listing
""", title="🔄 Workflow", border_style="green", box=box.ROUNDED))
    else:
        print(json.dumps(sample, indent=2))


def demo_ai_integration() -> None:
    print_header("AI INTEGRATION SCRIPTS")

    results = [
        ("img_001.jpg", "Shepard Fairey", "94%", "green"),
        ("img_002.jpg", "Death NYC", "97%", "green"),
        ("img_003.jpg", "Banksy", "89%", "green"),
        ("img_004.jpg", "NEEDS REVIEW", "45%", "red"),
        ("img_005.jpg", "Mr. Brainwash", "91%", "green"),
    ]

    if RICH_AVAILABLE:
        console.print("[dim]Connects Google Sheets to Claude, GPT-4, Gemini for image classification[/dim]\n")

        table = Table(title="🤖 AI Classification Results", box=box.ROUNDED)
        table.add_column("Image", style="cyan")
        table.add_column("Classification")
        table.add_column("Confidence", justify="center")

        for img, result, conf, color in results:
            table.add_row(img, result, f"[{color}]{conf}[/{color}]")
        console.print(table)

        console.print("\n[green]Auto-sorted:[/green] 4  |  [yellow]Needs review:[/yellow] 1")
    else:
        for img, result, conf, _ in results:
            print(f"  {img}: {result} ({conf})")


def demo_news_engine() -> None:
    print_header("NEWS ENGINE")

    articles = [
        ("OpenAI Announces GPT-5", 94, "High"),
        ("AI Regulation Bill", 87, "High"),
        ("Gemini API Update", 82, "Medium"),
        ("AI in Healthcare Study", 76, "Medium"),
        ("Tech Stocks Rally", 68, "Low"),
    ]

    if RICH_AVAILABLE:
        console.print("[dim]Fetches and scores articles with 5 AI models for LinkedIn content[/dim]\n")

        table = Table(title="📰 Scored Articles (Top 5)", box=box.ROUNDED)
        table.add_column("Rank", justify="center", width=5)
        table.add_column("Title")
        table.add_column("Score", justify="center")
        table.add_column("Relevance", justify="center")

        for i, (title, score, rel) in enumerate(articles, 1):
            score_bar = "█" * (score // 10) + "░" * (10 - score // 10)
            rel_color = "green" if rel == "High" else "yellow" if rel == "Medium" else "dim"
            table.add_row(str(i), title, f"[cyan]{score_bar}[/cyan] {score}", f"[{rel_color}]{rel}[/{rel_color}]")
        console.print(table)
    else:
        for i, (title, score, rel) in enumerate(articles, 1):
            print(f"  {i}. [{score}] {title}")


def demo_sales_analytics() -> None:
    print_header("SALES CHANNEL ANALYTICS")

    channels = [
        ("eBay", "$12,450", 47, "75%"),
        ("Etsy", "$3,200", 18, "19%"),
        ("Poshmark", "$890", 8, "6%"),
    ]

    if RICH_AVAILABLE:
        table = Table(title="📊 Sales by Channel", box=box.ROUNDED)
        table.add_column("Channel", style="cyan")
        table.add_column("Revenue", justify="right", style="green")
        table.add_column("Units", justify="right")
        table.add_column("Share", justify="center")

        for channel, rev, units, share in channels:
            share_bar = "█" * (int(share[:-1]) // 10) + "░" * (10 - int(share[:-1]) // 10)
            table.add_row(channel, rev, str(units), f"[gold1]{share_bar}[/gold1] {share}")
        console.print(table)

        console.print("\n[bold]Total Revenue:[/bold] [green]$16,540[/green]  |  [bold]Best:[/bold] eBay (75%)")
    else:
        for channel, rev, units, share in channels:
            print(f"  {channel}: {rev} ({units} units)")


def demo_installation() -> None:
    print_header("HOW TO USE")

    if RICH_AVAILABLE:
        console.print(Panel("""
[bold]Installation:[/bold]
  1. Open Google Sheets
  2. Go to: Extensions → Apps Script
  3. Copy/paste any .gs file from this repo
  4. Click Save, then Run
  5. Grant permissions when prompted

[bold]API Keys (Script Properties):[/bold]
  • [cyan]CLAUDE_API_KEY[/cyan] - console.anthropic.com
  • [cyan]OPENAI_API_KEY[/cyan] - platform.openai.com

[dim]Scripts automatically add menu items to your Sheet.[/dim]
""", title="🚀 Quick Start", border_style="cyan", box=box.ROUNDED))
    else:
        print("1. Open Google Sheets → Extensions → Apps Script")
        print("2. Copy/paste .gs file")
        print("3. Add API keys to Script Properties")


def main() -> None:
    show_banner()

    if RICH_AVAILABLE:
        console.print("[dim]This repo contains 24+ Google Apps Scripts for inventory,")
        console.print("AI analysis, news scoring, and sales analytics.[/dim]\n")

    demo_3dsellers()
    demo_ai_integration()
    demo_news_engine()
    demo_sales_analytics()
    demo_installation()

    print_header("SCRIPT CATEGORIES")
    if RICH_AVAILABLE:
        table = Table(box=box.ROUNDED)
        table.add_column("Category", style="cyan")
        table.add_column("Scripts", justify="center", style="gold1")
        table.add_column("Description")

        categories = [
            ("3DSellers", "5", "Inventory sync, eBay/Etsy integration"),
            ("AI Integration", "8", "Claude, GPT-4, Gemini connectors"),
            ("News Engine", "4", "Multi-AI article scoring"),
            ("Utilities", "7+", "SKU gen, renaming, analytics"),
        ]
        for cat, count, desc in categories:
            table.add_row(cat, count, desc)
        console.print(table)

    print_header("DEMO COMPLETE")
    if RICH_AVAILABLE:
        console.print("[bold]To use:[/bold] Copy any .gs file → Google Sheets → Apps Script")


if __name__ == "__main__":
    main()
