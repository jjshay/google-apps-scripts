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

    samples = [
        {"sku": "1_DNYC_Marilyn", "artist": "Death NYC", "title": "Marilyn Monroe x Chanel", "price": 225.00, "status": "Ready"},
        {"sku": "2_DNYC_Snoopy", "artist": "Death NYC", "title": "Snoopy x Louis Vuitton", "price": 195.00, "status": "Ready"},
        {"sku": "3_SF_Hope", "artist": "Shepard Fairey", "title": "Hope (2008)", "price": 1200.00, "status": "Ready"},
        {"sku": "4_BK_Balloon", "artist": "Banksy", "title": "Balloon Girl", "price": 850.00, "status": "Ready"},
        {"sku": "5_KAWS_Comp", "artist": "KAWS", "title": "Companion (Grey)", "price": 450.00, "status": "Review"},
        {"sku": "6_MBW_Einstein", "artist": "Mr. Brainwash", "title": "Einstein", "price": 1800.00, "status": "Ready"},
    ]

    if RICH_AVAILABLE:
        console.print("[dim]Auto-populates 30+ fields from artwork images using Claude Vision AI[/dim]\n")

        table = Table(title="📦 Today's Processed Items (6 of 47)", box=box.ROUNDED)
        table.add_column("SKU", style="cyan")
        table.add_column("Artist", style="gold1")
        table.add_column("Title")
        table.add_column("Price", justify="right", style="green")
        table.add_column("Status", justify="center")

        for item in samples:
            status_color = "green" if item["status"] == "Ready" else "yellow"
            table.add_row(
                item["sku"],
                item["artist"],
                item["title"],
                f"${item['price']:,.2f}",
                f"[{status_color}]{item['status']}[/{status_color}]"
            )
        console.print(table)

        total_value = sum(s["price"] for s in samples)
        console.print(f"\n[bold]Batch Total:[/bold] [green]${total_value:,.2f}[/green]  |  [bold]Ready:[/bold] 5  |  [yellow]Review:[/yellow] 1")

        console.print(Panel("""
[bold]What it does:[/bold]
  1. Monitors Google Drive for new images
  2. Sends to Claude Vision API for analysis
  3. Auto-fills inventory sheet (artist, title, price, etc.)
  4. Syncs to 3DSellers for eBay/Etsy listing

[bold]Stats This Month:[/bold]
  • 47 items processed
  • 94% auto-classified correctly
  • 3 hours saved per day
""", title="🔄 Workflow", border_style="green", box=box.ROUNDED))
    else:
        for item in samples:
            print(f"  {item['sku']}: {item['artist']} - {item['title']} (${item['price']})")


def demo_ai_integration() -> None:
    print_header("AI INTEGRATION SCRIPTS")

    results = [
        ("DNYC-Marilyn-001.jpg", "Death NYC", "Screen Print", "97%", "green"),
        ("SF-Hope-002.jpg", "Shepard Fairey", "Screen Print", "94%", "green"),
        ("BK-Thrower-003.jpg", "Banksy", "Screen Print", "89%", "green"),
        ("Unknown-004.jpg", "NEEDS REVIEW", "Unknown", "45%", "red"),
        ("MBW-Einstein-005.jpg", "Mr. Brainwash", "Mixed Media", "91%", "green"),
        ("KAWS-Comp-006.jpg", "KAWS", "Vinyl Sculpture", "96%", "green"),
        ("JR-Face-007.jpg", "JR", "Lithograph", "88%", "green"),
        ("Invader-008.jpg", "Invader", "Screen Print", "93%", "green"),
    ]

    if RICH_AVAILABLE:
        console.print("[dim]Connects Google Sheets to Claude, GPT-4, Gemini for image classification[/dim]\n")

        table = Table(title="🤖 AI Classification Results (8 Images)", box=box.ROUNDED)
        table.add_column("Image", style="cyan")
        table.add_column("Artist", style="gold1")
        table.add_column("Medium", style="dim")
        table.add_column("Confidence", justify="center")

        for img, artist, medium, conf, color in results:
            conf_val = int(conf.replace('%', ''))
            bar = "█" * (conf_val // 10) + "░" * (10 - conf_val // 10)
            table.add_row(img, artist, medium, f"[{color}]{bar}[/{color}] {conf}")
        console.print(table)

        console.print("\n[green]Auto-sorted:[/green] 7  |  [yellow]Needs review:[/yellow] 1  |  [bold]Accuracy:[/bold] 87.5%")
    else:
        for img, artist, medium, conf, _ in results:
            print(f"  {img}: {artist} - {medium} ({conf})")


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
        ("eBay", 12450, 47, 75),
        ("Etsy", 3200, 18, 19),
        ("Poshmark", 890, 8, 6),
    ]

    if RICH_AVAILABLE:
        table = Table(title="📊 Sales by Channel (This Month)", box=box.ROUNDED)
        table.add_column("Channel", style="cyan")
        table.add_column("Revenue", justify="right", style="green")
        table.add_column("Units", justify="right")
        table.add_column("Avg Price", justify="right", style="dim")
        table.add_column("Share", justify="center")

        for channel, rev, units, share in channels:
            share_bar = "█" * (share // 10) + "░" * (10 - share // 10)
            avg = rev / units if units > 0 else 0
            table.add_row(channel, f"${rev:,}", str(units), f"${avg:.0f}", f"[gold1]{share_bar}[/gold1] {share}%")
        console.print(table)

        total_rev = sum(c[1] for c in channels)
        total_units = sum(c[2] for c in channels)

        # Top sellers
        top_sellers = [
            ("Shepard Fairey - Hope", "eBay", "$1,200"),
            ("Banksy - Thrower", "eBay", "$850"),
            ("KAWS - Companion", "Etsy", "$450"),
        ]

        top_table = Table(title="🏆 Top Sellers This Month", box=box.ROUNDED)
        top_table.add_column("Item", style="gold1")
        top_table.add_column("Channel", style="cyan")
        top_table.add_column("Price", justify="right", style="green")

        for item, channel, price in top_sellers:
            top_table.add_row(item, channel, price)
        console.print(top_table)

        stats = f"""
[bold]Monthly Summary[/bold]

[cyan]Total Revenue:[/cyan]   [green]${total_rev:,}[/green]
[cyan]Units Sold:[/cyan]      {total_units}
[cyan]Avg Order:[/cyan]       ${total_rev/total_units:.2f}

[bold]vs Last Month:[/bold]
  Revenue:  [green]↑ 23%[/green]
  Units:    [green]↑ 18%[/green]
  AOV:      [green]↑ 4%[/green]
"""
        console.print(Panel(stats, title="📈 Performance", border_style="green", box=box.ROUNDED))
    else:
        for channel, rev, units, share in channels:
            print(f"  {channel}: ${rev:,} ({units} units)")


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
