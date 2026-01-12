#!/usr/bin/env python3
"""Marketing Demo - Google Apps Scripts"""
import time
import sys

try:
    from rich.console import Console
    from rich.panel import Panel
    from rich.table import Table
    from rich import box
    console = Console()
except ImportError:
    print("Run: pip install rich")
    sys.exit(1)

def pause(seconds=2):
    time.sleep(seconds)

def clear():
    console.clear()

# SCENE 1: Hook
clear()
console.print("\n" * 5)
console.print("[bold yellow]         MANAGING INVENTORY IN SPREADSHEETS?[/bold yellow]", justify="center")
pause(2)

# SCENE 2: Problem
clear()
console.print("\n" * 3)
console.print(Panel("""
[bold red]MANUAL DATA ENTRY IS KILLING YOU:[/bold red]

   • Hours copying between sheets
   • No AI to help classify items
   • Manual eBay/Etsy sync
   • Mistakes everywhere

[dim]Let Google Sheets do the work.[/dim]
""", title="❌ Spreadsheet Hell", border_style="red", width=60), justify="center")
pause(3)

# SCENE 3: Solution
clear()
console.print("\n" * 3)
console.print(Panel("""
[bold green]24+ AUTOMATION SCRIPTS:[/bold green]

   ✓ AI classifies your images
   ✓ Auto-fills 30+ fields
   ✓ Syncs to 3DSellers
   ✓ Tracks sales across channels

[bold]Your sheets, supercharged.[/bold]
""", title="✅ Google Apps Scripts Collection", border_style="green", width=60), justify="center")
pause(3)

# SCENE 4: AI Demo
clear()
console.print("\n\n")
console.print("[bold cyan]              🤖 AI CLASSIFICATION[/bold cyan]", justify="center")
console.print()
pause(1)

console.print("[dim]              Dropping images into Drive...[/dim]", justify="center")
pause(1)

items = [
    ("DNYC-Marilyn.jpg", "Death NYC", "97%"),
    ("SF-Hope.jpg", "Shepard Fairey", "94%"),
    ("BK-Thrower.jpg", "Banksy", "89%"),
    ("KAWS-Comp.jpg", "KAWS", "96%"),
]

table = Table(box=box.ROUNDED, width=55)
table.add_column("Image", style="cyan")
table.add_column("AI Detected", style="gold1")
table.add_column("Confidence", justify="center")

for img, artist, conf in items:
    table.add_row(img, artist, f"[green]{conf}[/green]")
    console.clear()
    console.print("\n\n")
    console.print("[bold cyan]              🤖 AI CLASSIFICATION[/bold cyan]", justify="center")
    console.print()
    console.print(table, justify="center")
    pause(0.6)

pause(2)

# SCENE 5: Stats
clear()
console.print("\n\n")
console.print("[bold magenta]              📊 THIS MONTH'S STATS[/bold magenta]", justify="center")
console.print()

stats = Table(box=box.ROUNDED, width=50)
stats.add_column("Metric", style="cyan")
stats.add_column("Value", style="green", justify="right")
stats.add_row("Items Processed", "47")
stats.add_row("Hours Saved", "15+")
stats.add_row("AI Accuracy", "94%")
stats.add_row("Revenue Tracked", "$16,540")

console.print(stats, justify="center")
pause(3)

# SCENE 6: Script List
clear()
console.print("\n\n")
console.print("[bold yellow]              📁 24+ READY-TO-USE SCRIPTS[/bold yellow]", justify="center")
console.print()

scripts = [
    "3DSellers Sync",
    "Claude Vision AI",
    "GPT-4 Integration",
    "News Scoring Engine",
    "Sales Analytics",
    "SKU Generator",
    "...and 18 more!",
]

for script in scripts:
    console.print(f"[dim]              ✓ {script}[/dim]", justify="center")
    pause(0.3)

pause(2)

# SCENE 7: CTA
clear()
console.print("\n" * 4)
console.print("[bold yellow]         ⭐ SUPERCHARGE YOUR SHEETS ⭐[/bold yellow]", justify="center")
console.print()
console.print("[bold white]          github.com/jjshay/google-apps-scripts[/bold white]", justify="center")
console.print()
console.print("[dim]            Copy any .gs file → Apps Script → Run[/dim]", justify="center")
pause(3)
