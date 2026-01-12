#!/usr/bin/env python3
"""Google Apps Scripts - Marketing Demo"""
import time
import sys

try:
    from rich.console import Console
    from rich.panel import Panel
    from rich.table import Table
    from rich.align import Align
    from rich import box
except ImportError:
    print("Run: pip install rich")
    sys.exit(1)

console = Console()

def pause(s=1.5):
    time.sleep(s)

def step(text):
    console.print(f"\n[bold white on #1a1a2e]  {text}  [/]\n")
    pause(0.8)

# INTRO
console.clear()
console.print()
intro = Panel(
    Align.center("[bold yellow]GOOGLE APPS SCRIPTS[/]\n\n[white]Sheets Automation & Business Intelligence[/]"),
    border_style="cyan",
    width=60,
    padding=(1, 2)
)
console.print(intro)
pause(2)

# DEMO 1: INVENTORY SYNC
step("DEMO 1: AUTOMATED INVENTORY SYNC")

console.print("[bold]Running inventorySync()...[/]\n")
pause(0.8)

console.print("  Connecting to Sheets.....", end="")
pause(0.5)
console.print(" [green]Done[/]")

console.print("  Fetching eBay listings...", end="")
pause(0.6)
console.print(" [green]247 listings found[/]")

console.print("  Comparing inventory......", end="")
pause(0.5)
console.print(" [green]Done[/]")

pause(0.5)

sync = Panel(
    "[bold]Sync Results:[/]\n\n"
    "  [green]>[/] Items in sync:       [bold]234[/]\n"
    "  [yellow]![/] Quantity mismatches: [bold yellow]8[/]\n"
    "  [red]x[/] Missing from eBay:   [bold red]5[/]\n\n"
    "[dim]Auto-updated 8 quantities, flagged 5 for relisting[/]",
    title="[cyan]Sync Complete[/]",
    border_style="green",
    width=50
)
console.print(sync)
pause(2)

# DEMO 2: SALES ANALYTICS
step("DEMO 2: AUTOMATED SALES ANALYTICS")

console.print("[bold]Running generateSalesReport()...[/]\n")
pause(0.8)

console.print("  Pulling sales data.......", end="")
pause(0.5)
console.print(" [green]Done[/]")

console.print("  Calculating metrics......", end="")
pause(0.4)
console.print(" [green]Done[/]")

pause(0.5)

sales = Table(box=box.ROUNDED, width=55)
sales.add_column("Metric", style="white")
sales.add_column("Value", justify="right", style="cyan")
sales.add_column("vs Last Month", justify="right")

sales.add_row("Total Revenue", "$12,847.50", "[green]+18.3%[/]")
sales.add_row("Items Sold", "47", "[green]+12.0%[/]")
sales.add_row("Avg Sale Price", "$273.35", "[green]+5.6%[/]")
sales.add_row("Net Profit", "$8,892.40", "[green]+21.7%[/]")

console.print(sales)
pause(2)

# DEMO 3: EMAIL AUTOMATION
step("DEMO 3: CUSTOMER EMAIL AUTOMATION")

console.print("[bold]Running sendFollowUpEmails()...[/]\n")
pause(0.8)

console.print("  Querying recent buyers...", end="")
pause(0.5)
console.print(" [green]12 in queue[/]")

emails = [
    "john.d***@gmail.com",
    "sarah.m***@yahoo.com",
    "buyer_***@ebay.com",
]

console.print()
for email in emails:
    console.print(f"  [green]>[/] Sent to {email}")
    pause(0.15)

console.print(f"  [dim]... and 9 more emails sent[/]")
console.print(f"\n  [bold]Total:[/] 12 emails sent, 0 failed")
pause(1.5)

# DEMO 4: PRICE MONITORING
step("DEMO 4: COMPETITOR PRICE MONITORING")

console.print("[bold]Running priceMonitor()...[/]\n")
pause(0.8)

console.print("  Scraping competitors.....", end="")
pause(0.6)
console.print(" [green]156 listings analyzed[/]")

pause(0.5)

prices = Table(box=box.SIMPLE, width=55)
prices.add_column("Your Item", style="white")
prices.add_column("Your Price", justify="right")
prices.add_column("Market Avg", justify="right")
prices.add_column("Action")

prices.add_row("Omega Seamaster", "$4,899", "$4,650", "[yellow]Review[/]")
prices.add_row("Rolex Datejust", "$7,500", "$7,800", "[green]OK[/]")
prices.add_row("Warhol Lithograph", "$3,200", "$3,500", "[green]Raise[/]")

console.print(prices)
pause(1.5)

# DEMO 5: TRIGGERS
step("DEMO 5: AUTOMATED TRIGGERS")

triggers = Table(box=box.ROUNDED, width=55)
triggers.add_column("Function", style="white")
triggers.add_column("Schedule", style="cyan")
triggers.add_column("Last Run", style="dim")

triggers.add_row("inventorySync()", "Every 6 hours", "2 hours ago")
triggers.add_row("generateSalesReport()", "Daily 9:00 AM", "Yesterday")
triggers.add_row("sendFollowUpEmails()", "Daily 10:00 AM", "Yesterday")
triggers.add_row("priceMonitor()", "Weekly Monday", "5 days ago")
triggers.add_row("backupData()", "Daily 2:00 AM", "6 hours ago")

console.print(triggers)
console.print("\n  [green]All triggers active and running[/]")
pause(1.5)

# SUMMARY
step("AUTOMATION SUMMARY")

summary = Panel(
    Align.center(
        "[bold green]GOOGLE APPS SCRIPTS TOOLKIT[/]\n\n"
        "[bold]Scripts:[/] 5 automation modules\n"
        "[bold]Time Saved:[/] ~15 hours/week\n"
        "[bold]Error Reduction:[/] 94%"
    ),
    title="[bold yellow]OVERVIEW[/]",
    border_style="green",
    width=45
)
console.print(summary)
pause(2)

# FOOTER
console.print()
footer = Panel(
    Align.center(
        "[dim]JavaScript + Google APIs[/]\n"
        "[bold cyan]github.com/jjshay/google-apps-scripts[/]"
    ),
    title="[dim]Google Apps Scripts v2.0[/]",
    border_style="dim",
    width=50
)
console.print(footer)
pause(3)
