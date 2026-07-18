"""Generate realistic Excel-window screenshots of the Brew Haven workbook sheets."""
from pathlib import Path

import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
from matplotlib.patches import FancyBboxPatch, Rectangle
import numpy as np

ROOT = Path(__file__).parent
OUT = ROOT / "screenshots"
OUT.mkdir(exist_ok=True)
VERSION = "0.0.2"

# Excel-like colors
XL_GREEN = "#217346"
XL_HEADER = "#1B4332"
XL_ROW_ALT = "#F2F2F2"
XL_GRID = "#D0D0D0"
XL_TITLEBAR = "#FFFFFF"
XL_RIBBON = "#F3F3F3"
XL_TAB_ACTIVE = "#FFFFFF"
XL_TAB_INACTIVE = "#E1E1E1"


def draw_excel_chrome(fig, sheet_name, filename_label):
    """Draw Excel window chrome: title bar, ribbon strip, sheet tabs."""
    # Outer window
    fig.patch.set_facecolor("#B4B4B4")
    ax = fig.add_axes([0.02, 0.06, 0.96, 0.90])
    ax.set_xlim(0, 100)
    ax.set_ylim(0, 100)
    ax.axis("off")

    # Title bar
    ax.add_patch(Rectangle((0, 94), 100, 6, facecolor="#FFFFFF", edgecolor="#C8C8C8", linewidth=0.5))
    ax.text(1.5, 96.8, f"Brew_Haven_Cafe_Cost_Tracker.xlsx  -  Excel", fontsize=9, va="center", color="#333")
    ax.text(88, 96.8, "—  □  ✕", fontsize=9, va="center", color="#666", family="monospace")

    # Ribbon
    ax.add_patch(Rectangle((0, 88), 100, 6, facecolor=XL_RIBBON, edgecolor="#C8C8C8", linewidth=0.5))
    for i, tab in enumerate(["File", "Home", "Insert", "Page Layout", "Formulas", "Data", "Review", "View"]):
        weight = "bold" if tab == "Home" else "normal"
        color = XL_GREEN if tab == "Home" else "#444"
        ax.text(3 + i * 8, 91, tab, fontsize=7.5, va="center", color=color, fontweight=weight)
    ax.plot([0, 100], [88, 88], color=XL_GREEN, linewidth=2)

    # Formula bar
    ax.add_patch(Rectangle((0, 84), 100, 4, facecolor="#FFFFFF", edgecolor="#C8C8C8", linewidth=0.5))
    ax.text(1.5, 86, "fx", fontsize=8, va="center", color="#888", fontstyle="italic")
    ax.add_patch(Rectangle((5, 84.5), 90, 3, facecolor="#FFFFFF", edgecolor="#B0B0B0", linewidth=0.6))
    ax.text(6, 86, filename_label, fontsize=7.5, va="center", color="#333")

    # Sheet content area background
    ax.add_patch(Rectangle((0, 8), 100, 76, facecolor="#FFFFFF", edgecolor="#C8C8C8", linewidth=0.5))

    # Sheet tabs at bottom
    tabs = ["Overview", "Startup Costs", "Ongoing Expenses", "Monthly Ledger", "Profitability", "Regression"]
    x = 1
    for t in tabs:
        w = max(len(t) * 1.15 + 2, 10)
        active = t == sheet_name
        face = XL_TAB_ACTIVE if active else XL_TAB_INACTIVE
        ax.add_patch(Rectangle((x, 2.5), w, 4.5, facecolor=face, edgecolor="#A0A0A0", linewidth=0.6))
        ax.text(x + w / 2, 4.7, t, fontsize=6.5, ha="center", va="center",
                fontweight="bold" if active else "normal",
                color=XL_GREEN if active else "#555")
        if active:
            ax.plot([x, x + w], [7, 7], color=XL_GREEN, linewidth=2.5)
        x += w + 0.4

    ax.text(98.5, 0.8, f"v{VERSION}", fontsize=6.5, ha="right", va="center", color="#666")
    return ax


def draw_table(ax, headers, rows, col_widths, origin=(2, 78), row_h=4.2, money_cols=None, type_col=None):
    """Draw an Excel-like grid table inside the content area."""
    money_cols = money_cols or set()
    n_cols = len(headers)
    x = origin[0]
    # Column letters
    letters = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
    cx = x
    for i, w in enumerate(col_widths):
        ax.add_patch(Rectangle((cx, origin[1]), w, 3.2, facecolor="#E6E6E6", edgecolor=XL_GRID, linewidth=0.5))
        ax.text(cx + w / 2, origin[1] + 1.6, letters[i], fontsize=6.5, ha="center", va="center", color="#666")
        cx += w

    # Header row
    y = origin[1] - row_h
    cx = x
    for i, (h, w) in enumerate(zip(headers, col_widths)):
        ax.add_patch(Rectangle((cx, y), w, row_h, facecolor=XL_HEADER, edgecolor="#0D2818", linewidth=0.4))
        ax.text(cx + 0.6, y + row_h / 2, h, fontsize=7, va="center", color="white", fontweight="bold")
        cx += w

    # Row number gutter
    # Data rows
    for r_idx, row in enumerate(rows):
        y = origin[1] - row_h * (r_idx + 2)
        face = "#FFFFFF" if r_idx % 2 == 0 else XL_ROW_ALT
        cx = x
        for c_idx, (val, w) in enumerate(zip(row, col_widths)):
            cell_face = face
            text_color = "#222"
            # Cost type coloring
            if type_col is not None and c_idx == type_col:
                if str(val) == "Fixed":
                    cell_face = "#E8F5E9"
                elif str(val) == "Variable":
                    cell_face = "#FFF3E0"
                elif str(val) == "Mixed":
                    cell_face = "#E3F2FD"
            if "TOTAL" in str(row[1] if len(row) > 1 else "") or (c_idx == 1 and "TOTAL" in str(val)):
                cell_face = "#C8E6C9"
                text_color = "#1B4332"
            ax.add_patch(Rectangle((cx, y), w, row_h, facecolor=cell_face, edgecolor=XL_GRID, linewidth=0.4))
            align = "right" if c_idx in money_cols else "left"
            tx = cx + w - 0.5 if align == "right" else cx + 0.5
            ax.text(tx, y + row_h / 2, str(val), fontsize=6.8, va="center", ha=align, color=text_color)
            cx += w

    # Row numbers on left
    for r_idx in range(len(rows) + 1):
        y = origin[1] - row_h * (r_idx + 1)
        ax.add_patch(Rectangle((0.3, y), 1.5, row_h, facecolor="#E6E6E6", edgecolor=XL_GRID, linewidth=0.4))
        ax.text(1.05, y + row_h / 2, str(r_idx + 1) if r_idx > 0 else "", fontsize=6, ha="center", va="center", color="#666")


def screenshot_startup():
    fig = plt.figure(figsize=(13, 8.2), dpi=140)
    ax = draw_excel_chrome(fig, "Startup Costs", "Startup Costs!C16")
    ax.text(2, 81.5, "Startup Costs — Brew Haven Cafe", fontsize=11, fontweight="bold", color=XL_HEADER, va="center")

    headers = ["#", "Item", "Amount ($)", "Cost Type", "Category"]
    widths = [5, 38, 14, 14, 16]
    rows = [
        ["1", "Espresso machine (2-group)", "$8,500.00", "Fixed", "Equipment"],
        ["2", "Commercial grinders (x2)", "$1,600.00", "Fixed", "Equipment"],
        ["3", "Furniture & seating", "$5,200.00", "Fixed", "Furnishings"],
        ["4", "Lease deposit (2 months)", "$7,000.00", "Fixed", "Facilities"],
        ["5", "Build-out / renovations", "$9,500.00", "Fixed", "Facilities"],
        ["6", "Business licenses & permits", "$850.00", "Fixed", "Legal"],
        ["7", "Initial inventory (beans, milk, cups)", "$2,800.00", "Variable", "Inventory"],
        ["8", "POS system & tablets", "$1,400.00", "Fixed", "Technology"],
        ["9", "Exterior signage & branding", "$1,200.00", "Fixed", "Marketing"],
        ["10", "Opening marketing campaign", "$1,500.00", "Fixed", "Marketing"],
        ["11", "Smallwares (pitchers, scales, etc.)", "$650.00", "Fixed", "Equipment"],
        ["12", "Working capital buffer", "$5,000.00", "Fixed", "Cash reserve"],
        ["", "TOTAL STARTUP CAPITAL", "$45,200.00", "", ""],
    ]
    draw_table(ax, headers, rows, widths, origin=(2, 76), money_cols={2}, type_col=3)
    path = OUT / "06_excel_startup_costs.png"
    fig.savefig(path, dpi=140, bbox_inches="tight", facecolor=fig.get_facecolor())
    plt.close(fig)
    return path


def screenshot_ongoing():
    fig = plt.figure(figsize=(13, 8.2), dpi=140)
    ax = draw_excel_chrome(fig, "Ongoing Expenses", "Ongoing Expenses!D8")
    ax.text(2, 81.5, "Ongoing Monthly Expenses — Fixed / Variable / Mixed", fontsize=11, fontweight="bold", color=XL_HEADER, va="center")

    headers = ["#", "Expense", "Amount ($)", "Cost Type", "Category"]
    widths = [5, 38, 14, 14, 16]
    rows = [
        ["1", "Rent", "$3,500.00", "Fixed", "Facilities"],
        ["2", "Base salaries (manager + 2 baristas)", "$8,200.00", "Fixed", "Labor"],
        ["3", "Hourly labor (volume-driven)", "$2,100.00", "Variable", "Labor"],
        ["4", "Coffee beans & tea", "$1,680.00", "Variable", "COGS"],
        ["5", "Milk, syrups & dairy alternatives", "$920.00", "Variable", "COGS"],
        ["6", "Cups, lids, sleeves, napkins", "$480.00", "Variable", "COGS"],
        ["7", "Utilities (electricity, water, gas)", "$620.00", "Mixed", "Facilities"],
        ["8", "Insurance", "$275.00", "Fixed", "Insurance"],
        ["9", "Software (POS, payroll, accounting)", "$165.00", "Fixed", "Technology"],
        ["10", "Marketing / social ads", "$400.00", "Fixed", "Marketing"],
        ["11", "Maintenance & supplies", "$180.00", "Fixed", "Operations"],
        ["12", "Loan / equipment payment", "$450.00", "Fixed", "Debt"],
        ["", "TOTAL MONTHLY", "$18,970.00", "", ""],
    ]
    draw_table(ax, headers, rows, widths, origin=(2, 76), money_cols={2}, type_col=3)
    path = OUT / "07_excel_ongoing_expenses.png"
    fig.savefig(path, dpi=140, bbox_inches="tight", facecolor=fig.get_facecolor())
    plt.close(fig)
    return path


def screenshot_profitability():
    fig = plt.figure(figsize=(13, 8.2), dpi=140)
    ax = draw_excel_chrome(fig, "Profitability", "Profitability!B12")
    ax.text(2, 81.5, "Profitability Projection — Brew Haven Cafe", fontsize=11, fontweight="bold", color=XL_HEADER, va="center")

    headers = ["Metric", "Value"]
    widths = [52, 22]
    rows = [
        ["Average monthly cups", "4,204"],
        ["Average monthly revenue", "$23,123.33"],
        ["Average monthly total cost", "$16,677.50"],
        ["Average monthly profit", "$6,445.83"],
        ["Average profit margin", "27.9%"],
        ["Estimated fixed cost (regression intercept)", "$8,728.31"],
        ["Estimated variable cost per cup (slope)", "$1.89"],
        ["Average selling price per cup", "$5.50"],
        ["Contribution margin per cup", "$3.61"],
        ["Break-even volume (cups / month)", "2,418"],
    ]
    draw_table(ax, headers, rows, widths, origin=(2, 76), money_cols={1}, type_col=None)

    # Scenario table lower
    ax.text(2, 28, "Scenario projections (regression cost model + $5.50 avg price)", fontsize=9, fontweight="bold", color=XL_HEADER)
    headers2 = ["Scenario", "Cups / Month", "Projected Cost", "Projected Revenue", "Projected Profit"]
    widths2 = [16, 16, 18, 18, 18]
    rows2 = [
        ["Conservative", "3,500", "$15,346", "$19,250", "$3,904"],
        ["Base", "4,200", "$16,670", "$23,100", "$6,430"],
        ["Growth", "5,200", "$18,560", "$28,600", "$10,040"],
        ["Peak summer", "5,600", "$19,317", "$30,800", "$11,483"],
    ]
    draw_table(ax, headers2, rows2, widths2, origin=(2, 25), row_h=3.8, money_cols={2, 3, 4})
    path = OUT / "08_excel_profitability.png"
    fig.savefig(path, dpi=140, bbox_inches="tight", facecolor=fig.get_facecolor())
    plt.close(fig)
    return path


def screenshot_regression():
    fig = plt.figure(figsize=(13, 8.2), dpi=140)
    ax = draw_excel_chrome(fig, "Regression", "Regression!B7")
    ax.text(2, 81.5, "Simple Linear Regression — Forecasting Total Cost from Volume", fontsize=10.5, fontweight="bold", color=XL_HEADER, va="center")
    ax.text(2, 78.5, "Model: Total Cost (Y) = Intercept (a) + Slope (b) × Cups Sold (X)", fontsize=8, color="#2D6A4F", va="center")

    headers = ["Statistic", "Value", "Interpretation"]
    widths = [32, 16, 40]
    rows = [
        ["Intercept (a) — Fixed cost estimate", "$8,728.31", "When cups = 0, expected monthly cost ≈ $8,728"],
        ["Slope (b) — Variable cost / cup", "$1.8908", "Each extra cup adds ≈ $1.89 to total cost"],
        ["R-squared", "0.9995", "99.95% of cost variation explained by volume"],
        ["Correlation (r)", "0.9997", "Strong positive linear relationship"],
        ["p-value (slope)", "7.71E-18", "Slope is statistically significant (p < 0.05)"],
        ["Standard error (slope)", "0.0134", "Precision of the variable-cost estimate"],
        ["Observations (n)", "12", "Months of operating data"],
    ]
    draw_table(ax, headers, rows, widths, origin=(2, 74), row_h=4.0, money_cols={1})

    ax.text(2, 38, "3-Month Cost Forecast", fontsize=9, fontweight="bold", color=XL_HEADER)
    headers2 = ["Period", "Projected Cups", "Forecast Total Cost", "Implied Revenue @ $5.50", "Implied Profit"]
    widths2 = [12, 16, 20, 22, 18]
    rows2 = [
        ["Month +1", "5,200", "$18,560.41", "$28,600.00", "$10,039.59"],
        ["Month +2", "5,400", "$18,938.57", "$29,700.00", "$10,761.43"],
        ["Month +3", "5,600", "$19,316.73", "$30,800.00", "$11,483.27"],
    ]
    draw_table(ax, headers2, rows2, widths2, origin=(2, 35), row_h=4.0, money_cols={2, 3, 4})
    path = OUT / "09_excel_regression_output.png"
    fig.savefig(path, dpi=140, bbox_inches="tight", facecolor=fig.get_facecolor())
    plt.close(fig)
    return path


def screenshot_ledger():
    fig = plt.figure(figsize=(13, 8.2), dpi=140)
    ax = draw_excel_chrome(fig, "Monthly Ledger", "Monthly Ledger!E10")
    ax.text(2, 81.5, "12-Month Operating Ledger (Sample Actuals)", fontsize=11, fontweight="bold", color=XL_HEADER, va="center")

    headers = ["Month", "Cups Sold (Q)", "Total Cost ($)", "Revenue ($)", "Profit ($)", "Cost/Cup", "Margin"]
    widths = [10, 14, 15, 14, 14, 12, 10]
    data = [
        ("Jan", 3200, 14820, 17600),
        ("Feb", 3450, 15240, 18975),
        ("Mar", 3800, 15890, 20900),
        ("Apr", 4100, 16450, 22550),
        ("May", 4450, 17120, 24475),
        ("Jun", 4800, 17800, 26400),
        ("Jul", 5100, 18350, 28050),
        ("Aug", 4950, 18120, 27225),
        ("Sep", 4300, 16880, 23650),
        ("Oct", 4000, 16290, 22000),
        ("Nov", 3700, 15710, 20350),
        ("Dec", 4600, 17460, 25300),
    ]
    rows = []
    for m, q, c, r in data:
        p = r - c
        rows.append([
            m,
            f"{q:,}",
            f"${c:,.2f}",
            f"${r:,.2f}",
            f"${p:,.2f}",
            f"${c/q:.2f}",
            f"{p/r*100:.1f}%",
        ])
    draw_table(ax, headers, rows, widths, origin=(2, 76), row_h=3.55, money_cols={2, 3, 4, 5})
    path = OUT / "10_excel_monthly_ledger.png"
    fig.savefig(path, dpi=140, bbox_inches="tight", facecolor=fig.get_facecolor())
    plt.close(fig)
    return path


def main():
    paths = [
        screenshot_startup(),
        screenshot_ongoing(),
        screenshot_ledger(),
        screenshot_profitability(),
        screenshot_regression(),
    ]
    for p in paths:
        print(f"Wrote {p}")


if __name__ == "__main__":
    main()
