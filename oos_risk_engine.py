"""
OTIF & Out-of-Stock (OOS) Risk Engine for Wholesale Distributors (WDs)
======================================================================
Pipeline:
    1. Load WD inventory snapshot (CSV)
    2. Compute Days_of_Cover and OOS_Risk_Level
    3. Isolate 'Critical Risk' and 'Warning' SKUs
    4. Export a colour-coded Excel action report   (openpyxl)
    5. Export a one-page Word summary by WD        (python-docx)

Usage:
    python oos_risk_engine.py [path_to_csv]
"""

import sys
from datetime import date
from pathlib import Path

import numpy as np
import pandas as pd
from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt, RGBColor
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

# --------------------------------------------------------------------------- #
# Configuration
# --------------------------------------------------------------------------- #
INPUT_CSV = Path(sys.argv[1]) if len(sys.argv) > 1 else Path("wd_inventory_data.csv")
EXCEL_REPORT = Path("Daily_OOS_Action_Report.xlsx")
WORD_REPORT = Path("OOS_Risk_Summary_by_WD.docx")

WARNING_BUFFER_DAYS = 2  # 'Warning' = cover is within this many days above lead time
REQUIRED_COLS = [
    "WD_Location", "SKU_ID", "Product_Category", "Current_Stock_Qty",
    "Avg_Daily_Demand", "Transit_Lead_Time_Days", "Last_Delivery_OTIF_Status",
]

RISK_ORDER = {"Critical Risk": 0, "Warning": 1, "Safe": 2}


# --------------------------------------------------------------------------- #
# Step 1 & 2: Load data and calculate risk
# --------------------------------------------------------------------------- #
def load_data(path: Path) -> pd.DataFrame:
    df = pd.read_csv(path)
    missing = [c for c in REQUIRED_COLS if c not in df.columns]
    if missing:
        raise ValueError(f"Input file is missing required columns: {missing}")
    return df


def classify_risk(days_of_cover: float, lead_time: float) -> str:
    """Return the OOS risk level for a single SKU-WD combination."""
    if pd.isna(days_of_cover):
        return "Critical Risk"  # no demand data / zero demand -> flag for manual review
    if days_of_cover <= lead_time:
        return "Critical Risk"
    if days_of_cover <= lead_time + WARNING_BUFFER_DAYS:
        return "Warning"
    return "Safe"


def add_risk_columns(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    # Guard against divide-by-zero: zero demand -> NaN cover (handled in classify_risk)
    demand = df["Avg_Daily_Demand"].replace(0, np.nan)
    df["Days_of_Cover"] = (df["Current_Stock_Qty"] / demand).round(2)
    df["OOS_Risk_Level"] = [
        classify_risk(doc, lt)
        for doc, lt in zip(df["Days_of_Cover"], df["Transit_Lead_Time_Days"])
    ]
    return df


# --------------------------------------------------------------------------- #
# Step 3: Filter to actionable SKUs
# --------------------------------------------------------------------------- #
def filter_at_risk(df: pd.DataFrame) -> pd.DataFrame:
    at_risk = df[df["OOS_Risk_Level"].isin(["Critical Risk", "Warning"])].copy()
    at_risk["_rank"] = at_risk["OOS_Risk_Level"].map(RISK_ORDER)
    at_risk = (
        at_risk.sort_values(["_rank", "Days_of_Cover"])  # worst first
        .drop(columns="_rank")
        .reset_index(drop=True)
    )
    return at_risk


# --------------------------------------------------------------------------- #
# Step 4: Excel report (openpyxl)
# --------------------------------------------------------------------------- #
def export_excel(at_risk: pd.DataFrame, path: Path) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "OOS Action List"

    red_fill = PatternFill(fill_type="solid", start_color="FF4C4C", end_color="FF4C4C")
    amber_fill = PatternFill(fill_type="solid", start_color="FFE699", end_color="FFE699")
    header_fill = PatternFill(fill_type="solid", start_color="1F3864", end_color="1F3864")
    thin = Side(style="thin", color="BFBFBF")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    # Title block
    ws["A1"] = "Daily OOS Action Report - Wholesale Distributor Network"
    ws["A1"].font = Font(name="Arial", size=14, bold=True)
    ws["A2"] = (
        f"Generated: {date.today():%d %b %Y}  |  SKUs at risk: {len(at_risk)}  |  "
        f"Critical: {(at_risk['OOS_Risk_Level'] == 'Critical Risk').sum()}  |  "
        f"Warning: {(at_risk['OOS_Risk_Level'] == 'Warning').sum()}"
    )
    ws["A2"].font = Font(name="Arial", size=10, italic=True, color="595959")

    header_row = 4
    cols = list(at_risk.columns)
    for c_idx, col in enumerate(cols, start=1):
        cell = ws.cell(row=header_row, column=c_idx, value=col)
        cell.font = Font(name="Arial", bold=True, color="FFFFFF", size=10)
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = border

    for r_offset, (_, row) in enumerate(at_risk.iterrows(), start=1):
        r_idx = header_row + r_offset
        is_critical = row["OOS_Risk_Level"] == "Critical Risk"
        for c_idx, col in enumerate(cols, start=1):
            value = row[col]
            if isinstance(value, np.generic):
                value = value.item()
            cell = ws.cell(row=r_idx, column=c_idx, value=value)
            cell.border = border
            cell.alignment = Alignment(horizontal="left" if isinstance(value, str) else "center")
            if is_critical:
                cell.fill = red_fill
                cell.font = Font(name="Arial", size=10, bold=True, color="FFFFFF")
            else:
                cell.fill = amber_fill
                cell.font = Font(name="Arial", size=10)

    # Cosmetics: widths, freeze panes, filter
    for c_idx, col in enumerate(cols, start=1):
        longest = max([len(str(col))] + [len(str(v)) for v in at_risk[col]])
        ws.column_dimensions[get_column_letter(c_idx)].width = min(longest + 4, 32)
    ws.freeze_panes = ws.cell(row=header_row + 1, column=1)
    ws.auto_filter.ref = f"A{header_row}:{get_column_letter(len(cols))}{header_row + len(at_risk)}"

    # Legend
    legend_row = header_row + len(at_risk) + 2
    ws.cell(row=legend_row, column=1, value="Legend").font = Font(name="Arial", bold=True, size=10)
    ws.cell(row=legend_row + 1, column=1, value="Critical Risk").fill = red_fill
    ws.cell(row=legend_row + 1, column=1).font = Font(name="Arial", size=10, bold=True, color="FFFFFF")
    ws.cell(row=legend_row + 1, column=2, value="Days of Cover <= Transit Lead Time: stock-out likely before replenishment lands")
    ws.cell(row=legend_row + 2, column=1, value="Warning").fill = amber_fill
    ws.cell(row=legend_row + 2, column=1).font = Font(name="Arial", size=10)
    ws.cell(row=legend_row + 2, column=2, value=f"Days of Cover within {WARNING_BUFFER_DAYS} days above Transit Lead Time: expedite / monitor")
    for r in (legend_row + 1, legend_row + 2):
        ws.cell(row=r, column=2).font = Font(name="Arial", size=9, color="595959")

    wb.save(path)


# --------------------------------------------------------------------------- #
# Step 5: Word summary (python-docx)
# --------------------------------------------------------------------------- #
def _shade(cell, hex_color: str) -> None:
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), hex_color)
    tc_pr.append(shd)


def _set_cell(cell, text, bold=False, color=None, align=WD_ALIGN_PARAGRAPH.CENTER) -> None:
    cell.text = ""
    p = cell.paragraphs[0]
    p.alignment = align
    run = p.add_run(str(text))
    run.font.name = "Arial"
    run.font.size = Pt(10)
    run.bold = bold
    if color:
        run.font.color.rgb = RGBColor.from_string(color)


def export_word(df_all: pd.DataFrame, at_risk: pd.DataFrame, path: Path) -> None:
    summary = (
        at_risk.pivot_table(
            index="WD_Location", columns="OOS_Risk_Level",
            values="SKU_ID", aggfunc="count", fill_value=0,
        )
        .reindex(columns=["Critical Risk", "Warning"], fill_value=0)
    )
    summary["Total At Risk"] = summary.sum(axis=1)
    summary = summary.sort_values(["Critical Risk", "Total At Risk"], ascending=False)

    total_skus = len(df_all)
    total_risk = len(at_risk)
    n_crit = int(summary["Critical Risk"].sum())
    n_warn = int(summary["Warning"].sum())
    worst_wd = summary.index[0]

    doc = Document()
    normal = doc.styles["Normal"]
    normal.font.name = "Arial"
    normal.font.size = Pt(10.5)

    title = doc.add_heading("Daily OOS Risk Summary by WD", level=0)
    title.alignment = WD_ALIGN_PARAGRAPH.LEFT
    sub = doc.add_paragraph(f"Logistics & Warehousing  |  Report date: {date.today():%d %B %Y}")
    sub.runs[0].font.color.rgb = RGBColor(0x59, 0x59, 0x59)
    sub.runs[0].italic = True

    doc.add_heading("Executive Summary", level=1)
    doc.add_paragraph(
        f"Of {total_skus} SKU-WD combinations monitored, {total_risk} "
        f"({total_risk / total_skus:.0%}) are at risk of stocking out before the next replenishment arrives: "
        f"{n_crit} Critical and {n_warn} Warning. {worst_wd} carries the highest exposure with "
        f"{int(summary.loc[worst_wd, 'Critical Risk'])} Critical SKUs and should be prioritised for expedited dispatch."
    )

    doc.add_heading("SKUs at Risk by WD Location", level=1)
    table = doc.add_table(rows=1, cols=4)
    table.style = "Table Grid"
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    for i, h in enumerate(["WD Location", "Critical Risk", "Warning", "Total At Risk"]):
        _set_cell(table.rows[0].cells[i], h, bold=True, color="FFFFFF")
        _shade(table.rows[0].cells[i], "1F3864")

    for wd, r in summary.iterrows():
        cells = table.add_row().cells
        _set_cell(cells[0], wd, align=WD_ALIGN_PARAGRAPH.LEFT)
        _set_cell(cells[1], int(r["Critical Risk"]), bold=r["Critical Risk"] > 0,
                  color="C00000" if r["Critical Risk"] > 0 else None)
        _set_cell(cells[2], int(r["Warning"]))
        _set_cell(cells[3], int(r["Total At Risk"]), bold=True)

    cells = table.add_row().cells
    _set_cell(cells[0], "Network Total", bold=True, align=WD_ALIGN_PARAGRAPH.LEFT)
    _set_cell(cells[1], n_crit, bold=True)
    _set_cell(cells[2], n_warn, bold=True)
    _set_cell(cells[3], total_risk, bold=True)
    for c in cells:
        _shade(c, "D9E2F3")

    doc.add_heading("Recommended Actions", level=1)
    for text in [
        "Critical Risk: escalate to transport / WD owner today; consider expedited or split dispatch.",
        "Warning: confirm dispatch is on schedule and review safety stock norms for repeat offenders.",
        "Refer to Daily_OOS_Action_Report.xlsx for the SKU-level action list.",
    ]:
        doc.add_paragraph(text, style="List Bullet")

    doc.save(path)


# --------------------------------------------------------------------------- #
# Main
# --------------------------------------------------------------------------- #
def main() -> None:
    df = load_data(INPUT_CSV)
    df = add_risk_columns(df)

    print("=" * 70)
    print("PROCESSED DATAFRAME - FIRST 5 ROWS")
    print("=" * 70)
    with pd.option_context("display.max_columns", None, "display.width", 200):
        print(df.head())

    print("\nRisk distribution:")
    print(df["OOS_Risk_Level"].value_counts().to_string())

    at_risk = filter_at_risk(df)
    export_excel(at_risk, EXCEL_REPORT)
    export_word(df, at_risk, WORD_REPORT)

    print(f"\n[OK] {len(at_risk)} at-risk SKUs exported -> {EXCEL_REPORT}")
    print(f"[OK] WD summary document generated      -> {WORD_REPORT}")


if __name__ == "__main__":
    main()
