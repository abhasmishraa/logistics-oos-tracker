# OTIF & Out-of-Stock Risk Engine for Wholesale Distributors

**Automated OOS early-warning and daily action reporting for a Logistics & Warehousing network.**

A lightweight Python pipeline that converts a raw WD (Wholesale Distributor) inventory snapshot into a prioritised, colour-coded action list for the Logistics Manager, replacing manual spreadsheet checks with a repeatable, one-command daily report.

---

## Business Problem

Wholesale Distributors sit between the plant/warehouse and the retail market. When a WD runs out of a fast-moving SKU, the cost is immediate: lost secondary sales, lower OTIF performance, firefighting freight and strained distributor relationships. Most stock-outs are visible days in advance, but only if someone compares **stock cover** against **replenishment lead time** for every SKU-WD pair, every day.

This tool automates that comparison.

## Core Logic

| Metric | Definition |
|---|---|
| **Days of Cover** | `Current_Stock_Qty / Avg_Daily_Demand` |
| **Critical Risk** | `Days of Cover <= Transit Lead Time`: stock will run out before the next delivery can land |
| **Warning** | `Days of Cover` is within 2 days above `Transit Lead Time`: no buffer for any delay |
| **Safe** | All other SKUs |

The 2-day warning buffer is a single configurable constant (`WARNING_BUFFER_DAYS`) so it can be aligned with each category's service-level policy.

## Business Impact

- **Optimises inventory norms.** Days of Cover benchmarked against lead time shows where stock is too thin (stock-out exposure) and, by extension, where cover is far above need (working-capital tied up). This gives a data-backed basis to revisit safety-stock and norm settings by SKU, category and WD.
- **Reduces WD OOS position.** SKUs are flagged *before* they stock out, giving the team a window to expedite dispatch, rebalance stock between locations or escalate to transport. Reporting is ranked worst-first so effort goes where the risk is highest.
- **Automates pipeline monitoring.** One command produces the Excel action list and a Word summary by WD, replacing manual pivots and daily status compilation.
- **Links OTIF to OOS.** The dataset carries each WD's last-delivery OTIF status, so recurring late deliveries can be viewed together with the SKUs they are putting at risk.
- **Decision-ready output.** The Logistics Manager receives a filtered action list (Critical rows in red, Warning rows in amber) instead of the full 50+ line inventory.

## Outputs

| File | Audience | Content |
|---|---|---|
| `Daily_OOS_Action_Report.xlsx` | Logistics Manager / planners | Only Critical + Warning SKUs, sorted worst-first. Critical rows have a red fill, Warning rows amber. Filters, frozen header and legend included. |
| `OOS_Risk_Summary_by_WD.docx` | Management / review meetings | One-page summary: executive note, count of at-risk SKUs by WD Location (Critical / Warning / Total) and recommended actions. |

### Sample run (dummy data, 50 SKU-WD combinations)

```
Safe             24
Critical Risk    14
Warning          12
```

26 of 50 SKUs flagged for action; the WD summary ranks locations by Critical exposure so escalation can be targeted.

> The sample data is synthetic and exists only to demonstrate the logic. Figures above are not real business results.

## Input Data Schema

| Column | Description |
|---|---|
| `WD_Location` | Wholesale Distributor location |
| `SKU_ID` | Stock keeping unit identifier |
| `Product_Category` | Category of the SKU |
| `Current_Stock_Qty` | Units currently on hand at the WD |
| `Avg_Daily_Demand` | Average units sold per day |
| `Transit_Lead_Time_Days` | Days for replenishment to reach the WD |
| `Last_Delivery_OTIF_Status` | `On-Time` / `Delayed` |

## Quick Start

```bash
pip install -r requirements.txt
python oos_risk_engine.py                      # uses wd_inventory_data.csv
python oos_risk_engine.py my_latest_stock.csv  # or point at your own extract
```

The script prints the first five processed rows and the risk distribution, then writes both reports to the working directory.

## Project Structure

```
├── oos_risk_engine.py          # Main pipeline: load > calculate > filter > report
├── generate_dummy_data.py      # Seeded generator for the sample dataset
├── wd_inventory_data.csv       # Sample input (50 rows, synthetic)
├── Daily_OOS_Action_Report.xlsx
├── OOS_Risk_Summary_by_WD.docx
├── requirements.txt
└── README.md
```

## Design Notes

- **Validation:** the script checks for required columns before processing.
- **Edge cases:** zero demand is handled without a divide-by-zero error and is flagged for manual review rather than silently marked safe.
- **Separation of concerns:** load, risk calculation, filtering and each report export are separate functions, making it easy to plug in a new data source or output.
- **Tech stack:** Python, Pandas, NumPy, openpyxl, python-docx.

## Roadmap

- Replace the CSV with a direct ERP/WMS or SAP extract
- Add recommended replenishment quantity (to reach target cover) per flagged SKU
- Weight risk by SKU velocity or revenue contribution
- Trend `Days_of_Cover` over time and track repeat offenders
- Schedule the run (cron / Task Scheduler) with e-mail delivery of the reports

---

*Built as a practical demonstration of supply chain analytics for Logistics & Warehousing: inventory norms, OTIF and OOS control, and process automation.*
