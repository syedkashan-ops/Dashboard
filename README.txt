DIVISION 10 FUEL + DIVISION 20 LUBRICANT DASHBOARD

1. Put/keep the three Excel input files in the Input folder.
2. Run: python dashboard.py (see IT_SETUP.txt for installation/setup).
3. Open Output\Result.xlsx after generation completes.

NEW NAVIGATION
- DIVISION 10 — FUEL opens the original fuel dashboard.
- DIVISION 20 — LUBRICANT opens the lubricant Mat Group Text dashboard.
- Lubricant area buttons open Lubricant_Area_A00 through Lubricant_Area_A08.
- Material-group buttons remain area-aware and open summary + detail for only the selected area.

DIVISION 20 LUBRICANT TABLES
1. Fiscal Year-to-Date Mat Group Summary
2. Current Month-to-Date Mat Group Summary
3. Selected Month Analysis
4. Selected Quarter Analysis
5. Previous 6 Complete Months Average Sales

Mat Group Text values appear in the Product column. Material-group targets are shown as N/A because the supplied target workbook does not contain Mat Group Text-level objectives.

ANALYTICS UPDATE
- All Analytics charts use Billed Qty in SKU only (not Net Value).
- Added Area vs Billed Qty (SKU).
- Added Date-wise Sales — Billed Qty (SKU).
- Layout was reorganized to reduce clutter and top-outlet tables show Top 20 outlets.

MONTHLY AREA MATRIX
- Dashboard button: MONTHLY AREA MATRIX.
- Product dropdown: PMG, HSD, R95, MOGAS.
- MOGAS is calculated as PMG + R95.
- Month dropdown follows the fiscal-year sequence July to June.
- Rows show A00 to A08; columns show days 1 to 31.
- Values and the linked graph use Billed Qty in SKU only.

UPDATE 30-Jul-2026
- Sold to Party is written as a number when numeric, removing leading zeros.
- Monthly_Area_Matrix now includes LY adjusted benchmark, TY vs LY variance, month objective, and required per-day sales.
- Main dashboard now includes current-month product averages with red/green trend and LY full-month sales vs current-month objective.

AUTOMATIC DAILY SAP MERGE
1. Keep the two master sales workbooks and target workbook in Input.
2. Put each newly downloaded SAP sales file in Daily_Input.
3. Run: python dashboard.py (see IT_SETUP.txt for installation/setup).
4. The program appends only new rows to the latest/current-year master file.
5. Duplicate records are skipped using SAP billing-document and line details.
6. A pre-update copy is stored in Backup.
7. Imported files are moved to Daily_Input\Processed after a successful merge.
8. Output\Result.xlsx is then regenerated automatically.

The daily SAP file can contain only the new day or a cumulative period. Existing rows will not be added again.

COST CENTER / OUTLET MAPPING (V16)

1. Maintain ONLY Input/Outlet_Master.xlsx, sheet "Outlet Mapping".
2. Do not edit "Cost Center Summary" manually; it is regenerated automatically from Outlet Mapping on every run.
3. Cost Center is the stable outlet identity. All old/new outlet names and Sold-to Party codes with the same Cost Center are combined as one outlet.
4. When a name or Sold-to Party changes, append a new row with the SAME Cost Center. Do not delete the old row; it remains a historical alias.
5. The last row for a Cost Center becomes the current display Outlet Name and Sold-to Party in the dashboard.
6. Where Cost Center is blank or #N/A, the dashboard uses Sold-to Party as the outlet identity.
7. A Sold-to Party not yet present in the mapping is also retained using Sold-to Party fallback; its sales are not excluded. Add it later to the mapping to merge it into the correct Cost Center.


ANALYTICS UPDATE (V17)
- Analytics now shows Top 20 outlets for each product and combined ranking.
- Daily Date-wise Sales Trend (Date vs Billed Qty in thousands of SKU) is restored.

OBJECTIVE FILES (TWO INDEPENDENT SOURCES)
- Keep exactly two objective workbooks in Input. Their filenames can be anything.
- The outlet-level workbook is detected by Area, Name, Cost Center and Jul-Jun columns. It supplies only individual outlet objectives.
- The area-level workbook is detected by its Sales Area summary. It supplies Main, DDM, Area, Analytics and aggregate objectives.
- Replace either workbook with a new file using any filename, but remove the old matching copy first.
- Outlet objectives are never summed to produce area objectives, so targets are not duplicated.
- See Input/README_OBJECTIVES.txt for the step-by-step replacement process.

