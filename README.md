# Amazon Sales & Order MIS Report (Excel)

A formula-driven Excel MIS workbook on **128,969 Amazon India apparel order lines** (Mar–Jun 2022): daily, weekly and monthly reports, state / category / channel breakdowns, live data-quality checks, pivots with slicers, and a one-page dashboard with a month selector.

- **Business question:** where does revenue come from, and where is it leaking through cancellations?
- **What I built:** 13 sheets of live formulas (SUMIFS / COUNTIFS / lookups), a cleaned data sheet, 11 data-quality checks, 3 pivots on one shared cache with 2 slicers, and a VBA button that exports the dashboard to PDF.
- **Headline finding:** gross revenue is **₹7.17 Cr**, **14.2%** of lines are cancelled, and Merchant-fulfilled / Standard orders cancel at **17.5%** vs **12.9%** for Amazon-fulfilled / Expedited.

![Dashboard](images/dashboard.png)

## Key findings
- **Gross revenue ₹7.17 Cr** (71,670,759); cancellation rate **14.2%** (18,329 of 128,969 lines).
- **Average daily revenue fell ~16% from April to June** (₹874K → ₹738K). Lines per day dropped ~20% (≈1,636 → ≈1,300), while average line value rose 5.5% (₹626 → ₹660): a volume problem, not a price problem.
- **Merchant-fulfilled / Standard orders cancel at 17.5%** (6,861 of 39,277 lines), vs **12.9%** for Amazon-fulfilled / Expedited (11,420 of 88,609). Amazon / Standard has only 1,083 lines, so its 4.4% is not a reliable comparison.
- **Maharashtra is 17.1% of revenue**; the **South zone is 39.7%**; the top 10 states bring in ₹5.65 Cr (78.8%).
- **Set + kurta make up 77.0% of revenue** (49.9% + 27.1%).

## Recommendations
These are hypotheses grounded in the numbers above, not proven causes.
| Finding | Recommendation |
|---|---|
| Merchant / Standard cancels at 17.5% vs 12.9% for Amazon / Expedited | Test moving the top SKUs (Set, kurta) to Amazon fulfilment, or tighten merchant dispatch targets. If Merchant matched the Amazon / Expedited rate, about 1,799 fewer lines would be cancelled, worth about ₹11.7 lakh (1,799 × ₹650 average revenue per non-cancelled Merchant line; illustrative upper bound). |
| Daily revenue −16% Apr→Jun with line value up 5.5% | Look at demand and availability (stock-outs, traffic, ads) before touching price. |
| Set + kurta = 77% of revenue; top 10 states = 79% | Prioritise stock and fulfilment capacity for these; review very small categories (Dupatta: 3 lines, Saree: 164). |
| 229 non-cancelled lines have a blank amount | Fix at source: revenue is understated and cannot be recovered in Excel. |

## Metric definitions
| Metric | Definition in this workbook |
|---|---|
| Order lines | One row = one SKU in an order (128,969). Unique Order IDs: 120,378. |
| Gross revenue | Sum of Amount on lines whose status group is **not** Cancelled. It still includes returned and lost lines. |
| Net revenue | Gross revenue − value of Returned lines (₹1,384,559 → net ₹70,286,200). |
| Cancel % | Cancelled lines ÷ all lines. |
| Return % | Returned lines ÷ (lines − cancelled lines). The Dashboard KPI shows **Merchant orders only** (6.5%), because return statuses exist only for Merchant-fulfilled orders. The MIS sheets show the all-orders rate (1.9%). |
| Avg daily revenue | Gross revenue ÷ days of data (partial months handled via the Days of Data column). |

## Dataset
[Amazon Sale Report – Kaggle](https://www.kaggle.com/datasets/thedevastator/unlock-profits-with-e-commerce-sales-data). 128,975 raw rows; 128,969 after cleaning. The raw CSV is not in this repo.

## How it's built
### Workbook structure
| Sheet | What it does |
|---|---|
| **Dashboard** | 8 KPIs driven by a month dropdown (All / yyyy-mm), top state / category / channel callouts, 4 charts |
| **Monthly / Weekly / Daily MIS** | Order lines, units, gross revenue, cancellations, returns, net revenue, avg line value, growth %, days of data |
| **State MIS** | 38 rows (36 states/UTs + APO + Unknown) ranked by revenue, revenue share with data bars, zone lookup, zone summary |
| **Category MIS** | Category metrics + Category × Month revenue matrix (colour scale) |
| **Channel MIS** | Amazon-fulfilled vs Merchant-fulfilled × service level |
| **Pivot Analysis** | 3 PivotTables on one shared cache, 2 slicers, GETPIVOTDATA reconciliation block |
| **Data Quality** | 11 live PASS / FAIL / INFO checks + cleaning log |
| **Data** | Cleaned order lines; Status Group, Clean State and the INDEX/MATCH check column are formulas |
| **Mappings** | 13 raw statuses → 6 status groups; 47 state lookup keys (from 69 raw spellings) → clean state + zone |
| **Settings** | Alert thresholds (yellow input cells) that drive red / amber flags |
| **Guide** | Formula walkthrough and practice tasks |

### Key formulas
| Purpose | Cell | Formula |
|---|---|---|
| SUMIFS | `Monthly MIS!E5` | `=SUMIFS(d_Amt,d_Month,$A5,d_Group,"<>Cancelled")` |
| COUNTIFS | `Monthly MIS!F5` | `=COUNTIFS(d_Month,$A5,d_Group,"Cancelled")` |
| Growth % | `Monthly MIS!N6` | `=IF(B5<7,"n/a (prior partial)",IFERROR(M6/M5-1,""))` |
| State cleaning | `Data!P2` | `=IFERROR(VLOOKUP(TRIM(UPPER(O2)),StateMap,2,FALSE),"UNKNOWN")` |
| Data-quality check | `Data Quality!C12` | `=IF(ABS('Monthly MIS'!E9-SUMIFS(d_Amt,d_Group,"<>Cancelled"))<1,"OK","MISMATCH")` |

`d_Amt`, `d_Month`, `d_Group` etc. are named ranges over `Data!…2:…128970`.

### Data quality results
- Unmapped statuses: **0** (PASS). Revenue reconciles between Monthly MIS and Data: **OK** (PASS).
- 33 lines with an unknown / blank state; 7,792 blank amounts, of which **229 are on non-cancelled lines** (revenue understated); 106 non-cancelled lines with Qty = 0.

## Data cleaning
- Removed 6 exact duplicate rows.
- Converted text dates (mm-dd-yy) to real dates, and added Month and Week Start (Monday) keys.
- Mapped 69 raw state spellings (e.g. `RJ`, `Rajshthan`, `rajasthan`) to 36 states/UTs plus APO and Unknown, using `VLOOKUP(TRIM(UPPER()))` against a 47-key mapping table.
- Grouped 13 order statuses into 6 groups: Delivered, Shipped, Pending, Cancelled, Returned, Lost/Damaged.
- Kept blank amounts blank (not zero) and flagged 229 non-cancelled lines with no amount.

## Pivot Analysis & Slicers
The **Pivot Analysis** sheet (after Channel MIS) holds three PivotTables built on **one shared pivot cache** (`Data!A1:Q128970`), so the data is cached once, one refresh updates all three, and one slicer can filter every pivot.

| Pivot | Rows × Columns | Values | Filter |
|---|---|---|---|
| `ptCategoryMonth` | Category × Month | Sum of Amount (₹), `#,##0` | Status Group ≠ Cancelled |
| `ptTopStates` | State (Clean), sorted descending | Sum of Amount (₹) + Count of Order ID | Status Group ≠ Cancelled, Top 10 |
| `ptChannel` | Fulfilment × Status Group | Count of Order ID | none |

- **Slicers:** Fulfilment and Month, each connected to all three pivots. Pivots refresh on open.
- **Reconciliation block:** `GETPIVOTDATA` minus the formula reports; all four differences are **0 / PASS**:
  - Pivot 1 grand total ₹71,670,759 = `Category MIS` total = `Dashboard!C8`.
  - All 36 Category × Month cells match the `Category MIS` matrix.
  - The May-2022 column ₹23,952,062 = `Monthly MIS` May gross revenue.

## Lookup cross-check (INDEX/MATCH, XLOOKUP-ready)
- **`Data!R` "Status Group (INDEX/MATCH)"** re-derives each status group independently of the VLOOKUP in column F: `=IFERROR(INDEX(Mappings!$B$5:$B$17,MATCH(E2,Mappings!$A$5:$A$17,0)),"UNMAPPED")`. It runs in Excel 2016; in Excel 2021 / 365 the equivalent is `=XLOOKUP(E2,Mappings!$A$5:$A$17,Mappings!$B$5:$B$17,"UNMAPPED")`.
- **Data Quality check #11** counts mismatches between the two lookups: `=SUMPRODUCT(--(Data!F2:F128970<>Data!R2:R128970))`. Result: **0 / PASS** across 128,969 rows.

## Automation (VBA)
The `.xlsm` version contains module **`modMIS`** and a **Refresh & Export PDF** button on the Dashboard (next to F5). Each click recalculates, refreshes all pivots, and writes a dated PDF of the Dashboard next to the workbook.

```vba
Sub RefreshAndExportMIS()
    Dim f As String
    Application.CalculateFull                       ' recalc every formula
    ThisWorkbook.RefreshAll                         ' refresh the pivot cache (all 3 pivots)
    f = ThisWorkbook.Path & "\MIS_Dashboard_" & Format(Date, "yyyy-mm-dd") & ".pdf"
    ThisWorkbook.Sheets("Dashboard").ExportAsFixedFormat Type:=xlTypePDF, Filename:=f
    MsgBox "Dashboard exported: " & f
End Sub
```

## Limitations
- Return and Delivered statuses exist only for Merchant-fulfilled orders, so the Return % KPI is shown for Merchant orders only, and net revenue is overstated for Amazon-fulfilled orders.
- March has one day of data (31-Mar) and the first / last weeks are partial. Growth % therefore uses average daily revenue.
- "Order Lines" counts SKU lines, not unique orders (120,378 unique Order IDs).
- **Fixed ranges:** the report formulas use named ranges fixed to `Data` rows 2–128970, and the pivots read `Data!A1:Q128970`. New rows need those ranges extended (or `Data` converted to an Excel Table); the data-quality checks use the same ranges, so they would not flag rows added below row 128970.

## Files
| File | Contents |
|---|---|
| `Amazon_Sales_MIS_Report.xlsx` | Full workbook, **no macros** (safe to open anywhere) |
| `Amazon_Sales_MIS_Report.xlsm` | Same workbook **plus** the `modMIS` macro and the export button |

## How to use
1. Open the `.xlsx` in Excel 2016 or later. Formulas calculate on open; if Excel opens it in Protected View, click Enable Editing so the pivots refresh.
2. Pick a month in **Dashboard!C5**.
3. Change thresholds in **Settings** to change the alert flags.
4. For the automation, open the `.xlsm`, click Enable Content, then click **Refresh & Export PDF** on the Dashboard.

## About
Built by Vitthal Misal · [LinkedIn](https://www.linkedin.com/in/vitthal-misal-analyst) · [GitHub: Vitthal38](https://github.com/Vitthal38)
