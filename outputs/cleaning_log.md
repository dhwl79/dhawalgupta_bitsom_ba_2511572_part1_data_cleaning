# Cleaning Log — Retail Order Data

**Source file:** `data/raw_orders.xlsx` (sheet `raw_orders`, 932 rows, unchanged — verified
byte-for-byte identical to the original export at every step)
**Output file:** `data/cleaned_orders.xlsx` (912 rows after removing exact duplicates)
**Supporting reports:** `outputs/data_quality_report.xlsx`, `outputs/pivot_summary.xlsx`

All figures in this log are reproducible: every number in `data_quality_report.xlsx` and
`pivot_summary.xlsx` is a live Excel formula (COUNTIF/COUNTIFS/SUMIFS) computed against the
cleaned data, not a hardcoded value.

---

## 1. Issues found in the raw data

| # | Issue | Field(s) | Count |
|---|---|---|---|
| 1 | Inconsistent text casing/spacing | `customer_name`, `segment`, `region`, `category`, `sub_category`, `ship_mode`, `payment_status`, `order_status` | ~430 cells across these 8 fields (see `outputs/data_quality_report.xlsx`) |
| 2 | Mixed date formats within the same column | `order_date`, `ship_date` | All 932 rows (4 formats mixed) |
| 3 | Missing values | `region` (25), `ship_mode` (21), `discount` (18) | 64 cells (post-dedup counts) |
| 4 | Discount stored as percentage text instead of decimal | `discount` | 8 rows (`'70%'`, `'85%'`) |
| 5 | Negative discounts | `discount` | 15 rows (post-dedup) |
| 6 | Unusually high discounts (>50%) | `discount` | 15 rows (post-dedup) |
| 7 | Ship date earlier than order date | `order_date`/`ship_date` | 21 rows (post-dedup) |
| 8 | Reported `sales` ≠ recalculated sales | `sales` | 54 rows |
| 9 | Reported `profit` ≠ recalculated profit | `profit` | 54 rows |
| 10 | Exact duplicate rows | entire row | 20 rows |
| 11 | Same `order_id` with conflicting data | `order_id` + others | 12 order_ids / 24 rows |
| 12 | `customer_id` linked to more than one `customer_name` | `customer_id`, `customer_name` | 42 customer_ids (observed, not corrected — see Limitations) |

No missing dates, no unrecognized date text, no invalid dates (e.g. day=32), and no special
characters or genuine synonym variants (e.g. `'Tech'` vs `'Technology'`) were found anywhere
in the dataset.

## 2. Cleaning actions performed

### Text fields (Task 2)
`customer_name`, `segment`, `region`, `state`, `city`, `category`, `sub_category`,
`ship_mode`, `payment_status`, `order_status` are cleaned with a live Excel formula —
not pre-cleaned and pasted as static text:
```
=PROPER(TRIM(SUBSTITUTE(<raw cell>,CHAR(160)," ")))
```
- `SUBSTITUTE(...,CHAR(160)," ")` — replaces any non-breaking space with a normal one
  (defensive; none were actually found, but `TRIM` alone can't catch them).
- `TRIM(...)` — removes leading/trailing spaces and collapses repeated spaces to one.
- `PROPER(...)` — standardizes casing to Title Case.

The original raw text lives in the `Raw_Text_Fields` sheet of `cleaned_orders.xlsx`, so the
cleaning is fully auditable.

### Dates (Task 3)
`order_date` and `ship_date` mixed 4 formats in the same column: `YYYY-MM-DD`,
`MM/DD/YYYY`, `DD-MM-YYYY`, and `DD Mon YYYY` (e.g. `21 Jul 2024`). Each format was confirmed
unambiguous before parsing (the slash format never had a first component above 12, confirming
month-first; the dash format never had a second component above 12, confirming day-first).
Parsed dates were cross-checked against the year embedded in `order_id` (e.g.
`ORD-2024-10001` → 2024) with **zero mismatches** across all 932 rows.

Excel/LibreOffice's own automatic date interpretation was **not** relied on for parsing,
because it's locale-dependent and can't reliably distinguish `DD-MM-YYYY` from `MM-DD-YYYY`
without ambiguity. Dates were parsed programmatically with explicit format detection, then
loaded as real Excel date values. `shipping_delay_days`, `order_month`, `order_year`, and the
ship-before-order check are all live formulas operating on those date values.

### Discounts (Task 5)
Percentage-text values (`'70%'`) converted to decimals (`0.70`). Missing values defaulted to
0 (valid only because `quantity`/`unit_price`/`sales`/`cost`/`profit` were all present for
those rows). Negative and >50% values are flagged as invalid and kept visible in `discount`
for audit, but a separate `cleaned_discount` column treats them as 0% for use in
`calculated_sales`.

### Duplicates (Task 4)
- **Exact duplicates** (identical across every field) — 20 rows removed, first occurrence
  kept. Removed rows logged in `Removed_Duplicates_Log` (in `cleaned_orders.xlsx`) for audit.
- **Conflicting duplicates** (same `order_id`, different data) — 24 rows across 12 order_ids
  — **not removed**. No reliable way to know which copy is correct, so all copies are kept
  and flagged `Duplicate Order ID Conflict` / `data_quality_flag = Invalid` for manual review.

## 3. Business rules applied (Task 5)

| Rule Area | Required Action | How it was applied |
|---|---|---|
| Missing region | Fill as Unknown, flag | `region` filled `'Unknown'`; flagged in `original_data_flags` |
| Missing ship_mode | Fill as Unknown, flag | `ship_mode` filled `'Unknown'`; flagged in `original_data_flags` |
| Missing discount | Treat as 0 only if other sales fields valid | Defaulted to 0 — verified `quantity`/`unit_price`/`sales`/`cost`/`profit` were all present for every one of the 18 affected rows |
| Negative discount | Flag as invalid | Flagged `Invalid Discount (Negative)`; kept in `discount`, zeroed only in `cleaned_discount` |
| Discount above allowed range | Flag as invalid | Allowed range set at 0%–50% (raw valid data tops out at 25%); flagged `Invalid Discount (Too High)` |
| Cancelled orders | Must not contribute to completed-sales summary | `pivot_summary.xlsx` sales/profit pivots filter to `order_status="Completed"` |
| Failed payments | Must not contribute to completed-sales summary | Same pivots additionally filter to `payment_status="Paid"` |
| Refunded orders | Must be separately summarized | `pivot_summary.xlsx` sheet `5_Refunded_Cancelled_Failed` summarizes refunded orders (and cancelled/failed) by region, completely separate from the completed-sales totals |
| Ship date before order date | Flag as invalid shipping record | Flagged `Invalid Shipping Record (Ship Before Order)` when `shipping_delay_days < 0` |

## 4. Assumptions made

- **"Completed sales" = `order_status = "Completed"` AND `payment_status = "Paid"`.** The
  source rule said not to include "non-completed/failed/refunded" records; since 2 rows had
  `Completed` + `Failed` (both also flagged as conflicting duplicates), the stricter
  AND-filter was used to be unambiguous.
- **Allowed discount range = 0%–50%.** Not stated explicitly in the source rules; chosen
  because the clean, non-anomalous data tops out at 25%, so 50% is a generous cutoff above
  which a value is treated as a data error rather than a real promotion.
- **Calculation-mismatch tolerance = 1 currency unit.** Differences this size or smaller are
  treated as rounding noise, not genuine mismatches.
- **`profit_margin` uses `calculated_profit / calculated_sales`**, not the originally
  reported `profit`/`sales`, so the ratio is internally consistent with the other calculated
  columns even on rows where the reported figures don't reconcile.
- **`data_quality_flag` severity tiers:** `Invalid` = any discount/date/calculation/duplicate
  problem; `Warning` = only a missing value that was safely defaulted; `Clean` = no issues.

## 5. Records removed

- **20 exact duplicate rows** — removed from `cleaned_orders.xlsx`, logged in full in
  `Removed_Duplicates_Log` for audit. No other rows were removed for any reason; every other
  issue (invalid discount, bad date sequence, calculation mismatch, conflicting duplicate
  order_id) was flagged and kept, never deleted.

## 6. Records flagged

Out of 912 cleaned rows: **756 Clean**, **44 Warning**, **112 Invalid**
(see `outputs/data_quality_report.xlsx`, sheet `7_Final_Clean_vs_Flagged`, for the live,
reproducible breakdown). Specific flag counts:

| Flag | Count |
|---|---|
| Missing Region (filled Unknown) | 25 |
| Missing Ship Mode (filled Unknown) | 21 |
| Missing Discount (defaulted 0) | 18 |
| Invalid Discount (Negative) | 15 |
| Invalid Discount (Too High) | 15 |
| Invalid Shipping Record (Ship Before Order) | 21 |
| Sales Mismatch | 54 |
| Profit Mismatch | 54 |
| Duplicate Order ID Conflict | 24 |

(Rows can carry more than one flag, so these don't sum to 112.)

## 7. Limitations of this cleaning process

- **Conflicting duplicates were not resolved**, only flagged. Picking a "correct" version
  between two conflicting copies of the same order would require checking the source system —
  guessing risked silently introducing wrong sales/profit figures into the cleaned dataset.
- **42 `customer_id` values map to more than one `customer_name`** in the raw data. This
  wasn't covered by the stated business rules and there's no reliable way to tell which name
  is correct for a given ID, so it was left as-is and isn't reflected in any flag.
- **Pivot summaries are SUMIFS/COUNTIFS-based "pivot-style" tables, not native interactive
  Excel PivotTable objects** — `openpyxl` (the library used to build these files
  programmatically) does not reliably support writing functional PivotTable objects that
  Excel can refresh. The analytical output is equivalent, and the underlying `Data` sheet in
  `pivot_summary.xlsx` can be used to build a true interactive PivotTable manually (Insert →
  PivotTable) if needed.
- **`data_quality_report.xlsx` and `pivot_summary.xlsx` each contain their own embedded copy**
  of the cleaned dataset (sheet `Data`) rather than linking live to `cleaned_orders.xlsx` —
  cross-file formula references in Excel are path-dependent and fragile when a project folder
  is moved or shared, so each report is self-contained and will compute correctly on its own.
  All three files are generated from the same single cleaning run, so the numbers agree.
- **Sales/profit mismatches were investigated but not "fixed"** — the original reported
  `sales`/`profit` are kept as the source of truth in those columns; `calculated_sales` /
  `calculated_profit` show what the rule-based calculation produces, for comparison. Deciding
  which figure is "right" is a business judgment outside the scope of this cleaning pass.
