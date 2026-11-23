# Build a real PivotTable on this data (2 minutes)

`Monthly_Trend` is a **formula-driven pivot** (SUMIFS), so it updates automatically and never needs refreshing.
A native Excel PivotTable is still the fastest way to explore the data, so here is how to add one yourself:

1. Open `output/BranchPulse_MIS_Report.xlsx` and go to the **Data_Policies** sheet.
2. Click any cell in the data, then **Insert > PivotTable > New Worksheet**.
3. Drag **Month (formula)** to *Rows*, **LOB** to *Columns*, **Premium** to *Values* (Sum).
4. Right-click a Month, **Group...** by *Months* and *Years*.
5. Add **Branch_ID** to *Filters*, then **PivotTable Analyze > Insert Slicer** for *Channel* or *Policy_Type*.
6. Compare your numbers with `Monthly_Trend`. They should match to the rupee.

## VLOOKUP used in this workbook

| Where | Formula | Purpose |
|---|---|---|
| `Data_Claims!G` | `=IFERROR(VLOOKUP(B2, Data_Policies!$A$2:$D$6109, 4, FALSE), "NOT FOUND")` | Finds which branch each claim belongs to |
| `Branch_Summary!B:E` | `=VLOOKUP($A2, Branch_Master!$A$2:$E$13, 2..5, FALSE)` | Pulls branch name, region, manager and target |

`FALSE` forces an exact match. `IFERROR` turns a failed lookup into a visible `NOT FOUND`, which the `DQ_Report` sheet counts (it should be 0).
