# Excel Essential Guide

A practical Excel reference for formulas, tables, charts, attendance tracking, reporting, and everyday business tasks.

This repository is designed for beginners who want to learn the Excel functions commonly used in operational reports, dashboards, IT teams, HR trackers, project reporting, and administrative work.

## Repository Contents

```text
excel-essential-guide/
├── README.md
├── FORMULAS.md
├── CHARTS.md
├── SHORTCUTS.md
├── CONTRIBUTING.md
├── LICENSE
├── .gitignore
├── docs/
│   └── EXCEL-TABLES.md
└── examples/
    └── member-availability-example.xlsx
```

## Start Here

Download the workbook from the `examples` folder and open it in Microsoft Excel.

The sample workbook contains:

- Two months on separate sheets
- Monday–Thursday availability records
- Up to 20 members per day
- An automatic available-member count
- A summary sheet
- Daily availability charts

## Most Important Formula for the Sample

To count how many members are available in one row:

```excel
=COUNTA(C5:V5)
```

`COUNTA` counts all non-empty cells between `C5` and `V5`.

If the range contains 16 names and 4 empty cells, the result is `16`.

## Formula Categories

| Category | Main Functions |
|---|---|
| Totals | `SUM`, `SUMIF`, `SUMIFS` |
| Counting | `COUNT`, `COUNTA`, `COUNTIF`, `COUNTIFS` |
| Conditions | `IF`, `AND`, `OR`, `IFERROR` |
| Averages | `AVERAGE`, `AVERAGEIF`, `AVERAGEIFS` |
| Lookups | `XLOOKUP`, `VLOOKUP`, `INDEX`, `MATCH` |
| Text | `LEFT`, `RIGHT`, `MID`, `LEN`, `TRIM`, `CONCAT` |
| Dates | `TODAY`, `NOW`, `DATE`, `DAY`, `MONTH`, `YEAR`, `TEXT` |
| Dynamic arrays | `FILTER`, `UNIQUE`, `SORT` |
| Statistics | `MAX`, `MIN`, `LARGE`, `SMALL` |

See [FORMULAS.md](FORMULAS.md) for explanations and examples.

## Quick Attendance Examples

Count names in one day's row:

```excel
=COUNTA(C5:V5)
```

Count how many cells say `Present`:

```excel
=COUNTIF(B2:B100,"Present")
```

Count staff who were present on Monday:

```excel
=COUNTIFS(B2:B100,"Monday",C2:C100,"Present")
```

Find average availability while ignoring zero:

```excel
=AVERAGEIF(W5:W100,">0")
```

Find the highest daily availability:

```excel
=MAX(W5:W100)
```

Find the lowest positive availability:

```excel
=MINIFS(W5:W100,W5:W100,">0")
```

## Excel Tables

Select your data and press:

```text
Ctrl + T
```

Confirm **My table has headers**, then click **OK**.

Tables automatically provide:

- Filtering and sorting
- Consistent formatting
- Automatic formula expansion
- Structured references
- Easier chart and PivotTable creation

Example structured formula:

```excel
=COUNTA([@[Member 1]:[Member 20]])
```

## Charts

To create a daily availability chart:

1. Select the date column.
2. Hold `Ctrl`.
3. Select the available-count column.
4. Open **Insert**.
5. Choose a **Line chart** or **Clustered column chart**.
6. Rename the chart to `Daily Member Availability`.

To reuse the chart for another sheet:

1. Copy and paste the chart.
2. Right-click the copied chart.
3. Select **Select Data**.
4. Edit the series values.
5. Set the new month range, for example:

```excel
='August 2026'!$W$5:$W$22
```

6. Edit the horizontal axis labels:

```excel
='August 2026'!$A$5:$A$22
```

See [CHARTS.md](CHARTS.md) for detailed instructions.

## Recommended Learning Order

1. Cell references
2. `SUM`
3. `COUNT`, `COUNTA`
4. `COUNTIF`, `COUNTIFS`
5. `IF`
6. `SUMIF`, `SUMIFS`
7. `AVERAGE`
8. Excel Tables
9. Charts
10. `XLOOKUP`
11. `IFERROR`
12. `FILTER`, `UNIQUE`, `SORT`
13. PivotTables and PivotCharts

## Important Cell Reference Types

Relative reference:

```excel
=A2+B2
```

Changes when copied.

Absolute reference:

```excel
=A2*$F$1
```

`$F$1` remains fixed.

Mixed references:

```excel
=$A2
=A$2
```

Only the row or column remains fixed.

## Compatibility Notes

Some functions require newer versions of Excel:

- `XLOOKUP`
- `FILTER`
- `UNIQUE`
- `SORT`
- `MINIFS`
- `MAXIFS`

For older versions, use `VLOOKUP`, `INDEX` + `MATCH`, filters, and PivotTables.

## Repository Purpose

This repository is for learning and reference. The sample data uses fictional names and contains no real employee information.

## Author

Built and maintained by **sxxiif**.

## License

This project is available under the MIT License.
