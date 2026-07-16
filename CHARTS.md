# Excel Chart Instructions

## Create a Daily Availability Chart

Assume:

- Dates are in column `A`
- Available counts are in column `W`

### Method 1: Select Two Non-Adjacent Columns

1. Select the dates, for example `A5:A22`.
2. Hold `Ctrl`.
3. Select the counts, for example `W5:W22`.
4. Open **Insert**.
5. Select **Line with Markers** or **Clustered Column**.
6. Change the title to `Daily Member Availability`.

## Copy a Chart for Another Sheet

1. Click the existing chart.
2. Press `Ctrl + C`.
3. Open the other monthly sheet.
4. Press `Ctrl + V`.
5. Right-click the copied chart.
6. Select **Select Data**.

### Change the Values

Under **Legend Entries (Series)**:

1. Select `Available Count`.
2. Click **Edit**.
3. Enter the new series range:

```excel
='August 2026'!$W$5:$W$22
```

### Change the Dates

Under **Horizontal (Category) Axis Labels**:

1. Click **Edit**.
2. Enter the new date range:

```excel
='August 2026'!$A$5:$A$22
```

3. Click **OK** twice.

## Common Error

Do not enter two separate ranges in the single **Chart data range** box like this:

```excel
='August 2026'!$A$5:$A$22,='August 2026'!$W$5:$W$22
```

Excel may reject that formula.

Edit the series values and category labels separately instead.

## Recommended Chart Types

| Data | Recommended chart |
|---|---|
| Daily availability over time | Line chart |
| Compare daily totals | Column chart |
| Compare departments | Bar chart |
| Share of one total | Pie chart |
| Actual vs target | Combo chart |

## Add Data Labels

1. Click the chart.
2. Click the `+` button beside the chart.
3. Tick **Data Labels**.
4. Choose **Above** or **Outside End**.

## Update the Chart Title

Click the chart title and type:

```text
Daily Member Availability — August 2026
```

## Make the Vertical Axis Start at Zero

1. Right-click the vertical axis.
2. Select **Format Axis**.
3. Under **Bounds**, set **Minimum** to `0`.
4. Set **Maximum** to `20` when the maximum team size is 20.

This prevents misleading visual differences.
