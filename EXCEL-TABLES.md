# Excel Tables Guide

## Convert Data into a Table

1. Click any cell inside the data.
2. Press `Ctrl + T`.
3. Confirm the selected range.
4. Tick **My table has headers**.
5. Click **OK**.

## Why Tables Are Useful

- Filters appear automatically.
- Formatting expands automatically.
- Formulas copy down automatically.
- Charts and PivotTables are easier to maintain.
- Structured formulas are easier to read.

## Rename a Table

1. Click inside the table.
2. Open **Table Design**.
3. Locate **Table Name**.
4. Enter a name without spaces, such as:

```text
AvailabilityTable
```

## Structured References

Standard formula:

```excel
=SUM(W5:W22)
```

Table formula:

```excel
=SUM(AvailabilityTable[Available Count])
```

Current-row count:

```excel
=COUNTA([@[Member 1]:[Member 20]])
```

## Total Row

1. Click inside the table.
2. Open **Table Design**.
3. Tick **Total Row**.
4. Use the dropdown in the total row to choose Sum, Average, Count, Maximum, or Minimum.
