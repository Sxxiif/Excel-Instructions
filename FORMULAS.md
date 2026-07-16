# Essential Excel Formulas

This manual explains the Excel formulas worth learning first.

---

## 1. SUM

Adds numbers.

```excel
=SUM(A2:A20)
```

Example use: total tasks, hours, expenses, or sales.

---

## 2. COUNT

Counts cells containing numbers.

```excel
=COUNT(A2:A20)
```

Text and empty cells are ignored.

---

## 3. COUNTA

Counts all non-empty cells.

```excel
=COUNTA(C5:V5)
```

Best use: count the number of member names entered for one day.

---

## 4. COUNTBLANK

Counts empty cells.

```excel
=COUNTBLANK(C5:V5)
```

Useful for finding how many member slots are unused.

---

## 5. COUNTIF

Counts cells matching one condition.

```excel
=COUNTIF(B2:B100,"Present")
```

Other examples:

```excel
=COUNTIF(C2:C100,">10")
=COUNTIF(D2:D100,"Ahmed")
=COUNTIF(E2:E100,"<>")
```

`<>` means not empty.

---

## 6. COUNTIFS

Counts rows matching multiple conditions.

```excel
=COUNTIFS(B2:B100,"Monday",C2:C100,"Present")
```

All conditions must be true.

---

## 7. SUMIF

Adds values matching one condition.

```excel
=SUMIF(A2:A100,"Ahmed",B2:B100)
```

This adds values in column B where column A equals Ahmed.

---

## 8. SUMIFS

Adds values matching multiple conditions.

```excel
=SUMIFS(C2:C100,A2:A100,"Team 1",B2:B100,"Completed")
```

The first range is the range to add.

---

## 9. AVERAGE

Calculates the arithmetic average.

```excel
=AVERAGE(D2:D20)
```

---

## 10. AVERAGEIF

Calculates an average for matching values.

```excel
=AVERAGEIF(W5:W100,">0")
```

This ignores zero availability values.

---

## 11. MAX

Returns the highest value.

```excel
=MAX(W5:W100)
```

---

## 12. MIN

Returns the lowest value.

```excel
=MIN(W5:W100)
```

When blank rows are represented as zero, use:

```excel
=MINIFS(W5:W100,W5:W100,">0")
```

---

## 13. IF

Returns one result when a condition is true and another when it is false.

```excel
=IF(W5>=15,"Good Coverage","Low Coverage")
```

Syntax:

```excel
=IF(condition,value_if_true,value_if_false)
```

---

## 14. AND

Returns `TRUE` only when all conditions are true.

```excel
=IF(AND(B2="Present",C2>10),"Approved","Review")
```

---

## 15. OR

Returns `TRUE` when at least one condition is true.

```excel
=IF(OR(B2="Leave",B2="Sick Leave"),"Absent","Working")
```

---

## 16. IFERROR

Replaces errors with a cleaner result.

```excel
=IFERROR(A2/B2,0)
```

Lookup example:

```excel
=IFERROR(XLOOKUP(A2,E:E,F:F),"Not Found")
```

---

## 17. XLOOKUP

Finds a value and returns a corresponding result.

```excel
=XLOOKUP(A2,E2:E100,F2:F100)
```

Example: find an employee ID and return the employee name.

Exact match is the default.

---

## 18. VLOOKUP

An older lookup function.

```excel
=VLOOKUP(A2,E2:H100,3,FALSE)
```

`FALSE` requests an exact match.

---

## 19. INDEX + MATCH

A flexible lookup combination.

```excel
=INDEX(F2:F100,MATCH(A2,E2:E100,0))
```

`MATCH` locates the row. `INDEX` returns the value.

---

## 20. FILTER

Returns only matching records.

```excel
=FILTER(A2:D100,C2:C100="Present")
```

Requires Microsoft 365 or a recent Excel version.

---

## 21. UNIQUE

Returns distinct values.

```excel
=UNIQUE(A2:A100)
```

Useful for unique names, teams, or statuses.

---

## 22. SORT

Sorts a range dynamically.

```excel
=SORT(A2:A100)
```

Combined example:

```excel
=SORT(UNIQUE(A2:A100))
```

---

## 23. TEXT

Displays a value in a specific format.

```excel
=TEXT(A2,"dd-mmm-yyyy")
```

Other formats:

```excel
=TEXT(A2,"dddd")
=TEXT(A2,"mmm")
=TEXT(B2,"0.0%")
```

---

## 24. TODAY

Returns the current date.

```excel
=TODAY()
```

---

## 25. NOW

Returns the current date and time.

```excel
=NOW()
```

---

## 26. DATE

Creates a valid date.

```excel
=DATE(2026,7,16)
```

---

## 27. DAY, MONTH, and YEAR

Extract parts of a date.

```excel
=DAY(A2)
=MONTH(A2)
=YEAR(A2)
```

---

## 28. LEFT

Returns characters from the left.

```excel
=LEFT(A2,3)
```

---

## 29. RIGHT

Returns characters from the right.

```excel
=RIGHT(A2,2)
```

---

## 30. MID

Returns characters from the middle.

```excel
=MID(A2,3,5)
```

---

## 31. LEN

Counts characters.

```excel
=LEN(A2)
```

---

## 32. TRIM

Removes unnecessary spaces.

```excel
=TRIM(A2)
```

---

## 33. CONCAT

Joins text.

```excel
=CONCAT(A2," ",B2)
```

Alternative:

```excel
=A2&" "&B2
```

---

## 34. ROUND

Rounds a number.

```excel
=ROUND(A2,2)
```

Other versions:

```excel
=ROUNDUP(A2,0)
=ROUNDDOWN(A2,0)
```

---

## 35. SUBTOTAL

Calculates totals that respond to filters.

```excel
=SUBTOTAL(9,W5:W100)
```

Common function numbers:

| Number | Calculation |
|---:|---|
| 1 | Average |
| 2 | Count |
| 3 | Count non-empty |
| 4 | Maximum |
| 5 | Minimum |
| 9 | Sum |

---

# Attendance and Availability Formulas

Count names in one day:

```excel
=COUNTA(C5:V5)
```

Count empty member slots:

```excel
=COUNTBLANK(C5:V5)
```

Check whether coverage is enough:

```excel
=IF(COUNTA(C5:V5)>=15,"Enough","Need More")
```

Total availability entries for a month:

```excel
=SUM(W5:W100)
```

Average available people per working day:

```excel
=AVERAGEIF(W5:W100,">0")
```

Highest available count:

```excel
=MAX(W5:W100)
```

Lowest positive available count:

```excel
=MINIFS(W5:W100,W5:W100,">0")
```

Number of working dates:

```excel
=COUNT(A5:A100)
```

---

# Excel Table References

When data is converted into an Excel Table using `Ctrl + T`, formulas become easier to read.

Count member names in the current row:

```excel
=COUNTA([@[Member 1]:[Member 20]])
```

Sum an entire table column:

```excel
=SUM(AvailabilityTable[Available Count])
```

Count records matching a status:

```excel
=COUNTIF(AvailabilityTable[Status],"Present")
```

---

# Formula Symbols

| Symbol | Meaning |
|---|---|
| `=` | Starts a formula |
| `:` | Range, such as `A2:A20` |
| `,` | Separates formula arguments |
| `$` | Locks a row or column |
| `"` | Surrounds text conditions |
| `>` | Greater than |
| `<` | Less than |
| `>=` | Greater than or equal |
| `<=` | Less than or equal |
| `<>` | Not equal |
| `&` | Joins text |

> Some regional Excel installations use semicolons `;` instead of commas `,`.
