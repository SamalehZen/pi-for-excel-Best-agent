---
name: power-user-patterns
description: "Advanced Excel techniques and power-user patterns via Office.js. Use when the user needs LAMBDA functions, LET expressions, dynamic arrays, structured references, named ranges, array formulas, data validation, advanced filtering, or any technique beyond standard formulas."
compatibility: HyperFix runtime. Works with all models. No external dependencies.
---

# Power User Patterns

Advanced Excel techniques implemented via Office.js for expert-level spreadsheet work.

---

## 1. LAMBDA — Custom Reusable Functions

LAMBDA creates custom functions that can be named and reused like built-in functions.

### Define in a formula
```
=LAMBDA(x, x^2 + 2*x + 1)(5)           → 36
=LAMBDA(price, tax, price*(1+tax))(100, 0.2)  → 120
```

### Assign to a Named Range for reuse
Via Office.js:
```typescript
// Create a named LAMBDA function
context.workbook.names.add("PYTHAGOREAN", "=LAMBDA(a, b, SQRT(a^2 + b^2))");
context.workbook.names.add("TAX_CALC", "=LAMBDA(price, rate, price * (1 + rate))");
context.workbook.names.add("GRADE", '=LAMBDA(score, IFS(score>=90,"A", score>=80,"B", score>=70,"C", score>=60,"D", TRUE,"F"))');

await context.sync();
// Now users can type =PYTHAGOREAN(3, 4) → 5 anywhere in the workbook
```

### Recursive LAMBDA
```
// Factorial
=LAMBDA(n, IF(n<=1, 1, n * Factorial(n-1)))
// Assign to name "Factorial", then =Factorial(5) → 120
```

### LAMBDA with MAP, REDUCE, SCAN
```
=MAP(array, LAMBDA(x, x^2))               → square every element
=REDUCE(0, array, LAMBDA(acc, x, acc + x)) → sum (custom aggregation)
=SCAN(0, array, LAMBDA(acc, x, acc + x))   → running total
=BYCOL(array, LAMBDA(col, SUM(col)))       → sum each column
=BYROW(array, LAMBDA(row, MAX(row)))       → max of each row
=MAKEARRAY(5, 3, LAMBDA(r, c, r * c))     → multiplication table
```

---

## 2. LET — Named Variables in Formulas

Eliminates redundant calculations and improves readability.

### Basic LET
```
=LET(
  revenue, SUM(B2:B100),
  costs, SUM(C2:C100),
  profit, revenue - costs,
  margin, IF(revenue=0, 0, profit/revenue),
  TEXT(margin, "0.0%")
)
```

### LET with intermediate arrays
```
=LET(
  data, A2:D100,
  filtered, FILTER(data, INDEX(data,,3)>1000),
  sorted, SORT(filtered, 4, -1),
  TAKE(sorted, 10)
)
```
This filters rows where column C > 1000, sorts by column D descending, returns top 10.

### LET for DRY formulas
```
❌ Without LET (VLOOKUP computed 3 times):
=IF(VLOOKUP(A1,data,2,0)>100, VLOOKUP(A1,data,2,0)*1.1, VLOOKUP(A1,data,2,0)*0.9)

✅ With LET (computed once):
=LET(val, VLOOKUP(A1,data,2,0), IF(val>100, val*1.1, val*0.9))
```

---

## 3. Dynamic Arrays

Functions that return multiple values and "spill" into adjacent cells.

### Core Dynamic Array Functions
```
=FILTER(array, include, [if_empty])
=SORT(array, [sort_index], [sort_order])
=SORTBY(array, by_array, [order])
=UNIQUE(array, [by_col], [exactly_once])
=SEQUENCE(rows, [cols], [start], [step])
=RANDARRAY(rows, [cols], [min], [max], [integer])
```

### Spill Reference Operator (#)
Use `#` to reference the entire spill range of a dynamic formula:
```
Cell E1: =UNIQUE(A2:A100)          → spills unique values
Cell F1: =COUNTIF(A2:A100, E1#)   → counts each unique value using spill reference
Cell G1: =SORT(E1#)               → sorts the spill range
```

### Power Combos
```
// Unique sorted list
=SORT(UNIQUE(A2:A1000))

// Top 5 customers by revenue
=TAKE(SORTBY(UNIQUE(A2:A1000), SUMIF(A2:A1000, UNIQUE(A2:A1000), B2:B1000), -1), 5)

// Filter + Sort + Take
=LET(
  raw, A2:D1000,
  filtered, FILTER(raw, INDEX(raw,,2)="Active"),
  sorted, SORTBY(filtered, INDEX(filtered,,4), -1),
  TAKE(sorted, 20)
)

// Cross-join two lists
=LET(
  a, A2:A10,
  b, B2:B5,
  rows_a, ROWS(a),
  rows_b, ROWS(b),
  total, rows_a * rows_b,
  col1, INDEX(a, SEQUENCE(total,,1) - (SEQUENCE(total,,0) - MOD(SEQUENCE(total,,0), rows_b)) / rows_b * 0 + INT((SEQUENCE(total,,0))/rows_b) + 1),
  col2, INDEX(b, MOD(SEQUENCE(total,,0), rows_b) + 1),
  HSTACK(col1, col2)
)

// Unpivot with dynamic arrays
=LET(
  headers, B1:D1,
  labels, A2:A10,
  data, B2:D10,
  n_cols, COLUMNS(headers),
  n_rows, ROWS(labels),
  total, n_rows * n_cols,
  col1, INDEX(labels, INT((SEQUENCE(total,,0))/n_cols)+1),
  col2, INDEX(headers, , MOD(SEQUENCE(total,,0), n_cols)+1),
  col3, INDEX(data, INT((SEQUENCE(total,,0))/n_cols)+1, MOD(SEQUENCE(total,,0), n_cols)+1),
  HSTACK(col1, col2, col3)
)
```

---

## 4. Named Ranges

### Create via Office.js
```typescript
// Workbook-scoped name
context.workbook.names.add("TaxRate", "=Assumptions!$B$5");
context.workbook.names.add("DataRange", "=Data!$A$2:$D$1000");
context.workbook.names.add("Regions", '={"North","South","East","West"}');

// Sheet-scoped name
const sheet = context.workbook.worksheets.getItem("Revenue");
sheet.names.add("GrowthRate", "=Revenue!$B$3");

await context.sync();
```

### Dynamic Named Ranges (auto-expanding)
```typescript
// Using OFFSET+COUNTA (volatile but compatible)
context.workbook.names.add("DynamicData",
  "=OFFSET(Data!$A$1, 0, 0, COUNTA(Data!$A:$A), 4)"
);

// Using structured table reference (preferred)
// If data is in a Table named "SalesData":
// Just reference SalesData[Revenue] — auto-expands
```

### Use Cases for Named Ranges
| Name | Formula | Purpose |
|---|---|---|
| `TaxRate` | `=Assumptions!$B$5` | Central assumption |
| `StartDate` | `=Parameters!$B$2` | Report parameter |
| `Departments` | `=Admin!$A$2:$A$20` | Dropdown source |
| `CAGR` | LAMBDA formula | Custom function |
| `DataRange` | Table reference | Dynamic data source |

---

## 5. Structured Table References

When data is in an Excel Table, use structured references instead of cell addresses:

```
=SUM(SalesData[Revenue])                    → sum the Revenue column
=AVERAGE(SalesData[Margin])                 → average of Margin column
=SalesData[@Revenue]                        → current row's Revenue (in calculated column)
=SalesData[@Revenue] * SalesData[@Units]    → row-level calculation
=SalesData[[#Headers],[Revenue]]            → header text "Revenue"
=SalesData[[#Totals],[Revenue]]             → total row value
=SUMIFS(SalesData[Revenue], SalesData[Region], "North")
```

### Create Table via Office.js
```typescript
const range = sheet.getRange("A1:F100");
const table = sheet.tables.add(range, true);
table.name = "SalesData";

// Add calculated column
const newCol = table.columns.add(null, [
  ["Profit Margin"],
  ...Array.from({ length: 99 }, () => [null])
]);

// Set formula for calculated column using structured reference
const dataBody = table.columns.getItem("Profit Margin").getDataBodyRange();
dataBody.formulas = Array.from({ length: 99 }, () =>
  ["=IF([@Revenue]=0,0,([@Revenue]-[@COGS])/[@Revenue])"]
);
```

---

## 6. Advanced Data Validation

### Dependent Dropdowns (cascading)
```typescript
// First dropdown: Region
const regionRange = sheet.getRange("A2:A100");
regionRange.dataValidation.rule = {
  list: { inCellDropDown: true, source: "North,South,East,West" }
};

// Second dropdown: City depends on Region (using INDIRECT)
// Setup: Named ranges "North" = "NYC,Boston", "South" = "Miami,Atlanta", etc.
context.workbook.names.add("North", '"NYC,Boston,Chicago"');
context.workbook.names.add("South", '"Miami,Atlanta,Dallas"');

const cityRange = sheet.getRange("B2:B100");
cityRange.dataValidation.rule = {
  list: { inCellDropDown: true, source: "=INDIRECT(A2)" }
};
```

### Date Validation
```typescript
range.dataValidation.rule = {
  date: {
    formula1: "2026-01-01",
    formula2: "2026-12-31",
    operator: Excel.DataValidationOperator.between,
  }
};
range.dataValidation.errorAlert = {
  showAlert: true,
  title: "Invalid Date",
  message: "Please enter a date in 2026.",
  style: Excel.DataValidationAlertStyle.stop,
};
```

### Custom Formula Validation
```typescript
// No duplicates allowed
range.dataValidation.rule = {
  custom: { formula: "=COUNTIF(A:A,A2)<=1" }
};

// Must be a valid email
range.dataValidation.rule = {
  custom: { formula: '=AND(ISNUMBER(FIND("@",A2)),ISNUMBER(FIND(".",A2,FIND("@",A2))))' }
};

// Value must be greater than the cell above
range.dataValidation.rule = {
  custom: { formula: "=A2>A1" }
};
```

### Input Message (tooltip when cell selected)
```typescript
range.dataValidation.prompt = {
  showPrompt: true,
  title: "Enter Revenue",
  message: "Enter the quarterly revenue in USD. Must be positive.",
};
```

---

## 7. Array Formulas (Legacy CSE)

For older Excel versions without dynamic arrays, Ctrl+Shift+Enter (CSE) array formulas:

```
{=SUM(IF(A1:A100="Apple", B1:B100))}       → SUMIF equivalent
{=SUM((A1:A100="Apple")*(B1:B100>10)*C1:C100)}  → multi-criteria sum
{=INDEX(B1:B100, MATCH(1, (A1:A100="Apple")*(C1:C100="Large"), 0))}  → multi-criteria lookup
{=SUM(LARGE(A1:A100, ROW(INDIRECT("1:5"))))}    → sum top 5 values
```

Via Office.js, write these as regular formulas — Excel handles the array context:
```typescript
range.formulas = [['=SUM(IF(A1:A100="Apple", B1:B100))']];
```

---

## 8. Performance Optimization

### Volatile Functions to Avoid
These recalculate on EVERY change in the workbook:
```
OFFSET()     → use INDEX() instead
INDIRECT()   → use structured references or INDEX
NOW()        → cache in a cell, update manually
TODAY()      → same — use sparingly
RAND()       → use RANDARRAY() once, paste values
INFO()       → rarely needed
```

### Calculation Best Practices
| Pattern | Slow | Fast |
|---|---|---|
| Lookup | `VLOOKUP` in unsorted data | `XLOOKUP` or `INDEX/MATCH` |
| Whole column ref | `A:A` (1M rows) | `A2:A1000` (explicit range) |
| Nested IFs | 10 levels deep | `IFS()` or `SWITCH()` |
| Repeated calc | Same VLOOKUP in 5 formulas | `LET()` to compute once |
| Conditional sum | Array formula `{=SUM(IF(...))}` | `SUMIFS()` native function |
| Multiple criteria | Nested ANDs | `FILTER()` with `*` operator |

### Turn Off Screen Updating During Bulk Operations
```typescript
// Office.js handles this automatically via context.sync() batching
// But minimize sync() calls:

// ❌ Slow: sync after each operation
for (const row of data) {
  sheet.getRange(`A${i}`).values = [[row[0]]];
  await context.sync(); // DON'T DO THIS IN A LOOP
}

// ✅ Fast: batch all operations, sync once
const range = sheet.getRange(`A1:D${data.length}`);
range.values = data;
await context.sync(); // ONE sync for all data
```

---

## 9. Error Handling Patterns

### Comprehensive Error Wrapper
```
=LET(
  result, <your_formula>,
  IF(ISERROR(result),
    SWITCH(ERROR.TYPE(result),
      1, "NULL!",
      2, "DIV/0!",
      3, "VALUE!",
      4, "REF!",
      5, "NAME?",
      6, "NUM!",
      7, "N/A",
      "ERROR"),
    result
  )
)
```

### Safe Division
```
=IF(B1=0, 0, A1/B1)                        → returns 0
=IFERROR(A1/B1, 0)                          → catches all errors
=IF(OR(B1=0, ISBLANK(B1)), "-", A1/B1)    → returns dash for display
```

### Safe Lookup
```
=IFNA(XLOOKUP(key, range, return), "Not found")
=IFERROR(INDEX(MATCH(...)), "Missing")
```

### Cascading Fallback
```
=LET(
  v1, XLOOKUP(A1, primary_table, primary_values),
  v2, XLOOKUP(A1, backup_table, backup_values),
  IFNA(v1, IFNA(v2, "No match in any source"))
)
```

---

## 10. Useful Patterns Cheat Sheet

| Task | Formula |
|---|---|
| Running total | `=SUM($B$2:B2)` |
| Rank without ties | `=RANK(A2,$A$2:$A$100)+COUNTIF($A$2:A2,A2)-1` |
| Extract unique + count | `=UNIQUE(A2:A100)` + `=COUNTIF(A2:A100, E2#)` |
| Reverse a list | `=SORTBY(A2:A100, SEQUENCE(ROWS(A2:A100)), -1)` |
| Random sample of n rows | `=TAKE(SORTBY(data, RANDARRAY(ROWS(data))), n)` |
| Find last non-empty in column | `=LOOKUP(2,1/(A:A<>""),A:A)` |
| Nth occurrence lookup | `=INDEX(B:B,SMALL(IF(A2:A1000="Apple",ROW(A2:A1000)),N))` |
| Concatenate with condition | `=TEXTJOIN(", ",TRUE,IF(B2:B100="Active",A2:A100,""))` |
| Week start date (Monday) | `=A1-WEEKDAY(A1,3)` |
| Fiscal quarter | `=CHOOSE(MONTH(A1),3,3,3,4,4,4,1,1,1,2,2,2)` or `="Q"&INT((MONTH(A1)+2)/3)` |
| Remove duplicates keeping last | `=FILTER(data, COUNTIF(OFFSET(key,1,0,ROWS(key)-ROW(key)+ROW(TAKE(key,1)),1),key)=0)` |
| Generate date series | `=SEQUENCE(12,1,DATE(2026,1,1),30)` with date format |
