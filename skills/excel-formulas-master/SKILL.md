---
name: excel-formulas-master
description: "Complete reference of 500+ Excel functions organized by category. Use this skill whenever the user asks about formulas, calculations, lookups, text manipulation, date operations, or any Excel function. Provides syntax, best-practice usage, common pitfalls, and Office.js implementation patterns for every function category."
compatibility: HyperFix runtime. Works with all models. No external dependencies.
metadata:
  docs: https://support.microsoft.com/en-us/office/excel-functions-by-category-5f91f4e9-7b42-46d2-9bd1-63f26a86c0eb
---

# Excel Formulas Master Reference

Complete reference for writing, applying, and debugging Excel formulas via Office.js in HyperFix.

## Core Principle

Always write formulas into cells via Office.js — never compute values in code and paste results. The spreadsheet must stay dynamic.

```typescript
// CORRECT — formula stays live
range.formulas = [["=SUM(A1:A100)"]];

// WRONG — hardcoded dead value
range.values = [[5000]];
```

## Writing Formulas via Office.js

```typescript
// Single formula
const range = sheet.getRange("B10");
range.formulas = [["=SUM(B2:B9)"]];

// Array of formulas across a row
const row = sheet.getRange("B2:E2");
row.formulas = [["=A2*1.1", "=A2*1.2", "=A2*1.3", "=A2*1.4"]];

// Fill a column with relative formulas
const col = sheet.getRange("C2:C100");
col.formulas = Array.from({ length: 99 }, (_, i) => [`=A${i+2}+B${i+2}`]);

// Read existing formulas
range.load("formulas");
await context.sync();
console.log(range.formulas); // [["=SUM(B2:B9)"]]
```

---

## 1. Lookup & Reference Functions

The most critical category. Master these first.

### XLOOKUP (preferred over VLOOKUP)
```
=XLOOKUP(lookup_value, lookup_array, return_array, [if_not_found], [match_mode], [search_mode])
```
- **Use when:** Finding a value in any direction (replaces VLOOKUP + HLOOKUP + INDEX/MATCH)
- **match_mode:** 0 = exact, -1 = exact or next smaller, 1 = exact or next larger, 2 = wildcard
- **search_mode:** 1 = first-to-last, -1 = last-to-first, 2 = binary ascending, -2 = binary descending
- Example: `=XLOOKUP("Apple", A:A, C:C, "Not found")`

### VLOOKUP
```
=VLOOKUP(lookup_value, table_array, col_index_num, [range_lookup])
```
- **Always use FALSE** for exact match (4th argument)
- **Limitation:** Can only look right. Use XLOOKUP or INDEX/MATCH instead when looking left.
- **Pitfall:** col_index_num is 1-based. Inserting a column breaks it.
- Example: `=VLOOKUP(A2, Products!A:D, 3, FALSE)`

### HLOOKUP
```
=HLOOKUP(lookup_value, table_array, row_index_num, [range_lookup])
```
- Same as VLOOKUP but searches horizontally (first row)
- Rarely needed — prefer XLOOKUP

### INDEX + MATCH (classic power combo)
```
=INDEX(return_range, MATCH(lookup_value, lookup_range, 0))
```
- **Use when:** XLOOKUP unavailable, or you need 2D lookup
- **2D lookup:** `=INDEX(data, MATCH(row_val, row_range, 0), MATCH(col_val, col_range, 0))`
- MATCH 3rd arg: 0 = exact, 1 = less than (sorted asc), -1 = greater than (sorted desc)
- Example: `=INDEX(C2:C100, MATCH("Apple", A2:A100, 0))`

### INDEX
```
=INDEX(array, row_num, [col_num])
```
- Returns the value at a specific position in a range
- With 2 ranges: `=INDEX((A1:B10, D1:E10), 3, 2, 2)` — area_num selects which range

### MATCH
```
=MATCH(lookup_value, lookup_array, [match_type])
```
- Returns the **position** (not the value) of the match
- **Always use 0** for exact match unless data is sorted

### XMATCH
```
=XMATCH(lookup_value, lookup_array, [match_mode], [search_mode])
```
- Modern replacement for MATCH with wildcard and binary search support

### OFFSET
```
=OFFSET(reference, rows, cols, [height], [width])
```
- Returns a range shifted from a starting cell
- **Warning:** Volatile — recalculates on every change. Use INDEX when possible.

### INDIRECT
```
=INDIRECT(ref_text, [a1])
```
- Converts a text string to a cell reference
- Useful for dynamic sheet references: `=INDIRECT("'"&A1&"'!B2")`
- **Warning:** Volatile. Breaks if sheet/cell is renamed.

### CHOOSE
```
=CHOOSE(index_num, value1, value2, ...)
```
- Returns the nth value from a list
- Example: `=CHOOSE(MONTH(A1), "Q1","Q1","Q1","Q2","Q2","Q2","Q3","Q3","Q3","Q4","Q4","Q4")`

### ROW / COLUMN / ROWS / COLUMNS
```
=ROW([reference])          → row number of a cell
=COLUMN([reference])       → column number
=ROWS(array)               → count of rows in range
=COLUMNS(array)            → count of columns in range
```

### ADDRESS
```
=ADDRESS(row_num, col_num, [abs_num], [a1], [sheet_text])
```
- Creates a cell address as text: `=ADDRESS(1,3)` → `"$C$1"`

---

## 2. Logical Functions

### IF
```
=IF(logical_test, value_if_true, value_if_false)
```
- **Nested IFs:** Avoid deep nesting (>3 levels). Use IFS or SWITCH instead.
- Example: `=IF(A1>100, "High", IF(A1>50, "Medium", "Low"))`

### IFS
```
=IFS(condition1, value1, condition2, value2, ..., TRUE, default_value)
```
- Cleaner than nested IFs. Evaluates conditions in order, returns first TRUE match.
- **Always end with TRUE, default_value** as fallback.

### SWITCH
```
=SWITCH(expression, value1, result1, value2, result2, ..., default)
```
- Matches a single value against a list. Cleaner than IF chains for exact matching.
- Example: `=SWITCH(A1, "FR","France", "US","United States", "Unknown")`

### AND / OR / NOT / XOR
```
=AND(condition1, condition2, ...)   → TRUE if ALL are true
=OR(condition1, condition2, ...)    → TRUE if ANY is true
=NOT(logical)                       → inverts TRUE/FALSE
=XOR(condition1, condition2, ...)   → TRUE if ODD number are true
```

### IFERROR / IFNA
```
=IFERROR(value, value_if_error)     → catches ALL errors
=IFNA(value, value_if_na)           → catches only #N/A
```
- **Prefer IFNA** when wrapping lookups — it won't hide real formula errors.
- Example: `=IFNA(VLOOKUP(A1, data, 2, FALSE), "Not found")`

### LET
```
=LET(name1, value1, name2, value2, ..., calculation)
```
- Define named variables inside a formula. Reduces repetition and improves readability.
- Example: `=LET(tax, 0.2, price, A1, price * (1 + tax))`

### LAMBDA
```
=LAMBDA(parameter1, parameter2, ..., calculation)
```
- Create custom reusable functions. Assign to a Name for reuse.
- Example: `=LAMBDA(x, y, x^2 + y^2)(3, 4)` → 25

---

## 3. Math & Trigonometry Functions

### Aggregation
```
=SUM(range)                    → total
=SUMIF(range, criteria, sum_range)    → conditional sum
=SUMIFS(sum_range, criteria_range1, criteria1, ...)   → multi-criteria sum
=SUMPRODUCT(array1, array2, ...)      → sum of element-wise products
```
- **SUMPRODUCT trick:** `=SUMPRODUCT((A1:A100="Apple")*(B1:B100>10))` counts with multiple criteria
- **SUMIFS:** criteria_range and criteria come in pairs. Sum_range is FIRST argument.

### Rounding
```
=ROUND(number, num_digits)      → standard rounding
=ROUNDUP(number, num_digits)    → always rounds up
=ROUNDDOWN(number, num_digits)  → always rounds down (truncates)
=MROUND(number, multiple)       → round to nearest multiple
=CEILING(number, significance)  → round up to multiple
=FLOOR(number, significance)    → round down to multiple
=INT(number)                    → round down to integer
=TRUNC(number, [num_digits])    → truncate without rounding
```
- **num_digits:** positive = decimal places, 0 = integer, negative = tens/hundreds
- Example: `=ROUND(3.14159, 2)` → 3.14, `=ROUND(1234, -2)` → 1200

### Other Math
```
=ABS(number)                    → absolute value
=MOD(number, divisor)           → remainder after division
=POWER(number, power)           → or use ^ operator
=SQRT(number)                   → square root
=LOG(number, [base])            → logarithm (default base 10)
=LN(number)                     → natural logarithm
=EXP(number)                    → e raised to power
=RAND()                         → random 0-1 (volatile)
=RANDBETWEEN(bottom, top)       → random integer in range
=PI()                           → 3.14159265358979
=SIGN(number)                   → -1, 0, or 1
=PRODUCT(range)                 → multiply all values
=QUOTIENT(numerator, denominator) → integer division
=GCD(number1, number2, ...)     → greatest common divisor
=LCM(number1, number2, ...)     → least common multiple
```

### Array Math
```
=SEQUENCE(rows, [cols], [start], [step])    → generates sequence array
=RANDARRAY(rows, [cols], [min], [max], [whole_number])  → random array
```

---

## 4. Statistical Functions

### Central Tendency
```
=AVERAGE(range)                → arithmetic mean
=AVERAGEIF(range, criteria, [average_range])
=AVERAGEIFS(average_range, criteria_range1, criteria1, ...)
=MEDIAN(range)                 → middle value
=MODE(range)                   → most frequent value (single)
=MODE.MULT(range)              → all modes (array)
=GEOMEAN(range)                → geometric mean (for growth rates)
=HARMEAN(range)                → harmonic mean
=TRIMMEAN(range, percent)      → mean excluding outlier %
```

### Dispersion
```
=STDEV(range)                  → sample standard deviation (STDEV.S)
=STDEVP(range)                 → population standard deviation (STDEV.P)
=VAR(range)                    → sample variance (VAR.S)
=VARP(range)                   → population variance (VAR.P)
=MAX(range) / =MIN(range)
=MAXIFS(range, criteria_range, criteria, ...)
=MINIFS(range, criteria_range, criteria, ...)
=LARGE(range, k)               → kth largest value
=SMALL(range, k)               → kth smallest value
```

### Counting
```
=COUNT(range)                  → count numbers only
=COUNTA(range)                 → count non-empty cells
=COUNTBLANK(range)             → count empty cells
=COUNTIF(range, criteria)      → count matching criteria
=COUNTIFS(range1, criteria1, range2, criteria2, ...)
```
- **Criteria syntax:** `">100"`, `"<>0"`, `"Apple"`, `"*text*"` (wildcard), `">"&B1` (cell ref)

### Ranking & Percentiles
```
=RANK(number, ref, [order])              → rank (0=descending, 1=ascending)
=RANK.EQ(number, ref, [order])           → same, ties get same rank
=RANK.AVG(number, ref, [order])          → ties get average rank
=PERCENTILE(array, k)                    → kth percentile (k = 0 to 1)
=PERCENTRANK(array, x)                   → percentile rank of a value
=QUARTILE(array, quart)                  → 0=min, 1=Q1, 2=median, 3=Q3, 4=max
```

### Correlation & Regression
```
=CORREL(array1, array2)                  → Pearson correlation (-1 to 1)
=SLOPE(known_ys, known_xs)               → slope of linear regression
=INTERCEPT(known_ys, known_xs)           → y-intercept
=RSQ(known_ys, known_xs)                 → R-squared (coefficient of determination)
=LINEST(known_ys, known_xs, [const], [stats])  → full regression statistics (array)
=TREND(known_ys, known_xs, new_xs)       → predicted values from linear trend
=FORECAST(x, known_ys, known_xs)         → single forecast point
=FORECAST.LINEAR(x, known_ys, known_xs)  → same, explicit name
=FORECAST.ETS(target_date, values, timeline)  → exponential smoothing forecast
```

### Distribution Functions
```
=NORM.DIST(x, mean, stdev, cumulative)       → normal distribution
=NORM.INV(probability, mean, stdev)          → inverse normal
=NORM.S.DIST(z, cumulative)                  → standard normal
=T.DIST(x, deg_freedom, cumulative)          → Student's t-distribution
=CONFIDENCE.NORM(alpha, stdev, size)         → confidence interval
=CONFIDENCE.T(alpha, stdev, size)            → t-based confidence interval
```

---

## 5. Text Functions

### Extraction
```
=LEFT(text, [num_chars])           → first n characters
=RIGHT(text, [num_chars])          → last n characters
=MID(text, start_num, num_chars)   → substring from position
=TEXTBEFORE(text, delimiter, [instance_num])   → text before delimiter
=TEXTAFTER(text, delimiter, [instance_num])    → text after delimiter
=TEXTSPLIT(text, col_delimiter, [row_delimiter])  → split into array
```

### Transformation
```
=UPPER(text)                       → UPPERCASE
=LOWER(text)                       → lowercase
=PROPER(text)                      → Title Case
=TRIM(text)                        → remove extra spaces
=CLEAN(text)                       → remove non-printable chars
=SUBSTITUTE(text, old, new, [instance])  → replace text
=REPLACE(old_text, start, num_chars, new_text)  → replace by position
=REPT(text, number_times)          → repeat text
=CONCAT(text1, text2, ...)         → join text (or & operator)
=TEXTJOIN(delimiter, ignore_empty, text1, text2, ...)  → join with separator
```

### Information
```
=LEN(text)                         → character count
=FIND(find_text, within_text, [start])    → position (case-sensitive)
=SEARCH(find_text, within_text, [start])  → position (case-insensitive, wildcards)
=EXACT(text1, text2)               → case-sensitive comparison
=ISNUMBER(SEARCH("word", A1))      → check if text contains "word"
```

### Conversion
```
=VALUE(text)                       → convert text to number
=TEXT(value, format_text)          → format number as text
=NUMBERVALUE(text, [decimal_sep], [group_sep])  → locale-aware text-to-number
=CHAR(number)                      → character from ASCII code
=CODE(text)                        → ASCII code of first character
```

**TEXT format codes:**
```
=TEXT(A1, "0.00")          → "1234.50"
=TEXT(A1, "#,##0")         → "1,235"
=TEXT(A1, "$#,##0.00")     → "$1,234.50"
=TEXT(A1, "0.0%")          → "15.5%"
=TEXT(A1, "yyyy-mm-dd")    → "2026-03-29"
=TEXT(A1, "dddd")          → "Sunday"
=TEXT(A1, "mmmm yyyy")     → "March 2026"
```

---

## 6. Date & Time Functions

### Current Date/Time
```
=TODAY()                           → current date (volatile)
=NOW()                             → current date + time (volatile)
```

### Date Construction & Extraction
```
=DATE(year, month, day)            → create date
=YEAR(date) / =MONTH(date) / =DAY(date)
=HOUR(time) / =MINUTE(time) / =SECOND(time)
=TIME(hour, minute, second)        → create time
=DATEVALUE(date_text)              → text to date
=TIMEVALUE(time_text)              → text to time
```

### Date Arithmetic
```
=EDATE(start_date, months)         → add/subtract months
=EOMONTH(start_date, months)       → end of month after adding months
=DATEDIF(start, end, unit)         → difference ("Y","M","D","YM","MD","YD")
=DAYS(end_date, start_date)        → days between dates
=DAYS360(start, end, [method])     → days on 360-day year (finance)
=NETWORKDAYS(start, end, [holidays])  → working days between dates
=WORKDAY(start, days, [holidays])  → date after n working days
=WEEKDAY(date, [return_type])      → day of week (1=Sunday default)
=WEEKNUM(date, [return_type])      → week number of year
=ISOWEEKNUM(date)                  → ISO week number
```

### Age / Tenure Calculation
```
=DATEDIF(birth_date, TODAY(), "Y")                           → years
=DATEDIF(A1,TODAY(),"Y")&" years "&DATEDIF(A1,TODAY(),"YM")&" months"  → full age
```

---

## 7. Financial Functions

### Time Value of Money
```
=PV(rate, nper, pmt, [fv], [type])         → present value
=FV(rate, nper, pmt, [pv], [type])         → future value
=PMT(rate, nper, pv, [fv], [type])         → periodic payment
=NPER(rate, pmt, pv, [fv], [type])         → number of periods
=RATE(nper, pmt, pv, [fv], [type], [guess]) → interest rate per period
=IPMT(rate, per, nper, pv, [fv], [type])   → interest portion of payment
=PPMT(rate, per, nper, pv, [fv], [type])   → principal portion of payment
```
- **type:** 0 = end of period (default), 1 = beginning of period
- **rate must match period:** monthly payments → rate/12, nper*12

### Investment Analysis
```
=NPV(rate, value1, value2, ...)            → net present value (cash flows at END of periods)
=XNPV(rate, values, dates)                 → NPV with irregular dates
=IRR(values, [guess])                      → internal rate of return
=XIRR(values, dates, [guess])             → IRR with irregular dates
=MIRR(values, finance_rate, reinvest_rate) → modified IRR
```
- **NPV pitfall:** Does NOT include time-0 investment. Use `=NPV(rate, CF1:CFn) + CF0`

### Depreciation
```
=SLN(cost, salvage, life)                  → straight-line
=DB(cost, salvage, life, period, [month])  → declining balance
=DDB(cost, salvage, life, period, [factor]) → double declining balance
=SYD(cost, salvage, life, per)             → sum-of-years digits
```

### Bond & Securities
```
=ACCRINT(issue, first_interest, settlement, rate, par, frequency)
=PRICE(settlement, maturity, rate, yld, redemption, frequency)
=YIELD(settlement, maturity, rate, pr, redemption, frequency)
=DURATION(settlement, maturity, coupon, yld, frequency)
=MDURATION(settlement, maturity, coupon, yld, frequency)
```

---

## 8. Information Functions

```
=ISBLANK(value)          → TRUE if cell is empty
=ISNUMBER(value)         → TRUE if number
=ISTEXT(value)           → TRUE if text
=ISLOGICAL(value)        → TRUE if TRUE/FALSE
=ISERROR(value)          → TRUE if any error
=ISNA(value)             → TRUE if #N/A
=ISFORMULA(reference)    → TRUE if cell contains formula
=TYPE(value)             → 1=number, 2=text, 4=logical, 16=error, 64=array
=CELL("type", reference) → cell info ("b"=blank, "l"=label, "v"=value)
=FORMULATEXT(reference)  → shows formula as text string
```

---

## 9. Dynamic Array Functions (Microsoft 365)

These spill results automatically into adjacent cells.

```
=FILTER(array, include, [if_empty])        → filter rows by condition
=SORT(array, [sort_index], [sort_order], [by_col])  → sort array
=SORTBY(array, by_array1, [order1], ...)   → sort by another column
=UNIQUE(array, [by_col], [exactly_once])   → remove duplicates
=SEQUENCE(rows, [cols], [start], [step])   → number sequence
=RANDARRAY(rows, [cols], [min], [max], [integer])
=CHOOSECOLS(array, col_num1, ...)          → select columns
=CHOOSEROWS(array, row_num1, ...)          → select rows
=TOCOL(array, [ignore], [scan_by_column])  → flatten to single column
=TOROW(array, [ignore], [scan_by_column])  → flatten to single row
=WRAPCOLS(vector, wrap_count, [pad_with])  → wrap into columns
=WRAPROWS(vector, wrap_count, [pad_with])  → wrap into rows
=TAKE(array, rows, [cols])                 → first/last n rows/cols
=DROP(array, rows, [cols])                 → remove first/last n rows/cols
=EXPAND(array, rows, [cols], [pad_with])   → expand array size
=VSTACK(array1, array2, ...)              → stack vertically
=HSTACK(array1, array2, ...)              → stack horizontally
```

**Power combo examples:**
```
=SORT(UNIQUE(A2:A100))                              → sorted unique list
=FILTER(A2:D100, B2:B100>1000, "No results")        → filter by condition
=SORTBY(FILTER(A:D, C:C="Active"), D:D, -1)         → filter then sort descending
=UNIQUE(CHOOSECOLS(data, 1, 3))                      → unique values from specific columns
```

---

## 10. Database Functions

```
=DSUM(database, field, criteria)
=DAVERAGE(database, field, criteria)
=DCOUNT(database, field, criteria)
=DCOUNTA(database, field, criteria)
=DMAX(database, field, criteria)
=DMIN(database, field, criteria)
=DGET(database, field, criteria)
```
- **database:** range including headers
- **field:** column name as text or column number
- **criteria:** separate range with headers + conditions

---

## 11. Engineering Functions

```
=CONVERT(number, from_unit, to_unit)       → unit conversion
=BIN2DEC(number) / =DEC2BIN(number)        → binary ↔ decimal
=HEX2DEC(number) / =DEC2HEX(number)        → hex ↔ decimal
=OCT2DEC(number) / =DEC2OCT(number)        → octal ↔ decimal
=COMPLEX(real, imaginary)                   → create complex number
=IMREAL(complex) / =IMAGINARY(complex)      → extract parts
```

**CONVERT units:** "m","km","mi","ft","in","cm","kg","g","lbm","oz","C","F","K","hr","mn","sec","l","gal","m2","ft2"

---

## 12. Error Prevention Checklist

Before delivering any formula solution via Office.js:

1. **Zero #REF! errors** — verify all cell references exist
2. **Zero #DIV/0!** — wrap divisions: `=IF(B1=0, 0, A1/B1)` or `=IFERROR(A1/B1, 0)`
3. **Zero #N/A** — wrap lookups: `=IFNA(XLOOKUP(...), "Not found")`
4. **Zero #VALUE!** — check data types match (numbers vs text)
5. **Zero #NAME?** — verify function names are spelled correctly
6. **No volatile functions** unless necessary (OFFSET, INDIRECT, NOW, TODAY, RAND)
7. **Absolute references** where needed (`$A$1` vs `A1`)
8. **Consistent formula patterns** across rows/columns — no manual overrides in middle of a range

## 13. Office.js Formula Patterns

### Apply formula to entire column
```typescript
const range = sheet.getRange("C2:C1000");
range.formulas = Array.from({ length: 999 }, (_, i) => [`=A${i+2}*B${i+2}`]);
```

### Apply VLOOKUP across a range
```typescript
const range = sheet.getRange("D2:D100");
range.formulas = Array.from({ length: 99 }, (_, i) =>
  [`=IFERROR(VLOOKUP(A${i+2},Sheet2!A:C,3,FALSE),"")`]
);
```

### Read formula results
```typescript
const range = sheet.getRange("A1:D10");
range.load(["values", "formulas", "numberFormat"]);
await context.sync();
// range.values = computed results
// range.formulas = formula strings
```

### Set number format alongside formula
```typescript
const range = sheet.getRange("B2:B100");
range.formulas = Array.from({ length: 99 }, (_, i) => [`=A${i+2}*0.2`]);
range.numberFormat = Array.from({ length: 99 }, () => ["$#,##0.00"]);
```
