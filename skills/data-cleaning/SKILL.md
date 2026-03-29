---
name: data-cleaning
description: "Data cleaning and transformation workflows for Excel via Office.js. Use when the user needs to fix messy data: remove duplicates, trim whitespace, fix formatting, standardize text, split/merge columns, handle missing values, validate data types, or restructure malformed tables."
compatibility: HyperFix runtime. Works with all models. No external dependencies.
---

# Data Cleaning & Transformation

Systematic workflows for cleaning messy spreadsheet data via Office.js in HyperFix.

---

## Diagnostic Phase — Always Start Here

Before cleaning, read the data and diagnose issues:

```typescript
// Step 1: Read the data
const usedRange = sheet.getUsedRange();
usedRange.load(["values", "rowCount", "columnCount", "numberFormat"]);
await context.sync();

// Step 2: Analyze
const data = usedRange.values;
const headers = data[0];
const rowCount = usedRange.rowCount;
const colCount = usedRange.columnCount;
```

### Common Issues to Detect
| Issue | How to Detect | Formula Check |
|---|---|---|
| Leading/trailing spaces | `=LEN(A1)<>LEN(TRIM(A1))` | TRIM won't match original |
| Non-printable characters | `=LEN(A1)<>LEN(CLEAN(A1))` | CLEAN changes length |
| Numbers stored as text | `=ISTEXT(A1)` when expecting number | TYPE() returns 2 |
| Text stored as numbers | `=ISNUMBER(A1)` when expecting text | TYPE() returns 1 |
| Empty vs blank | `=ISBLANK(A1)` vs `=A1=""` | ISBLANK only for truly empty |
| Duplicates | `=COUNTIF(A:A, A1)>1` | Count > 1 means duplicate |
| Inconsistent case | `=EXACT(A1, UPPER(A1))` | Fails if mixed case |
| Dates as text | `=ISNUMBER(A1)` should be TRUE for dates | DATEVALUE() needed |

---

## 1. Remove Duplicates

### Via Office.js — Programmatic dedup
```typescript
// Read all data
const range = sheet.getUsedRange();
range.load("values");
await context.sync();

const data = range.values;
const headers = data[0];
const seen = new Set<string>();
const uniqueRows = [headers];

for (let i = 1; i < data.length; i++) {
  const key = JSON.stringify(data[i]);
  if (!seen.has(key)) {
    seen.add(key);
    uniqueRows.push(data[i]);
  }
}

// Clear and rewrite
range.clear();
const newRange = sheet.getRange(`A1:${String.fromCharCode(64 + headers.length)}${uniqueRows.length}`);
newRange.values = uniqueRows;
```

### Via Formula — Flag duplicates first
```
Column helper: =IF(COUNTIF($A$2:A2, A2)>1, "DUPLICATE", "")
```
Then filter and delete flagged rows.

### Dedup by specific column only
```typescript
const keyCol = 0; // column A as key
const seen = new Set();
const uniqueRows = [data[0]];
for (let i = 1; i < data.length; i++) {
  const key = String(data[i][keyCol]).trim().toLowerCase();
  if (!seen.has(key)) {
    seen.add(key);
    uniqueRows.push(data[i]);
  }
}
```

---

## 2. Clean Text Data

### Trim whitespace
```typescript
// Formula approach — add helper column
range.formulas = data.map((_, i) => [`=TRIM(CLEAN(A${i + 1}))`]);

// Direct via Office.js — read, clean in JS, write back
const cleaned = data.map(row =>
  row.map(cell => typeof cell === "string" ? cell.trim().replace(/\s+/g, " ") : cell)
);
range.values = cleaned;
```

### Standardize text case
```
=UPPER(A1)      → "JOHN SMITH"
=LOWER(A1)      → "john smith"
=PROPER(A1)     → "John Smith"
```

```typescript
// Proper case via Office.js
const cleaned = data.map(row =>
  row.map(cell =>
    typeof cell === "string"
      ? cell.toLowerCase().replace(/\b\w/g, c => c.toUpperCase())
      : cell
  )
);
```

### Remove specific characters
```
=SUBSTITUTE(A1, "-", "")                    → remove dashes
=SUBSTITUTE(SUBSTITUTE(A1, "(", ""), ")", "") → remove parentheses
=CLEAN(A1)                                   → remove non-printable
```

### Fix common data entry errors
```
=SUBSTITUTE(A1, "  ", " ")         → double spaces to single (may need nesting)
=SUBSTITUTE(A1, CHAR(160), " ")    → non-breaking space to normal space
=TRIM(CLEAN(SUBSTITUTE(A1, CHAR(160), " ")))  → comprehensive text clean
```

---

## 3. Split & Merge Columns

### Split full name into first/last
```
=TEXTBEFORE(A1, " ")              → first name
=TEXTAFTER(A1, " ")               → last name (if only 2 parts)
=LEFT(A1, FIND(" ", A1)-1)        → first name (legacy)
=MID(A1, FIND(" ", A1)+1, 100)    → rest after first space
```

### Split by delimiter (modern)
```
=TEXTSPLIT(A1, ",")               → split by comma into columns
=TEXTSPLIT(A1, ",", CHAR(10))     → split by comma (cols) and newline (rows)
```

### Split address components
```
Street:  =TEXTBEFORE(A1, ",")
City:    =TRIM(TEXTBEFORE(TEXTAFTER(A1, ","), ","))
State:   =TRIM(TEXTAFTER(TEXTAFTER(A1, ","), ","))
```

### Merge columns
```
=TEXTJOIN(" ", TRUE, A1, B1, C1)              → "John Michael Smith" (skip blanks)
=A1 & ", " & B1 & " " & C1                    → "Smith, John Michael"
=TEXTJOIN(CHAR(10), TRUE, A1:A5)              → multiline in one cell
```

### Office.js split implementation
```typescript
const data = range.values;
const splitData = data.map(row => {
  const parts = String(row[0]).split(",").map(s => s.trim());
  return [...parts, ...row.slice(1)];
});

// Write to expanded range
const newRange = sheet.getRange(`A1`).getResizedRange(
  splitData.length - 1, splitData[0].length - 1
);
newRange.values = splitData;
```

---

## 4. Handle Missing Values

### Detect missing values
```
=IF(ISBLANK(A1), "MISSING", A1)
=IF(OR(ISBLANK(A1), A1="", A1="N/A", A1="n/a", A1="-"), "MISSING", A1)
```

### Fill strategies

**Forward fill** (use last known value):
```
=IF(ISBLANK(A2), B1, A2)     → in helper column B, drag down
```

**Fill with average:**
```
=IF(ISBLANK(A1), AVERAGE(A:A), A1)
```

**Fill with zero:**
```typescript
const cleaned = data.map(row =>
  row.map(cell => (cell === null || cell === "" || cell === undefined) ? 0 : cell)
);
range.values = cleaned;
```

**Fill with "N/A" text:**
```typescript
const cleaned = data.map(row =>
  row.map(cell => (cell === null || cell === "" || cell === undefined) ? "N/A" : cell)
);
```

### Remove rows with missing values
```typescript
const headers = data[0];
const filtered = [headers, ...data.slice(1).filter(row =>
  row.every(cell => cell !== null && cell !== "" && cell !== undefined)
)];
```

---

## 5. Fix Data Types

### Numbers stored as text
```
=VALUE(A1)                    → convert text "123" to number 123
=A1*1                         → quick convert (multiply by 1)
=A1+0                         → quick convert (add 0)
```

```typescript
// Via Office.js
const fixed = data.map(row =>
  row.map(cell => {
    if (typeof cell === "string") {
      const num = Number(cell.replace(/[,$\s]/g, ""));
      return isNaN(num) ? cell : num;
    }
    return cell;
  })
);
range.values = fixed;
```

### Text dates to real dates
```
=DATEVALUE(A1)                → if text like "March 29, 2026"
=DATE(LEFT(A1,4), MID(A1,6,2), RIGHT(A1,2))  → if text like "2026-03-29"
```

```typescript
// Parse common date formats in JS
const parseDate = (s: string): number | string => {
  const d = new Date(s);
  if (!isNaN(d.getTime())) {
    // Excel serial date: days since 1900-01-01
    const epoch = new Date(1899, 11, 30);
    return Math.floor((d.getTime() - epoch.getTime()) / 86400000);
  }
  return s;
};
```

### Currency text to numbers
```
=VALUE(SUBSTITUTE(SUBSTITUTE(A1, "$", ""), ",", ""))
```

### Percentage text to numbers
```
=VALUE(SUBSTITUTE(A1, "%", "")) / 100
```

---

## 6. Standardize & Validate

### Standardize categories
```typescript
const mapping: Record<string, string> = {
  "us": "United States", "usa": "United States", "u.s.": "United States", "united states": "United States",
  "uk": "United Kingdom", "u.k.": "United Kingdom", "united kingdom": "United Kingdom",
  "fr": "France", "france": "France",
};

const standardized = data.map(row =>
  row.map((cell, col) => {
    if (col === targetCol && typeof cell === "string") {
      return mapping[cell.toLowerCase().trim()] || cell;
    }
    return cell;
  })
);
```

### Data validation via Office.js
```typescript
// Only allow numbers between 0 and 100
const validation = range.dataValidation;
validation.rule = {
  wholeNumber: {
    formula1: 0,
    formula2: 100,
    operator: Excel.DataValidationOperator.between,
  },
};
validation.errorAlert = {
  showAlert: true,
  title: "Invalid Entry",
  message: "Please enter a number between 0 and 100.",
  style: Excel.DataValidationAlertStyle.stop,
};

// Dropdown list validation
const listValidation = range.dataValidation;
listValidation.rule = {
  list: {
    inCellDropDown: true,
    source: "High,Medium,Low",
  },
};
```

### Email validation formula
```
=AND(ISNUMBER(FIND("@",A1)), ISNUMBER(FIND(".",A1,FIND("@",A1))), LEN(A1)>5)
```

### Phone number standardization
```
=SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(A1," ",""),"-",""),"(",""),")","")
→ strips all formatting from phone numbers
```

---

## 7. Restructure Malformed Tables

### Headers in wrong row
```typescript
// Skip first N junk rows, use row N+1 as headers
const headerRow = 3; // 0-indexed
const headers = data[headerRow];
const cleanData = [headers, ...data.slice(headerRow + 1).filter(row =>
  row.some(cell => cell !== null && cell !== "")
)];
```

### Unpivot / Melt (wide to long)
Transform:
```
| Product | Jan | Feb | Mar |
| Apple   | 100 | 150 | 200 |
```
Into:
```
| Product | Month | Sales |
| Apple   | Jan   | 100   |
| Apple   | Feb   | 150   |
| Apple   | Mar   | 200   |
```

```typescript
const headers = data[0];
const months = headers.slice(1);
const longData = [["Product", "Month", "Sales"]];

for (let i = 1; i < data.length; i++) {
  const product = data[i][0];
  for (let j = 0; j < months.length; j++) {
    longData.push([product, months[j], data[i][j + 1]]);
  }
}
```

### Pivot (long to wide)
Transform long format back to wide:
```typescript
const categories = [...new Set(data.slice(1).map(r => r[0]))];
const periods = [...new Set(data.slice(1).map(r => r[1]))];
const pivotHeaders = ["Product", ...periods];
const pivotData = [pivotHeaders];

for (const cat of categories) {
  const row = [cat];
  for (const period of periods) {
    const match = data.find(r => r[0] === cat && r[1] === period);
    row.push(match ? match[2] : 0);
  }
  pivotData.push(row);
}
```

---

## 8. Cleaning Pipeline Template

Standard order for a complete data cleaning workflow:

```
Step 1: READ     → Load data, identify shape (rows, cols, headers)
Step 2: DIAGNOSE → Count blanks, duplicates, type mismatches per column
Step 3: HEADERS  → Fix or create proper headers
Step 4: TRIM     → Remove whitespace, non-printable characters
Step 5: TYPES    → Convert text-numbers, text-dates to proper types
Step 6: MISSING  → Handle blanks (fill, remove, or flag)
Step 7: DEDUP    → Remove duplicate rows
Step 8: STANDARDIZE → Normalize categories, case, codes
Step 9: VALIDATE → Apply data validation rules
Step 10: FORMAT  → Apply number formats, column widths, header styling
```

### Report summary after cleaning
Always tell the user what was changed:
```
Cleaned dataset:
- Rows: 1,000 → 947 (53 duplicates removed)
- Blanks filled: 23 cells (forward-fill in column C)
- Type fixes: 156 text-to-number conversions in column D
- Trimmed: 89 cells with extra whitespace
- Standardized: 34 country names normalized
```
