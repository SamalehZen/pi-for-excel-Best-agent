---
name: financial-modeling
description: "Financial modeling standards and best practices for Excel. Use when building DCF models, LBO, amortization schedules, budgets, P&L, balance sheets, or any financial analysis. Includes color-coding conventions, number formatting, formula construction rules, and Office.js implementation."
compatibility: HyperFix runtime. Works with all models. No external dependencies.
metadata:
  adapted-from: anthropics/skills/xlsx (Anthropic official skill, adapted for Office.js)
---

# Financial Modeling Standards for HyperFix

Professional financial modeling conventions adapted for Office.js execution. Based on industry standards (investment banking, FP&A, audit).

---

## Core Rules

1. **Always use formulas** — never hardcode calculated values
2. **One assumption per cell** — no magic numbers buried in formulas
3. **Consistent structure** — time flows left-to-right, line items top-to-bottom
4. **Document sources** — every hardcoded input must cite its origin
5. **Zero formula errors** — #REF!, #DIV/0!, #VALUE!, #N/A, #NAME? are unacceptable

---

## Color-Coding Standards

Industry-standard color conventions for financial models. Apply via Office.js `font.color` and `fill.color`.

| Element | Font Color | RGB | Use |
|---|---|---|---|
| **Hardcoded inputs** | Blue | `#0000FF` | Numbers the user types/changes (assumptions, scenarios) |
| **Formulas** | Black | `#000000` | ALL calculated cells |
| **Cross-sheet links** | Green | `#008000` | Formulas pulling from other sheets in same workbook |
| **External links** | Red | `#FF0000` | Links to other files |
| **Key assumptions** | Yellow background | `#FFFF00` fill | Cells needing attention or update |
| **Check cells** | — | — | Cells that should equal zero if model balances |

### Office.js Implementation
```typescript
// Blue input cell
const inputRange = sheet.getRange("B5");
inputRange.format.font.color = "#0000FF";

// Formula cell — keep black
const formulaRange = sheet.getRange("C5");
formulaRange.formulas = [["=B5*(1+$B$3)"]];
formulaRange.format.font.color = "#000000";

// Key assumption with yellow background
const assumptionRange = sheet.getRange("B3");
assumptionRange.format.fill.color = "#FFFF00";
assumptionRange.format.font.color = "#0000FF";

// Cross-sheet link in green
const linkRange = sheet.getRange("D5");
linkRange.formulas = [["=Revenue!C5"]];
linkRange.format.font.color = "#008000";
```

---

## Number Formatting Standards

### Required Formats
| Data Type | Format Code | Example | Notes |
|---|---|---|---|
| **Currency** | `$#,##0` or `$#,##0.00` | $1,234 | Always specify units in headers: "Revenue ($mm)" |
| **Currency (neg)** | `$#,##0;($#,##0);"-"` | ($500) | Parentheses for negatives, dash for zero |
| **Percentage** | `0.0%` | 15.5% | One decimal by default |
| **Multiples** | `0.0"x"` | 8.5x | For EV/EBITDA, P/E ratios |
| **Years** | `0` or `@` | 2026 | Format as text or integer — never `#,##0` (shows 2,026) |
| **Dates** | `yyyy-mm-dd` | 2026-03-29 | ISO format for consistency |
| **Count** | `#,##0` | 1,234 | Thousands separator |
| **Basis points** | `0 "bps"` | 150 bps | For interest rate changes |
| **Negative numbers** | `#,##0;(#,##0);"-"` | (500) | Parentheses, never minus sign |

### Office.js Implementation
```typescript
// Apply financial number format
const revenueRange = sheet.getRange("B5:F5");
revenueRange.numberFormat = [["$#,##0;($#,##0);\"-\"", "$#,##0;($#,##0);\"-\"", "$#,##0;($#,##0);\"-\"", "$#,##0;($#,##0);\"-\"", "$#,##0;($#,##0);\"-\""]];

// Percentage format
const marginRange = sheet.getRange("B8:F8");
marginRange.numberFormat = [Array(5).fill("0.0%")];

// Multiple format
const multipleRange = sheet.getRange("B20:F20");
multipleRange.numberFormat = [Array(5).fill("0.0\"x\"")];

// Year format (avoid comma)
const yearRange = sheet.getRange("B1:F1");
yearRange.numberFormat = [Array(5).fill("0")];
```

---

## Formula Construction Rules

### 1. Separate Assumptions
Place ALL assumptions in dedicated cells. Never embed numbers in formulas.

```
❌ WRONG: =B5*1.05          → what is 1.05?
✅ RIGHT: =B5*(1+$B$3)      → $B$3 is labeled "Growth Rate" and colored blue
```

### 2. One Operation Per Row When Possible
Break complex calculations into intermediate rows for auditability.

```
Row 5: Revenue          =Prior_Revenue * (1 + Growth_Rate)
Row 6: COGS             =Revenue * COGS_Margin
Row 7: Gross Profit     =Revenue - COGS
Row 8: Gross Margin %   =Gross_Profit / Revenue
```

### 3. Consistent Formulas Across Periods
Every projection column should use the same formula structure. Never manually override one period.

### 4. Error Protection
```
=IF(Revenue=0, 0, EBITDA/Revenue)           → prevent #DIV/0!
=IFERROR(INDEX(MATCH(...)), "Missing")      → prevent #N/A
```

### 5. Source Documentation for Hardcoded Values
Every hardcoded number must have a comment or adjacent cell with:
```
Source: [Document], [Date], [Reference], [URL if applicable]

Examples:
- "Source: Company 10-K, FY2024, Page 45, Revenue Note"
- "Source: Bloomberg Terminal, 2026-03-15, AAPL US Equity"
- "Source: Management guidance, Q4 2025 earnings call"
- "Source: Industry average, Damodaran dataset, Jan 2026"
```

---

## Common Financial Models

### 1. Revenue Build
```
Headers:        | FY2024A | FY2025E | FY2026E | FY2027E |
Units Sold      |  1,000  | =prior*(1+growth) | ... | ... |
Price per Unit  |  $50    | =prior*(1+price_inflation) | ... |
Revenue         | =units*price | ... | ... | ... |
YoY Growth %    | — | =(current-prior)/prior | ... | ... |
```

### 2. Three-Statement Model Structure
```
Sheet 1: Assumptions    → all blue inputs
Sheet 2: Income Statement → formulas in black
Sheet 3: Balance Sheet   → formulas with green cross-sheet links
Sheet 4: Cash Flow       → formulas with green cross-sheet links
Sheet 5: DCF Valuation   → formulas referencing all above
Sheet 6: Sensitivity     → DATA TABLE scenarios
```

### 3. DCF Template Formulas
```
Free Cash Flow:    =EBIT*(1-Tax_Rate) + D&A - CapEx - Change_in_WC
Terminal Value:    =FCF_final * (1+g) / (WACC-g)
Enterprise Value:  =NPV(WACC, FCF_range) + TV/(1+WACC)^n
Equity Value:      =EV - Net_Debt
Per Share:         =Equity_Value / Diluted_Shares
```

### 4. Amortization Schedule
```
Period | Begin Balance | Payment | Interest | Principal | End Balance
  1    | =Loan_Amount | =PMT(rate,nper,-PV) | =Begin*rate | =Payment-Interest | =Begin-Principal
  2    | =Prior_End   | =same   | =Begin*rate | =Payment-Interest | =Begin-Principal
```

### 5. Sensitivity Table (DATA TABLE)
Use Excel's built-in Data Table feature for 1-way and 2-way sensitivity analysis:
```
=Target_cell in top-left corner
Row input: WACC values across top
Column input: Growth rate values down left side
→ Select entire table → Data → What-If → Data Table
```

---

## Office.js Model Building Patterns

### Create a full worksheet with headers + formulas + formatting
```typescript
async function buildIncomeStatement(context: Excel.RequestContext) {
  const sheet = context.workbook.worksheets.add("Income Statement");

  // Headers
  const headers = sheet.getRange("A1:F1");
  headers.values = [["", "FY2024A", "FY2025E", "FY2026E", "FY2027E", "FY2028E"]];
  headers.format.font.bold = true;
  headers.format.fill.color = "#D9E1F2";
  headers.numberFormat = [Array(6).fill("@")]; // text format for years

  // Line items
  const labels = ["Revenue", "COGS", "Gross Profit", "Gross Margin %",
                   "OpEx", "EBITDA", "EBITDA Margin %",
                   "D&A", "EBIT", "Interest", "EBT", "Tax", "Net Income"];
  const labelRange = sheet.getRange(`A2:A${labels.length + 1}`);
  labelRange.values = labels.map(l => [l]);

  // Historical data (blue — hardcoded inputs)
  const histData = sheet.getRange("B2:B3");
  histData.values = [[10000], [4000]];
  histData.format.font.color = "#0000FF";
  histData.numberFormat = [["$#,##0"], ["$#,##0"]];

  // Projection formulas (black)
  const projRevenue = sheet.getRange("C2:F2");
  projRevenue.formulas = [["=B2*(1+Assumptions!$B$2)", "=C2*(1+Assumptions!$B$2)", "=D2*(1+Assumptions!$B$2)", "=E2*(1+Assumptions!$B$2)"]];
  projRevenue.format.font.color = "#000000";
  projRevenue.numberFormat = [Array(4).fill("$#,##0")];

  await context.sync();
}
```

### Batch-apply color coding
```typescript
// Color all input cells blue in a range
const inputs = sheet.getRange("B5:B20");
inputs.format.font.color = "#0000FF";

// Color all formula cells black
const formulas = sheet.getRange("C5:F20");
formulas.format.font.color = "#000000";

// Highlight check row
const checkRow = sheet.getRange("B25:F25");
checkRow.format.fill.color = "#E2EFDA";
checkRow.numberFormat = [Array(5).fill("0.00")];
```

---

## Validation Checklist

Before delivering any financial model:

- [ ] All inputs are blue, all formulas are black, cross-sheet links are green
- [ ] All numbers have appropriate formatting (currency, %, multiples)
- [ ] Years display as 2024, not 2,024
- [ ] Negative numbers use parentheses, not minus signs
- [ ] Zero values display as "-" (dash)
- [ ] No hardcoded numbers inside formulas — all assumptions in separate cells
- [ ] Balance sheet balances (Assets = Liabilities + Equity) — add check row
- [ ] Cash flow reconciles to balance sheet cash change — add check row
- [ ] All check cells equal zero (or show "OK"/"ERROR")
- [ ] Sources documented for every hardcoded assumption
- [ ] No formula errors (#REF!, #DIV/0!, #VALUE!, #N/A, #NAME?)
- [ ] Formulas are consistent across all projection periods
