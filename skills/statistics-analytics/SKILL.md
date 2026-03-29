---
name: statistics-analytics
description: "Statistical analysis and data analytics in Excel via Office.js. Use when the user needs descriptive statistics, correlation, regression, forecasting, hypothesis testing, distributions, or data interpretation. Provides both formula-based and Office.js approaches with guidance on interpreting results."
compatibility: HyperFix runtime. Works with all models. No external dependencies.
---

# Statistics & Analytics

Statistical analysis workflows for Excel via Office.js. Covers descriptive stats, inferential analysis, forecasting, and result interpretation.

---

## 1. Descriptive Statistics — Quick Summary

Build a complete stats summary for any numeric column:

### Formula Block
```
Mean:           =AVERAGE(data)
Median:         =MEDIAN(data)
Mode:           =MODE.SNGL(data)
Std Dev:        =STDEV.S(data)         (sample)
Variance:       =VAR.S(data)           (sample)
Min:            =MIN(data)
Max:            =MAX(data)
Range:          =MAX(data)-MIN(data)
Count:          =COUNT(data)
Q1:             =QUARTILE.INC(data, 1)
Q3:             =QUARTILE.INC(data, 3)
IQR:            =QUARTILE.INC(data, 3) - QUARTILE.INC(data, 1)
Skewness:       =SKEW(data)
Kurtosis:       =KURT(data)
Coeff of Var:   =STDEV.S(data)/AVERAGE(data)
Std Error:      =STDEV.S(data)/SQRT(COUNT(data))
```

### Office.js Implementation — Auto Stats Panel
```typescript
async function buildStatsPanel(context: Excel.RequestContext, dataAddress: string, outputAddress: string) {
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  const output = sheet.getRange(outputAddress);

  const labels = [
    ["Statistic", "Value"],
    ["Count", `=COUNT(${dataAddress})`],
    ["Mean", `=AVERAGE(${dataAddress})`],
    ["Median", `=MEDIAN(${dataAddress})`],
    ["Std Deviation", `=STDEV.S(${dataAddress})`],
    ["Variance", `=VAR.S(${dataAddress})`],
    ["Min", `=MIN(${dataAddress})`],
    ["Max", `=MAX(${dataAddress})`],
    ["Range", `=MAX(${dataAddress})-MIN(${dataAddress})`],
    ["Q1", `=QUARTILE.INC(${dataAddress},1)`],
    ["Q3", `=QUARTILE.INC(${dataAddress},3)`],
    ["IQR", `=QUARTILE.INC(${dataAddress},3)-QUARTILE.INC(${dataAddress},1)`],
    ["Skewness", `=SKEW(${dataAddress})`],
    ["Kurtosis", `=KURT(${dataAddress})`],
  ];

  const range = sheet.getRange(outputAddress).getResizedRange(labels.length - 1, 1);
  range.formulas = labels;
  range.getRow(0).format.font.bold = true;
  range.getColumn(1).numberFormat = Array.from({ length: labels.length }, () => ["#,##0.00"]);
  range.format.autofitColumns();

  await context.sync();
}
```

### Interpreting Descriptive Stats
| Metric | Interpretation |
|---|---|
| Mean vs Median | If mean >> median → right skew (outliers pulling up). Use median for "typical" value. |
| Std Dev | Low relative to mean → data is clustered. High → spread out. |
| Skewness | ~0 = symmetric, >0 = right tail, <0 = left tail. \|skew\| > 1 = highly skewed |
| Kurtosis | ~3 = normal (Excel returns excess, so ~0). >0 = heavy tails, <0 = light tails |
| IQR | Robust measure of spread. Outlier = below Q1-1.5×IQR or above Q3+1.5×IQR |
| CV (Coeff of Var) | <0.1 = low variability, 0.1-0.3 = moderate, >0.3 = high |

---

## 2. Correlation Analysis

### Pairwise Correlation
```
=CORREL(range1, range2)          → Pearson correlation coefficient
```

**Interpretation:**
| r value | Strength | Meaning |
|---|---|---|
| 0.9 to 1.0 | Very strong positive | Variables move together strongly |
| 0.7 to 0.9 | Strong positive | Clear positive relationship |
| 0.4 to 0.7 | Moderate positive | Noticeable positive trend |
| 0.0 to 0.4 | Weak/none | Little to no linear relationship |
| Negative | Same scale, inverse | Variables move in opposite directions |

### Correlation Matrix via Office.js
```typescript
async function buildCorrelationMatrix(context: Excel.RequestContext, dataRange: string, headers: string[], outputCell: string) {
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  const n = headers.length;

  // Headers
  const output = sheet.getRange(outputCell);
  const cornerCell = output;
  cornerCell.values = [["Correlation"]];
  cornerCell.format.font.bold = true;

  for (let i = 0; i < n; i++) {
    // Column headers
    output.getOffsetRange(0, i + 1).values = [[headers[i]]];
    output.getOffsetRange(0, i + 1).format.font.bold = true;
    // Row headers
    output.getOffsetRange(i + 1, 0).values = [[headers[i]]];
    output.getOffsetRange(i + 1, 0).format.font.bold = true;
  }

  // Correlation formulas
  for (let i = 0; i < n; i++) {
    for (let j = 0; j < n; j++) {
      const colI = String.fromCharCode(65 + i); // assumes data starts at A
      const colJ = String.fromCharCode(65 + j);
      output.getOffsetRange(i + 1, j + 1).formulas =
        [[`=CORREL(${colI}2:${colI}1000,${colJ}2:${colJ}1000)`]];
      output.getOffsetRange(i + 1, j + 1).numberFormat = [["0.00"]];
    }
  }

  await context.sync();
}
```

### Conditional formatting for correlation matrix
```typescript
// Red for negative, white for zero, green for positive
const matrixRange = sheet.getRange("B2:E5"); // adjust size
const cfNeg = matrixRange.conditionalFormats.add(Excel.ConditionalFormatType.colorScale);
cfNeg.colorScale.criteria = {
  minimum: { color: "#FF0000", type: Excel.ConditionalFormatColorCriterionType.lowestValue },
  midpoint: { color: "#FFFFFF", type: Excel.ConditionalFormatColorCriterionType.number, value: "0" },
  maximum: { color: "#00B050", type: Excel.ConditionalFormatColorCriterionType.highestValue },
};
```

---

## 3. Regression Analysis

### Simple Linear Regression via Formulas
```
Slope:          =SLOPE(known_ys, known_xs)
Intercept:      =INTERCEPT(known_ys, known_xs)
R-squared:      =RSQ(known_ys, known_xs)
Std Error:      =STEYX(known_ys, known_xs)
Equation:       y = SLOPE * x + INTERCEPT
Prediction:     =FORECAST.LINEAR(new_x, known_ys, known_xs)
```

### Full Regression Output with LINEST
```
=LINEST(known_ys, known_xs, TRUE, TRUE)
```
Returns a 5×2 array (select 5 rows × 2 cols, enter as array formula):
```
Row 1: slope, intercept
Row 2: std error of slope, std error of intercept
Row 3: R², std error of y
Row 4: F-statistic, degrees of freedom
Row 5: regression SS, residual SS
```

### Interpreting Regression
| Metric | Good Value | Meaning |
|---|---|---|
| R² | > 0.7 for social science, > 0.9 for physical | % of variance explained |
| p-value (from F-stat) | < 0.05 | Model is statistically significant |
| Std Error | Low relative to predictions | Predictions are precise |
| Slope | Depends on context | Change in Y per unit change in X |

### Add Trendline to Data
```typescript
// Create scatter data then add regression line
const xs = data.slice(1).map(r => r[0]);
const ys = data.slice(1).map(r => r[1]);

// Regression line endpoints
const minX = Math.min(...xs);
const maxX = Math.max(...xs);

// Use FORECAST to compute line points
const trendSheet = sheet.getRange("E1:F3");
trendSheet.values = [
  ["Trend X", "Trend Y"],
  [minX, null],
  [maxX, null],
];
trendSheet.getCell(1, 1).formulas = [[`=FORECAST.LINEAR(E2,B2:B${data.length},A2:A${data.length})`]];
trendSheet.getCell(2, 1).formulas = [[`=FORECAST.LINEAR(E3,B2:B${data.length},A2:A${data.length})`]];
```

---

## 4. Forecasting

### FORECAST.ETS (Exponential Smoothing — best for time series)
```
=FORECAST.ETS(target_date, values, timeline, [seasonality], [data_completion], [aggregation])
```
- **seasonality:** 0 = auto-detect, 1 = none, or specific number (12 for monthly)
- Handles trends AND seasonality automatically

### FORECAST.ETS.SEASONALITY
```
=FORECAST.ETS.SEASONALITY(values, timeline)   → detected period length
```

### FORECAST.ETS.CONFINT
```
=FORECAST.ETS.CONFINT(target_date, values, timeline, [confidence_level])
```
- Default 95% confidence. Returns the interval width.
- Upper bound = forecast + confint, Lower bound = forecast - confint

### TREND (linear projection)
```
=TREND(known_ys, known_xs, new_xs)
```
- Array formula: projects Y values for new X values using linear fit

### GROWTH (exponential projection)
```
=GROWTH(known_ys, known_xs, new_xs)
```
- Like TREND but fits exponential curve: y = b × m^x

### Moving Averages
```
3-period MA:   =AVERAGE(A1:A3)     → drag down
7-period MA:   =AVERAGE(A1:A7)     → drag down
Weighted MA:   =SUMPRODUCT(A1:A3, {0.5,0.3,0.2})
```

---

## 5. Frequency & Distribution Analysis

### FREQUENCY
```
=FREQUENCY(data_array, bins_array)
```
- Returns an array of counts for each bin. One more result than bins.
- Example bins: {10, 20, 30, 40, 50} → counts for ≤10, 11-20, 21-30, 31-40, 41-50, >50

### Histogram via Office.js
```typescript
// Create frequency bins
const bins = [0, 10, 20, 30, 40, 50, 60, 70, 80, 90, 100];
const binRange = sheet.getRange(`G2:G${bins.length + 1}`);
binRange.values = bins.map(b => [b]);

// FREQUENCY formula
const freqRange = sheet.getRange(`H2:H${bins.length + 2}`);
freqRange.formulas = [[`=FREQUENCY(A2:A1000,G2:G${bins.length + 1})`]];

// Create chart from frequency data
const chart = sheet.charts.add(Excel.ChartType.columnClustered, freqRange);
chart.title.text = "Distribution";
```

### Normal Distribution Check
```
Skewness:       =SKEW(data)                → should be near 0
Kurtosis:       =KURT(data)                → should be near 0 (excess)
```
If |skew| < 0.5 and |kurt| < 1 → approximately normal.

---

## 6. Hypothesis Testing Formulas

### Z-Test (large samples, known σ)
```
Z-score:        =(sample_mean - population_mean) / (sigma / SQRT(n))
p-value (2-tail): =2*(1-NORM.S.DIST(ABS(z_score), TRUE))
```

### T-Test (small samples, unknown σ)
```
T-score:        =(sample_mean - hypothesized_mean) / (sample_stdev / SQRT(n))
p-value (2-tail): =T.DIST.2T(ABS(t_score), n-1)
```

### Two-Sample T-Test
```
=T.TEST(array1, array2, tails, type)
```
- **tails:** 1 = one-tail, 2 = two-tail
- **type:** 1 = paired, 2 = equal variance, 3 = unequal variance (Welch's)

### Chi-Square Test
```
=CHISQ.TEST(actual_range, expected_range)    → p-value
=CHISQ.INV.RT(alpha, degrees_freedom)        → critical value
```

### Interpreting p-values
| p-value | Conclusion |
|---|---|
| < 0.01 | Very strong evidence against null hypothesis |
| 0.01 - 0.05 | Strong evidence (standard threshold) |
| 0.05 - 0.10 | Weak evidence (sometimes accepted) |
| > 0.10 | Insufficient evidence — fail to reject null |

---

## 7. Comparative Analysis Formulas

### Year-over-Year Growth
```
=IF(prior=0, "", (current-prior)/ABS(prior))
```

### CAGR (Compound Annual Growth Rate)
```
=(end_value/start_value)^(1/years) - 1
```

### Weighted Average
```
=SUMPRODUCT(values, weights) / SUM(weights)
```

### Pareto Analysis (80/20)
```
Step 1: Sort data descending by value
Step 2: Cumulative sum: =SUM($B$2:B2)
Step 3: Cumulative %: =cumulative_sum / SUM($B$2:$B$100)
Step 4: Flag where cumulative % crosses 80%
```

### Index/Benchmark Comparison
```
Indexed value:  =current_value / base_value * 100
Variance:       =actual - budget
Variance %:     =IF(budget=0, "", (actual-budget)/ABS(budget))
```

---

## 8. Best Practices

### Always tell the user:
1. **What test/method** you used and why
2. **Sample size** — is it large enough? (n > 30 for z-tests, n > 5 per group for t-tests)
3. **Assumptions** — normality, equal variance, independence
4. **Confidence level** — default 95% unless specified
5. **Practical significance** vs statistical significance — a tiny difference can be "significant" with huge n

### When to use what:
| Goal | Method | Key Formula |
|---|---|---|
| Summarize data | Descriptive stats | AVERAGE, MEDIAN, STDEV |
| Find relationships | Correlation | CORREL |
| Predict from relationship | Regression | LINEST, FORECAST |
| Predict time series | ETS Forecasting | FORECAST.ETS |
| Compare two groups | T-test | T.TEST |
| Compare proportions | Chi-square | CHISQ.TEST |
| Rank items | Ranking | RANK, PERCENTILE |
| Find patterns | Frequency analysis | FREQUENCY |
