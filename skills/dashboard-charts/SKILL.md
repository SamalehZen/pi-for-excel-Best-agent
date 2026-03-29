---
name: dashboard-charts
description: "Dashboard creation, charts, conditional formatting, and data visualization in Excel via Office.js. Use when the user wants to create charts, format tables, build dashboards, add conditional formatting, sparklines, or any visual presentation of data."
compatibility: HyperFix runtime. Works with all models. No external dependencies.
---

# Dashboards & Charts

Creating professional data visualizations and dashboards via Office.js in HyperFix.

---

## 1. Charts via Office.js

### Create a Chart
```typescript
const sheet = context.workbook.worksheets.getActiveWorksheet();
const dataRange = sheet.getRange("A1:D10");

const chart = sheet.charts.add(
  Excel.ChartType.columnClustered,  // chart type
  dataRange,                         // data source
  Excel.ChartSeriesBy.columns       // series by columns or rows
);

chart.title.text = "Monthly Revenue";
chart.legend.position = Excel.ChartLegendPosition.bottom;
chart.setPosition("F1", "N20"); // top-left cell, bottom-right cell

await context.sync();
```

### Chart Types Available
| Type | Excel.ChartType | Best For |
|---|---|---|
| Column (vertical bars) | `columnClustered` | Comparing categories |
| Stacked Column | `columnStacked` | Part-to-whole over categories |
| 100% Stacked Column | `columnStacked100` | Proportion comparison |
| Bar (horizontal) | `barClustered` | Long category labels |
| Line | `line` | Trends over time |
| Line with Markers | `lineMarkers` | Trends with data points visible |
| Area | `area` | Cumulative totals over time |
| Stacked Area | `areaStacked` | Part-to-whole over time |
| Pie | `pie` | Simple part-to-whole (≤6 slices) |
| Doughnut | `doughnut` | Part-to-whole with center label |
| Scatter (XY) | `xyscatter` | Relationship between 2 variables |
| Scatter with Lines | `xyscatterLines` | Connected scatter points |
| Combo | `columnClustered` + line series | Dual-axis (revenue + margin %) |
| Waterfall | `waterfall` | Bridge charts (P&L walk) |
| Funnel | `funnel` | Pipeline/conversion stages |
| Treemap | `treemap` | Hierarchical proportions |
| Sunburst | `sunburst` | Multi-level hierarchy |

### Customize Chart Appearance
```typescript
// Title
chart.title.text = "Q1 2026 Performance";
chart.title.format.font.size = 14;
chart.title.format.font.bold = true;
chart.title.format.font.color = "#333333";

// Axes
const valueAxis = chart.axes.valueAxis;
valueAxis.title.text = "Revenue ($)";
valueAxis.minimum = 0;
valueAxis.numberFormat = "$#,##0";
valueAxis.majorGridlines.visible = true;
valueAxis.majorGridlines.format.line.color = "#E0E0E0";

const categoryAxis = chart.axes.categoryAxis;
categoryAxis.title.text = "Month";

// Legend
chart.legend.visible = true;
chart.legend.position = Excel.ChartLegendPosition.bottom;

// Chart area
chart.format.fill.setSolidColor("#FFFFFF");
chart.plotArea.format.fill.setSolidColor("#FAFAFA");

// Size and position
chart.height = 300;
chart.width = 500;
```

### Series Customization
```typescript
const series = chart.series.getItemAt(0);
series.name = "Revenue";
series.format.fill.setSolidColor("#4472C4");

const series2 = chart.series.getItemAt(1);
series2.name = "Target";
series2.format.fill.setSolidColor("#ED7D31");
series2.format.line.color = "#ED7D31";

// Data labels
series.dataLabels.showValue = true;
series.dataLabels.numberFormat = "$#,##0";
series.dataLabels.format.font.size = 9;
```

### Combo Chart (Column + Line)
```typescript
const chart = sheet.charts.add(Excel.ChartType.columnClustered, dataRange);

// Change second series to line
const lineSeries = chart.series.getItemAt(1);
lineSeries.chartType = Excel.ChartType.line;
lineSeries.axisGroup = Excel.ChartAxisGroup.secondary; // secondary Y axis

// Format secondary axis
const secondaryAxis = chart.axes.getItem(Excel.ChartAxisType.value, Excel.ChartAxisGroup.secondary);
secondaryAxis.title.text = "Margin %";
secondaryAxis.numberFormat = "0%";
```

---

## 2. Conditional Formatting

### Color Scales (Heatmap)
```typescript
const range = sheet.getRange("B2:M13");
const cf = range.conditionalFormats.add(Excel.ConditionalFormatType.colorScale);
cf.colorScale.criteria = {
  minimum: { color: "#F8696B", type: Excel.ConditionalFormatColorCriterionType.lowestValue },
  midpoint: { color: "#FFEB84", type: Excel.ConditionalFormatColorCriterionType.percentile, value: "50" },
  maximum: { color: "#63BE7B", type: Excel.ConditionalFormatColorCriterionType.highestValue },
};
```

### Data Bars
```typescript
const range = sheet.getRange("C2:C20");
const cf = range.conditionalFormats.add(Excel.ConditionalFormatType.dataBar);
cf.dataBar.barDirection = Excel.ConditionalDataBarDirection.leftToRight;
cf.dataBar.positiveFormat.fillColor = "#4472C4";
cf.dataBar.negativeFormat.fillColor = "#FF0000";
cf.dataBar.showDataBarOnly = false; // show numbers too
```

### Icon Sets
```typescript
const range = sheet.getRange("D2:D20");
const cf = range.conditionalFormats.add(Excel.ConditionalFormatType.iconSet);
cf.iconSet.style = Excel.IconSet.threeTrafficLights1;
cf.iconSet.criteria = [
  { type: Excel.ConditionalFormatIconRuleType.number, value: 0, operator: Excel.ConditionalIconCriterionOperator.greaterThanOrEqual },
  { type: Excel.ConditionalFormatIconRuleType.number, value: 50, operator: Excel.ConditionalIconCriterionOperator.greaterThanOrEqual },
  { type: Excel.ConditionalFormatIconRuleType.number, value: 80, operator: Excel.ConditionalIconCriterionOperator.greaterThanOrEqual },
];
```

### Cell Value Rules
```typescript
// Highlight cells > 100 in green
const cf = range.conditionalFormats.add(Excel.ConditionalFormatType.cellValue);
cf.cellValue.format.fill.color = "#C6EFCE";
cf.cellValue.format.font.color = "#006100";
cf.cellValue.rule = {
  formula1: "100",
  operator: Excel.ConditionalCellValueOperator.greaterThan,
};

// Highlight negatives in red
const cfNeg = range.conditionalFormats.add(Excel.ConditionalFormatType.cellValue);
cfNeg.cellValue.format.fill.color = "#FFC7CE";
cfNeg.cellValue.format.font.color = "#9C0006";
cfNeg.cellValue.rule = {
  formula1: "0",
  operator: Excel.ConditionalCellValueOperator.lessThan,
};
```

### Top/Bottom Rules
```typescript
// Top 10 values
const cf = range.conditionalFormats.add(Excel.ConditionalFormatType.topBottom);
cf.topBottom.format.fill.color = "#C6EFCE";
cf.topBottom.rule = {
  rank: 10,
  type: Excel.ConditionalTopBottomCriterionType.topItems,
};
```

### Text Contains
```typescript
const cf = range.conditionalFormats.add(Excel.ConditionalFormatType.containsText);
cf.textComparison.format.fill.color = "#FFEB9C";
cf.textComparison.format.font.color = "#9C6500";
cf.textComparison.rule = {
  operator: Excel.ConditionalTextOperator.contains,
  text: "Overdue",
};
```

---

## 3. Table Formatting

### Create an Excel Table
```typescript
const range = sheet.getRange("A1:E20");
const table = sheet.tables.add(range, true); // true = has headers
table.name = "SalesData";
table.style = "TableStyleMedium2"; // built-in style

// Auto-filter is enabled by default on tables
table.showFilterButton = true;
table.showTotals = true;

// Total row functions
table.columns.getItemAt(1).totalsCalculation = Excel.TotalsCalculation.sum;
table.columns.getItemAt(2).totalsCalculation = Excel.TotalsCalculation.average;
table.columns.getItemAt(3).totalsCalculation = Excel.TotalsCalculation.count;
```

### Available Table Styles
```
TableStyleLight1-21      → light themes
TableStyleMedium1-28     → medium themes (most professional)
TableStyleDark1-11       → dark themes
```

### Format Headers
```typescript
const headerRow = sheet.getRange("A1:F1");
headerRow.format.font.bold = true;
headerRow.format.font.color = "#FFFFFF";
headerRow.format.fill.color = "#4472C4";
headerRow.format.horizontalAlignment = Excel.HorizontalAlignment.center;
headerRow.format.rowHeight = 30;
headerRow.format.font.size = 11;
```

### Alternating Row Colors (banded rows without table)
```typescript
const dataRows = sheet.getRange("A2:F100");
for (let i = 0; i < 99; i++) {
  if (i % 2 === 0) {
    const row = sheet.getRange(`A${i + 2}:F${i + 2}`);
    row.format.fill.color = "#D9E2F3";
  }
}
```

### Column Width & Row Height
```typescript
// Auto-fit columns
sheet.getRange("A:F").format.autofitColumns();

// Manual column width
sheet.getRange("A:A").format.columnWidth = 120;
sheet.getRange("B:E").format.columnWidth = 80;

// Row height
sheet.getRange("1:1").format.rowHeight = 35; // header row
```

### Borders
```typescript
const range = sheet.getRange("A1:F20");
range.format.borders.getItem(Excel.BorderIndex.edgeTop).style = Excel.BorderLineStyle.thin;
range.format.borders.getItem(Excel.BorderIndex.edgeBottom).style = Excel.BorderLineStyle.thin;
range.format.borders.getItem(Excel.BorderIndex.edgeLeft).style = Excel.BorderLineStyle.thin;
range.format.borders.getItem(Excel.BorderIndex.edgeRight).style = Excel.BorderLineStyle.thin;
range.format.borders.getItem(Excel.BorderIndex.insideHorizontal).style = Excel.BorderLineStyle.thin;
range.format.borders.getItem(Excel.BorderIndex.insideVertical).style = Excel.BorderLineStyle.thin;

// All borders color
const borderItems = [
  Excel.BorderIndex.edgeTop, Excel.BorderIndex.edgeBottom,
  Excel.BorderIndex.edgeLeft, Excel.BorderIndex.edgeRight,
  Excel.BorderIndex.insideHorizontal, Excel.BorderIndex.insideVertical,
];
for (const item of borderItems) {
  range.format.borders.getItem(item).color = "#D0D0D0";
}
```

### Freeze Panes
```typescript
// Freeze top row and first column
sheet.freezePanes.freezeRows(1);
sheet.freezePanes.freezeColumns(1);

// Or freeze at a specific cell (everything above and left is frozen)
sheet.freezePanes.freezeAt(sheet.getRange("B2"));
```

---

## 4. Dashboard Layout Patterns

### KPI Header Row
```typescript
// Build KPI summary cards across the top
const kpis = [
  { label: "Total Revenue", formula: "=SUM(Data!C2:C1000)", format: "$#,##0" },
  { label: "Avg Order Value", formula: "=AVERAGE(Data!C2:C1000)", format: "$#,##0.00" },
  { label: "Total Orders", formula: "=COUNTA(Data!A2:A1000)", format: "#,##0" },
  { label: "Growth Rate", formula: "=(SUM(Data!C500:C1000)-SUM(Data!C2:C499))/SUM(Data!C2:C499)", format: "0.0%" },
];

for (let i = 0; i < kpis.length; i++) {
  const col = i * 3; // space between KPIs
  const labelCell = sheet.getCell(0, col);
  const valueCell = sheet.getCell(1, col);

  labelCell.values = [[kpis[i].label]];
  labelCell.format.font.size = 9;
  labelCell.format.font.color = "#666666";

  valueCell.formulas = [[kpis[i].formula]];
  valueCell.numberFormat = [[kpis[i].format]];
  valueCell.format.font.size = 20;
  valueCell.format.font.bold = true;
}
```

### Dashboard Structure Template
```
Row 1-2:    KPI summary cards (large numbers)
Row 3:      Blank separator
Row 4-18:   Main chart (left) + Secondary chart (right)
Row 19:     Blank separator
Row 20-35:  Detail table with filters
Row 36:     Blank separator
Row 37-45:  Trend chart (full width)
```

### Professional Color Palette
```typescript
const COLORS = {
  primary:    "#4472C4",   // blue
  secondary:  "#ED7D31",   // orange
  success:    "#70AD47",   // green
  danger:     "#FF0000",   // red
  warning:    "#FFC000",   // amber
  neutral:    "#A5A5A5",   // gray
  background: "#F2F2F2",   // light gray
  text:       "#333333",   // dark gray
  white:      "#FFFFFF",
};

// 6-color chart palette
const CHART_PALETTE = ["#4472C4", "#ED7D31", "#A5A5A5", "#FFC000", "#5B9BD5", "#70AD47"];
```

---

## 5. Number Format Quick Reference

| Type | Format Code | Example |
|---|---|---|
| Integer | `#,##0` | 1,234 |
| 2 decimals | `#,##0.00` | 1,234.56 |
| Currency | `$#,##0.00` | $1,234.56 |
| Currency (neg) | `$#,##0;($#,##0)` | ($500) |
| Percentage | `0.0%` | 15.5% |
| Date | `yyyy-mm-dd` | 2026-03-29 |
| Date (display) | `mmm dd, yyyy` | Mar 29, 2026 |
| Short date | `mm/dd/yy` | 03/29/26 |
| Time | `hh:mm:ss` | 14:30:00 |
| Millions | `$#,##0,,"M"` | $1M |
| Thousands | `$#,##0,"K"` | $1,234K |
| Accounting | `_($* #,##0_)` | $ 1,234 |
| Text | `@` | (as-is) |
| Custom neg/zero | `#,##0;(#,##0);"-"` | 1,234 / (500) / - |

---

## 6. Design Principles for Dashboards

1. **One page, one story** — each dashboard sheet answers ONE question
2. **KPIs at the top** — most important numbers immediately visible
3. **Charts middle** — visual patterns and trends
4. **Detail table bottom** — drill-down data for those who want it
5. **Consistent colors** — same color = same meaning across all charts
6. **Minimal gridlines** — remove distracting lines, use whitespace
7. **Clear labels** — every chart has a title, axis labels, and units
8. **No 3D charts** — they distort data perception. Always use 2D.
9. **Max 6 slices in pie charts** — group small items into "Other"
10. **Right chart for right data** — time series → line, comparison → bar, proportion → pie/treemap
