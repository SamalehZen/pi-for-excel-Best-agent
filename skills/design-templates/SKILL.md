---
name: design-templates
description: Apply and manage Excel design templates. Use when the user asks to format a sheet with a professional design, apply a template, or extract/reuse a design from an existing spreadsheet.
compatibility: Requires HyperFix core tools. Uses apply_template, format_cells, read_range, get_workbook_overview, and write_cells.
---

# Design Templates

Use this skill when the user wants to apply a professional design to a worksheet, pick from bundled templates, or recreate a design from an uploaded spreadsheet.

## Overview

The `apply_template` tool supports two application modes:

- **full** — creates a complete template on a blank sheet: structure, sample data, and formatting. Best when the user starts from scratch.
- **design_only** — applies only the visual design (colors, fonts, borders) to a sheet that already has data. Auto-detects title, header, data, and total rows.

### Bundled templates

| ID | Category | Fonts | Palette |
|---|---|---|---|
| `monthly-time-sheet` | Timesheet | Calibri + Aptos Narrow | Dark navy / teal / orange |
| `meeting-attendance` | Attendance | Calibri | Blue / lavender |
| `sales-balanced-scorecard` | Sales | Avenir Book | Monochrome warm gray |
| `sales-forecast-12m` | Sales | Helvetica | Green / orange / yellow / blue |
| `sales-contest-tracker` | Sales | Helvetica | Blue / neon green / dark gray |
| `daily-sales-report` | Sales | Franklin Gothic Book | Teal / purple / green |

---

## Workflow: Applying a template (full mode)

Use this flow when the user wants a new sheet built from scratch.

1. **Determine intent.** Ask or infer what kind of document the user needs (timesheet, attendance, sales report, forecast, etc.).

2. **Show available templates.** Call `apply_template` with action "list" to display all bundled and user-created templates with their IDs, categories, fonts, and primary colors.

3. **Preview a template (optional).** If the user wants more detail before committing, call `apply_template` with action "preview" and the chosen template_id. This returns the full palette, typography, structure, and zone layout.

4. **Apply the template.** Call `apply_template` with action "apply", the chosen template_id, and mode "full". This creates the complete sheet with headers, sample data rows, formatting, and any meta fields.

5. **Customize after applying.** The user may want to:
   - Change meta field values (e.g. update the employee name or reporting period) using `write_cells`
   - Add data to empty rows using `write_cells`
   - Insert or remove rows using `modify_structure`
   - Adjust specific colors or fonts using `format_cells`
   - Modify column widths or row heights

---

## Workflow: Applying design to existing data (design_only mode)

Use this flow when the user already has data on the sheet and wants it professionally styled.

1. **Understand the sheet.** Call `get_workbook_overview` to see all sheets, row counts, and column counts.

2. **Inspect current formatting.** Call `read_range` in detailed mode on a sample area (e.g. the first 20 rows) to see the current structure and any existing formatting.

3. **Suggest a template.** Based on the data type, recommend a matching template (see template matching heuristics below). Or show the full list and let the user pick.

4. **Apply design only.** Call `apply_template` with action "apply", the chosen template_id, and mode "design_only". The tool auto-detects which rows are title, header, data, and total rows.

5. **Handle detection overrides.** If the auto-detection gets the structure wrong (e.g. it misidentifies the header row), the user can specify explicit row numbers. The override parameters are:
   - `header_row` — the row containing column headers
   - `data_start_row` — the first data row
   - `data_end_row` — the last data row
   - `total_row` — the summary/totals row
   - `title_row` — the title/heading row at the top

6. **Fine-tune.** After applying, use `format_cells` to adjust any specific cells that need different treatment (e.g. a merged title area, conditional accent colors, or special borders).

---

## Workflow: Extracting a design from an existing sheet (import design)

This is the power-user workflow. The user uploads an .xlsx file and says "I want my sheets to look like this." There is no single tool that extracts a full design automatically, so the assistant must manually read and recreate the design using existing tools.

### Step 1: Understand the uploaded file

Call `get_workbook_overview` to see the sheet names, dimensions, and general layout of the uploaded workbook.

### Step 2: Extract formatting details

Call `read_range` with mode "detailed" on the key areas of the source sheet:

- **Title row** — read the first 1-2 rows to capture font family, font size, font color, fill color, bold/italic state, and any merges.
- **Header row** — read the row containing column headers to capture font, colors, fill, and borders.
- **Sample data rows** — read 4-6 data rows to detect font, colors, alternating row patterns, and border styles.
- **Total/summary row** — read the last content row to capture bold state, fill, borders, and any special formatting.

### Step 3: Document the extracted design

Organize what you found into a design specification:

- **Color palette**: title bg/fg, header bg/fg, label bg/fg, accent bg/fg, alternate row bg, total bg/fg
- **Typography**: font family, sizes for title, header, and body text
- **Layout patterns**: alternating row colors, accent columns, merged cell ranges, border placement

### Step 4: Apply the extracted design to the target sheet

Use `format_cells` calls to recreate the design on the target sheet:

- Apply title formatting to the title row (font family, size, color, fill, bold)
- Apply header formatting to the header row (font, colors, fill, borders)
- Apply data formatting to all data rows (font, colors, alternating fill if detected)
- Apply total formatting to the total/summary row (font, bold, borders, fill)
- Set column widths and row heights to match the source if needed

### Step 5: Suggest saving for reuse

For designs the user wants to reuse frequently, suggest saving the design specification as a custom instruction via the `instructions` tool. This way the assistant can recreate it on future sheets without re-extracting.

---

## Best practices

- Always call `apply_template` with action "list" before suggesting a specific template. This ensures you show the current set of available templates, including any user-created ones.
- For design_only mode, always call `get_workbook_overview` first to understand the data layout before applying.
- When auto-detection of structure seems wrong, tell the user they can override with specific row numbers (header_row, data_start_row, data_end_row, total_row, title_row).
- After applying a template, offer to customize specific aspects: colors, fonts, column widths, or meta field values.
- When the user describes a desired style without referencing a template, check whether any bundled template matches their description before building formatting from scratch with `format_cells`.
- For multi-sheet workbooks, apply templates one sheet at a time. Specify the target sheet name with the `sheet` parameter.
- Prefer design_only mode when the sheet already has meaningful data — full mode overwrites everything.

---

## Template matching heuristics

When the user describes what they need, match their intent to the right template:

| User says... | Suggested template |
|---|---|
| "timesheet", "time tracking", "hours", "attendance log" | `monthly-time-sheet` |
| "meeting attendance", "roll call", "who attended" | `meeting-attendance` |
| "scorecard", "KPI", "balanced scorecard", "OKR" | `sales-balanced-scorecard` |
| "forecast", "projection", "12 month", "annual plan" | `sales-forecast-12m` |
| "contest", "competition", "leaderboard", "tracker" | `sales-contest-tracker` |
| "daily report", "invoice", "receipt", "sales log" | `daily-sales-report` |
| Wants something custom | Extract design from upload or build with `format_cells` |

If the user's request does not clearly match a single template, show the full list with `apply_template` action "list" and let them choose. If they have an uploaded file they want to replicate, follow the import design workflow above.
