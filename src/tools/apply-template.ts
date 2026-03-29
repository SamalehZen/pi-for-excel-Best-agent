/**
 * apply_template — List, preview, and apply design templates to Excel worksheets.
 *
 * Uses the template registry to manage bundled and user-created design templates.
 * Supports two application modes:
 * - "full": create complete template with structure + sample data + formatting
 * - "design_only": apply only the visual design to existing data
 */

import { Type, type Static } from "@sinclair/typebox";
import type { AgentTool, AgentToolResult } from "@mariozechner/pi-agent-core";
import { excelRun, qualifiedAddress } from "../excel/helpers.js";
import { getWorkbookChangeAuditLog } from "../audit/workbook-change-audit.js";
import { getErrorMessage } from "../utils/errors.js";
import {
  NON_CHECKPOINTED_MUTATION_NOTE,
  NON_CHECKPOINTED_MUTATION_REASON,
  recoveryCheckpointUnavailable,
} from "./recovery-metadata.js";
import { finalizeMutationOperation } from "./mutation/finalize.js";
import { appendMutationResultNote } from "./mutation/result-note.js";
import type { MutationFinalizeDependencies } from "./mutation/types.js";
import type {
  ApplyTemplateListDetails,
  ApplyTemplatePreviewDetails,
  ApplyTemplateApplyDetails,
  ApplyTemplateDetails,
} from "./tool-details.js";
import {
  listTemplates,
  getTemplateById,
  loadUserTemplates,
  type TemplateSettingsStore,
  type TemplateDefinition,
  type TemplateSummary,
  type TemplateColumn,
  type TemplateZone,
} from "../templates/index.js";
import { getAppStorage } from "@mariozechner/pi-web-ui/dist/storage/app-storage.js";

function StringEnum<T extends string[]>(values: [...T], opts?: { description?: string }) {
  return Type.Union(
    values.map((v) => Type.Literal(v)),
    opts,
  );
}

const schema = Type.Object({
  action: StringEnum(
    ["list", "preview", "apply", "gallery"],
    { description: "list = show available templates, preview = show template design details, apply = apply template to sheet, gallery = open visual template gallery for user selection." },
  ),
  template_id: Type.Optional(
    Type.String({ description: "Template ID (required for preview/apply). Use list to see available IDs." }),
  ),
  mode: Type.Optional(
    StringEnum(["full", "design_only"], {
      description: '"full" = create complete template with structure + sample data + formatting (for blank sheets). "design_only" = apply only the visual design to existing data. Default: "full".',
    }),
  ),
  sheet: Type.Optional(
    Type.String({ description: "Target sheet name. Defaults to active sheet." }),
  ),
  header_row: Type.Optional(
    Type.Number({ description: 'Override auto-detected header row for "design_only" mode.' }),
  ),
  data_start_row: Type.Optional(
    Type.Number({ description: 'Override auto-detected data start row for "design_only" mode.' }),
  ),
  data_end_row: Type.Optional(
    Type.Number({ description: 'Override auto-detected data end row for "design_only" mode.' }),
  ),
  total_row: Type.Optional(
    Type.Number({ description: 'Override auto-detected total row for "design_only" mode.' }),
  ),
  title_row: Type.Optional(
    Type.Number({ description: 'Override auto-detected title row for "design_only" mode.' }),
  ),
});

type Params = Static<typeof schema>;

const mutationFinalizeDependencies: MutationFinalizeDependencies = {
  appendAuditEntry: (entry) => getWorkbookChangeAuditLog().append(entry),
};

function colLetterToIndex(col: string): number {
  let idx = 0;
  for (let i = 0; i < col.length; i++) {
    idx = idx * 26 + (col.toUpperCase().charCodeAt(i) - 64);
  }
  return idx;
}

function indexToColLetter(index: number): string {
  let result = "";
  let n = index;
  while (n > 0) {
    const mod = (n - 1) % 26;
    result = String.fromCharCode(65 + mod) + result;
    n = Math.floor((n - 1) / 26);
  }
  return result;
}

function getColumnRange(columns: readonly TemplateColumn[], row: number): string {
  if (columns.length === 0) return `A${row}`;
  const firstCol = columns[0].col;
  const lastCol = columns[columns.length - 1].col;
  return `${firstCol}${row}:${lastCol}${row}`;
}

function getDataSlotRows(zones: readonly TemplateZone[]): number[] {
  const rows: number[] = [];
  for (const zone of zones) {
    if (zone.type === "data" || zone.type === "data_alternate") {
      if (Array.isArray(zone.rows)) {
        for (let r = zone.rows[0]; r <= zone.rows[1]; r++) {
          rows.push(r);
        }
      } else {
        rows.push(zone.rows);
      }
    }
  }
  return rows;
}

async function tryLoadUserTemplates(): Promise<TemplateDefinition[]> {
  try {
    const storage = getAppStorage();
    const settings: TemplateSettingsStore = {
      get: (key: string) => storage.settings.get(key),
      set: (key: string, value: unknown) => storage.settings.set(key, value),
    };
    return await loadUserTemplates(settings);
  } catch {
    return [];
  }
}

function buildTemplateListTable(templates: TemplateSummary[]): string {
  const lines: string[] = [
    "| ID | Name | Category | Font | Primary Color | Columns |",
    "|---|---|---|---|---|---|",
  ];
  for (const t of templates) {
    lines.push(`| ${t.id} | ${t.name} | ${t.category} | ${t.fontFamily} | ${t.primaryColor} | ${t.columnCount} |`);
  }
  return lines.join("\n");
}

function buildPreviewMarkdown(template: TemplateDefinition): string {
  const { palette, typography } = template.design;
  const { structure } = template;
  const lines: string[] = [];

  lines.push(`## ${template.name}`);
  lines.push("");
  lines.push(`**Category:** ${template.category}`);
  lines.push(`**Description:** ${template.description}`);
  lines.push("");

  lines.push("### Palette");
  lines.push(`- Title: bg ${palette.titleBg}, fg ${palette.titleFg}`);
  lines.push(`- Header: bg ${palette.headerBg}, fg ${palette.headerFg}`);
  lines.push(`- Label: bg ${palette.labelBg}, fg ${palette.labelFg}`);
  lines.push(`- Accent: bg ${palette.accentBg}, fg ${palette.accentFg}`);
  lines.push(`- Total: bg ${palette.totalBg}, fg ${palette.totalFg}`);
  if (palette.alternateBg) {
    lines.push(`- Alternate row: ${palette.alternateBg}`);
  }
  lines.push("");

  lines.push("### Typography");
  lines.push(`- Font: ${typography.fontFamily}${typography.titleFontFamily ? ` (title: ${typography.titleFontFamily})` : ""}`);
  lines.push(`- Title size: ${typography.titleSize}pt`);
  lines.push(`- Header size: ${typography.headerSize}pt`);
  lines.push(`- Body size: ${typography.bodySize}pt`);
  lines.push("");

  lines.push("### Structure");
  lines.push(`- Title row: ${structure.titleRow} — "${structure.title}"`);
  if (structure.metaFields.length > 0) {
    lines.push(`- Meta fields: ${structure.metaFields.map((m) => m.label).join(", ")}`);
  }
  lines.push(`- Header row: ${structure.headerRow}`);
  lines.push(`- Columns (${structure.columns.length}): ${structure.columns.map((c) => c.header).join(", ")}`);
  lines.push(`- Sample data rows: ${structure.sampleData.length}`);
  if (structure.totalRow) {
    lines.push(`- Total row: ${structure.totalRow.row}${structure.totalRow.label ? ` — "${structure.totalRow.label}"` : ""}`);
  }
  lines.push(`- Span: ${structure.columnSpan}, total rows: ${structure.totalRows}`);
  lines.push("");

  lines.push("### Zones");
  for (const zone of structure.zones) {
    const rowsStr = Array.isArray(zone.rows) ? `${zone.rows[0]}–${zone.rows[1]}` : String(zone.rows);
    lines.push(`- ${zone.type}: row(s) ${rowsStr}`);
  }

  return lines.join("\n");
}

async function executeList(): Promise<AgentToolResult<ApplyTemplateListDetails>> {
  const userTemplates = await tryLoadUserTemplates();
  const templates = listTemplates(userTemplates.length > 0 ? userTemplates : undefined);
  const table = buildTemplateListTable(templates);

  return {
    content: [{ type: "text", text: `**${templates.length} templates available:**\n\n${table}` }],
    details: {
      kind: "apply_template_list",
      count: templates.length,
      templateIds: templates.map((t) => t.id),
    },
  };
}

async function executeGallery(): Promise<AgentToolResult<ApplyTemplateListDetails>> {
  const userTemplates = await tryLoadUserTemplates();
  const templates = listTemplates(userTemplates.length > 0 ? userTemplates : undefined);

  try {
    const { showTemplateGallery } = await import("../ui/template-gallery-host.js");
    const galleryResult = await new Promise<{ templateId: string; xlsxFile: string; templateName: string } | null>((resolve) => {
      showTemplateGallery({
        onTemplateSelected: (templateId, xlsxFile, templateName) => {
          resolve({ templateId, xlsxFile, templateName });
        },
        onClosed: () => {
          resolve(null);
        },
      });
    });

    if (!galleryResult) {
      return {
        content: [{ type: "text", text: "Template gallery closed without selection." }],
        details: {
          kind: "apply_template_list",
          count: templates.length,
          templateIds: templates.map((t) => t.id),
        },
      };
    }

    return {
      content: [{ type: "text", text: `User selected template **"${galleryResult.templateName}"** (ID: \`${galleryResult.templateId}\`). Use \`action: "apply"\` with \`template_id: "${galleryResult.templateId}"\` to apply it.` }],
      details: {
        kind: "apply_template_list",
        count: templates.length,
        templateIds: templates.map((t) => t.id),
      },
    };
  } catch {
    const table = buildTemplateListTable(templates);
    return {
      content: [{ type: "text", text: `Gallery UI unavailable. Showing template list instead.\n\n**${templates.length} templates available:**\n\n${table}` }],
      details: {
        kind: "apply_template_list",
        count: templates.length,
        templateIds: templates.map((t) => t.id),
      },
    };
  }
}

async function executePreview(params: Params): Promise<AgentToolResult<ApplyTemplatePreviewDetails>> {
  if (!params.template_id) {
    return {
      content: [{ type: "text", text: "Error: `template_id` is required for preview. Use `action: \"list\"` to see available IDs." }],
      details: {
        kind: "apply_template_preview",
        templateId: "",
        templateName: "",
        category: "",
      },
    };
  }

  const userTemplates = await tryLoadUserTemplates();
  const template = getTemplateById(params.template_id, userTemplates.length > 0 ? userTemplates : undefined);
  if (!template) {
    return {
      content: [{ type: "text", text: `Error: Template "${params.template_id}" not found. Use \`action: "list"\` to see available IDs.` }],
      details: {
        kind: "apply_template_preview",
        templateId: params.template_id,
        templateName: "",
        category: "",
      },
    };
  }

  const markdown = buildPreviewMarkdown(template);
  return {
    content: [{ type: "text", text: markdown }],
    details: {
      kind: "apply_template_preview",
      templateId: template.id,
      templateName: template.name,
      category: template.category,
    },
  };
}

async function applyFull(
  toolCallId: string,
  params: Params,
  template: TemplateDefinition,
): Promise<AgentToolResult<ApplyTemplateApplyDetails>> {
  const { design, structure } = template;
  const { palette, typography } = design;

  const result = await excelRun(async (context) => {
    const sheet = params.sheet
      ? context.workbook.worksheets.getItem(params.sheet)
      : context.workbook.worksheets.getActiveWorksheet();
    sheet.load("name");

    const titleRow = structure.titleRow;
    const lastColLetter = structure.columns.length > 0
      ? structure.columns[structure.columns.length - 1].col
      : "A";

    const titleCell = sheet.getRange(`A${titleRow}`);
    titleCell.values = [[structure.title]];
    const titleMergeRange = sheet.getRange(`A${titleRow}:${lastColLetter}${titleRow}`);
    titleMergeRange.merge();
    titleMergeRange.format.font.name = typography.titleFontFamily ?? typography.fontFamily;
    titleMergeRange.format.font.size = typography.titleSize;
    titleMergeRange.format.font.bold = design.titleBold;
    titleMergeRange.format.font.color = palette.titleFg;
    titleMergeRange.format.fill.color = palette.titleBg;
    if (design.titleRowHeight) {
      titleMergeRange.format.rowHeight = design.titleRowHeight;
    }

    for (const meta of structure.metaFields) {
      const labelCell = sheet.getRange(`${meta.labelCol}${meta.row}`);
      labelCell.values = [[meta.label]];
      labelCell.format.font.bold = true;
      labelCell.format.fill.color = palette.labelBg;
      labelCell.format.font.color = palette.labelFg;
      labelCell.format.font.name = typography.fontFamily;
      labelCell.format.font.size = typography.bodySize;

      const valueCell = sheet.getRange(`${meta.valueCol}${meta.row}`);
      valueCell.values = [[meta.placeholder]];
      valueCell.format.fill.color = palette.labelBg;
      valueCell.format.font.color = palette.labelFg;
      valueCell.format.font.name = typography.fontFamily;
      valueCell.format.font.size = typography.bodySize;
    }

    const metaRows = new Set(structure.metaFields.map((m) => m.row));
    for (const metaRow of metaRows) {
      const metaRowRange = sheet.getRange(`A${metaRow}:${lastColLetter}${metaRow}`);
      metaRowRange.format.fill.color = palette.labelBg;
    }

    const writeColumnHeaders = (row: number) => {
      for (const col of structure.columns) {
        const headerCell = sheet.getRange(`${col.col}${row}`);
        headerCell.values = [[col.header]];
        headerCell.format.font.bold = true;
        headerCell.format.fill.color = palette.headerBg;
        headerCell.format.font.color = palette.headerFg;
        headerCell.format.font.size = typography.headerSize;
        headerCell.format.font.name = typography.fontFamily;
        if (col.alignment) {
          headerCell.format.horizontalAlignment = col.alignment;
        }
      }
      if (design.headerRowHeight) {
        const rowRange = sheet.getRange(`A${row}`);
        rowRange.format.rowHeight = design.headerRowHeight;
      }
    };

    writeColumnHeaders(structure.headerRow);

    for (const zone of structure.zones) {
      const zoneRow = Array.isArray(zone.rows) ? zone.rows[0] : zone.rows;
      if (zone.type === "column_header" && zoneRow !== structure.headerRow) {
        writeColumnHeaders(zoneRow);
      }
      if (zone.type === "section_header") {
        const sectionRange = sheet.getRange(`A${zoneRow}:${lastColLetter}${zoneRow}`);
        sectionRange.format.fill.color = palette.headerBg;
        sectionRange.format.font.color = palette.headerFg;
        sectionRange.format.font.bold = true;
        sectionRange.format.font.name = typography.fontFamily;
      }
      if (zone.type === "accent") {
        const accentRange = sheet.getRange(`A${zoneRow}:${lastColLetter}${zoneRow}`);
        accentRange.format.fill.color = palette.accentBg;
        accentRange.format.font.color = palette.accentFg;
        accentRange.format.font.bold = true;
        accentRange.format.font.name = typography.fontFamily;
      }
    }

    const dataSlotRows = getDataSlotRows(structure.zones);
    const fallbackStartRow = structure.headerRow + 1;

    for (let rowIdx = 0; rowIdx < structure.sampleData.length; rowIdx++) {
      const rowData = structure.sampleData[rowIdx];
      const currentRow = rowIdx < dataSlotRows.length
        ? dataSlotRows[rowIdx]
        : fallbackStartRow + rowIdx;

      for (let colIdx = 0; colIdx < structure.columns.length && colIdx < rowData.length; colIdx++) {
        const col = structure.columns[colIdx];
        const cell = sheet.getRange(`${col.col}${currentRow}`);
        const value = rowData[colIdx];
        cell.values = [[value ?? ""]];
        cell.format.font.name = typography.fontFamily;
        cell.format.font.size = typography.bodySize;

        if (col.alignment) {
          cell.format.horizontalAlignment = col.alignment;
        }
        if (col.isBold) {
          cell.format.font.bold = true;
        }
        if (col.isAccent) {
          cell.format.fill.color = palette.accentBg;
          cell.format.font.color = palette.accentFg;
        }
        if (design.alternatingRows && palette.alternateBg && rowIdx % 2 === 1) {
          if (!col.isAccent) {
            cell.format.fill.color = palette.alternateBg;
          }
        }
      }

      if (design.defaultRowHeight) {
        const rowRange = sheet.getRange(`A${currentRow}`);
        rowRange.format.rowHeight = design.defaultRowHeight;
      }
    }

    if (design.defaultRowHeight) {
      for (let i = structure.sampleData.length; i < dataSlotRows.length; i++) {
        const rowRange = sheet.getRange(`A${dataSlotRows[i]}`);
        rowRange.format.rowHeight = design.defaultRowHeight;
      }
    }

    if (structure.totalRow) {
      const totalRowNum = structure.totalRow.row;
      if (structure.totalRow.label && structure.totalRow.labelCol) {
        const totalLabelCell = sheet.getRange(`${structure.totalRow.labelCol}${totalRowNum}`);
        totalLabelCell.values = [[structure.totalRow.label]];
      }
      const totalRowRange = sheet.getRange(getColumnRange(structure.columns, totalRowNum));
      totalRowRange.format.font.bold = true;
      totalRowRange.format.fill.color = palette.totalBg;
      totalRowRange.format.font.color = palette.totalFg;
      totalRowRange.format.font.name = typography.fontFamily;
      totalRowRange.format.font.size = typography.bodySize;

      for (const [colLetter, value] of Object.entries(structure.totalRow.values)) {
        const totalCell = sheet.getRange(`${colLetter}${totalRowNum}`);
        totalCell.values = [[value]];
      }
    }

    for (const col of structure.columns) {
      if (col.width) {
        const colRange = sheet.getRange(`${col.col}:${col.col}`);
        colRange.format.columnWidth = col.width * 7.2;
      }
    }

    await context.sync();

    const endRow = structure.totalRow ? structure.totalRow.row : structure.totalRows;
    const address = `A${titleRow}:${lastColLetter}${endRow}`;

    return { sheetName: sheet.name, address };
  });

  const fullAddr = qualifiedAddress(result.sheetName, result.address);

  const toolResult: AgentToolResult<ApplyTemplateApplyDetails> = {
    content: [
      {
        type: "text",
        text: `Applied template **${template.name}** (full mode) to **${fullAddr}**.`,
      },
    ],
    details: {
      kind: "apply_template_apply",
      templateId: template.id,
      templateName: template.name,
      mode: "full",
      address: fullAddr,
      recovery: recoveryCheckpointUnavailable(NON_CHECKPOINTED_MUTATION_REASON),
    },
  };

  appendMutationResultNote(toolResult, NON_CHECKPOINTED_MUTATION_NOTE);
  return toolResult;
}

interface DetectedStructure {
  titleRow: number | undefined;
  headerRow: number | undefined;
  dataStartRow: number | undefined;
  dataEndRow: number | undefined;
  totalRow: number | undefined;
  usedColCount: number;
}

async function detectSheetStructure(
  sheet: Excel.Worksheet,
  context: Excel.RequestContext,
): Promise<DetectedStructure> {
  const usedRange = sheet.getUsedRange();
  usedRange.load("rowIndex,rowCount,columnCount,values");
  await context.sync();

  const values = usedRange.values as unknown[][];
  const startRow = usedRange.rowIndex + 1;
  const rowCount = usedRange.rowCount;
  const colCount = usedRange.columnCount;

  let titleRow: number | undefined;
  let headerRow: number | undefined;
  let totalRow: number | undefined;

  if (rowCount >= 1) {
    const firstRow = values[0];
    const nonEmpty = firstRow.filter((v) => v !== null && v !== undefined && v !== "");
    if (nonEmpty.length <= 2) {
      titleRow = startRow;
    }
  }

  const scanLimit = Math.min(rowCount, 8);
  for (let i = (titleRow !== undefined ? 1 : 0); i < scanLimit; i++) {
    const row = values[i];
    const nonEmpty = row.filter((v) => v !== null && v !== undefined && v !== "");
    const textCells = row.filter((v) => typeof v === "string" && v.length > 0);
    if (nonEmpty.length >= Math.ceil(colCount * 0.5) && textCells.length >= Math.ceil(nonEmpty.length * 0.5)) {
      headerRow = startRow + i;
      break;
    }
  }

  const lastRowValues = values[rowCount - 1];
  if (lastRowValues) {
    for (const cell of lastRowValues) {
      if (typeof cell === "string") {
        const upper = cell.toUpperCase().trim();
        if (upper === "TOTAL" || upper === "TOTALS" || upper === "SUM" || upper.startsWith("TOTAL ")) {
          totalRow = startRow + rowCount - 1;
          break;
        }
      }
    }
  }

  const dataStartRow = headerRow !== undefined ? headerRow + 1 : (titleRow !== undefined ? titleRow + 1 : startRow);
  const dataEndRow = totalRow !== undefined ? totalRow - 1 : startRow + rowCount - 1;

  return {
    titleRow,
    headerRow,
    dataStartRow: dataStartRow <= dataEndRow ? dataStartRow : undefined,
    dataEndRow: dataStartRow <= dataEndRow ? dataEndRow : undefined,
    totalRow,
    usedColCount: colCount,
  };
}

async function applyDesignOnly(
  toolCallId: string,
  params: Params,
  template: TemplateDefinition,
): Promise<AgentToolResult<ApplyTemplateApplyDetails>> {
  const { design } = template;
  const { palette, typography } = design;

  const result = await excelRun(async (context) => {
    const sheet = params.sheet
      ? context.workbook.worksheets.getItem(params.sheet)
      : context.workbook.worksheets.getActiveWorksheet();
    sheet.load("name");

    const detected = await detectSheetStructure(sheet, context);

    const titleRow = params.title_row ?? detected.titleRow;
    const headerRow = params.header_row ?? detected.headerRow;
    const dataStartRow = params.data_start_row ?? detected.dataStartRow;
    const dataEndRow = params.data_end_row ?? detected.dataEndRow;
    const totalRow = params.total_row ?? detected.totalRow;

    const usedRange = sheet.getUsedRange();
    usedRange.load("columnIndex,columnCount,rowIndex,rowCount");
    await context.sync();

    const firstColLetter = indexToColLetter(usedRange.columnIndex + 1);
    const lastColLetter = indexToColLetter(usedRange.columnIndex + usedRange.columnCount);
    const firstUsedRow = usedRange.rowIndex + 1;
    const lastUsedRow = firstUsedRow + usedRange.rowCount - 1;

    const allRange = sheet.getRange(`${firstColLetter}${firstUsedRow}:${lastColLetter}${lastUsedRow}`);
    allRange.format.font.name = typography.fontFamily;
    allRange.format.font.size = typography.bodySize;

    if (titleRow) {
      const titleRange = sheet.getRange(`${firstColLetter}${titleRow}:${lastColLetter}${titleRow}`);
      titleRange.format.font.name = typography.titleFontFamily ?? typography.fontFamily;
      titleRange.format.font.size = typography.titleSize;
      titleRange.format.font.bold = design.titleBold;
      titleRange.format.font.color = palette.titleFg;
      titleRange.format.fill.color = palette.titleBg;
      if (design.titleRowHeight) {
        titleRange.format.rowHeight = design.titleRowHeight;
      }
    }

    if (titleRow && headerRow && headerRow - titleRow > 1) {
      for (let r = titleRow + 1; r < headerRow; r++) {
        const metaRange = sheet.getRange(`${firstColLetter}${r}:${lastColLetter}${r}`);
        metaRange.format.fill.color = palette.labelBg;
        metaRange.format.font.color = palette.labelFg;
      }
    }

    if (headerRow) {
      const headerRange = sheet.getRange(`${firstColLetter}${headerRow}:${lastColLetter}${headerRow}`);
      headerRange.format.font.bold = true;
      headerRange.format.fill.color = palette.headerBg;
      headerRange.format.font.color = palette.headerFg;
      headerRange.format.font.size = typography.headerSize;
      headerRange.format.font.name = typography.fontFamily;
      if (design.headerRowHeight) {
        headerRange.format.rowHeight = design.headerRowHeight;
      }
    }

    if (dataStartRow && dataEndRow && design.alternatingRows && palette.alternateBg) {
      for (let r = dataStartRow; r <= dataEndRow; r++) {
        if ((r - dataStartRow) % 2 === 1) {
          const rowRange = sheet.getRange(`${firstColLetter}${r}:${lastColLetter}${r}`);
          rowRange.format.fill.color = palette.alternateBg;
        }
      }
    }

    if (dataStartRow && dataEndRow && design.defaultRowHeight) {
      for (let r = dataStartRow; r <= dataEndRow; r++) {
        const rowRange = sheet.getRange(`A${r}`);
        rowRange.format.rowHeight = design.defaultRowHeight;
      }
    }

    if (totalRow) {
      const totalRange = sheet.getRange(`${firstColLetter}${totalRow}:${lastColLetter}${totalRow}`);
      totalRange.format.font.bold = true;
      totalRange.format.fill.color = palette.totalBg;
      totalRange.format.font.color = palette.totalFg;
    }

    const lastColIdx = usedRange.columnIndex + usedRange.columnCount;
    const firstColIdx = usedRange.columnIndex + 1;

    for (const col of template.structure.columns) {
      const colIdx = colLetterToIndex(col.col);
      if (colIdx < firstColIdx || colIdx > lastColIdx) continue;
      if (col.width) {
        const colRange = sheet.getRange(`${col.col}:${col.col}`);
        colRange.format.columnWidth = col.width * 7.2;
      }
      if (col.isAccent && dataStartRow && dataEndRow) {
        const accentRange = sheet.getRange(`${col.col}${dataStartRow}:${col.col}${dataEndRow}`);
        accentRange.format.fill.color = palette.accentBg;
        accentRange.format.font.color = palette.accentFg;
      }
    }

    await context.sync();

    const address = `${firstColLetter}${firstUsedRow}:${lastColLetter}${lastUsedRow}`;
    return {
      sheetName: sheet.name,
      address,
      detectedTitleRow: detected.titleRow,
      detectedHeaderRow: detected.headerRow,
      detectedDataRows: detected.dataStartRow && detected.dataEndRow
        ? [detected.dataStartRow, detected.dataEndRow] as [number, number]
        : undefined,
      detectedTotalRow: detected.totalRow,
    };
  });

  const fullAddr = qualifiedAddress(result.sheetName, result.address);
  const detections: string[] = [];
  if (result.detectedTitleRow) detections.push(`title row ${result.detectedTitleRow}`);
  if (result.detectedHeaderRow) detections.push(`header row ${result.detectedHeaderRow}`);
  if (result.detectedDataRows) detections.push(`data rows ${result.detectedDataRows[0]}–${result.detectedDataRows[1]}`);
  if (result.detectedTotalRow) detections.push(`total row ${result.detectedTotalRow}`);
  const detectSummary = detections.length > 0
    ? `\n\nDetected structure: ${detections.join(", ")}.`
    : "\n\nNo structure detected; formatting applied to full used range.";

  const toolResult: AgentToolResult<ApplyTemplateApplyDetails> = {
    content: [
      {
        type: "text",
        text: `Applied template **${template.name}** (design_only mode) to **${fullAddr}**.${detectSummary}`,
      },
    ],
    details: {
      kind: "apply_template_apply",
      templateId: template.id,
      templateName: template.name,
      mode: "design_only",
      address: fullAddr,
      detectedTitleRow: result.detectedTitleRow,
      detectedHeaderRow: result.detectedHeaderRow,
      detectedDataRows: result.detectedDataRows,
      detectedTotalRow: result.detectedTotalRow,
      recovery: recoveryCheckpointUnavailable(NON_CHECKPOINTED_MUTATION_REASON),
    },
  };

  appendMutationResultNote(toolResult, NON_CHECKPOINTED_MUTATION_NOTE);
  return toolResult;
}

async function executeApply(
  toolCallId: string,
  params: Params,
): Promise<AgentToolResult<ApplyTemplateApplyDetails>> {
  if (!params.template_id) {
    return {
      content: [{ type: "text", text: "Error: `template_id` is required for apply. Use `action: \"list\"` to see available IDs." }],
      details: {
        kind: "apply_template_apply",
        templateId: "",
        templateName: "",
        mode: params.mode === "design_only" ? "design_only" : "full",
      },
    };
  }

  const userTemplates = await tryLoadUserTemplates();
  const template = getTemplateById(params.template_id, userTemplates.length > 0 ? userTemplates : undefined);
  if (!template) {
    return {
      content: [{ type: "text", text: `Error: Template "${params.template_id}" not found. Use \`action: "list"\` to see available IDs.` }],
      details: {
        kind: "apply_template_apply",
        templateId: params.template_id,
        templateName: "",
        mode: params.mode === "design_only" ? "design_only" : "full",
      },
    };
  }

  const mode = params.mode ?? "full";
  if (mode === "design_only") {
    return applyDesignOnly(toolCallId, params, template);
  }
  return applyFull(toolCallId, params, template);
}

export function createApplyTemplateTool(): AgentTool<typeof schema, ApplyTemplateDetails> {
  return {
    name: "apply_template",
    label: "Apply Template",
    description:
      "List, preview, and apply design templates to worksheets. " +
      "10 bundled templates (timesheet, attendance, scorecard, forecast, contest tracker, daily report, lead tracking, schedule, resource planning, work planner, goal tracking). " +
      'Mode "full" creates complete structure + sample data + formatting on a blank sheet. ' +
      'Mode "design_only" detects existing data layout and applies only visual formatting.',
    parameters: schema,
    execute: async (
      toolCallId: string,
      params: Params,
    ): Promise<AgentToolResult<ApplyTemplateDetails>> => {
      try {
        if (params.action === "list") {
          return await executeList();
        }

        if (params.action === "gallery") {
          return await executeGallery();
        }

        if (params.action === "preview") {
          return await executePreview(params);
        }

        const result = await executeApply(toolCallId, params);

        const isError = result.details?.templateName === "";
        await finalizeMutationOperation(mutationFinalizeDependencies, {
          auditEntry: {
            toolName: "apply_template",
            toolCallId,
            blocked: isError,
            outputAddress: isError ? (params.sheet ?? "") : (result.details?.address ?? ""),
            changedCount: 0,
            changes: [],
            summary: isError
              ? `error: template "${params.template_id ?? "unknown"}" not applied`
              : `applied template "${params.template_id ?? "unknown"}" (${params.mode ?? "full"} mode)`,
          },
        });

        return result;
      } catch (e: unknown) {
        const message = getErrorMessage(e);

        await finalizeMutationOperation(mutationFinalizeDependencies, {
          auditEntry: {
            toolName: "apply_template",
            toolCallId,
            blocked: true,
            outputAddress: params.sheet ?? "",
            changedCount: 0,
            changes: [],
            summary: `error: ${message}`,
          },
        });

        return {
          content: [{ type: "text", text: `Error applying template: ${message}` }],
          details: {
            kind: "apply_template_apply",
            templateId: params.template_id ?? "",
            templateName: "",
            mode: params.mode === "design_only" ? "design_only" : "full",
          },
        };
      }
    },
  };
}
