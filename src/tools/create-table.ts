import type { AgentTool, AgentToolResult } from "@mariozechner/pi-agent-core";
import { Type, type Static } from "@sinclair/typebox";
import type { CreateTableDetails } from "./tool-details.js";
import { excelRun, getRange, qualifiedAddress } from "../excel/helpers.js";
import { getErrorMessage } from "../utils/errors.js";

const schema = Type.Object({
  range: Type.String({
    description: 'Data range for the table, e.g. "A1:D10" or "Sheet1!A1:D10". First row is treated as headers by default.',
  }),
  has_headers: Type.Optional(Type.Boolean({
    description: "Whether first row contains headers. Default: true.",
  })),
  table_name: Type.Optional(Type.String({
    description: "Optional name for the table. Auto-generated if omitted.",
  })),
  style: Type.Optional(Type.String({
    description: 'Table style name, e.g. "TableStyleMedium2", "TableStyleMedium9", "TableStyleLight1". Default: "TableStyleMedium2".',
  })),
});

type Params = Static<typeof schema>;

const DEFAULT_TABLE_STYLE = "TableStyleMedium2";

interface RangeBounds {
  rowIndex: number;
  columnIndex: number;
  rowCount: number;
  columnCount: number;
}

interface CreateTableRunResult {
  tableName: string;
  sheetName: string;
  address: string;
  rowCount: number;
  columnCount: number;
  style: string;
}

function rangesOverlap(a: RangeBounds, b: RangeBounds): boolean {
  const aLastRow = a.rowIndex + a.rowCount - 1;
  const aLastColumn = a.columnIndex + a.columnCount - 1;
  const bLastRow = b.rowIndex + b.rowCount - 1;
  const bLastColumn = b.columnIndex + b.columnCount - 1;

  return (
    a.rowIndex <= bLastRow &&
    aLastRow >= b.rowIndex &&
    a.columnIndex <= bLastColumn &&
    aLastColumn >= b.columnIndex
  );
}

export function createCreateTableTool(): AgentTool<typeof schema, CreateTableDetails> {
  return {
    name: "create_table",
    label: "Create Table",
    description:
      "Create a native Excel table from a worksheet range with headers, autofilter, and styling.",
    parameters: schema,
    execute: async (
      _toolCallId: string,
      params: Params,
    ): Promise<AgentToolResult<CreateTableDetails>> => {
      try {
        const result = await excelRun<CreateTableRunResult>(async (context) => {
          const { sheet, range } = getRange(context, params.range);
          sheet.load("name");
          range.load("address,rowIndex,columnIndex,rowCount,columnCount");

          const existingTables = sheet.tables;
          existingTables.load("items/name");

          let existingNamedTable: Excel.Table | null = null;
          if (params.table_name) {
            existingNamedTable = context.workbook.tables.getItemOrNullObject(params.table_name);
            existingNamedTable.load("isNullObject");
          }

          await context.sync();

          if (range.rowCount < 1 || range.columnCount < 1) {
            throw new Error("range must contain at least one cell.");
          }

          if (params.table_name && existingNamedTable && !existingNamedTable.isNullObject) {
            throw new Error(`A table named "${params.table_name}" already exists.`);
          }

          const tableRanges = existingTables.items.map((table) => {
            const tableRange = table.getRange();
            tableRange.load("address,rowIndex,columnIndex,rowCount,columnCount");
            return { table, tableRange };
          });

          await context.sync();

          for (const existing of tableRanges) {
            if (
              rangesOverlap(
                {
                  rowIndex: range.rowIndex,
                  columnIndex: range.columnIndex,
                  rowCount: range.rowCount,
                  columnCount: range.columnCount,
                },
                {
                  rowIndex: existing.tableRange.rowIndex,
                  columnIndex: existing.tableRange.columnIndex,
                  rowCount: existing.tableRange.rowCount,
                  columnCount: existing.tableRange.columnCount,
                },
              )
            ) {
              throw new Error(
                `Range ${qualifiedAddress(sheet.name, range.address)} overlaps existing table "${existing.table.name}" at ${qualifiedAddress(sheet.name, existing.tableRange.address)}.`,
              );
            }
          }

          const table = sheet.tables.add(range, params.has_headers ?? true);
          if (params.table_name) {
            table.name = params.table_name;
          }
          table.style = params.style ?? DEFAULT_TABLE_STYLE;
          table.highlightFirstColumn = false;
          table.highlightLastColumn = false;
          table.showBandedRows = true;
          table.showBandedColumns = false;
          table.showFilterButton = true;

          table.load("name,style");
          const tableRange = table.getRange();
          tableRange.load("address");
          const rowCount = table.rows.getCount();
          const columnCount = table.columns.getCount();
          await context.sync();

          return {
            tableName: table.name,
            sheetName: sheet.name,
            address: qualifiedAddress(sheet.name, tableRange.address),
            rowCount: rowCount.value,
            columnCount: columnCount.value,
            style: table.style,
          };
        });

        const lines = [
          `📋 Created table **${result.tableName}** on **${result.sheetName}** at \`${result.address}\`.`,
          `- Columns: ${result.columnCount}`,
          `- Data rows: ${result.rowCount}`,
          `- Style: ${result.style}`,
        ];

        return {
          content: [{ type: "text", text: lines.join("\n") }],
          details: {
            kind: "create_table",
            tableName: result.tableName,
            sheetName: result.sheetName,
            address: result.address,
          },
        };
      } catch (e: unknown) {
        return {
          content: [{ type: "text", text: `Error creating table: ${getErrorMessage(e)}` }],
          details: { kind: "create_table" },
        };
      }
    },
  };
}
