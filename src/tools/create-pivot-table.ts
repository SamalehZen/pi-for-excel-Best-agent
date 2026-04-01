import type { AgentTool, AgentToolResult } from "@mariozechner/pi-agent-core";
import { Type, type Static } from "@sinclair/typebox";
import type { CreatePivotTableDetails } from "./tool-details.js";
import { excelRun, getRange, qualifiedAddress } from "../excel/helpers.js";
import { getErrorMessage } from "../utils/errors.js";

const schema = Type.Object({
  source_range: Type.String({
    description: 'Source data range with headers, e.g. "A1:D20" or "DataSheet!A1:D20".',
  }),
  target_cell: Type.String({
    description: 'Cell where to place the pivot table, e.g. "G1" or "PivotSheet!A1". Recommended: use a different sheet.',
  }),
  name: Type.Optional(Type.String({
    description: "Pivot table name. Auto-generated if omitted.",
  })),
  rows: Type.Array(Type.String(), {
    description: 'Column header names to use as row labels, e.g. ["Category", "Product"].',
  }),
  values: Type.Array(Type.String(), {
    description: 'Column header names to summarize, e.g. ["Sales", "Quantity"].',
  }),
  columns: Type.Optional(Type.Array(Type.String(), {
    description: 'Column header names for column labels, e.g. ["Region"].',
  })),
  filters: Type.Optional(Type.Array(Type.String(), {
    description: 'Column header names for filter area, e.g. ["Year"].',
  })),
  agg_func: Type.Optional(Type.String({
    description: 'Aggregation function for value fields: "sum", "count", "average", "max", "min". Default: "sum".',
  })),
});

type Params = Static<typeof schema>;

const VALID_AGGREGATIONS = ["sum", "count", "average", "max", "min"] as const;
type SupportedAggregation = (typeof VALID_AGGREGATIONS)[number];

const AGGREGATION_FUNCTION_MAP: Record<SupportedAggregation, Excel.AggregationFunction> = {
  sum: Excel.AggregationFunction.sum,
  count: Excel.AggregationFunction.count,
  average: Excel.AggregationFunction.average,
  max: Excel.AggregationFunction.max,
  min: Excel.AggregationFunction.min,
};

interface CreatePivotTableRunResult {
  pivotTableName: string;
  sheetName: string;
  targetAddress: string;
  sourceAddress: string;
}

function normalizeText(value: string): string {
  return value.trim().toLowerCase();
}

function cleanHierarchyName(value: string): string {
  return value
    .trim()
    .replace(/\s+\((sum|average|avg|count|min|max)\)\s*$/iu, "")
    .trim();
}

function dedupeNames(values: readonly string[]): string[] {
  const seen = new Set<string>();
  const result: string[] = [];

  for (const value of values) {
    const cleaned = cleanHierarchyName(value);
    const key = cleaned.toLowerCase();
    if (key.length === 0 || seen.has(key)) continue;
    seen.add(key);
    result.push(cleaned);
  }

  return result;
}

function buildHierarchyLookup(
  hierarchies: readonly Excel.PivotHierarchy[],
): Map<string, Excel.PivotHierarchy> {
  const map = new Map<string, Excel.PivotHierarchy>();
  for (const hierarchy of hierarchies) {
    map.set(hierarchy.name.trim().toLowerCase(), hierarchy);
  }
  return map;
}

function resolveHierarchy(
  value: string,
  hierarchyLookup: ReadonlyMap<string, Excel.PivotHierarchy>,
): Excel.PivotHierarchy | undefined {
  return hierarchyLookup.get(cleanHierarchyName(value).toLowerCase());
}

function requireHierarchy(
  value: string,
  hierarchyLookup: ReadonlyMap<string, Excel.PivotHierarchy>,
): Excel.PivotHierarchy {
  const hierarchy = resolveHierarchy(value, hierarchyLookup);
  if (!hierarchy) {
    throw new Error(`Pivot hierarchy "${cleanHierarchyName(value)}" could not be resolved.`);
  }
  return hierarchy;
}

function nextPivotTableName(existingNames: readonly string[]): string {
  const normalized = new Set(existingNames.map((name) => name.toLowerCase()));
  let index = 1;
  while (normalized.has(`pivottable${String(index)}`.toLowerCase())) {
    index += 1;
  }
  return `PivotTable${String(index)}`;
}

export function createCreatePivotTableTool(): AgentTool<typeof schema, CreatePivotTableDetails> {
  return {
    name: "create_pivot_table",
    label: "Create Pivot Table",
    description:
      "Create a native Excel PivotTable from a source range with row, column, value, and filter hierarchies.",
    parameters: schema,
    execute: async (
      _toolCallId: string,
      params: Params,
    ): Promise<AgentToolResult<CreatePivotTableDetails>> => {
      try {
        if (params.rows.length === 0) {
          return {
            content: [{ type: "text", text: "Error: rows must include at least one field name." }],
            details: { kind: "create_pivot_table" },
          };
        }

        if (params.values.length === 0) {
          return {
            content: [{ type: "text", text: "Error: values must include at least one field name." }],
            details: { kind: "create_pivot_table" },
          };
        }

        const rawAggregation = params.agg_func;
        const aggregationKey = normalizeText(params.agg_func || "sum");
        if (!VALID_AGGREGATIONS.includes(aggregationKey as SupportedAggregation)) {
          return {
            content: [{
              type: "text",
              text: `Error: invalid agg_func "${rawAggregation ?? ""}". Valid values: ${VALID_AGGREGATIONS.join(", ")}.`,
            }],
            details: { kind: "create_pivot_table" },
          };
        }

        const result = await excelRun<CreatePivotTableRunResult>(async (context) => {
          const { sheet: sourceSheet, range: sourceRange } = getRange(context, params.source_range);
          const { sheet: targetSheet, range: targetRange } = getRange(context, params.target_cell);
          sourceSheet.load("name");
          sourceRange.load("address,rowCount,columnCount");
          targetSheet.load("name");
          targetRange.load("address");

          let existingNamedPivot: Excel.PivotTable | null = null;
          const workbookPivotTables = context.workbook.pivotTables;
          if (params.name) {
            existingNamedPivot = workbookPivotTables.getItemOrNullObject(params.name);
            existingNamedPivot.load("isNullObject");
          } else {
            workbookPivotTables.load("items/name");
          }

          await context.sync();

          if (sourceRange.rowCount < 2) {
            throw new Error("source_range must include a header row and at least one data row.");
          }

          if (sourceRange.columnCount < 2) {
            throw new Error("source_range must include at least two columns.");
          }

          if (params.name && existingNamedPivot && !existingNamedPivot.isNullObject) {
            throw new Error(`A pivot table named "${params.name}" already exists.`);
          }

          const pivotTableName = params.name ?? nextPivotTableName(workbookPivotTables.items.map((item) => item.name));
          const pivotTable = targetSheet.pivotTables.add(pivotTableName, sourceRange, targetRange);
          pivotTable.hierarchies.load("items/name");
          await context.sync();

          const availableFields = pivotTable.hierarchies.items.map((hierarchy) => hierarchy.name);
          const hierarchyLookup = buildHierarchyLookup(pivotTable.hierarchies.items);

          const rowNames = dedupeNames(params.rows);
          const columnNames = dedupeNames(params.columns ?? []);
          const filterNames = dedupeNames(params.filters ?? []);
          const valueNames = dedupeNames(params.values);

          const requestedNames = [...rowNames, ...columnNames, ...filterNames, ...valueNames];
          const missingFields = requestedNames.filter((name) => !resolveHierarchy(name, hierarchyLookup));

          if (missingFields.length > 0) {
            pivotTable.delete();
            await context.sync();
            throw new Error(
              `Unknown pivot field${missingFields.length === 1 ? "" : "s"}: ${missingFields.join(", ")}. Available fields: ${availableFields.join(", ")}.`,
            );
          }

          for (const rowName of rowNames) {
            pivotTable.rowHierarchies.add(requireHierarchy(rowName, hierarchyLookup));
          }

          for (const columnName of columnNames) {
            pivotTable.columnHierarchies.add(requireHierarchy(columnName, hierarchyLookup));
          }

          for (const filterName of filterNames) {
            pivotTable.filterHierarchies.add(requireHierarchy(filterName, hierarchyLookup));
          }

          for (const valueName of valueNames) {
            const dataHierarchy = pivotTable.dataHierarchies.add(requireHierarchy(valueName, hierarchyLookup));
            dataHierarchy.summarizeBy = AGGREGATION_FUNCTION_MAP[aggregationKey as SupportedAggregation];
          }

          return {
            pivotTableName,
            sheetName: targetSheet.name,
            targetAddress: qualifiedAddress(targetSheet.name, targetRange.address),
            sourceAddress: qualifiedAddress(sourceSheet.name, sourceRange.address),
          };
        });

        const lines = [
          `📊 Created pivot table **${result.pivotTableName}** on **${result.sheetName}** at \`${result.targetAddress}\`.`,
          `- Source: \`${result.sourceAddress}\``,
          `- Rows: ${dedupeNames(params.rows).join(", ")}`,
          params.columns && params.columns.length > 0 ? `- Columns: ${dedupeNames(params.columns).join(", ")}` : undefined,
          `- Values: ${dedupeNames(params.values).join(", ")} (${aggregationKey})`,
          params.filters && params.filters.length > 0 ? `- Filters: ${dedupeNames(params.filters).join(", ")}` : undefined,
        ].filter((line): line is string => typeof line === "string");

        return {
          content: [{ type: "text", text: lines.join("\n") }],
          details: {
            kind: "create_pivot_table",
            pivotTableName: result.pivotTableName,
            sheetName: result.sheetName,
            sourceAddress: result.sourceAddress,
          },
        };
      } catch (e: unknown) {
        return {
          content: [{ type: "text", text: `Error creating pivot table: ${getErrorMessage(e)}` }],
          details: { kind: "create_pivot_table" },
        };
      }
    },
  };
}
