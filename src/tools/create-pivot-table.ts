import type { AgentTool, AgentToolResult } from "@mariozechner/pi-agent-core";
import { Type, type Static } from "@sinclair/typebox";
import type { CreatePivotTableDetails } from "./tool-details.js";
import { excelRun, getRange, qualifiedAddress } from "../excel/helpers.js";
import { getErrorMessage } from "../utils/errors.js";

const schema = Type.Object({
  action: Type.Optional(Type.Union([
    Type.Literal("create"),
    Type.Literal("update"),
    Type.Literal("delete"),
  ], {
    description: 'Action to perform. Default: "create". Use "update" to modify an existing pivot, "delete" to remove one.',
  })),
  pivot_name: Type.Optional(Type.String({
    description: 'Name of existing pivot table. Required for "update" and "delete" actions. Use get_workbook_overview to find pivot names.',
  })),
  source_range: Type.Optional(Type.String({
    description: 'Source data range. Required for "create".',
  })),
  target_cell: Type.Optional(Type.String({
    description: 'Cell where to place the pivot table, e.g. "G1" or "PivotSheet!A1". Recommended: use a different sheet.',
  })),
  name: Type.Optional(Type.String({
    description: "Pivot table name. Auto-generated if omitted. For update, renames the existing pivot table.",
  })),
  rows: Type.Optional(Type.Array(Type.String(), {
    description: 'Column header names to use as row labels, e.g. ["Category", "Product"].',
  })),
  values: Type.Optional(Type.Array(Type.String(), {
    description: 'Column header names to summarize, e.g. ["Sales", "Quantity"].',
  })),
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
type PivotAction = NonNullable<Params["action"]>;

const VALID_AGGREGATIONS = ["sum", "count", "average", "max", "min"] as const;
type SupportedAggregation = (typeof VALID_AGGREGATIONS)[number];

let _aggregationFunctionMap: Record<SupportedAggregation, Excel.AggregationFunction> | null = null;
function getAggregationFunctionMap(): Record<SupportedAggregation, Excel.AggregationFunction> {
  if (!_aggregationFunctionMap) {
    _aggregationFunctionMap = {
      sum: Excel.AggregationFunction.sum,
      count: Excel.AggregationFunction.count,
      average: Excel.AggregationFunction.average,
      max: Excel.AggregationFunction.max,
      min: Excel.AggregationFunction.min,
    };
  }
  return _aggregationFunctionMap;
}

const DEFAULT_ACTION: PivotAction = "create";
const DEFAULT_AGGREGATION: SupportedAggregation = "sum";

interface PivotTableRunResult {
  pivotTableName: string;
  sheetName: string;
  targetAddress?: string;
  sourceAddress?: string;
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

function buildDetails(
  action: PivotAction,
  extra: Omit<CreatePivotTableDetails, "kind" | "action"> = {},
): CreatePivotTableDetails {
  return {
    kind: "create_pivot_table",
    action,
    ...extra,
  };
}

function resultText(text: string, details: CreatePivotTableDetails): AgentToolResult<CreatePivotTableDetails> {
  return {
    content: [{ type: "text", text }],
    details,
  };
}

function errorResult(
  action: PivotAction,
  text: string,
  extra: Omit<CreatePivotTableDetails, "kind" | "action"> = {},
): AgentToolResult<CreatePivotTableDetails> {
  return resultText(text, buildDetails(action, extra));
}

function resolveAggregationKey(rawValue: string): SupportedAggregation | null {
  const key = normalizeText(rawValue);
  return VALID_AGGREGATIONS.includes(key as SupportedAggregation)
    ? key as SupportedAggregation
    : null;
}

async function findPivotTableByName(
  context: Excel.RequestContext,
  pivotName: string,
): Promise<Excel.PivotTable> {
  const pivotTables = context.workbook.pivotTables;
  pivotTables.load("items/name,items/worksheet/name");
  await context.sync();

  const normalizedName = normalizeText(pivotName);
  const matches = pivotTables.items.filter((pivot) => normalizeText(pivot.name) === normalizedName);

  if (matches.length === 0) {
    throw new Error(`Pivot table "${pivotName}" not found. Use get_workbook_overview to list available objects.`);
  }

  if (matches.length > 1) {
    throw new Error(`Multiple pivot tables named "${pivotName}" were found. Use get_workbook_overview to identify the exact pivot.`);
  }

  return matches[0];
}

function formatFieldLine(label: string, values: string[] | undefined): string | undefined {
  if (!values) return undefined;
  return `- ${label}: ${values.length > 0 ? values.join(", ") : "(cleared)"}`;
}

export function createCreatePivotTableTool(): AgentTool<typeof schema, CreatePivotTableDetails> {
  return {
    name: "create_pivot_table",
    label: "Manage Pivot Table",
    description:
      "Create, update, or delete a native Excel PivotTable with row, column, value, and filter hierarchies.",
    parameters: schema,
    execute: async (
      _toolCallId: string,
      params: Params,
    ): Promise<AgentToolResult<CreatePivotTableDetails>> => {
      const action = params.action ?? DEFAULT_ACTION;
      const pivotName = params.pivot_name?.trim();

      try {
        if (params.agg_func !== undefined && !resolveAggregationKey(params.agg_func)) {
          return errorResult(
            action,
            `Error: invalid agg_func "${params.agg_func}". Valid values: ${VALID_AGGREGATIONS.join(", ")}.`,
            { pivotTableName: pivotName },
          );
        }

        if (action === "create") {
          const sourceRangeRef = params.source_range;
          if (!sourceRangeRef) {
            return errorResult(action, "Error: source_range is required for create.");
          }

          const targetCellRef = params.target_cell;
          if (!targetCellRef) {
            return errorResult(action, "Error: target_cell is required for create.");
          }

          const rowParams = params.rows;
          if (!rowParams || rowParams.length === 0) {
            return errorResult(action, "Error: rows must include at least one field name.");
          }

          const valueParams = params.values;
          if (!valueParams || valueParams.length === 0) {
            return errorResult(action, "Error: values must include at least one field name.");
          }

          const aggregationKey = resolveAggregationKey(params.agg_func ?? DEFAULT_AGGREGATION) ?? DEFAULT_AGGREGATION;
          const result = await excelRun<PivotTableRunResult>(async (context) => {
            const { sheet: sourceSheet, range: sourceRange } = getRange(context, sourceRangeRef);
            const { sheet: targetSheet, range: targetRange } = getRange(context, targetCellRef);
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

            const rowNames = dedupeNames(rowParams);
            const columnNames = dedupeNames(params.columns ?? []);
            const filterNames = dedupeNames(params.filters ?? []);
            const valueNames = dedupeNames(valueParams);

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
              dataHierarchy.summarizeBy = getAggregationFunctionMap()[aggregationKey];
            }

            return {
              pivotTableName,
              sheetName: targetSheet.name,
              targetAddress: qualifiedAddress(targetSheet.name, targetRange.address),
              sourceAddress: qualifiedAddress(sourceSheet.name, sourceRange.address),
            };
          });

          const lines = [
            `📊 Created pivot table **${result.pivotTableName}** on **${result.sheetName}** at \`${result.targetAddress ?? "position"}\`.`,
            `- Source: \`${result.sourceAddress ?? sourceRangeRef}\``,
            `- Rows: ${dedupeNames(rowParams).join(", ")}`,
            params.columns && params.columns.length > 0 ? `- Columns: ${dedupeNames(params.columns).join(", ")}` : undefined,
            `- Values: ${dedupeNames(valueParams).join(", ")} (${aggregationKey})`,
            params.filters && params.filters.length > 0 ? `- Filters: ${dedupeNames(params.filters).join(", ")}` : undefined,
          ].filter((line): line is string => typeof line === "string");

          return resultText(lines.join("\n"), buildDetails(action, {
            pivotTableName: result.pivotTableName,
            sheetName: result.sheetName,
            sourceAddress: result.sourceAddress,
          }));
        }

        if (action === "update") {
          if (!pivotName) {
            return errorResult(
              action,
              'Error: pivot_name is required for update. Use get_workbook_overview to find pivot names.',
            );
          }

          if (params.source_range !== undefined) {
            return errorResult(
              action,
              "Error: updating source_range for an existing pivot table is not supported by Office.js. Delete and recreate the pivot table with the new source_range.",
              { pivotTableName: pivotName },
            );
          }

          if (params.target_cell !== undefined) {
            return errorResult(
              action,
              "Error: moving an existing pivot table with target_cell is not supported by this tool. Delete and recreate the pivot table at the new location.",
              { pivotTableName: pivotName },
            );
          }

          const updatesRequested = params.rows !== undefined
            || params.columns !== undefined
            || params.filters !== undefined
            || params.values !== undefined
            || params.agg_func !== undefined
            || params.name !== undefined;

          if (!updatesRequested) {
            return errorResult(action, "Error: no update parameters were provided.", {
              pivotTableName: pivotName,
            });
          }

          const aggregationKey = params.agg_func ? resolveAggregationKey(params.agg_func) : null;
          const result = await excelRun<PivotTableRunResult>(async (context) => {
            const pivot = await findPivotTableByName(context, pivotName);
            pivot.load("name,worksheet/name");
            pivot.hierarchies.load("items/name");

            if (params.rows !== undefined) pivot.rowHierarchies.load("items/name");
            if (params.columns !== undefined) pivot.columnHierarchies.load("items/name");
            if (params.filters !== undefined) pivot.filterHierarchies.load("items/name");
            if (params.values !== undefined || params.agg_func !== undefined) {
              pivot.dataHierarchies.load("items/name,items/summarizeBy");
            }

            const workbookPivotTables = context.workbook.pivotTables;
            if (params.name && normalizeText(params.name) !== normalizeText(pivotName)) {
              workbookPivotTables.load("items/name");
            }

            await context.sync();

            const hierarchyLookup = buildHierarchyLookup(pivot.hierarchies.items);
            const availableFields = pivot.hierarchies.items.map((hierarchy) => hierarchy.name);
            const rowNames = params.rows !== undefined ? dedupeNames(params.rows) : undefined;
            const columnNames = params.columns !== undefined ? dedupeNames(params.columns) : undefined;
            const filterNames = params.filters !== undefined ? dedupeNames(params.filters) : undefined;
            const valueNames = params.values !== undefined ? dedupeNames(params.values) : undefined;

            const requestedNames = [
              ...(rowNames ?? []),
              ...(columnNames ?? []),
              ...(filterNames ?? []),
              ...(valueNames ?? []),
            ];
            const missingFields = requestedNames.filter((name) => !resolveHierarchy(name, hierarchyLookup));
            if (missingFields.length > 0) {
              throw new Error(
                `Unknown pivot field${missingFields.length === 1 ? "" : "s"}: ${missingFields.join(", ")}. Available fields: ${availableFields.join(", ")}.`,
              );
            }

            if (params.name && normalizeText(params.name) !== normalizeText(pivot.name)) {
              const nextName = params.name;
              const duplicate = workbookPivotTables.items.find((item) => normalizeText(item.name) === normalizeText(nextName));
              if (duplicate) {
                throw new Error(`A pivot table named "${params.name}" already exists.`);
              }
              pivot.name = params.name;
            }

            if (rowNames !== undefined) {
              for (let i = pivot.rowHierarchies.items.length - 1; i >= 0; i -= 1) {
                pivot.rowHierarchies.remove(pivot.rowHierarchies.items[i]);
              }
              for (const field of rowNames) {
                pivot.rowHierarchies.add(requireHierarchy(field, hierarchyLookup));
              }
            }

            if (columnNames !== undefined) {
              for (let i = pivot.columnHierarchies.items.length - 1; i >= 0; i -= 1) {
                pivot.columnHierarchies.remove(pivot.columnHierarchies.items[i]);
              }
              for (const field of columnNames) {
                pivot.columnHierarchies.add(requireHierarchy(field, hierarchyLookup));
              }
            }

            if (filterNames !== undefined) {
              for (let i = pivot.filterHierarchies.items.length - 1; i >= 0; i -= 1) {
                pivot.filterHierarchies.remove(pivot.filterHierarchies.items[i]);
              }
              for (const field of filterNames) {
                pivot.filterHierarchies.add(requireHierarchy(field, hierarchyLookup));
              }
            }

            if (valueNames !== undefined) {
              for (let i = pivot.dataHierarchies.items.length - 1; i >= 0; i -= 1) {
                pivot.dataHierarchies.remove(pivot.dataHierarchies.items[i]);
              }
              for (const field of valueNames) {
                const hierarchy = pivot.dataHierarchies.add(requireHierarchy(field, hierarchyLookup));
                if (aggregationKey) {
                  hierarchy.summarizeBy = getAggregationFunctionMap()[aggregationKey];
                }
              }
            } else if (aggregationKey) {
              for (const hierarchy of pivot.dataHierarchies.items) {
                hierarchy.summarizeBy = getAggregationFunctionMap()[aggregationKey];
              }
            }

            await context.sync();

            return {
              pivotTableName: params.name ?? pivot.name,
              sheetName: pivot.worksheet.name,
            };
          });

          const lines = [
            `📊 Updated pivot table **${result.pivotTableName}** on **${result.sheetName}**.`,
            params.name !== undefined ? `- Name: ${result.pivotTableName}` : undefined,
            formatFieldLine("Rows", params.rows !== undefined ? dedupeNames(params.rows) : undefined),
            formatFieldLine("Columns", params.columns !== undefined ? dedupeNames(params.columns) : undefined),
            formatFieldLine("Filters", params.filters !== undefined ? dedupeNames(params.filters) : undefined),
            params.values !== undefined
              ? `- Values: ${dedupeNames(params.values).length > 0 ? dedupeNames(params.values).join(", ") : "(cleared)"}${aggregationKey ? ` (${aggregationKey})` : ""}`
              : aggregationKey
                ? `- Aggregation: ${aggregationKey}`
                : undefined,
          ].filter((line): line is string => typeof line === "string");

          return resultText(lines.join("\n"), buildDetails(action, {
            pivotTableName: result.pivotTableName,
            sheetName: result.sheetName,
          }));
        }

        if (!pivotName) {
          return errorResult(
            action,
            'Error: pivot_name is required for delete. Use get_workbook_overview to find pivot names.',
          );
        }

        const result = await excelRun<PivotTableRunResult>(async (context) => {
          const pivot = await findPivotTableByName(context, pivotName);
          pivot.load("name,worksheet/name");
          await context.sync();

          const resolvedPivotName = pivot.name;
          const sheetName = pivot.worksheet.name;
          pivot.delete();
          await context.sync();

          return {
            pivotTableName: resolvedPivotName,
            sheetName,
          };
        });

        return resultText(
          `📊 Deleted pivot table **${result.pivotTableName}** from **${result.sheetName}**.`,
          buildDetails(action, {
            pivotTableName: result.pivotTableName,
            sheetName: result.sheetName,
          }),
        );
      } catch (e: unknown) {
        const verb = action === "create" ? "creating" : action === "update" ? "updating" : "deleting";
        return errorResult(action, `Error ${verb} pivot table: ${getErrorMessage(e)}`, {
          pivotTableName: pivotName,
        });
      }
    },
  };
}
