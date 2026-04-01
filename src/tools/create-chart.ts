import type { AgentTool, AgentToolResult } from "@mariozechner/pi-agent-core";
import { Type, type Static } from "@sinclair/typebox";
import type { CreateChartDetails } from "./tool-details.js";
import { excelRun, getRange, qualifiedAddress } from "../excel/helpers.js";
import { getErrorMessage } from "../utils/errors.js";

const schema = Type.Object({
  action: Type.Optional(Type.Union([
    Type.Literal("create"),
    Type.Literal("update"),
    Type.Literal("delete"),
  ], {
    description: 'Action to perform. Default: "create". Use "update" to modify an existing chart, "delete" to remove one.',
  })),
  chart_name: Type.Optional(Type.String({
    description: 'Name or ID of existing chart. Required for "update" and "delete" actions. Use get_workbook_overview to find chart names.',
  })),
  data_range: Type.Optional(Type.String({
    description: 'Data range for the chart. Required for "create", optional for "update".',
  })),
  chart_type: Type.Optional(Type.String({
    description: 'Chart type: "column", "bar", "line", "pie", "scatter", "area", "doughnut", "radar". Required for "create", optional for "update".',
  })),
  target_cell: Type.Optional(Type.String({
    description: 'Cell where to place the chart, e.g. "F1". If omitted, placed below the data.',
  })),
  title: Type.Optional(Type.String({
    description: "Chart title text.",
  })),
  x_axis: Type.Optional(Type.String({
    description: "X-axis (category axis) label text.",
  })),
  y_axis: Type.Optional(Type.String({
    description: "Y-axis (value axis) label text.",
  })),
  series_by: Type.Optional(Type.String({
    description: '"auto", "rows", or "columns". How data series are arranged. Default: "auto".',
  })),
  width: Type.Optional(Type.Number({
    description: "Chart width in pixels. Default: 450.",
  })),
  height: Type.Optional(Type.Number({
    description: "Chart height in pixels. Default: 300.",
  })),
  show_legend: Type.Optional(Type.Boolean({
    description: "Show chart legend. Default: true.",
  })),
  legend_position: Type.Optional(Type.String({
    description: 'Legend position: "top", "bottom", "left", "right". Default: "right".',
  })),
  show_data_labels: Type.Optional(Type.Boolean({
    description: "Show data labels on chart points. Default: false.",
  })),
});

type Params = Static<typeof schema>;
type ChartAction = NonNullable<Params["action"]>;

const VALID_CHART_TYPES = ["column", "bar", "line", "pie", "scatter", "area", "doughnut", "radar"] as const;
type SupportedChartType = (typeof VALID_CHART_TYPES)[number];

const CHART_TYPE_MAP: Record<SupportedChartType, Excel.ChartType> = {
  column: Excel.ChartType.columnClustered,
  bar: Excel.ChartType.barClustered,
  line: Excel.ChartType.line,
  pie: Excel.ChartType.pie,
  scatter: Excel.ChartType.xyscatter,
  area: Excel.ChartType.area,
  doughnut: Excel.ChartType.doughnut,
  radar: Excel.ChartType.radar,
};

const VALID_SERIES_BY = ["auto", "rows", "columns"] as const;
type SupportedSeriesBy = (typeof VALID_SERIES_BY)[number];

const SERIES_BY_MAP: Record<SupportedSeriesBy, Excel.ChartSeriesBy> = {
  auto: Excel.ChartSeriesBy.auto,
  rows: Excel.ChartSeriesBy.rows,
  columns: Excel.ChartSeriesBy.columns,
};

const VALID_LEGEND_POSITIONS = ["top", "bottom", "left", "right"] as const;
type SupportedLegendPosition = (typeof VALID_LEGEND_POSITIONS)[number];

const LEGEND_POSITION_MAP: Record<SupportedLegendPosition, Excel.ChartLegendPosition> = {
  top: Excel.ChartLegendPosition.top,
  bottom: Excel.ChartLegendPosition.bottom,
  left: Excel.ChartLegendPosition.left,
  right: Excel.ChartLegendPosition.right,
};

const AXISLESS_CHART_TYPES = new Set<SupportedChartType>(["pie", "doughnut"]);

const DEFAULT_ACTION: ChartAction = "create";
const DEFAULT_SERIES_BY: SupportedSeriesBy = "auto";
const DEFAULT_LEGEND_POSITION: SupportedLegendPosition = "right";
const DEFAULT_WIDTH = 450;
const DEFAULT_HEIGHT = 300;

interface CreateChartRunResult {
  sheetName: string;
  chartName: string;
  chartAddress?: string;
  sourceAddress?: string;
  chartType?: string;
  legendVisible?: boolean;
  legendPosition?: string;
  axisTitlesSkipped?: boolean;
}

function normalizeText(value: string): string {
  return value.trim().toLowerCase();
}

function buildDetails(
  action: ChartAction,
  extra: Omit<CreateChartDetails, "kind" | "action"> = {},
): CreateChartDetails {
  return {
    kind: "create_chart",
    action,
    ...extra,
  };
}

function resultText(text: string, details: CreateChartDetails): AgentToolResult<CreateChartDetails> {
  return {
    content: [{ type: "text", text }],
    details,
  };
}

function errorResult(
  action: ChartAction,
  text: string,
  extra: Omit<CreateChartDetails, "kind" | "action"> = {},
): AgentToolResult<CreateChartDetails> {
  return resultText(text, buildDetails(action, extra));
}

function resolveChartTypeKey(rawValue: string): SupportedChartType | null {
  const key = normalizeText(rawValue);
  return VALID_CHART_TYPES.includes(key as SupportedChartType)
    ? key as SupportedChartType
    : null;
}

function resolveSeriesByKey(rawValue: string): SupportedSeriesBy | null {
  const key = normalizeText(rawValue);
  return VALID_SERIES_BY.includes(key as SupportedSeriesBy)
    ? key as SupportedSeriesBy
    : null;
}

function resolveLegendPositionKey(rawValue: string): SupportedLegendPosition | null {
  const key = normalizeText(rawValue);
  return VALID_LEGEND_POSITIONS.includes(key as SupportedLegendPosition)
    ? key as SupportedLegendPosition
    : null;
}

function inferSupportedChartType(value: Excel.ChartType | string | undefined): SupportedChartType | null {
  const key = normalizeText(String(value ?? ""));
  if (key.includes("doughnut")) return "doughnut";
  if (key.includes("scatter")) return "scatter";
  if (key.includes("column")) return "column";
  if (key.includes("line")) return "line";
  if (key.includes("pie")) return "pie";
  if (key.includes("area")) return "area";
  if (key.includes("radar")) return "radar";
  if (key.includes("bar")) return "bar";
  return null;
}

function positiveNumberError(label: string, value: number | undefined): string | null {
  if (value === undefined) return null;
  if (!Number.isFinite(value) || value <= 0) {
    return `Error: ${label} must be a positive number.`;
  }
  return null;
}

function hasChartTitleText(value: string): boolean {
  return value.trim().length > 0;
}

async function findChartByName(
  context: Excel.RequestContext,
  chartName: string,
): Promise<{ sheet: Excel.Worksheet; chart: Excel.Chart }> {
  const worksheets = context.workbook.worksheets;
  worksheets.load("items/name");
  await context.sync();

  for (const sheet of worksheets.items) {
    sheet.charts.load("items/name");
  }
  await context.sync();

  const normalizedName = normalizeText(chartName);
  const matches: Array<{ sheet: Excel.Worksheet; chart: Excel.Chart }> = [];

  for (const sheet of worksheets.items) {
    for (const chart of sheet.charts.items) {
      if (normalizeText(chart.name) === normalizedName) {
        matches.push({ sheet, chart });
      }
    }
  }

  if (matches.length === 0) {
    throw new Error(`Chart "${chartName}" not found. Use get_workbook_overview to list available objects.`);
  }

  if (matches.length > 1) {
    throw new Error(`Multiple charts named "${chartName}" were found. Use get_workbook_overview to identify the exact chart.`);
  }

  return matches[0];
}

export function createCreateChartTool(): AgentTool<typeof schema, CreateChartDetails> {
  return {
    name: "create_chart",
    label: "Manage Chart",
    description:
      "Create, update, or delete an Excel chart with configurable chart type, data range, title, axis labels, legend, data labels, and placement.",
    parameters: schema,
    execute: async (
      _toolCallId: string,
      params: Params,
    ): Promise<AgentToolResult<CreateChartDetails>> => {
      const action = params.action ?? DEFAULT_ACTION;
      const chartName = params.chart_name?.trim();

      try {
        const widthError = positiveNumberError("width", params.width);
        if (widthError) {
          return errorResult(action, widthError, { chartName });
        }

        const heightError = positiveNumberError("height", params.height);
        if (heightError) {
          return errorResult(action, heightError, { chartName });
        }

        if (params.series_by !== undefined && !resolveSeriesByKey(params.series_by)) {
          return errorResult(
            action,
            `Error: invalid series_by "${params.series_by}". Valid values: ${VALID_SERIES_BY.join(", ")}.`,
            { chartName },
          );
        }

        if (params.legend_position !== undefined && !resolveLegendPositionKey(params.legend_position)) {
          return errorResult(
            action,
            `Error: invalid legend_position "${params.legend_position}". Valid positions: ${VALID_LEGEND_POSITIONS.join(", ")}.`,
            { chartName },
          );
        }

        if (params.chart_type !== undefined && !resolveChartTypeKey(params.chart_type)) {
          return errorResult(
            action,
            `Error: invalid chart_type "${params.chart_type}". Valid types: ${VALID_CHART_TYPES.join(", ")}.`,
            { chartName },
          );
        }

        if (action === "create") {
          const dataRangeRef = params.data_range;
          if (!dataRangeRef) {
            return errorResult(action, "Error: data_range is required for create.");
          }

          if (!params.chart_type) {
            return errorResult(action, "Error: chart_type is required for create.");
          }

          const chartTypeKey = resolveChartTypeKey(params.chart_type);
          if (!chartTypeKey) {
            return errorResult(
              action,
              `Error: invalid chart_type "${params.chart_type}". Valid types: ${VALID_CHART_TYPES.join(", ")}.`,
            );
          }
          const seriesByKey = resolveSeriesByKey(params.series_by ?? DEFAULT_SERIES_BY) ?? DEFAULT_SERIES_BY;
          const legendPositionKey = resolveLegendPositionKey(params.legend_position ?? DEFAULT_LEGEND_POSITION) ?? DEFAULT_LEGEND_POSITION;
          const width = params.width ?? DEFAULT_WIDTH;
          const height = params.height ?? DEFAULT_HEIGHT;

          const result = await excelRun<CreateChartRunResult>(async (context) => {
            const { sheet: sourceSheet, range: sourceRange } = getRange(context, dataRangeRef);
            sourceSheet.load("name");
            sourceRange.load("address,rowIndex,columnIndex,rowCount,columnCount");
            await context.sync();

            let chartSheet = sourceSheet;
            let targetRange: Excel.Range;

            if (params.target_cell) {
              const targetRef = params.target_cell.includes("!")
                ? params.target_cell
                : qualifiedAddress(sourceSheet.name, params.target_cell);
              const target = getRange(context, targetRef);
              chartSheet = target.sheet;
              chartSheet.load("name");
              targetRange = target.range;
              targetRange.load("address");
            } else {
              targetRange = sourceSheet.getCell(
                sourceRange.rowIndex + sourceRange.rowCount + 1,
                sourceRange.columnIndex,
              );
              targetRange.load("address");
            }

            await context.sync();

            if (sourceRange.rowCount < 2) {
              throw new Error("data_range must include a header row and at least one data row.");
            }

            if (sourceRange.columnCount < 2) {
              throw new Error("data_range must include at least two columns.");
            }

            if (chartTypeKey === "scatter" && sourceRange.columnCount < 2) {
              throw new Error("scatter charts require at least two columns of source data.");
            }

            const chart = chartSheet.charts.add(
              CHART_TYPE_MAP[chartTypeKey],
              sourceRange,
              SERIES_BY_MAP[seriesByKey],
            );

            chart.setPosition(targetRange);
            chart.width = width;
            chart.height = height;
            chart.legend.visible = params.show_legend ?? true;
            chart.legend.position = LEGEND_POSITION_MAP[legendPositionKey];
            chart.dataLabels.showValue = params.show_data_labels ?? false;

            if (params.title !== undefined) {
              chart.title.text = params.title;
              chart.title.visible = hasChartTitleText(params.title);
            }

            let axisTitlesSkipped = false;
            if (!AXISLESS_CHART_TYPES.has(chartTypeKey)) {
              if (params.x_axis !== undefined) {
                chart.axes.categoryAxis.title.text = params.x_axis;
                chart.axes.categoryAxis.title.visible = hasChartTitleText(params.x_axis);
              }
              if (params.y_axis !== undefined) {
                chart.axes.valueAxis.title.text = params.y_axis;
                chart.axes.valueAxis.title.visible = hasChartTitleText(params.y_axis);
              }
            } else if (params.x_axis !== undefined || params.y_axis !== undefined) {
              axisTitlesSkipped = true;
            }

            chart.load("name");
            await context.sync();

            return {
              sheetName: chartSheet.name,
              sourceAddress: qualifiedAddress(sourceSheet.name, sourceRange.address),
              chartName: chart.name,
              chartAddress: qualifiedAddress(chartSheet.name, targetRange.address),
              chartType: chartTypeKey ?? undefined,
              legendVisible: params.show_legend ?? true,
              legendPosition: legendPositionKey,
              axisTitlesSkipped,
            };
          });

          const lines = [
            `📊 Created ${result.chartType ?? "chart"} chart **${result.chartName}** on **${result.sheetName}** at \`${result.chartAddress ?? "position"}\`.`,
            result.sourceAddress ? `- Source: \`${result.sourceAddress}\`` : undefined,
            params.title !== undefined ? `- Title: ${params.title || "(cleared)"}` : undefined,
            params.x_axis !== undefined && !result.axisTitlesSkipped ? `- X-axis: ${params.x_axis || "(cleared)"}` : undefined,
            params.y_axis !== undefined && !result.axisTitlesSkipped ? `- Y-axis: ${params.y_axis || "(cleared)"}` : undefined,
            result.legendVisible !== undefined
              ? `- Legend: ${result.legendVisible ? (result.legendPosition ? `shown (${result.legendPosition})` : "shown") : "hidden"}`
              : undefined,
            params.show_data_labels !== undefined ? `- Data labels: ${params.show_data_labels ? "shown" : "hidden"}` : undefined,
            result.axisTitlesSkipped ? "- Axis titles: skipped because pie and doughnut charts do not support axis titles." : undefined,
          ].filter((line): line is string => typeof line === "string");

          return resultText(lines.join("\n"), buildDetails(action, {
            chartType: result.chartType,
            sheetName: result.sheetName,
            chartName: result.chartName,
          }));
        }

        if (action === "update") {
          if (!chartName) {
            return errorResult(
              action,
              'Error: chart_name is required for update. Use get_workbook_overview to find chart names.',
            );
          }

          if (params.series_by !== undefined && !params.data_range) {
            return errorResult(
              action,
              "Error: series_by can only be used with data_range when updating a chart.",
              { chartName },
            );
          }

          if (
            params.data_range === undefined
            && params.chart_type === undefined
            && params.target_cell === undefined
            && params.title === undefined
            && params.x_axis === undefined
            && params.y_axis === undefined
            && params.width === undefined
            && params.height === undefined
            && params.show_legend === undefined
            && params.legend_position === undefined
            && params.show_data_labels === undefined
          ) {
            return errorResult(action, "Error: no update parameters were provided.", { chartName });
          }

          const chartTypeKey = params.chart_type ? resolveChartTypeKey(params.chart_type) : null;
          const legendPositionKey = params.legend_position ? resolveLegendPositionKey(params.legend_position) : null;
          const updateResult = await excelRun<CreateChartRunResult>(async (context) => {
            const located = await findChartByName(context, chartName);
            const chartSheet = located.sheet;
            const chart = located.chart;
            chartSheet.load("name");
            chart.load("name,chartType");
            await context.sync();

            if (chartTypeKey) {
              chart.chartType = CHART_TYPE_MAP[chartTypeKey];
            }

            let sourceAddress: string | undefined;
            if (params.data_range) {
              const { sheet: sourceSheet, range: sourceRange } = getRange(context, params.data_range);
              sourceSheet.load("name");
              sourceRange.load("address,rowCount,columnCount");
              await context.sync();

              if (sourceRange.rowCount < 2) {
                throw new Error("data_range must include a header row and at least one data row.");
              }

              if (sourceRange.columnCount < 2) {
                throw new Error("data_range must include at least two columns.");
              }

              if ((chartTypeKey ?? inferSupportedChartType(chart.chartType)) === "scatter" && sourceRange.columnCount < 2) {
                throw new Error("scatter charts require at least two columns of source data.");
              }

              const seriesByKey = resolveSeriesByKey(params.series_by ?? DEFAULT_SERIES_BY) ?? DEFAULT_SERIES_BY;
              chart.setData(sourceRange, SERIES_BY_MAP[seriesByKey]);
              sourceAddress = qualifiedAddress(sourceSheet.name, sourceRange.address);
            }

            let chartAddress: string | undefined;
            if (params.target_cell) {
              const targetRef = params.target_cell.includes("!")
                ? params.target_cell
                : qualifiedAddress(chartSheet.name, params.target_cell);
              const target = getRange(context, targetRef);
              target.sheet.load("name");
              target.range.load("address");
              await context.sync();

              if (normalizeText(target.sheet.name) !== normalizeText(chartSheet.name)) {
                throw new Error(`target_cell must be on the chart worksheet (${chartSheet.name}) when updating a chart.`);
              }

              chart.setPosition(target.range);
              chartAddress = qualifiedAddress(chartSheet.name, target.range.address);
            }

            if (params.title !== undefined) {
              chart.title.text = params.title;
              chart.title.visible = hasChartTitleText(params.title);
            }

            const effectiveChartType = chartTypeKey ?? inferSupportedChartType(chart.chartType);
            let axisTitlesSkipped = false;
            if (!effectiveChartType || !AXISLESS_CHART_TYPES.has(effectiveChartType)) {
              if (params.x_axis !== undefined) {
                chart.axes.categoryAxis.title.text = params.x_axis;
                chart.axes.categoryAxis.title.visible = hasChartTitleText(params.x_axis);
              }
              if (params.y_axis !== undefined) {
                chart.axes.valueAxis.title.text = params.y_axis;
                chart.axes.valueAxis.title.visible = hasChartTitleText(params.y_axis);
              }
            } else if (params.x_axis !== undefined || params.y_axis !== undefined) {
              axisTitlesSkipped = true;
            }

            if (params.width !== undefined) chart.width = params.width;
            if (params.height !== undefined) chart.height = params.height;
            if (params.show_legend !== undefined) chart.legend.visible = params.show_legend;
            if (legendPositionKey) chart.legend.position = LEGEND_POSITION_MAP[legendPositionKey];
            if (params.show_data_labels !== undefined) chart.dataLabels.showValue = params.show_data_labels;

            await context.sync();

            return {
              sheetName: chartSheet.name,
              chartName: chart.name,
              chartAddress,
              sourceAddress,
              chartType: chartTypeKey ?? effectiveChartType ?? String(chart.chartType),
              legendVisible: params.show_legend,
              legendPosition: legendPositionKey ?? undefined,
              axisTitlesSkipped,
            };
          });

          const lines = [
            `📊 Updated chart **${updateResult.chartName}** on **${updateResult.sheetName}**.`,
            updateResult.chartType && params.chart_type !== undefined ? `- Type: ${updateResult.chartType}` : undefined,
            updateResult.sourceAddress ? `- Source: \`${updateResult.sourceAddress}\`` : undefined,
            updateResult.chartAddress ? `- Position: \`${updateResult.chartAddress}\`` : undefined,
            params.title !== undefined ? `- Title: ${params.title || "(cleared)"}` : undefined,
            params.x_axis !== undefined && !updateResult.axisTitlesSkipped ? `- X-axis: ${params.x_axis || "(cleared)"}` : undefined,
            params.y_axis !== undefined && !updateResult.axisTitlesSkipped ? `- Y-axis: ${params.y_axis || "(cleared)"}` : undefined,
            params.width !== undefined ? `- Width: ${params.width}px` : undefined,
            params.height !== undefined ? `- Height: ${params.height}px` : undefined,
            params.show_legend !== undefined
              ? `- Legend: ${params.show_legend ? `shown${updateResult.legendPosition ? ` (${updateResult.legendPosition})` : ""}` : "hidden"}`
              : undefined,
            params.legend_position !== undefined && params.show_legend === undefined
              ? `- Legend position: ${params.legend_position}`
              : undefined,
            params.show_data_labels !== undefined ? `- Data labels: ${params.show_data_labels ? "shown" : "hidden"}` : undefined,
            updateResult.axisTitlesSkipped ? "- Axis titles: skipped because pie and doughnut charts do not support axis titles." : undefined,
          ].filter((line): line is string => typeof line === "string");

          return resultText(lines.join("\n"), buildDetails(action, {
            chartType: params.chart_type ?? updateResult.chartType,
            sheetName: updateResult.sheetName,
            chartName: updateResult.chartName,
          }));
        }

        if (!chartName) {
          return errorResult(
            action,
            'Error: chart_name is required for delete. Use get_workbook_overview to find chart names.',
          );
        }

        const deleteResult = await excelRun<CreateChartRunResult>(async (context) => {
          const located = await findChartByName(context, chartName);
          located.sheet.load("name");
          located.chart.load("name");
          await context.sync();

          const sheetName = located.sheet.name;
          const resolvedChartName = located.chart.name;
          located.chart.delete();
          await context.sync();

          return {
            sheetName,
            chartName: resolvedChartName,
          };
        });

        return resultText(
          `📊 Deleted chart **${deleteResult.chartName}** from **${deleteResult.sheetName}**.`,
          buildDetails(action, {
            sheetName: deleteResult.sheetName,
            chartName: deleteResult.chartName,
          }),
        );
      } catch (e: unknown) {
        const verb = action === "create" ? "creating" : action === "update" ? "updating" : "deleting";
        return errorResult(action, `Error ${verb} chart: ${getErrorMessage(e)}`, { chartName });
      }
    },
  };
}
