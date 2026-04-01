import type { AgentTool, AgentToolResult } from "@mariozechner/pi-agent-core";
import { Type, type Static } from "@sinclair/typebox";
import type { CreateChartDetails } from "./tool-details.js";
import { excelRun, getRange, qualifiedAddress } from "../excel/helpers.js";
import { getErrorMessage } from "../utils/errors.js";

const schema = Type.Object({
  data_range: Type.String({
    description: 'Data range for the chart, e.g. "A1:D10" or "Sheet1!A1:D10". Must include headers in the first row.',
  }),
  chart_type: Type.String({
    description: 'Chart type: "column", "bar", "line", "pie", "scatter", "area", "doughnut", "radar". Default: "column".',
  }),
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

const DEFAULT_CHART_TYPE: SupportedChartType = "column";
const DEFAULT_SERIES_BY: SupportedSeriesBy = "auto";
const DEFAULT_LEGEND_POSITION: SupportedLegendPosition = "right";
const DEFAULT_WIDTH = 450;
const DEFAULT_HEIGHT = 300;

interface CreateChartRunResult {
  sheetName: string;
  sourceAddress: string;
  chartName: string;
  chartAddress: string;
}

function normalizeText(value: string): string {
  return value.trim().toLowerCase();
}

export function createCreateChartTool(): AgentTool<typeof schema, CreateChartDetails> {
  return {
    name: "create_chart",
    label: "Create Chart",
    description:
      "Create an Excel chart from a data range with configurable chart type, title, axis labels, legend, data labels, and placement.",
    parameters: schema,
    execute: async (
      _toolCallId: string,
      params: Params,
    ): Promise<AgentToolResult<CreateChartDetails>> => {
      try {
        const rawChartType = params.chart_type;
        const chartTypeKey = normalizeText(params.chart_type || DEFAULT_CHART_TYPE);
        if (!VALID_CHART_TYPES.includes(chartTypeKey as SupportedChartType)) {
          return {
            content: [{
              type: "text",
              text: `Error: invalid chart_type "${rawChartType ?? ""}". Valid types: ${VALID_CHART_TYPES.join(", ")}.`,
            }],
            details: { kind: "create_chart" },
          };
        }

        const rawSeriesBy = params.series_by;
        const seriesByKey = normalizeText(params.series_by || DEFAULT_SERIES_BY);
        if (!VALID_SERIES_BY.includes(seriesByKey as SupportedSeriesBy)) {
          return {
            content: [{
              type: "text",
              text: `Error: invalid series_by "${rawSeriesBy ?? ""}". Valid values: ${VALID_SERIES_BY.join(", ")}.`,
            }],
            details: { kind: "create_chart" },
          };
        }

        const rawLegendPosition = params.legend_position;
        const legendPositionKey = normalizeText(params.legend_position || DEFAULT_LEGEND_POSITION);
        if (!VALID_LEGEND_POSITIONS.includes(legendPositionKey as SupportedLegendPosition)) {
          return {
            content: [{
              type: "text",
              text: `Error: invalid legend_position "${rawLegendPosition ?? ""}". Valid positions: ${VALID_LEGEND_POSITIONS.join(", ")}.`,
            }],
            details: { kind: "create_chart" },
          };
        }

        const width = params.width ?? DEFAULT_WIDTH;
        const height = params.height ?? DEFAULT_HEIGHT;

        if (!Number.isFinite(width) || width <= 0) {
          return {
            content: [{ type: "text", text: "Error: width must be a positive number." }],
            details: { kind: "create_chart" },
          };
        }

        if (!Number.isFinite(height) || height <= 0) {
          return {
            content: [{ type: "text", text: "Error: height must be a positive number." }],
            details: { kind: "create_chart" },
          };
        }

        const result = await excelRun<CreateChartRunResult>(async (context) => {
          const { sheet: sourceSheet, range: sourceRange } = getRange(context, params.data_range);
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

          if ((chartTypeKey as SupportedChartType) === "scatter" && sourceRange.columnCount < 2) {
            throw new Error("scatter charts require at least two columns of source data.");
          }

          const chart = chartSheet.charts.add(
            CHART_TYPE_MAP[chartTypeKey as SupportedChartType],
            sourceRange,
            SERIES_BY_MAP[seriesByKey as SupportedSeriesBy],
          );

          chart.setPosition(targetRange);
          chart.width = width;
          chart.height = height;
          chart.legend.visible = params.show_legend ?? true;
          chart.legend.position = LEGEND_POSITION_MAP[legendPositionKey as SupportedLegendPosition];

          if (params.title) {
            chart.title.text = params.title;
            chart.title.visible = true;
          }

          if (params.show_data_labels) {
            chart.dataLabels.showValue = true;
          }

          if (!AXISLESS_CHART_TYPES.has(chartTypeKey as SupportedChartType)) {
            if (params.x_axis) {
              chart.axes.categoryAxis.title.text = params.x_axis;
              chart.axes.categoryAxis.title.visible = true;
            }
            if (params.y_axis) {
              chart.axes.valueAxis.title.text = params.y_axis;
              chart.axes.valueAxis.title.visible = true;
            }
          }

          chart.load("name");
          await context.sync();

          return {
            sheetName: chartSheet.name,
            sourceAddress: qualifiedAddress(sourceSheet.name, sourceRange.address),
            chartName: chart.name,
            chartAddress: qualifiedAddress(chartSheet.name, targetRange.address),
          };
        });

        const legendVisible = params.show_legend ?? true;
        const lines = [
          `📊 Created ${chartTypeKey} chart **${result.chartName}** on **${result.sheetName}** at \`${result.chartAddress}\`.`,
          `- Source: \`${result.sourceAddress}\``,
          params.title ? `- Title: ${params.title}` : undefined,
          params.x_axis ? `- X-axis: ${params.x_axis}` : undefined,
          params.y_axis ? `- Y-axis: ${params.y_axis}` : undefined,
          `- Legend: ${legendVisible ? `shown (${legendPositionKey})` : "hidden"}`,
          params.show_data_labels ? "- Data labels: shown" : undefined,
        ].filter((line): line is string => typeof line === "string");

        return {
          content: [{ type: "text", text: lines.join("\n") }],
          details: {
            kind: "create_chart",
            chartType: chartTypeKey,
            sheetName: result.sheetName,
            chartName: result.chartName,
          },
        };
      } catch (e: unknown) {
        return {
          content: [{ type: "text", text: `Error creating chart: ${getErrorMessage(e)}` }],
          details: { kind: "create_chart" },
        };
      }
    },
  };
}
