import type { AgentTool, AgentToolResult } from "@mariozechner/pi-agent-core";
import { Type, type Static } from "@sinclair/typebox";
import { excelRun, getRange, parseRangeRef, qualifiedAddress } from "../excel/helpers.js";
import { getErrorMessage } from "../utils/errors.js";
import type { RangeOperationsDetails } from "./tool-details.js";

const schema = Type.Object({
  action: Type.Union([
    Type.Literal("copy"),
    Type.Literal("delete"),
    Type.Literal("merge"),
    Type.Literal("unmerge"),
  ], {
    description: 'Operation to perform: "copy", "delete", "merge", or "unmerge".',
  }),
  range: Type.String({
    description: 'Source/target range, e.g. "A1:D5" or "Sheet1!A1:D5".',
  }),
  target_cell: Type.Optional(Type.String({
    description: 'Target cell for "copy" action, e.g. "F1" or "Sheet2!A1". Top-left corner of the paste area.',
  })),
  copy_type: Type.Optional(Type.String({
    description: '"all" (values+formatting), "values" (values only), "formats" (formatting only), "formulas" (formulas+formatting). Default: "all".',
  })),
  shift_direction: Type.Optional(Type.String({
    description: 'For "delete" action: "up" or "left". How remaining cells shift. Default: "up".',
  })),
  merge_across: Type.Optional(Type.Boolean({
    description: "For \"merge\" action: if true, merge each row separately instead of the entire range. Default: false.",
  })),
});

type Params = Static<typeof schema>;

const COPY_TYPE_MAP = {
  all: Excel.RangeCopyType.all,
  values: Excel.RangeCopyType.values,
  formats: Excel.RangeCopyType.formats,
  formulas: Excel.RangeCopyType.formulas,
} as const;

type SupportedCopyType = keyof typeof COPY_TYPE_MAP;

const DELETE_SHIFT_MAP = {
  up: Excel.DeleteShiftDirection.up,
  left: Excel.DeleteShiftDirection.left,
} as const;

type SupportedDeleteShift = keyof typeof DELETE_SHIFT_MAP;

interface CopyRangeResult {
  sourceRange: string;
  targetRange: string;
  sheetName: string;
  copyType: SupportedCopyType;
}

interface SingleRangeResult {
  range: string;
  sheetName: string;
}

function normalizeText(value: string): string {
  return value.trim().toLowerCase();
}

function normalizeCopyType(value: string | undefined): SupportedCopyType | null {
  const normalized = normalizeText(value ?? "all");
  return normalized in COPY_TYPE_MAP ? normalized as SupportedCopyType : null;
}

function normalizeShiftDirection(value: string | undefined): SupportedDeleteShift | null {
  const normalized = normalizeText(value ?? "up");
  return normalized in DELETE_SHIFT_MAP ? normalized as SupportedDeleteShift : null;
}

function ensureSingleCell(address: string, paramName: string): void {
  const parsed = parseRangeRef(address);
  if (parsed.address.includes(":") || parsed.address.includes(",") || parsed.address.includes(";")) {
    throw new Error(`${paramName} must be a single cell, e.g. "F1" or "Sheet2!A1".`);
  }
}

async function copyRange(params: Params): Promise<AgentToolResult<RangeOperationsDetails>> {
  const targetCell = params.target_cell;
  if (!targetCell) {
    return {
      content: [{ type: "text", text: "Error: target_cell is required for action=\"copy\"." }],
      details: {
        kind: "range_operations",
        action: "copy",
      },
    };
  }

  ensureSingleCell(targetCell, "target_cell");

  const copyType = normalizeCopyType(params.copy_type);
  if (!copyType) {
    return {
      content: [{ type: "text", text: `Error: invalid copy_type "${params.copy_type ?? ""}". Valid values: ${Object.keys(COPY_TYPE_MAP).join(", ")}.` }],
      details: {
        kind: "range_operations",
        action: "copy",
      },
    };
  }

  const result = await excelRun<CopyRangeResult>(async (context) => {
    const { sheet: sourceSheet, range: sourceRange } = getRange(context, params.range);
    sourceSheet.load("name");
    sourceRange.load("address,rowCount,columnCount");
    await context.sync();

    const parsedTarget = parseRangeRef(targetCell);
    const targetRef = parsedTarget.sheet
      ? targetCell
      : qualifiedAddress(sourceSheet.name, parsedTarget.address);
    const { sheet: targetSheet, range: targetAnchor } = getRange(context, targetRef);
    targetSheet.load("name");
    targetAnchor.load("address");
    const targetRange = targetAnchor.getResizedRange(sourceRange.rowCount - 1, sourceRange.columnCount - 1);
    targetRange.load("address");
    await context.sync();

    targetAnchor.copyFrom(sourceRange, COPY_TYPE_MAP[copyType]);
    await context.sync();

    return {
      sourceRange: qualifiedAddress(sourceSheet.name, sourceRange.address),
      targetRange: qualifiedAddress(targetSheet.name, targetRange.address),
      sheetName: targetSheet.name,
      copyType,
    };
  });

  return {
    content: [{ type: "text", text: `Copied **${result.sourceRange}** to **${result.targetRange}** (${result.copyType}).` }],
    details: {
      kind: "range_operations",
      action: "copy",
      sourceRange: result.sourceRange,
      targetRange: result.targetRange,
      sheetName: result.sheetName,
    },
  };
}

async function deleteRange(params: Params): Promise<AgentToolResult<RangeOperationsDetails>> {
  const shiftDirection = normalizeShiftDirection(params.shift_direction);
  if (!shiftDirection) {
    return {
      content: [{ type: "text", text: `Error: invalid shift_direction "${params.shift_direction ?? ""}". Valid values: ${Object.keys(DELETE_SHIFT_MAP).join(", ")}.` }],
      details: {
        kind: "range_operations",
        action: "delete",
      },
    };
  }

  const result = await excelRun<SingleRangeResult>(async (context) => {
    const { sheet, range } = getRange(context, params.range);
    sheet.load("name");
    range.load("address");
    await context.sync();

    const fullRange = qualifiedAddress(sheet.name, range.address);
    range.delete(DELETE_SHIFT_MAP[shiftDirection]);
    await context.sync();

    return {
      range: fullRange,
      sheetName: sheet.name,
    };
  });

  return {
    content: [{ type: "text", text: `Deleted **${result.range}** and shifted remaining cells ${shiftDirection}.` }],
    details: {
      kind: "range_operations",
      action: "delete",
      sourceRange: result.range,
      sheetName: result.sheetName,
    },
  };
}

async function mergeRange(params: Params): Promise<AgentToolResult<RangeOperationsDetails>> {
  const result = await excelRun<SingleRangeResult>(async (context) => {
    const { sheet, range } = getRange(context, params.range);
    sheet.load("name");
    range.load("address");
    await context.sync();

    const fullRange = qualifiedAddress(sheet.name, range.address);
    range.merge(params.merge_across ?? false);
    await context.sync();

    return {
      range: fullRange,
      sheetName: sheet.name,
    };
  });

  const modeText = params.merge_across ? " across rows" : "";
  return {
    content: [{ type: "text", text: `Merged cells in **${result.range}**${modeText}.` }],
    details: {
      kind: "range_operations",
      action: "merge",
      sourceRange: result.range,
      sheetName: result.sheetName,
    },
  };
}

async function unmergeRange(params: Params): Promise<AgentToolResult<RangeOperationsDetails>> {
  const result = await excelRun<SingleRangeResult>(async (context) => {
    const { sheet, range } = getRange(context, params.range);
    sheet.load("name");
    range.load("address");
    await context.sync();

    const fullRange = qualifiedAddress(sheet.name, range.address);
    range.unmerge();
    await context.sync();

    return {
      range: fullRange,
      sheetName: sheet.name,
    };
  });

  return {
    content: [{ type: "text", text: `Unmerged cells in **${result.range}**.` }],
    details: {
      kind: "range_operations",
      action: "unmerge",
      sourceRange: result.range,
      sheetName: result.sheetName,
    },
  };
}

export function createRangeOperationsTool(): AgentTool<typeof schema, RangeOperationsDetails> {
  return {
    name: "range_operations",
    label: "Range Operations",
    description:
      "Copy, delete, merge, or unmerge ranges within a worksheet or across sheets.",
    parameters: schema,
    execute: async (
      _toolCallId: string,
      params: Params,
    ): Promise<AgentToolResult<RangeOperationsDetails>> => {
      try {
        switch (params.action) {
          case "copy":
            return await copyRange(params);
          case "delete":
            return await deleteRange(params);
          case "merge":
            return await mergeRange(params);
          case "unmerge":
            return await unmergeRange(params);
        }

        throw new Error("Unsupported range operation.");
      } catch (e: unknown) {
        return {
          content: [{ type: "text", text: `Error with range operations: ${getErrorMessage(e)}` }],
          details: {
            kind: "range_operations",
            action: params.action,
          },
        };
      }
    },
  };
}
