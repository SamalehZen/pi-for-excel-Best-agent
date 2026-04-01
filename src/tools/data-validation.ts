import type { AgentTool, AgentToolResult } from "@mariozechner/pi-agent-core";
import { Type, type Static } from "@sinclair/typebox";
import { excelRun, getRange, parseRangeRef, qualifiedAddress } from "../excel/helpers.js";
import { getErrorMessage } from "../utils/errors.js";
import type { DataValidationDetails } from "./tool-details.js";

const schema = Type.Object({
  action: Type.Union([Type.Literal("get"), Type.Literal("set"), Type.Literal("clear")], {
    description: '"get" to read existing rules, "set" to apply a rule, "clear" to remove rules.',
  }),
  range: Type.String({
    description: 'Target range, e.g. "B2:B20" or "Sheet1!C3:C50".',
  }),
  type: Type.Optional(Type.Union([
    Type.Literal("list"),
    Type.Literal("whole_number"),
    Type.Literal("decimal"),
    Type.Literal("date"),
    Type.Literal("text_length"),
    Type.Literal("custom"),
  ], {
    description: 'Validation type for "set" action.',
  })),
  list_items: Type.Optional(Type.Array(Type.String(), {
    description: 'Dropdown items for "list" type, e.g. ["Yes", "No", "Maybe"].',
  })),
  list_source: Type.Optional(Type.String({
    description: 'Cell range source for "list" type, e.g. "Sheet2!A1:A10". Alternative to list_items.',
  })),
  operator: Type.Optional(Type.String({
    description: 'Comparison operator for number/date/text_length: "between", "notBetween", "equalTo", "notEqualTo", "greaterThan", "lessThan", "greaterThanOrEqualTo", "lessThanOrEqualTo".',
  })),
  formula1: Type.Optional(Type.Union([Type.String(), Type.Number()], {
    description: "First value/formula for comparison. Required for number/date/text_length rules.",
  })),
  formula2: Type.Optional(Type.Union([Type.String(), Type.Number()], {
    description: 'Second value/formula. Required when operator is "between" or "notBetween".',
  })),
  custom_formula: Type.Optional(Type.String({
    description: 'Custom formula for "custom" type, e.g. "=LEN(A1)<=50".',
  })),
  input_message: Type.Optional(Type.String({
    description: "Prompt message shown when the cell is selected.",
  })),
  input_title: Type.Optional(Type.String({
    description: "Title for the input prompt.",
  })),
  error_message: Type.Optional(Type.String({
    description: "Error message shown when invalid data is entered.",
  })),
  error_title: Type.Optional(Type.String({
    description: "Title for the error alert.",
  })),
});

type Params = Static<typeof schema>;

const VALIDATION_TYPES = [
  "list",
  "whole_number",
  "decimal",
  "date",
  "text_length",
  "custom",
] as const;

type SupportedValidationType = (typeof VALIDATION_TYPES)[number];

const VALID_OPERATORS = [
  "between", "notBetween", "equalTo", "notEqualTo",
  "greaterThan", "lessThan", "greaterThanOrEqualTo", "lessThanOrEqualTo",
] as const;

type SupportedValidationOperator = (typeof VALID_OPERATORS)[number];

let _operatorMap: Record<SupportedValidationOperator, Excel.DataValidationOperator> | null = null;
function getOperatorMap(): Record<SupportedValidationOperator, Excel.DataValidationOperator> {
  if (!_operatorMap) {
    _operatorMap = {
      between: Excel.DataValidationOperator.between,
      notBetween: Excel.DataValidationOperator.notBetween,
      equalTo: Excel.DataValidationOperator.equalTo,
      notEqualTo: Excel.DataValidationOperator.notEqualTo,
      greaterThan: Excel.DataValidationOperator.greaterThan,
      lessThan: Excel.DataValidationOperator.lessThan,
      greaterThanOrEqualTo: Excel.DataValidationOperator.greaterThanOrEqualTo,
      lessThanOrEqualTo: Excel.DataValidationOperator.lessThanOrEqualTo,
    };
  }
  return _operatorMap;
}

const OPERATOR_MAP_KEYS = VALID_OPERATORS;

const BETWEEN_OPERATORS = new Set<SupportedValidationOperator>(["between", "notBetween"]);
const COMPARISON_TYPES = new Set<SupportedValidationType>([
  "whole_number",
  "decimal",
  "date",
  "text_length",
]);

interface ValidationSummary {
  ruleCount: number;
  lines: string[];
}

interface GetValidationResult {
  sheetName: string;
  address: string;
  summary: ValidationSummary;
}

interface SetValidationResult {
  sheetName: string;
  address: string;
  description: string;
}

function normalizeText(value: string): string {
  return value.trim();
}

function normalizeOperator(value: string | undefined): SupportedValidationOperator | null {
  if (!value) return null;
  const normalized = normalizeText(value);
  return normalized in getOperatorMap()
    ? normalized as SupportedValidationOperator
    : null;
}

function stringifyValue(value: string | number | undefined): string {
  if (value === undefined) return "";
  return typeof value === "number" ? value.toString() : value;
}

function qualifyRangeSource(source: string, sheetName: string): string {
  const parsed = parseRangeRef(source);
  if (parsed.sheet) return source;
  if (!/^\$?[A-Z]+\$?\d+(?::\$?[A-Z]+\$?\d+)?$/iu.test(parsed.address)) {
    return source;
  }
  return qualifiedAddress(sheetName, parsed.address);
}

function isAddressLike(value: unknown): value is { address: string } {
  return typeof value === "object"
    && value !== null
    && "address" in value
    && typeof value.address === "string";
}

function displayValidationValue(value: unknown): string {
  if (typeof value === "string" || typeof value === "number" || typeof value === "boolean") {
    return String(value);
  }
  if (isAddressLike(value)) {
    return value.address;
  }
  return "(complex value)";
}

function describeRule(rule: Excel.DataValidationRule): ValidationSummary {
  if (rule.list) {
    const lines = [
      "- Type: list",
      `- Source: ${displayValidationValue(rule.list.source)}`,
      `- In-cell dropdown: ${rule.list.inCellDropDown ? "yes" : "no"}`,
    ];
    return { ruleCount: 1, lines };
  }

  if (rule.wholeNumber) {
    return {
      ruleCount: 1,
      lines: describeComparisonRule("whole number", rule.wholeNumber.operator, rule.wholeNumber.formula1, rule.wholeNumber.formula2),
    };
  }

  if (rule.decimal) {
    return {
      ruleCount: 1,
      lines: describeComparisonRule("decimal", rule.decimal.operator, rule.decimal.formula1, rule.decimal.formula2),
    };
  }

  if (rule.date) {
    return {
      ruleCount: 1,
      lines: describeComparisonRule("date", rule.date.operator, rule.date.formula1, rule.date.formula2),
    };
  }

  if (rule.textLength) {
    return {
      ruleCount: 1,
      lines: describeComparisonRule("text length", rule.textLength.operator, rule.textLength.formula1, rule.textLength.formula2),
    };
  }

  if (rule.custom) {
    return {
      ruleCount: 1,
      lines: [
        "- Type: custom",
        `- Formula: ${rule.custom.formula}`,
      ],
    };
  }

  return { ruleCount: 0, lines: [] };
}

function describeComparisonRule(
  label: string,
  operator: Excel.BasicDataValidation["operator"],
  formula1: unknown,
  formula2?: unknown,
): string[] {
  const lines = [
    `- Type: ${label}`,
    `- Operator: ${String(operator)}`,
    `- Value 1: ${displayValidationValue(formula1)}`,
  ];

  if (formula2 !== undefined && formula2 !== "") {
    lines.push(`- Value 2: ${displayValidationValue(formula2)}`);
  }

  return lines;
}

function buildRule(params: Params, sheetName: string): { rule: Excel.DataValidationRule; description: string } {
  const type = params.type;
  if (!type || !VALIDATION_TYPES.includes(type)) {
    throw new Error(`type is required for action="set". Valid types: ${VALIDATION_TYPES.join(", ")}.`);
  }

  if (type === "list") {
    const items = params.list_items?.filter((item) => item.trim().length > 0) ?? [];
    const source = params.list_source ? qualifyRangeSource(params.list_source, sheetName) : "";
    if (items.length === 0 && source.length === 0) {
      throw new Error("For type=\"list\", provide either list_items or list_source.");
    }
    if (items.length > 0 && source.length > 0) {
      throw new Error("For type=\"list\", provide either list_items or list_source, not both.");
    }

    const resolvedSource = source.length > 0 ? source : items.join(",");
    return {
      rule: {
        list: {
          inCellDropDown: true,
          source: resolvedSource,
        },
      },
      description: source.length > 0
        ? `list from \`${resolvedSource}\``
        : `list (${items.join(", ")})`,
    };
  }

  if (type === "custom") {
    if (!params.custom_formula || params.custom_formula.trim().length === 0) {
      throw new Error("custom_formula is required for type=\"custom\".");
    }

    return {
      rule: {
        custom: {
          formula: params.custom_formula,
        },
      },
      description: `custom formula \`${params.custom_formula}\``,
    };
  }

  if (!COMPARISON_TYPES.has(type)) {
    throw new Error(`Unsupported validation type "${type}".`);
  }

  const operator = normalizeOperator(params.operator);
  if (!operator) {
    throw new Error(`operator is required for type="${type}". Valid values: ${Object.keys(getOperatorMap()).join(", ")}.`);
  }

  if (params.formula1 === undefined || params.formula1 === "") {
    throw new Error(`formula1 is required for type="${type}".`);
  }

  if (BETWEEN_OPERATORS.has(operator) && (params.formula2 === undefined || params.formula2 === "")) {
    throw new Error(`formula2 is required when operator="${operator}".`);
  }

  const descriptionParts = [
    type.replaceAll("_", " "),
    operator,
    stringifyValue(params.formula1),
    params.formula2 !== undefined && params.formula2 !== "" ? stringifyValue(params.formula2) : undefined,
  ].filter((part): part is string => part !== undefined && part.length > 0);

  switch (type) {
    case "whole_number":
      return {
        rule: {
          wholeNumber: {
            formula1: params.formula1,
            operator: getOperatorMap()[operator],
            ...(params.formula2 !== undefined && params.formula2 !== "" ? { formula2: params.formula2 } : {}),
          },
        },
        description: descriptionParts.join(" "),
      };
    case "decimal":
      return {
        rule: {
          decimal: {
            formula1: params.formula1,
            operator: getOperatorMap()[operator],
            ...(params.formula2 !== undefined && params.formula2 !== "" ? { formula2: params.formula2 } : {}),
          },
        },
        description: descriptionParts.join(" "),
      };
    case "date":
      if (typeof params.formula1 !== "string") {
        throw new Error("For type=\"date\", formula1 must be a date string, cell reference, or formula.");
      }
      if (params.formula2 !== undefined && params.formula2 !== "" && typeof params.formula2 !== "string") {
        throw new Error("For type=\"date\", formula2 must be a date string, cell reference, or formula.");
      }
      return {
        rule: {
          date: {
            formula1: params.formula1,
            operator: getOperatorMap()[operator],
            ...(params.formula2 !== undefined && params.formula2 !== "" ? { formula2: params.formula2 } : {}),
          },
        },
        description: descriptionParts.join(" "),
      };
    case "text_length":
      return {
        rule: {
          textLength: {
            formula1: params.formula1,
            operator: getOperatorMap()[operator],
            ...(params.formula2 !== undefined && params.formula2 !== "" ? { formula2: params.formula2 } : {}),
          },
        },
        description: descriptionParts.join(" "),
      };
  }

  throw new Error("Unsupported validation type.");
}

function appendPromptAndErrorLines(
  lines: string[],
  prompt: Excel.DataValidationPrompt,
  errorAlert: Excel.DataValidationErrorAlert,
): void {
  if (prompt.title || prompt.message) {
    const parts = [prompt.title, prompt.message].filter((value): value is string => typeof value === "string" && value.length > 0);
    if (parts.length > 0) lines.push(`- Prompt: ${parts.join(" — ")}`);
  }

  if (errorAlert.title || errorAlert.message) {
    const parts = [errorAlert.title, errorAlert.message].filter((value): value is string => typeof value === "string" && value.length > 0);
    if (parts.length > 0) lines.push(`- Error alert: ${parts.join(" — ")}`);
  }
}

async function getValidation(params: Params): Promise<AgentToolResult<DataValidationDetails>> {
  const result = await excelRun<GetValidationResult>(async (context) => {
    const { sheet, range } = getRange(context, params.range);
    sheet.load("name");
    range.load("address");
    range.dataValidation.load("rule,prompt,errorAlert");
    await context.sync();

    const summary = describeRule(range.dataValidation.rule);
    appendPromptAndErrorLines(summary.lines, range.dataValidation.prompt, range.dataValidation.errorAlert);

    return {
      sheetName: sheet.name,
      address: range.address,
      summary,
    };
  });

  const fullRange = qualifiedAddress(result.sheetName, result.address);
  const text = result.summary.ruleCount > 0
    ? [
      `Data validation for **${fullRange}**:`,
      ...result.summary.lines,
    ].join("\n")
    : `No data validation rule found on **${fullRange}**.`;

  return {
    content: [{ type: "text", text }],
    details: {
      kind: "data_validation",
      action: "get",
      range: fullRange,
      sheetName: result.sheetName,
      ruleCount: result.summary.ruleCount,
    },
  };
}

async function clearValidation(params: Params): Promise<AgentToolResult<DataValidationDetails>> {
  const result = await excelRun<SetValidationResult>(async (context) => {
    const { sheet, range } = getRange(context, params.range);
    sheet.load("name");
    range.load("address");
    await context.sync();

    range.dataValidation.clear();
    await context.sync();

    return {
      sheetName: sheet.name,
      address: range.address,
      description: "cleared",
    };
  });

  const fullRange = qualifiedAddress(result.sheetName, result.address);
  return {
    content: [{ type: "text", text: `Cleared data validation on **${fullRange}**.` }],
    details: {
      kind: "data_validation",
      action: "clear",
      range: fullRange,
      sheetName: result.sheetName,
    },
  };
}

async function setValidation(params: Params): Promise<AgentToolResult<DataValidationDetails>> {
  const result = await excelRun<SetValidationResult>(async (context) => {
    const { sheet, range } = getRange(context, params.range);
    sheet.load("name");
    range.load("address");
    await context.sync();

    const built = buildRule(params, sheet.name);
    range.dataValidation.rule = built.rule;

    if (params.input_title !== undefined || params.input_message !== undefined) {
      if (params.input_title !== undefined) range.dataValidation.prompt.title = params.input_title;
      if (params.input_message !== undefined) range.dataValidation.prompt.message = params.input_message;
      range.dataValidation.prompt.showPrompt = true;
    }

    if (params.error_title !== undefined || params.error_message !== undefined) {
      if (params.error_title !== undefined) range.dataValidation.errorAlert.title = params.error_title;
      if (params.error_message !== undefined) range.dataValidation.errorAlert.message = params.error_message;
      range.dataValidation.errorAlert.showAlert = true;
    }

    await context.sync();

    return {
      sheetName: sheet.name,
      address: range.address,
      description: built.description,
    };
  });

  const fullRange = qualifiedAddress(result.sheetName, result.address);
  return {
    content: [{ type: "text", text: `Set data validation on **${fullRange}**: ${result.description}.` }],
    details: {
      kind: "data_validation",
      action: "set",
      range: fullRange,
      sheetName: result.sheetName,
      ruleCount: 1,
    },
  };
}

export function createDataValidationTool(): AgentTool<typeof schema, DataValidationDetails> {
  return {
    name: "data_validation",
    label: "Data Validation",
    description:
      "Read, apply, or clear Excel data validation rules including dropdown lists, numeric/date constraints, text length, and custom formulas.",
    parameters: schema,
    execute: async (
      _toolCallId: string,
      params: Params,
    ): Promise<AgentToolResult<DataValidationDetails>> => {
      try {
        if (params.action === "get") {
          return await getValidation(params);
        }

        if (params.action === "clear") {
          return await clearValidation(params);
        }

        return await setValidation(params);
      } catch (e: unknown) {
        return {
          content: [{ type: "text", text: `Error with data validation: ${getErrorMessage(e)}` }],
          details: {
            kind: "data_validation",
            action: params.action,
          },
        };
      }
    },
  };
}
