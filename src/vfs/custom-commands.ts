import { defineCommand, type CustomCommand } from "just-bash/browser";

import { computeRangeAddress, excelRun, getRange, parseRangeRef } from "../excel/helpers.js";
import { getErrorMessage } from "../utils/errors.js";

function parseCsv(text: string): string[][] {
  const rows: string[][] = [];
  let row: string[] = [];
  let cell = "";
  let inQuotes = false;

  for (let i = 0; i < text.length; i += 1) {
    const ch = text[i] ?? "";
    const next = text[i + 1] ?? "";

    if (inQuotes) {
      if (ch === '"' && next === '"') {
        cell += '"';
        i += 1;
        continue;
      }

      if (ch === '"') {
        inQuotes = false;
        continue;
      }

      cell += ch;
      continue;
    }

    if (ch === '"') {
      inQuotes = true;
      continue;
    }

    if (ch === ",") {
      row.push(cell);
      cell = "";
      continue;
    }

    if (ch === "\n") {
      row.push(cell.replace(/\r$/u, ""));
      const hasContent = row.some((value) => value.length > 0);
      if (hasContent) rows.push(row);
      row = [];
      cell = "";
      continue;
    }

    cell += ch;
  }

  if (cell.length > 0 || row.length > 0) {
    row.push(cell.replace(/\r$/u, ""));
    const hasContent = row.some((value) => value.length > 0);
    if (hasContent) rows.push(row);
  }

  return rows;
}

function coerceValue(raw: string): string | number | boolean {
  const trimmed = raw.trim();

  if (trimmed === "") return "";
  if (trimmed.toLowerCase() === "true") return true;
  if (trimmed.toLowerCase() === "false") return false;

  const num = Number(trimmed);
  if (!Number.isNaN(num)) return num;

  return raw;
}

function resolveCommandPath(ctx: { cwd: string; fs: { resolvePath(base: string, path: string): string } }, path: string): string {
  return path.startsWith("/") ? path : ctx.fs.resolvePath(ctx.cwd, path);
}

async function ensureParentDir(
  fs: {
    mkdir(path: string, options?: { recursive?: boolean }): Promise<void>;
  },
  path: string,
): Promise<void> {
  const dir = path.substring(0, path.lastIndexOf("/"));
  if (!dir || dir === "/") return;

  try {
    await fs.mkdir(dir, { recursive: true });
  } catch {
  }
}

function createCsvToSheetCommand(): CustomCommand {
  return defineCommand("csv-to-sheet", async (args, ctx) => {
    const force = args.includes("--force") || args.includes("-f");
    const positional = args.filter((arg) => arg !== "--force" && arg !== "-f");

    if (positional.length < 2) {
      return {
        stdout: "",
        stderr: "Usage: csv-to-sheet <file> <sheetName> [startCell] [--force]\n  file      - Path to CSV file in VFS\n  sheetName - Target sheet name\n  startCell - Top-left cell, default A1\n  --force   - Overwrite existing cell data",
        exitCode: 1,
      };
    }

    const [filePath, sheetName, startCell = "A1"] = positional;
    const upperStartCell = startCell.toUpperCase();
    if (!/^[A-Z]+\d+$/u.test(upperStartCell)) {
      return {
        stdout: "",
        stderr: `Invalid start cell: ${startCell}`,
        exitCode: 1,
      };
    }

    try {
      const resolvedPath = resolveCommandPath(ctx, filePath);
      const content = await ctx.fs.readFile(resolvedPath);
      const rows = parseCsv(content);

      if (rows.length === 0) {
        return {
          stdout: "",
          stderr: "CSV file is empty",
          exitCode: 1,
        };
      }

      const maxCols = Math.max(...rows.map((row) => row.length));
      const values: (string | number | boolean)[][] = rows.map((row) => {
        const padded = [...row];
        while (padded.length < maxCols) padded.push("");
        return padded.map(coerceValue);
      });

      const rangeAddr = computeRangeAddress(upperStartCell, rows.length, maxCols);

      await excelRun(async (context) => {
        const sheet = context.workbook.worksheets.getItem(sheetName);
        const range = sheet.getRange(rangeAddr);

        if (!force) {
          range.load("values");
          await context.sync();

          const existingValues = range.values as unknown[][];
          const hasData = existingValues.some((row) => row.some((value) => value !== "" && value !== null));
          if (hasData) {
            throw new Error("Target range contains existing data. Use --force to overwrite.");
          }
        }

        range.values = values;
        await context.sync();
      });

      return {
        stdout: `Imported ${rows.length} rows × ${maxCols} columns into "${sheetName}" at ${upperStartCell} (${rangeAddr}).`,
        stderr: "",
        exitCode: 0,
      };
    } catch (error: unknown) {
      return {
        stdout: "",
        stderr: getErrorMessage(error),
        exitCode: 1,
      };
    }
  });
}

function serializeCsvCell(cell: unknown): string {
  let text = "";

  if (typeof cell === "string") {
    text = cell;
  } else if (typeof cell === "number" || typeof cell === "boolean" || typeof cell === "bigint") {
    text = String(cell);
  } else if (cell instanceof Date) {
    text = cell.toISOString();
  } else if (cell !== null && cell !== undefined) {
    try {
      text = JSON.stringify(cell) ?? "";
    } catch {
      text = "[unserializable]";
    }
  }

  if (text.includes(",") || text.includes("\n") || text.includes('"')) {
    return `"${text.replace(/"/gu, '""')}"`;
  }

  return text;
}

function createSheetToCsvCommand(): CustomCommand {
  return defineCommand("sheet-to-csv", async (args, ctx) => {
    if (args.length < 1) {
      return {
        stdout: "",
        stderr: "Usage: sheet-to-csv <sheetName> [range] [file]\n  sheetName - Source sheet name\n  range     - Cell range, e.g. A1:D100 (optional, defaults to used range)\n  file      - Output file path (optional, prints to stdout if omitted)",
        exitCode: 1,
      };
    }

    const sheetName = args[0];
    let rangeAddr: string | undefined;
    let outFile: string | undefined;

    const looksLikeRange = (value: string): boolean => /^[A-Z]+\d+(:[A-Z]+\d+)?$/iu.test(value) || value.includes("!");

    if (args.length === 2) {
      if (looksLikeRange(args[1] ?? "")) {
        rangeAddr = args[1];
      } else {
        outFile = args[1];
      }
    } else if (args.length >= 3) {
      rangeAddr = args[1];
      outFile = args[2];
    }

    try {
      let csv = "";

      await excelRun(async (context) => {
        const sheet = context.workbook.worksheets.getItem(sheetName);
        let range: Excel.Range;

        if (rangeAddr) {
          const parsed = parseRangeRef(rangeAddr);
          if (parsed.sheet && parsed.sheet !== sheetName) {
            throw new Error(`Range sheet "${parsed.sheet}" does not match sheet "${sheetName}".`);
          }

          range = getRange(context, `${sheetName}!${parsed.address}`).range;
        } else {
          range = sheet.getUsedRangeOrNullObject();
          range.load("isNullObject");
          await context.sync();

          if (range.isNullObject) {
            throw new Error("Sheet is empty (no used range)");
          }
        }

        range.load("values");
        await context.sync();

        const rows = range.values as unknown[][];
        csv = rows.map((row) => row.map(serializeCsvCell).join(",")).join("\n");
      });

      if (outFile) {
        const resolvedPath = resolveCommandPath(ctx, outFile);
        await ensureParentDir(ctx.fs, resolvedPath);
        await ctx.fs.writeFile(resolvedPath, csv);
        return {
          stdout: `Exported to ${outFile}`,
          stderr: "",
          exitCode: 0,
        };
      }

      return {
        stdout: csv,
        stderr: "",
        exitCode: 0,
      };
    } catch (error: unknown) {
      return {
        stdout: "",
        stderr: getErrorMessage(error),
        exitCode: 1,
      };
    }
  });
}

export function getExcelCustomCommands(): CustomCommand[] {
  return [
    createCsvToSheetCommand(),
    createSheetToCsvCommand(),
  ];
}

export function getCustomCommandPromptSnippets(): string[] {
  return [
    "- `csv-to-sheet <file> <sheetName> [startCell] [--force]` — import CSV from VFS into Excel. Auto-coerces numbers and booleans. Use `--force` to overwrite existing cells.",
    "- `sheet-to-csv <sheetName> [range] [file]` — export a sheet/range to CSV in VFS. If `file` is omitted, prints CSV to stdout (pipeable).",
  ];
}
