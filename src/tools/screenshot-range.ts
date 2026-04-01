import { Type, type Static } from "@sinclair/typebox";
import type { AgentTool, AgentToolResult } from "@mariozechner/pi-agent-core";

import type { ScreenshotRangeDetails } from "./tool-details.js";
import { excelRun, getRange, parseRangeRef, qualifiedAddress } from "../excel/helpers.js";
import { getErrorMessage } from "../utils/errors.js";

const MAX_CAPTURE_ROWS = 100;
const MAX_CAPTURE_COLS = 50;
const MAX_EXPLANATION_CHARS = 50;

const HEADER_WIDTH = 40;
const HEADER_HEIGHT = 20;
const HEADER_BG = "#f0f0f0";
const HEADER_BORDER = "#c0c0c0";
const HEADER_FONT = "bold 11px Calibri, Arial, sans-serif";
const HEADER_TEXT_COLOR = "#333333";

const schema = Type.Object({
  range: Type.String({
    description: 'Range to capture, e.g. "A1:F20" or "Sheet1!B3:M30". Keep ranges reasonable (max ~50 columns × ~100 rows).',
  }),
  explanation: Type.Optional(
    Type.String({
      maxLength: MAX_EXPLANATION_CHARS,
      description: "Brief explanation of what you're inspecting (max 50 chars).",
    }),
  ),
});

type Params = Static<typeof schema>;

interface ScreenshotCaptureData {
  base64: string;
  colWidths: number[];
  rowHeights: number[];
  startRow: number;
  startCol: number;
  qualifiedAddr: string;
}

interface CompositedImageResult {
  base64: string;
  width: number;
  height: number;
}

function columnIndexToLetter(index: number): string {
  let letter = "";
  let temp = index;
  while (temp >= 0) {
    letter = String.fromCharCode((temp % 26) + 65) + letter;
    temp = Math.floor(temp / 26) - 1;
  }
  return letter;
}

function normalizeSize(value: number | undefined): number {
  return typeof value === "number" && Number.isFinite(value) && value > 0 ? value : 0;
}

function normalizeScreenshotError(message: string): string {
  const normalized = message.toLowerCase();
  if (
    normalized.includes("getimage")
    || (normalized.includes("image") && normalized.includes("not supported"))
    || (normalized.includes("range") && normalized.includes("image") && normalized.includes("unsupported"))
  ) {
    return "screenshot_range requires Excel range.getImage(), which is not supported in this Excel host/version.";
  }
  return message;
}

function loadImage(base64: string): Promise<HTMLImageElement> {
  return new Promise((resolve, reject) => {
    const image = new Image();
    image.onload = () => resolve(image);
    image.onerror = () => reject(new Error("Failed to decode the Excel range image."));
    image.src = `data:image/png;base64,${base64}`;
  });
}

function drawHeaderCell(
  ctx: CanvasRenderingContext2D,
  x: number,
  y: number,
  width: number,
  height: number,
  label?: string,
): void {
  if (width <= 0 || height <= 0) return;

  ctx.fillStyle = HEADER_BG;
  ctx.fillRect(x, y, width, height);
  ctx.strokeStyle = HEADER_BORDER;
  ctx.strokeRect(x, y, width, height);

  if (!label) return;

  ctx.save();
  ctx.beginPath();
  ctx.rect(x, y, width, height);
  ctx.clip();
  ctx.font = HEADER_FONT;
  ctx.fillStyle = HEADER_TEXT_COLOR;
  ctx.textAlign = "center";
  ctx.textBaseline = "middle";
  ctx.fillText(label, x + (width / 2), y + (height / 2));
  ctx.restore();
}

async function compositeWithHeaders(
  base64: string,
  colWidths: number[],
  rowHeights: number[],
  startRow: number,
  startCol: number,
): Promise<CompositedImageResult> {
  const image = await loadImage(base64);
  const canvas = document.createElement("canvas");
  canvas.width = HEADER_WIDTH + image.width;
  canvas.height = HEADER_HEIGHT + image.height;

  const ctx = canvas.getContext("2d");
  if (!ctx) {
    throw new Error("Canvas 2D context is unavailable in this Excel host.");
  }

  ctx.fillStyle = "#ffffff";
  ctx.fillRect(0, 0, canvas.width, canvas.height);
  ctx.drawImage(image, HEADER_WIDTH, HEADER_HEIGHT);

  const totalColWidth = colWidths.reduce((sum, width) => sum + normalizeSize(width), 0);
  const totalRowHeight = rowHeights.reduce((sum, height) => sum + normalizeSize(height), 0);
  const scaleX = totalColWidth > 0 ? image.width / totalColWidth : 1;
  const scaleY = totalRowHeight > 0 ? image.height / totalRowHeight : 1;
  const fallbackColWidth = colWidths.length > 0 ? image.width / colWidths.length : image.width;
  const fallbackRowHeight = rowHeights.length > 0 ? image.height / rowHeights.length : image.height;

  drawHeaderCell(ctx, 0, 0, HEADER_WIDTH, HEADER_HEIGHT);

  let x = HEADER_WIDTH;
  for (let i = 0; i < colWidths.length; i += 1) {
    const width = totalColWidth > 0 ? normalizeSize(colWidths[i]) * scaleX : fallbackColWidth;
    drawHeaderCell(ctx, x, 0, width, HEADER_HEIGHT, columnIndexToLetter(startCol + i));
    x += width;
  }

  let y = HEADER_HEIGHT;
  for (let i = 0; i < rowHeights.length; i += 1) {
    const height = totalRowHeight > 0 ? normalizeSize(rowHeights[i]) * scaleY : fallbackRowHeight;
    drawHeaderCell(ctx, 0, y, HEADER_WIDTH, height, String(startRow + i + 1));
    y += height;
  }

  const dataUrl = canvas.toDataURL("image/png");
  const compositedBase64 = dataUrl.split(",")[1];
  if (!compositedBase64) {
    throw new Error("Failed to encode the composited screenshot image.");
  }

  return {
    base64: compositedBase64,
    width: image.width,
    height: image.height,
  };
}

export function createScreenshotRangeTool(): AgentTool<typeof schema, ScreenshotRangeDetails> {
  return {
    name: "screenshot_range",
    label: "Screenshot Range",
    description: "Capture a visual screenshot of a cell range with row/column headers.",
    parameters: schema,
    execute: async (
      _toolCallId: string,
      params: Params,
    ): Promise<AgentToolResult<ScreenshotRangeDetails>> => {
      try {
        const parsedRange = parseRangeRef(params.range);
        if (parsedRange.address.includes(",")) {
          return {
            content: [{
              type: "text",
              text: "Error: screenshot_range expects a single contiguous range, not a multi-area reference.",
            }],
            details: {
              kind: "screenshot_range",
            },
          };
        }

        const capture = await excelRun<ScreenshotCaptureData>(async (context) => {
          const { sheet, range } = getRange(context, params.range);
          sheet.load("name");
          range.load("address,rowIndex,columnIndex,rowCount,columnCount");
          await context.sync();

          const qualifiedAddr = qualifiedAddress(sheet.name, range.address);
          if (range.rowCount > MAX_CAPTURE_ROWS || range.columnCount > MAX_CAPTURE_COLS) {
            throw new Error(
              `${qualifiedAddr} is ${range.rowCount} rows × ${range.columnCount} columns. screenshot_range works best on smaller ranges up to about ${MAX_CAPTURE_ROWS} rows × ${MAX_CAPTURE_COLS} columns. Try a smaller range.`,
            );
          }

          const imageResult = range.getImage();
          const cols: Excel.Range[] = [];
          for (let i = 0; i < range.columnCount; i += 1) {
            const col = range.getColumn(i);
            col.format.load("columnWidth");
            cols.push(col);
          }

          const rows: Excel.Range[] = [];
          for (let i = 0; i < range.rowCount; i += 1) {
            const row = range.getRow(i);
            row.format.load("rowHeight");
            rows.push(row);
          }

          await context.sync();

          return {
            base64: imageResult.value,
            colWidths: cols.map((col) => col.format.columnWidth),
            rowHeights: rows.map((row) => row.format.rowHeight),
            startRow: range.rowIndex,
            startCol: range.columnIndex,
            qualifiedAddr,
          };
        });

        const composited = await compositeWithHeaders(
          capture.base64,
          capture.colWidths,
          capture.rowHeights,
          capture.startRow,
          capture.startCol,
        );

        return {
          content: [
            { type: "text", text: `Screenshot of ${capture.qualifiedAddr}` },
            { type: "image", data: composited.base64, mimeType: "image/png" },
          ],
          details: {
            kind: "screenshot_range",
            address: capture.qualifiedAddr,
            width: composited.width,
            height: composited.height,
          },
        };
      } catch (error: unknown) {
        return {
          content: [{ type: "text", text: `Error: ${normalizeScreenshotError(getErrorMessage(error))}` }],
          details: {
            kind: "screenshot_range",
          },
        };
      }
    },
  };
}
