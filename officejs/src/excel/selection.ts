import { Bounds, addressFromBounds } from "../tools/protocol";
export interface SelectedRange {
  workbook: string; sheet: string; sheetId: string; address: string; bounds: Bounds;
  values: (string | number | boolean | null)[][]; formulas: (string | number | boolean | null)[][];
  dataMode: "full" | "metadata";
  formatMode?: "full" | "metadata";
  formats?: CellFormat[][];
  supplementalRanges?: SupplementalRange[];
  selectionKind?: "range" | "column" | "row";
  selectionLabel?: string;
}
export interface CellFormat {
  fill: string | null; fontColor: string | null; bold: boolean | null; italic: boolean | null;
  underline: string | null; numberFormat: string | null;
}
export interface SupplementalRange {
  workbook: string; sheet: string; sheetId: string; address: string; bounds: Bounds;
  values: (string | number | boolean | null)[][];
  formulas: (string | number | boolean | null)[][];
}
export const MAX_READ_CELLS = 100000;
export const MAX_FORMAT_CELLS = 10000;
export type ReadMode = "adaptive" | "metadata" | "format" | "full";
export async function readSelection(mode: ReadMode = "adaptive", target?: { sheetId: string; address: string }, includeFormats = false): Promise<SelectedRange> {
  return Excel.run(async (context) => {
    const range = target ? context.workbook.worksheets.getItem(target.sheetId).getRange(target.address) : context.workbook.getSelectedRange(), sheet = range.worksheet;
    range.load(["address", "rowIndex", "columnIndex", "rowCount", "columnCount"]);
    sheet.load(["name", "id"]);
    context.workbook.load("name");
    await context.sync();
    const fullColumn = range.rowIndex === 0 && range.rowCount === 1048576 && range.columnCount < 16384;
    const fullRow = range.columnIndex === 0 && range.columnCount === 16384 && range.rowCount < 1048576;
    const selectionKind = fullColumn ? "column" as const : fullRow ? "row" as const : "range" as const;
    const columnLabel = (index: number) => { let text = ""; for (let n = index + 1; n > 0; n = Math.floor((n - 1) / 26)) text = String.fromCharCode(65 + (n - 1) % 26) + text; return text; };
    const selectionLabel = fullColumn ? `${columnLabel(range.columnIndex)}:${columnLabel(range.columnIndex + range.columnCount - 1)}` : fullRow ? `${range.rowIndex + 1}:${range.rowIndex + range.rowCount}` : undefined;
    let effective = range;
    if (fullColumn || fullRow) {
      effective = range.getUsedRange(true);
      effective.load(["rowIndex", "columnIndex", "rowCount", "columnCount"]);
      await context.sync();
    }
    const cellCount = effective.rowCount * effective.columnCount;
    if (cellCount > MAX_READ_CELLS) throw new Error("一次最多处理 100000 个已用单元格；请缩小选区或分批处理");
    const shouldReadData = mode === "full" || (mode === "adaptive" && cellCount <= 100);
    const bounds = { row: effective.rowIndex, col: effective.columnIndex, rows: effective.rowCount, cols: effective.columnCount };
    // Excel may report just the anchor address for a merged range even while
    // values and dimensions describe the full rectangle. Preserve that rectangle.
    const base = { workbook: context.workbook.name, sheet: sheet.name, sheetId: sheet.id, address: addressFromBounds(bounds), bounds, selectionKind, ...(selectionLabel ? { selectionLabel } : {}) };
    if (mode === "metadata") return { ...base, values: [], formulas: [], dataMode: "metadata" as const, formatMode: "metadata" as const };
    if (mode === "format") {
      if (cellCount > MAX_FORMAT_CELLS) throw new Error(`一次最多读取 ${MAX_FORMAT_CELLS} 个单元格的格式；请缩小选区后重试`);
      const formats = await readFormats(context, effective);
      return { ...base, values: [], formulas: [], dataMode: "metadata" as const, formatMode: "full" as const, formats };
    }
    if (!shouldReadData) return { ...base, values: [], formulas: [], dataMode: "metadata" as const, formatMode: "metadata" as const };
    effective.load(["values", "formulas"]);
    await context.sync();
    if (!includeFormats) return { ...base, values: effective.values, formulas: effective.formulas, dataMode: "full" as const, formatMode: "metadata" as const };
    if (cellCount > MAX_FORMAT_CELLS) throw new Error(`一次最多读取 ${MAX_FORMAT_CELLS} 个单元格的格式；请缩小选区后重试`);
    const formats = await readFormats(context, effective);
    return { ...base, values: effective.values, formulas: effective.formulas, dataMode: "full" as const, formatMode: "full" as const, formats };
  });
}

export async function readSupplementalRanges(ranges: readonly { sheet: string; address: string }[]): Promise<SupplementalRange[]> {
  const results: SupplementalRange[] = [];
  for (const request of ranges) {
    const selected = await readSelection("full", { sheetId: request.sheet, address: request.address });
    results.push({ workbook: selected.workbook, sheet: selected.sheet, sheetId: selected.sheetId, address: selected.address, bounds: selected.bounds, values: selected.values, formulas: selected.formulas });
  }
  return results;
}

async function readFormats(context: Excel.RequestContext, effective: Excel.Range): Promise<CellFormat[][]> {
    const reads: { cell: Excel.Range; row: number; col: number }[] = [];
    for (let row = 0; row < effective.rowCount; row++) for (let col = 0; col < effective.columnCount; col++) {
      const cell = effective.getCell(row, col);
      cell.format.fill.load("color"); cell.format.font.load(["color", "bold", "italic", "underline"]); cell.load("numberFormat");
      reads.push({ cell, row, col });
    }
    await context.sync();
    const formats: CellFormat[][] = Array.from({ length: effective.rowCount }, () => Array.from({ length: effective.columnCount }, () => ({ fill: null, fontColor: null, bold: null, italic: null, underline: null, numberFormat: null })));
    for (const item of reads) {
      const numberFormat = item.cell.numberFormat as string[][];
      formats[item.row][item.col] = {
        fill: item.cell.format.fill.color ?? null,
        fontColor: item.cell.format.font.color ?? null,
        bold: item.cell.format.font.bold ?? null,
        italic: item.cell.format.font.italic ?? null,
        underline: item.cell.format.font.underline ?? null,
        numberFormat: numberFormat?.[0]?.[0] ?? null,
      };
    }
    return formats;
}
