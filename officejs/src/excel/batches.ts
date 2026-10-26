import { Bounds, CellMatrix, CellValue, TextTransform, WRITE_BATCH_CELLS, addressFromBounds } from "../tools/protocol";

export interface BatchProgress { completed: number; total: number; address: string; }
export interface BatchOptions { signal?: AbortSignal; onProgress?: (progress: BatchProgress) => void; }
export function checkCancellation(signal?: AbortSignal): void { if (signal?.aborted) throw new Error("已取消分批执行"); }

// Tile both axes: a single row can itself be wider than a host write batch.
export function* cellBatches(bounds: Bounds): Generator<Bounds> {
  const width = Math.min(bounds.cols, WRITE_BATCH_CELLS), height = Math.max(1, Math.floor(WRITE_BATCH_CELLS / width));
  for (let row = 0; row < bounds.rows; row += height) for (let col = 0; col < bounds.cols; col += width) {
    yield { row: bounds.row + row, col: bounds.col + col, rows: Math.min(height, bounds.rows - row), cols: Math.min(width, bounds.cols - col) };
  }
}
export function transformText(value: string, rule: TextTransform): string {
  if (rule.removeInvisible) value = value.replace(/[\u200b\u2060\ufeff]/g, "");
  if (rule.whitespace === "remove") value = value.replace(/\s/g, "");
  if (rule.whitespace === "trim") value = value.trim();
  if (rule.whitespace === "collapse") value = value.trim().replace(/\s+/g, " ");
  for (const replacement of rule.replacements ?? []) value = replacement.wholeCell
    ? value === replacement.find ? replacement.replace : value
    : value.split(replacement.find).join(replacement.replace);
  if (value.length > 32767) throw new Error("清理后的单元格文字超过 Excel 上限，已停止执行");
  return value;
}
export function literalValues(data: CellMatrix, forceText = false): (string | number | boolean)[][] {
  return data.map(row => row.map(value => value === null ? "" : typeof value === "string" && value !== "" && (forceText || /^[=+\-@']/.test(value)) ? `'${value}` : value));
}
function* boundedMatrices(data: CellMatrix, bounds: Bounds): Generator<{ bounds: Bounds; data: CellMatrix }> {
  // Limit string payload as well as cell count (UTF-8 is at most 3 bytes per UTF-16 unit).
  if (JSON.stringify(data).length <= 250000 || bounds.rows * bounds.cols === 1) { yield { bounds, data }; return; }
  if (bounds.rows > 1) {
    const rows = Math.floor(bounds.rows / 2);
    yield* boundedMatrices(data.slice(0, rows), { ...bounds, rows });
    yield* boundedMatrices(data.slice(rows), { ...bounds, row: bounds.row + rows, rows: bounds.rows - rows });
  } else {
    const cols = Math.floor(bounds.cols / 2);
    yield* boundedMatrices(data.map(row => row.slice(0, cols)), { ...bounds, cols });
    yield* boundedMatrices(data.map(row => row.slice(cols)), { ...bounds, col: bounds.col + cols, cols: bounds.cols - cols });
  }
}
export async function writeMatrix(context: Excel.RequestContext, sheet: Excel.Worksheet, bounds: Bounds, data: CellMatrix, formulas: boolean, options: BatchOptions = {}): Promise<void> {
  let completed = 0;
  for (const tile of cellBatches(bounds)) {
    const matrix = data.slice(tile.row - bounds.row, tile.row - bounds.row + tile.rows).map(row => row.slice(tile.col - bounds.col, tile.col - bounds.col + tile.cols));
    for (const batch of boundedMatrices(matrix, tile)) {
      checkCancellation(options.signal);
      const target = sheet.getRange(addressFromBounds(batch.bounds));
      if (formulas) target.formulas = batch.data as string[][]; else target.values = literalValues(batch.data, true);
      await context.sync();
      completed += batch.bounds.rows * batch.bounds.cols;
      options.onProgress?.({ completed, total: bounds.rows * bounds.cols, address: addressFromBounds(batch.bounds) });
      // Let the pane process cancellation between host submissions.
      await new Promise<void>(resolve => setTimeout(resolve, 0));
    }
  }
  checkCancellation(options.signal);
}
export async function cleanText(context: Excel.RequestContext, sheet: Excel.Worksheet, bounds: Bounds, rule: TextTransform, options: BatchOptions = {}): Promise<{ changed: number; skippedFormulas: number }> {
  let completed = 0, changed = 0, skippedFormulas = 0;
  for (const batch of cellBatches(bounds)) {
    checkCancellation(options.signal);
    const target = sheet.getRange(addressFromBounds(batch));
    target.load(["values", "formulas"]);
    await context.sync();
    const values = target.values as CellMatrix, formulas = target.formulas as CellMatrix;
    const writes: { bounds: Bounds; data: CellMatrix }[] = [];
    for (let row = 0; row < batch.rows; row++) {
      let run: { bounds: Bounds; data: CellMatrix } | undefined;
      for (let col = 0; col < batch.cols; col++) {
        const value = values[row][col], formula = formulas[row][col];
        const isFormula = typeof formula === "string" && formula.startsWith("=");
        if (isFormula) skippedFormulas++;
        const next: CellValue = !isFormula && typeof value === "string" ? transformText(value, rule) : value;
        if (next === value) { run = undefined; continue; }
        changed++;
        if (!run) { run = { bounds: { row: batch.row + row, col: batch.col + col, rows: 1, cols: 0 }, data: [[]] }; writes.push(run); }
        run.bounds.cols++; run.data[0].push(next);
      }
    }
    // Coalesce vertical runs too, so a 96015-row text column needs ~49 writes,
    // rather than one host command per cell. Formula gaps split the runs.
    const merged: typeof writes = [], previous = new Map<string, typeof writes[number]>();
    for (const write of writes) {
      const key = `${write.bounds.col}:${write.bounds.cols}`, above = previous.get(key);
      if (above && above.bounds.row + above.bounds.rows === write.bounds.row) { above.bounds.rows++; above.data.push(write.data[0]); }
      else { merged.push(write); previous.set(key, write); }
    }
    // Write only changed text constants; numeric values, formulas and their formats stay intact.
    let pending = 0, characters = 0;
    for (const write of merged) for (const part of boundedMatrices(write.data, write.bounds)) {
      checkCancellation(options.signal);
      const size = JSON.stringify(part.data).length;
      if (pending && (pending >= 100 || characters + size > 250000)) { await context.sync(); pending = characters = 0; }
      sheet.getRange(addressFromBounds(part.bounds)).values = literalValues(part.data, true);
      pending++; characters += size;
    }
    if (pending) await context.sync();
    completed += batch.rows * batch.cols;
    options.onProgress?.({ completed, total: bounds.rows * bounds.cols, address: addressFromBounds(batch) });
    await new Promise<void>(resolve => setTimeout(resolve, 0));
  }
  checkCancellation(options.signal);
  return { changed, skippedFormulas };
}
