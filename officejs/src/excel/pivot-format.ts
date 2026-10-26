import { addressFromBounds, Bounds, ConditionalRule, parseAddress } from "../tools/protocol";

type CellReference = { col: number; row: number; absoluteCol: boolean; absoluteRow: boolean };
type Operand = { value?: number | string; reference?: CellReference };

const comparison = /^(EqualTo|NotEqual|GreaterThan|GreaterThanOrEqual|LessThan|LessThanOrEqual|Between|NotBetween)$/;
const cellReference = /^(\$?)([A-Z]{1,3})(\$?)([1-9]\d{0,6})$/i;
const customRule = /^\s*=\s*(\$?[A-Z]{1,3}\$?[1-9]\d{0,6})\s*(<=|>=|<>|=|<|>)\s*(\$?[A-Z]{1,3}\$?[1-9]\d{0,6}|[-+]?\d+(?:\.\d+)?)\s*$/i;

function columnNumber(value: string): number {
  return [...value.toUpperCase()].reduce((total, char) => total * 26 + char.charCodeAt(0) - 64, 0) - 1;
}
function parseReference(value: string): CellReference | undefined {
  const match = cellReference.exec(value.trim());
  if (!match) return;
  return { col: columnNumber(match[2]), row: Number(match[4]) - 1, absoluteCol: Boolean(match[1]), absoluteRow: Boolean(match[3]) };
}
function parseOperand(value: string): Operand | undefined {
  const reference = parseReference(value);
  if (reference) return { reference };
  const number = Number(value);
  return Number.isFinite(number) ? { value: number } : undefined;
}
function compare(left: unknown, operator: string, right: unknown): boolean {
  const a = typeof left === "number" ? left : String(left ?? "");
  const b = typeof right === "number" ? right : String(right ?? "");
  switch (operator) {
    case "=": case "EqualTo": return a === b;
    case "<>": case "NotEqual": return a !== b;
    case "<": case "LessThan": return a < b;
    case "<=": case "LessThanOrEqual": return a <= b;
    case ">": case "GreaterThan": return a > b;
    case ">=": case "GreaterThanOrEqual": return a >= b;
    default: return false;
  }
}
function resolve(reference: CellReference, target: Bounds, row: number, col: number): string {
  const resolvedRow = reference.absoluteRow ? reference.row : reference.row + row - target.row;
  const resolvedCol = reference.absoluteCol ? reference.col : reference.col + col - target.col;
  return addressFromBounds({ row: resolvedRow, col: resolvedCol, rows: 1, cols: 1 });
}
function formatCell(cell: Excel.Range, rule: ConditionalRule): void {
  if (rule.fontColor !== undefined) cell.format.font.color = rule.fontColor;
  if (rule.fill !== undefined) cell.format.fill.color = rule.fill;
  if (rule.bold !== undefined) cell.format.font.bold = rule.bold;
}

export function supportsPivotConditional(rule: ConditionalRule): boolean {
  if (rule.type === "ContainsText") return true;
  if (rule.type === "CellValue") return Boolean(rule.formula1 && comparison.test(rule.operator ?? "") && !["Between", "NotBetween"].includes(rule.operator ?? ""));
  if (rule.type === "Custom") return Boolean(rule.formula1 && customRule.test(rule.formula1));
  return false;
}

/**
 * Excel's conditionalFormats collection can reject PivotTable ranges even though
 * direct font/fill changes are valid. For simple rules, evaluate the current
 * values after formula writes and apply the resulting direct format instead.
 */
export async function applyPivotConditional(context: Excel.RequestContext, sheet: Excel.Worksheet, target: Excel.Range, address: string, rule: ConditionalRule): Promise<void> {
  if (!supportsPivotConditional(rule)) throw new Error("透视表区域不支持此类条件格式，请改用直接字体或填充格式");
  const targetBounds = parseAddress(address);
  const references = new Map<string, Excel.Range>();
  const custom = rule.type === "Custom" ? customRule.exec(rule.formula1!) : undefined;
  const customLeft = custom ? parseOperand(custom[1]) : undefined;
  const customRight = custom ? parseOperand(custom[3]) : undefined;
  if (custom && (!customLeft || !customRight)) throw new Error("透视表条件格式公式过于复杂，请改用直接格式");
  const addReference = (operand?: Operand, row = targetBounds.row, col = targetBounds.col) => {
    if (!operand?.reference) return;
    const ref = resolve(operand.reference, targetBounds, row, col);
    if (!references.has(ref)) references.set(ref, sheet.getRange(ref));
  };
  if (custom) {
    for (let row = targetBounds.row; row < targetBounds.row + targetBounds.rows; row++) {
      for (let col = targetBounds.col; col < targetBounds.col + targetBounds.cols; col++) {
        addReference(customLeft, row, col); addReference(customRight, row, col);
      }
    }
  }
  target.load("values");
  references.forEach(range => range.load("values"));
  await context.sync();
  const values = target.values as unknown[][];
  const referenceValue = (operand: Operand | undefined, row: number, col: number): unknown => {
    if (operand?.value !== undefined) return operand.value;
    if (!operand?.reference) return undefined;
    const ref = resolve(operand.reference, targetBounds, row, col);
    return (references.get(ref)?.values as unknown[][] | undefined)?.[0]?.[0];
  };
  for (let row = 0; row < targetBounds.rows; row++) {
    for (let col = 0; col < targetBounds.cols; col++) {
      const absoluteRow = targetBounds.row + row, absoluteCol = targetBounds.col + col;
      let matches = false;
      if (rule.type === "ContainsText") matches = String(values[row]?.[col] ?? "").includes(rule.text!);
      else if (rule.type === "CellValue") matches = compare(values[row]?.[col], rule.operator!, Number(rule.formula1));
      else if (custom) matches = compare(referenceValue(customLeft, absoluteRow, absoluteCol), custom[2], referenceValue(customRight, absoluteRow, absoluteCol));
      if (matches) formatCell(target.getCell(row, col), rule);
    }
  }
}
