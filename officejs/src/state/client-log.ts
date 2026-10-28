import type { ExcelOperation } from "../tools/protocol";

export interface ClientLogContext {
  stage: string;
  error: string;
  selected?: { sheet: string; address: string };
  operations?: readonly ExcelOperation[];
}

function text(value: unknown, limit: number): string | undefined {
  if (typeof value !== "string") return undefined;
  const cleaned = value.replace(/[\u0000-\u001f\u007f]/g, " ").trim();
  return cleaned ? cleaned.slice(0, limit) : undefined;
}

function operationSummary(operation: ExcelOperation): Record<string, string> {
  const result: Record<string, string> = { kind: operation.kind };
  if ("sheet" in operation) {
    result.sheet = operation.sheet;
    result.address = operation.address;
  }
  if (operation.kind === "lookup_replace") {
    result.lookupSheet = operation.lookupSheet;
    result.lookupAddress = operation.lookupAddress;
  }
  return result;
}

export async function reportClientError(context: ClientLogContext): Promise<void> {
  const payload = {
    event: "client-error",
    stage: text(context.stage, 80),
    error: text(context.error, 1000),
    selected: context.selected ? { sheet: text(context.selected.sheet, 120), address: text(context.selected.address, 80) } : undefined,
    operations: context.operations?.slice(0, 20).map(operationSummary),
  };
  const endpoint = process.env.NODE_ENV === "development" ? "/__dev/client-log" : "/__local/client-log";
  try {
    await fetch(endpoint, { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify(payload), keepalive: true });
  } catch {
    // Diagnostics must never change the user-visible operation result.
  }
}
