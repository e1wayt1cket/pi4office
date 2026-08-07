/** Shared helpers for workbook recovery modules. */

export function isDictionaryValue(value: DynamicValue): value is DynamicObject {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}

export function defaultCreateId(prefix: string): string {
  const randomUuid = globalThis.crypto?.randomUUID;
  if (typeof randomUuid === "function") {
    return randomUuid.call(globalThis.crypto);
  }

  const randomChunk = Math.floor(Math.random() * 1_000_000)
    .toString(36)
    .padStart(4, "0");

  return `${prefix}_${Date.now().toString(36)}_${randomChunk}`;
}
