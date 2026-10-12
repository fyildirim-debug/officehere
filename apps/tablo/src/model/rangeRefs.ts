/**
 * Nesnelerin (grafik, mini grafik) metin olarak tuttuğu veri aralıklarını
 * satır/sütun ekleme-silme ve sayfa yeniden adlandırmada güncel tutar.
 */
import { formatQualifiedRange, parseQualifiedRange } from "@officehere/formula";
import { shiftRect } from "./rect";

function withAbs(text: string, abs: boolean): string {
  if (!abs) return text;
  const bang = text.lastIndexOf("!");
  const head = bang >= 0 ? text.slice(0, bang + 1) : "";
  return head + text.slice(bang + 1).replace(/([A-Z]{1,3})(\d+)/g, "$$$1$$$2");
}

/** Aralık metninin gösterdiği sayfa adı (nitelenmemişse sahibi). */
export function rangeSheetName(text: string, owner: string): string | null {
  const q = parseQualifiedRange(text.replace(/^=/, ""));
  return q ? (q.sheet ?? owner) : null;
}

/**
 * Hedef sayfadaki ekleme/silmeye göre aralığı kaydırır. Aralık hedef sayfayı
 * göstermiyorsa ya da tümüyle silindiyse metin değişmez.
 */
export function shiftRangeText(text: string, owner: string, target: string, axis: "row" | "col", at: number, count: number): string {
  const q = parseQualifiedRange(text.replace(/^=/, ""));
  if (!q) return text;
  const sheet = q.sheet ?? owner;
  if (sheet.toLocaleUpperCase("tr-TR") !== target.toLocaleUpperCase("tr-TR")) return text;
  const r = shiftRect(q.range, axis, at, count);
  if (!r) return text;
  return withAbs(formatQualifiedRange(q.sheet, r), text.includes("$"));
}

/** Sayfa yeniden adlandırıldığında nitelenmiş aralıktaki adı değiştirir. */
export function renameRangeText(text: string, oldName: string, newName: string): string {
  const q = parseQualifiedRange(text.replace(/^=/, ""));
  if (!q?.sheet || q.sheet.toLocaleUpperCase("tr-TR") !== oldName.toLocaleUpperCase("tr-TR")) return text;
  return withAbs(formatQualifiedRange(newName, q.range), text.includes("$"));
}
