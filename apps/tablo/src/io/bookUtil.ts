/** Biçim okuyucularının ortak yardımcıları. */
import type { WorkbookModel } from "../model/workbook";

/** Dosyadaki sayfa adını geçerli ve benzersiz bir ada çevirir (31 karakter, yasak karakterler "_"). */
export function safeSheetName(wb: WorkbookModel, raw: string): string {
  let base = (raw || "Sayfa").replace(/[\\/?*[\]:]/g, "_").replace(/^'+|'+$/g, "").trim().slice(0, 31) || "Sayfa";
  if (!wb.sheetByName(base)) return base;
  for (let n = 2; n < 10000; n++) {
    const suffix = ` (${n})`;
    const cand = base.slice(0, 31 - suffix.length) + suffix;
    if (!wb.sheetByName(cand)) return cand;
  }
  base = wb.nextSheetName();
  return base;
}

/** Bir hücreye yazılacak en çok hücre sayısı (bozuk/kötü niyetli dosyalara karşı). */
export const CELL_BUDGET = 4_000_000;

/** Okuma sırasında kullanıcıya anlaşılır hata. */
export class FormatError extends Error {
  constructor(message: string) {
    super(message);
    this.name = "FormatError";
  }
}
