/** Biçim okuyucularının ortak yardımcıları. */
import type { Rect } from "../model/types";
import type { SheetModel, WorkbookModel } from "../model/workbook";

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

/** Bu kadar hücreden büyük birleştirmeler kullanılan alana kırpılır (bozuk dosyada A1:XFD1048576 olabilir). */
export const BIG_MERGE_CELLS = 100_000;

/**
 * Büyük birleştirmeyi denetler: makul boyuttaki (en çok `limit` hücre, sayfa kenarına
 * dayanmayan) birleştirme olduğu gibi kalır; diğerleri sayfanın kullanılan alanına
 * kırpılır. Tek hücreye iner ya da yine çok büyükse null.
 */
export function clampMerge(r: Rect, used: { rows: number; cols: number }, limit = 1_000_000): Rect | null {
  const size = (x: Rect) => (x.r2 - x.r1 + 1) * (x.c2 - x.c1 + 1);
  if (size(r) <= limit && r.r2 < 1048575 && r.c2 < 16383) return r;
  const out = { r1: r.r1, c1: r.c1, r2: Math.min(r.r2, Math.max(r.r1, used.rows - 1)), c2: Math.min(r.c2, Math.max(r.c1, used.cols - 1)) };
  if (out.r1 === out.r2 && out.c1 === out.c2) return null;
  return size(out) > limit ? null : out;
}

/** Sayfanın kullanılan alanı (değerli ya da biçimli hücreler). */
export function usedExtent(wb: WorkbookModel, s: SheetModel): { rows: number; cols: number } {
  let rows = 1;
  let cols = 1;
  wb.engine.forEachCell(s.id, (r, c) => {
    if (r + 1 > rows) rows = r + 1;
    if (c + 1 > cols) cols = c + 1;
  });
  for (const k of s.styles.keys()) {
    const r = Math.floor(k / 16384);
    const c = k % 16384;
    if (r + 1 > rows) rows = r + 1;
    if (c + 1 > cols) cols = c + 1;
  }
  return { rows, cols };
}

/** Sayfadaki dev birleştirmeleri kullanılan alana kırpar (clampMerge); değişiklik varsa uyarı metni. */
export function fitMerges(wb: WorkbookModel, s: SheetModel): string | null {
  let used: { rows: number; cols: number } | undefined;
  let note: string | null = null;
  s.merges = s.merges.flatMap((m) => {
    if ((m.r2 - m.r1 + 1) * (m.c2 - m.c1 + 1) <= BIG_MERGE_CELLS && m.r2 < 1048575 && m.c2 < 16383) return [m];
    used ??= usedExtent(wb, s);
    const c = clampMerge(m, used);
    if (c === m) return [m];
    note = c ? "çok büyük birleştirilmiş alanlar (kullanılan alana kırpıldı)" : "çok büyük birleştirilmiş alanlar";
    return c ? [c] : [];
  });
  return note;
}

/** Okuma sırasında kullanıcıya anlaşılır hata. */
export class FormatError extends Error {
  /** Teknik ayrıntı (özgün kitaplık iletisi): "Ayrıntılar" bölümünde ve günlükte gösterilir. */
  detail?: string;
  constructor(message: string, detail?: string) {
    super(message);
    this.name = "FormatError";
    if (detail) this.detail = detail;
  }
}
