/** Veri > Gelişmiş: listeyi yerinde filtreler ya da uyan satırları başka bir yere kopyalar. */
import type { Controller } from "../controller";
import type { Rect } from "../model/types";
import { advancedMatches } from "../model/advancedFilter";
import { recomputeFilters } from "./data";

export interface AdvancedFilterOptions {
  list: Rect;
  criteria: Rect | null;
  /** Kopyalanacak yerin sol üst hücresi (yoksa yerinde filtre). */
  copyTo?: { row: number; col: number };
  unique: boolean;
}

/** Hata iletisi ya da null. */
export function applyAdvancedFilter(ctl: Controller, o: AdvancedFilterOptions): string | null {
  const { wb, sheet } = ctl;
  if (o.list.r2 <= o.list.r1) return "Liste aralığı başlık satırı ve en az bir veri satırı içermelidir.";
  if (o.criteria && o.criteria.r2 <= o.criteria.r1) return "Ölçüt aralığı başlık satırı ve en az bir ölçüt satırı içermelidir.";
  if (!o.copyTo) {
    ctl.ops.sheetOp(sheet, "Gelişmiş Filtre", () => {
      sheet.advFilter = { range: { ...o.list }, criteria: o.criteria ? { ...o.criteria } : null, unique: o.unique };
      recomputeFilters(ctl);
    });
    const shown = advancedMatches(wb, sheet, o.list, o.criteria, o.unique).length;
    ctl.status = `${(o.list.r2 - o.list.r1).toLocaleString("tr-TR")} kayıttan ${shown.toLocaleString("tr-TR")} tanesi bulundu`;
    ctl.emit();
    return null;
  }
  const rows = [o.list.r1, ...advancedMatches(wb, sheet, o.list, o.criteria, o.unique)];
  const t = o.copyTo;
  const width = o.list.c2 - o.list.c1 + 1;
  if (t.row <= o.list.r2 && t.row + rows.length - 1 >= o.list.r1 && t.col <= o.list.c2 && t.col + width - 1 >= o.list.c1) return "Hedef aralık liste aralığıyla çakışamaz.";
  const entries: { row: number; col: number; input: string | number | boolean | null; style?: number | null }[] = [];
  rows.forEach((r, i) => {
    for (let c = 0; c < width; c++) {
      const v = wb.getValue(sheet, r, o.list.c1 + c);
      entries.push({ row: t.row + i, col: t.col + c, input: typeof v === "object" && v !== null ? null : (v as string | number | boolean | null), style: wb.styleIdAt(sheet, r, o.list.c1 + c) || null });
    }
  });
  ctl.ops.setInputs(sheet, entries, "Gelişmiş Filtre");
  ctl.selectRange({ r1: t.row, c1: t.col, r2: t.row + rows.length - 1, c2: t.col + width - 1 });
  return null;
}
