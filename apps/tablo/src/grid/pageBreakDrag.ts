/**
 * Sayfa Sonu Önizleme'de sayfa sonlarını ve yazdırma alanı kenarlarını
 * fareyle sürükleme: otomatik ya da el ile sayfa sonu yeni yerinde el ile
 * sayfa sonu olur; yazdırma alanının dışına sürüklenen son kaldırılır;
 * dış kenarı sürüklemek yazdırma alanını değiştirir.
 */
import type { Controller } from "../controller";
import { computePageBreaks } from "../io/print";
import type { PageSetup } from "../model/types";
import type { GridLayout } from "./layout";

export interface PageBreakDrag {
  axis: "row" | "col";
  /** Sürüklenen çizginin özgün yeri (satır/sütun kenarı). */
  from: number;
  kind: "break" | "start" | "end";
  /** Çizginin gideceği yer. */
  at: number;
}

const NEAR = 4;

/** Fare noktasındaki sayfa sonu ya da yazdırma alanı kenarı. */
export function pageBreakHit(ctl: Controller, L: GridLayout, x: number, y: number): PageBreakDrag | null {
  const sheet = ctl.sheet;
  if (sheet.viewMode !== "pageBreak" || x < L.headerW || y < L.headerH) return null;
  const pb = computePageBreaks(ctl.wb, sheet);
  const r = pb.range;
  if (!r) return null;
  if (x >= L.colX(r.c1) - NEAR && x <= L.colX(r.c2 + 1) + NEAR)
    for (const row of [r.r1, ...pb.rows])
      if (Math.abs(L.rowY(row) - y) <= NEAR) return { axis: "row", from: row, kind: row === r.r1 ? "start" : row === r.r2 + 1 ? "end" : "break", at: row };
  if (y >= L.rowY(r.r1) - NEAR && y <= L.rowY(r.r2 + 1) + NEAR)
    for (const col of [r.c1, ...pb.cols])
      if (Math.abs(L.colX(col) - x) <= NEAR) return { axis: "col", from: col, kind: col === r.c1 ? "start" : col === r.c2 + 1 ? "end" : "break", at: col };
  return null;
}

/** Fare imleci (sayfa sonu üzerindeyse). */
export function pageBreakCursor(ctl: Controller, L: GridLayout, x: number, y: number): string | null {
  const h = pageBreakHit(ctl, L, x, y);
  return h ? (h.axis === "row" ? "row-resize" : "col-resize") : null;
}

/** Sürüklerken çizgiyi en yakın satır/sütun kenarına taşır. */
export function pageBreakMove(L: GridLayout, d: PageBreakDrag, x: number, y: number): void {
  if (d.axis === "row") {
    const row = L.rowAtY(Math.max(L.headerH, y));
    d.at = y > L.rowY(row) + L.rowH(row) / 2 ? row + 1 : row;
  } else {
    const col = L.colAtX(Math.max(L.headerW, x));
    d.at = x > L.colX(col) + L.colW(col) / 2 ? col + 1 : col;
  }
}

/** Bırakınca sayfa sonunu ya da yazdırma alanını günceller (geri alınabilir). */
export function pageBreakDrop(ctl: Controller, d: PageBreakDrag): void {
  if (d.at === d.from) return;
  const sheet = ctl.sheet;
  const r = computePageBreaks(ctl.wb, sheet).range;
  if (!r) return;
  const set = (patch: Partial<PageSetup>, label: string) =>
    ctl.ops.sheetOp(sheet, label, () => {
      sheet.page = { ...sheet.page, ...patch };
    });
  const row = d.axis === "row";
  if (d.kind !== "break") {
    const area = { ...r };
    if (row) {
      if (d.kind === "start") area.r1 = Math.min(d.at, r.r2);
      else area.r2 = Math.max(d.at - 1, r.r1);
    } else if (d.kind === "start") area.c1 = Math.min(d.at, r.c2);
    else area.c2 = Math.max(d.at - 1, r.c1);
    set({ printArea: area }, "Yazdırma Alanı");
    return;
  }
  const key = row ? "rowBreaks" : "colBreaks";
  const start = row ? r.r1 : r.c1;
  const end = row ? r.r2 : r.c2;
  const list = (sheet.page[key] ?? []).filter((i) => i !== d.from);
  // Yazdırma alanının dışına sürüklenen sayfa sonu kaldırılır
  if (d.at > start && d.at <= end) list.push(d.at);
  set({ [key]: [...new Set(list)].sort((a, b) => a - b) }, d.at > start && d.at <= end ? "Sayfa Sonunu Taşı" : "Sayfa Sonunu Kaldır");
}

/** Sayfa Düzeni > Sonlar: etkin hücrenin üstüne / soluna sayfa sonu ekler ya da kaldırır; hepsini sıfırlar. */
export function insertPageBreak(ctl: Controller, remove: boolean): void {
  const sheet = ctl.sheet;
  const { row, col } = ctl.sel.active;
  const toggle = (list: number[] | undefined, i: number): number[] | undefined => {
    if (i <= 0) return list;
    const s = new Set(list ?? []);
    if (remove) s.delete(i);
    else s.add(i);
    return [...s].sort((a, b) => a - b);
  };
  ctl.ops.sheetOp(sheet, remove ? "Sayfa Sonunu Kaldır" : "Sayfa Sonu Ekle", () => {
    sheet.page = { ...sheet.page, rowBreaks: toggle(sheet.page.rowBreaks, row), colBreaks: toggle(sheet.page.colBreaks, col) };
  });
}

export function resetPageBreaks(ctl: Controller): void {
  const sheet = ctl.sheet;
  ctl.ops.sheetOp(sheet, "Tüm Sayfa Sonlarını Sıfırla", () => {
    sheet.page = { ...sheet.page, rowBreaks: undefined, colBreaks: undefined };
  });
}
