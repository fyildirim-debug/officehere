/**
 * Hücre içi onay kutuları (Ekle > Onay Kutusu): onay kutusu bir hücre biçimidir,
 * değeri DOĞRU/YANLIŞ'tır. Tıklama ya da Ara Çubuğu durumunu değiştirir; Delete
 * işaretli kutuları boşaltır, boş kutularda ikinci Delete kutuyu kaldırır.
 */
import { keyCol, keyRow } from "@officehere/formula";
import type { Controller } from "../controller";
import { intersection, isFullCols, isFullRows, rect } from "../model/rect";
import type { CellPos, Rect } from "../model/types";

/** Seçimdeki onay kutulu hücreler (tüm satır/sütun seçimleri kullanılan alana kırpılır). */
export function checkboxCells(ctl: Controller, ranges: Rect[] = ctl.sel.ranges): CellPos[] {
  const { wb, sheet } = ctl;
  const used = wb.usedRange(sheet);
  const out: CellPos[] = [];
  for (const r0 of ranges) {
    const r = isFullCols(r0) || isFullRows(r0) ? (used ? intersection(r0, used) : null) : r0;
    if (!r) continue;
    // Yalnız biçimli ya da değerli hücreler gezilir; satır/sütun stilleri de hesaba katılır
    const cols = r.c2 - r.c1 + 1;
    const rows = r.r2 - r.r1 + 1;
    if (rows * cols <= 200000) {
      for (let row = r.r1; row <= r.r2; row++) for (let col = r.c1; col <= r.c2; col++) if (wb.styleAt(sheet, row, col).checkbox) out.push({ row, col });
    } else for (const k of wb.cellKeysIn(sheet, r)) if (wb.styleAt(sheet, keyRow(k), keyCol(k)).checkbox) out.push({ row: keyRow(k), col: keyCol(k) });
  }
  return out;
}

export function isCheckboxCell(ctl: Controller, row: number, col: number): boolean {
  return !!ctl.wb.styleAt(ctl.sheet, row, col).checkbox;
}

/** Seçime onay kutusu ekler: boş hücreler YANLIŞ olur, DOĞRU/YANLIŞ ve diğer değerler korunur. */
export function insertCheckboxes(ctl: Controller): void {
  if (!ctl.guardSelection()) return;
  const { wb, sheet } = ctl;
  const ranges = ctl.sel.ranges.map((r) => ({ ...r }));
  const mark = ctl.ops.history.mark();
  ctl.ops.applyStyle(sheet, ranges, { checkbox: true }, "Onay Kutusu");
  // Tüm satır/sütunda YANLIŞ yalnız kullanılan alana yazılır
  const used = wb.usedRange(sheet);
  const entries: { row: number; col: number; input: boolean }[] = [];
  for (const r0 of ranges) {
    const r = isFullCols(r0) || isFullRows(r0) ? (used ? intersection(r0, used) : null) : r0;
    if (!r) continue;
    for (let row = r.r1; row <= Math.min(r.r2, r.r1 + 100000); row++)
      for (let col = r.c1; col <= Math.min(r.c2, r.c1 + 1000); col++) {
        if (sheet.rows.isHidden(row) && sheet.filterHidden.has(row)) continue;
        if (!wb.engine.hasInput(sheet.id, row, col)) entries.push({ row, col, input: false });
      }
  }
  if (entries.length) ctl.ops.setInputs(sheet, entries, "Onay Kutusu");
  ctl.ops.history.squash(mark, "Onay Kutusu");
  ctl.recordRepeat("Onay Kutusu", () => insertCheckboxes(ctl));
}

/** Onay kutularını değiştirir: etkin hücre işaretsizse hepsi işaretlenir, işaretliyse hepsi kaldırılır. */
export function toggleCheckboxes(ctl: Controller, cells: CellPos[] = checkboxCells(ctl)): boolean {
  if (!cells.length) return false;
  if (!ctl.guardSelection()) return true;
  const { wb, sheet } = ctl;
  const { row, col } = ctl.sel.active;
  const ref = cells.some((c) => c.row === row && c.col === col) ? { row, col } : cells[0];
  const next = wb.getValue(sheet, ref.row, ref.col) !== true;
  ctl.ops.setInputs(sheet, cells.map((c) => ({ row: c.row, col: c.col, input: next })), next ? "Onay Kutusunu İşaretle" : "Onay Kutusu İşaretini Kaldır");
  return true;
}

/** Tek bir onay kutusunu tıklamayla değiştirir. */
export function toggleCheckboxAt(ctl: Controller, row: number, col: number): void {
  if (ctl.isLocked(row, col)) {
    ctl.alert("Değiştirmeye çalıştığınız hücre korumalı bir sayfada. Değişiklik yapmak için Gözden Geçir sekmesinde Sayfa Korumasını Kaldır'ı tıklatın.");
    return;
  }
  const next = ctl.wb.getValue(ctl.sheet, row, col) !== true;
  ctl.ops.setInputs(ctl.sheet, [{ row, col, input: next }], next ? "Onay Kutusunu İşaretle" : "Onay Kutusu İşaretini Kaldır");
}

/**
 * Delete: seçimde işaretli onay kutusu varsa yalnız işaretleri kaldırılır (YANLIŞ);
 * hepsi boşsa onay kutuları kaldırılıp içerik temizlenir. Onay kutusu yoksa false.
 */
export function deleteOnCheckboxes(ctl: Controller): boolean {
  const cells = checkboxCells(ctl);
  if (!cells.length) return false;
  if (!ctl.guardSelection()) return true;
  const { wb, sheet } = ctl;
  const checked = cells.filter((c) => wb.getValue(sheet, c.row, c.col) === true);
  const mark = ctl.ops.history.mark();
  // Seçimdeki diğer hücrelerin içeriği her durumda temizlenir
  const others = ctl.sel.ranges;
  if (checked.length) {
    ctl.ops.clear(sheet, others, "contents");
    ctl.ops.setInputs(sheet, cells.map((c) => ({ row: c.row, col: c.col, input: false })), "İçeriği Temizle");
  } else {
    ctl.ops.clear(sheet, others, "contents");
    ctl.ops.applyStyle(sheet, cells.map((c) => rect(c.row, c.col)), { checkbox: undefined }, "Onay Kutusunu Kaldır");
  }
  ctl.ops.history.squash(mark, checked.length ? "İçeriği Temizle" : "Onay Kutusunu Kaldır");
  return true;
}

