/**
 * Otomatik sığdırma: sütun genişliği ve satır yüksekliği içeriğe göre.
 * Excel gibi elle boyutlandırılmamış satırlar yazı tipi boyutu ya da metin
 * kaydırma değişince kendiliğinden büyür/küçülür; biçimli sayı girilen
 * varsayılan genişlikteki sütun sayıyı gösterecek kadar genişler.
 */
import { fontStack } from "@officehere/core";
import { keyCol } from "@officehere/formula";
import type { Controller } from "../controller";
import { displayText } from "../model/display";
import { MAX_COLS } from "../model/rect";
import { fontCss, DEFAULT_SIZE } from "../model/styles";
import type { Rect } from "../model/types";
import type { SheetModel, WorkbookModel } from "../model/workbook";
import { measureText, wrapLines } from "./render";

/** Varsayılan satır yüksekliği (Calibri 11 = 15 pt = 20 px). */
const DEFAULT_ROW_PX = 20;

export function fitColumnWidth(ctl: Controller, col: number): number | null {
  const { wb, sheet } = ctl;
  let max = 0;
  wb.engine.forEachValue(sheet.id, (row, c, v) => {
    if (c !== col || v === null) return;
    if (sheet.mergeAt(row, c)) return;
    const st = wb.styleAt(sheet, row, c);
    if (st.wrap) return;
    const text = sheet.showFormulas ? wb.editText(sheet, row, c) : displayText(wb, sheet, row, c);
    const w = Math.max(...text.split("\n").map((l) => measureText(l, fontCss(st, 1, fontStack)))) + 8 + (st.indent ?? 0) * 9;
    if (w > max) max = w;
  });
  // Filtre düğmesi için yer
  if (sheet.autoFilter && col >= sheet.autoFilter.range.c1 && col <= sheet.autoFilter.range.c2) max += 18;
  return max > 0 ? Math.min(1790, Math.ceil(max)) : null;
}

/** Satırda içerik ya da stil taşıyan son sütun (satır yüksekliği taraması için üst sınır). */
function lastUsedCol(wb: WorkbookModel, sheet: SheetModel): number {
  let last = wb.dataExtent(sheet).cols - 1;
  for (const k of sheet.styles.keys()) {
    const c = keyCol(k);
    if (c > last) last = c;
  }
  return Math.min(last, 700);
}

/** Bir satırın içeriğe göre yüksekliği (px, %100). */
export function rowContentHeight(wb: WorkbookModel, sheet: SheetModel, row: number, lastCol = lastUsedCol(wb, sheet)): number {
  let max = DEFAULT_ROW_PX;
  for (let c = 0; c <= lastCol; c++) {
    const hasValue = wb.engine.hasInput(sheet.id, row, c) || !wb.isEmpty(sheet, row, c);
    const st = wb.styleAt(sheet, row, c);
    if (!hasValue && !st.size) continue;
    const m = sheet.mergeAt(row, c);
    if (m && (m.r1 !== m.r2 || m.c1 !== m.c2)) continue;
    const size = (st.size ?? DEFAULT_SIZE) * (96 / 72);
    const lineH = size * 1.2;
    let lines = 1;
    if (hasValue && st.wrap) {
      const text = displayText(wb, sheet, row, c);
      lines = Math.max(1, wrapLines(text, fontCss(st, 1, fontStack), Math.max(4, sheet.cols.size(c) - 6)).length);
    }
    let h = Math.ceil(lines * lineH + 2.4);
    // Döndürülmüş metin: yükseklik metin genişliğinin dikey bileşeni kadar büyür (45° → w·sin45 + h·cos45)
    const rot = st.rotation ?? 0;
    if (hasValue && rot && !st.wrap) {
      const text = displayText(wb, sheet, row, c);
      const w = measureText(text, fontCss(st, 1, fontStack));
      if (rot === 255) h = Math.ceil(text.length * lineH + 4);
      else {
        const a = ((rot > 90 ? rot - 90 : rot) * Math.PI) / 180;
        h = Math.ceil(w * Math.sin(a) + lineH * Math.cos(a) + 6);
      }
    }
    if (h > max) max = h;
  }
  return Math.min(545, max);
}

export function fitRowHeight(ctl: Controller, row: number): number {
  return rowContentHeight(ctl.wb, ctl.sheet, row);
}

/**
 * İçerik/stil değişen aralıklardaki otomatik satırların yeni yüksekliklerini
 * uygular. Değişen satırları [satır, eski boyut, yeni boyut] olarak döndürür
 * (geri alma kaydı için; undefined = varsayılan).
 */
export function adjustRowHeights(wb: WorkbookModel, sheet: SheetModel, rects: Rect[]): [number, number | undefined, number | undefined][] {
  const rows = new Set<number>();
  const ext = wb.dataExtent(sheet);
  for (const r of rects) {
    const r2 = Math.min(r.r2, r.c1 === 0 && r.c2 === MAX_COLS - 1 ? r.r2 : Math.max(r.r1, ext.rows - 1), r.r1 + 5000);
    for (let i = r.r1; i <= r2; i++) rows.add(i);
  }
  if (!rows.size || rows.size > 5000) return [];
  const lastCol = lastUsedCol(wb, sheet);
  const out: [number, number | undefined, number | undefined][] = [];
  for (const row of rows) {
    if (sheet.rows.isManual(row) || sheet.rows.isHidden(row)) continue;
    const h = rowContentHeight(wb, sheet, row, lastCol);
    const old = sheet.rows.isCustom(row) ? sheet.rows.rawSize(row) : undefined;
    const next = h === DEFAULT_ROW_PX ? undefined : h;
    if (old === next) continue;
    if (next === undefined) sheet.rows.reset(row);
    else sheet.rows.set(row, next, false);
    out.push([row, old, next]);
  }
  return out;
}

/** Biçimli sayı varsayılan genişlikteki sütuna sığmıyorsa sütunu genişletir (Excel davranışı). */
export function widenForNumber(wb: WorkbookModel, sheet: SheetModel, row: number, col: number): number | null {
  if (sheet.cols.isManual(col) || sheet.mergeAt(row, col)) return null;
  const v = wb.getValue(sheet, row, col);
  const st = wb.styleAt(sheet, row, col);
  if (typeof v !== "number" || !st.numFmt || st.numFmt === "General" || st.wrap || st.shrink) return null;
  const text = displayText(wb, sheet, row, col);
  const need = Math.ceil(measureText(text, fontCss(st, 1, fontStack)) + 8 + (st.indent ?? 0) * 9);
  if (need <= sheet.cols.size(col)) return null;
  return Math.min(1790, need);
}

export function autoFitCols(ctl: Controller, cols: number[]): void {
  const sheet = ctl.sheet;
  ctl.ops.sheetOp(sheet, "Sütun Genişliğini Otomatik Sığdır", () => {
    for (const c of cols) {
      const w = fitColumnWidth(ctl, c);
      if (w === null) sheet.cols.reset(c);
      else sheet.cols.set(c, w);
    }
  });
}

export function autoFitRows(ctl: Controller, rows: number[]): void {
  const sheet = ctl.sheet;
  ctl.ops.sheetOp(sheet, "Satır Yüksekliğini Otomatik Sığdır", () => {
    const lastCol = lastUsedCol(ctl.wb, sheet);
    for (const r of rows) {
      const h = rowContentHeight(ctl.wb, sheet, r, lastCol);
      // Otomatik sığdırılan satır yeniden "otomatik" olur
      if (h === DEFAULT_ROW_PX) sheet.rows.reset(r);
      else sheet.rows.set(r, h, false);
    }
  });
}

