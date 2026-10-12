/** Otomatik sığdır: sütun genişliği ve satır yüksekliği içeriğe göre. */
import { fontStack } from "@officehere/core";
import type { Controller } from "../controller";
import { displayText } from "../model/display";
import { fontCss, DEFAULT_SIZE } from "../model/styles";
import { measureText, wrapLines } from "./render";

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

export function fitRowHeight(ctl: Controller, row: number): number {
  const { wb, sheet } = ctl;
  let max = 20;
  wb.engine.forEachValue(sheet.id, (r, c, v) => {
    if (r !== row || v === null) return;
    const st = wb.styleAt(sheet, r, c);
    const size = (st.size ?? DEFAULT_SIZE) * (96 / 72);
    const lineH = size * 1.2;
    let lines = 1;
    const text = displayText(wb, sheet, r, c);
    if (st.wrap) lines = wrapLines(text, fontCss(st, 1, fontStack), Math.max(4, sheet.cols.size(c) - 6)).length;
    const h = Math.ceil(lines * lineH + 2.4);
    if (h > max) max = h;
  });
  // Stilli boş hücrelerin yazı tipi boyutu da hesaba katılır
  for (let c = 0; c < 64; c++) {
    const st = wb.styleAt(sheet, row, c);
    if (st.size && st.size > DEFAULT_SIZE) max = Math.max(max, Math.ceil(st.size * (96 / 72) * 1.2 + 2.4));
  }
  return Math.min(545, max);
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
    for (const r of rows) {
      const h = fitRowHeight(ctl, r);
      if (h === 20) sheet.rows.reset(r);
      else sheet.rows.set(r, h);
    }
  });
}
