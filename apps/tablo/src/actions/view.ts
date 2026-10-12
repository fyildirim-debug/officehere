/** Görünüm eylemleri: bölmeleri dondurma, yakınlaştırma, görünüm kipleri, notlar arasında gezinme. */
import { keyCol, keyRow } from "@officehere/formula";
import type { Controller } from "../controller";
import type { ViewMode } from "../model/types";

export function freeze(ctl: Controller, kind: "panes" | "topRow" | "firstCol" | "none"): void {
  const sheet = ctl.sheet;
  const { row, col } = ctl.sel.active;
  const next =
    kind === "none"
      ? { rows: 0, cols: 0 }
      : kind === "topRow"
        ? { rows: 1, cols: 0 }
        : kind === "firstCol"
          ? { rows: 0, cols: 1 }
          : row === 0 && col === 0
            ? { rows: sheet.scroll.row + 10, cols: 0 }
            : { rows: row, cols: col };
  ctl.ops.sheetOp(sheet, kind === "none" ? "Bölmeleri Çöz" : "Bölmeleri Dondur", () => {
    sheet.frozen = next;
    if (kind !== "none") sheet.scroll = { row: 0, col: 0 };
  });
}

export function isFrozen(ctl: Controller): boolean {
  return ctl.sheet.frozen.rows > 0 || ctl.sheet.frozen.cols > 0;
}

export function setZoom(ctl: Controller, z: number): void {
  ctl.sheet.zoom = Math.max(10, Math.min(400, Math.round(z)));
  ctl.emit();
}

/** Seçimi pencereye sığdıracak yakınlaştırma. */
export function zoomToSelection(ctl: Controller): void {
  const s = ctl.sheet;
  const r = ctl.range;
  const w = s.cols.offset(Math.min(r.c2, 16383) + 1) - s.cols.offset(r.c1);
  const h = s.rows.offset(Math.min(r.r2, 1048575) + 1) - s.rows.offset(r.r1);
  const host = document.querySelector(".tb-canvas-wrap") as HTMLElement | null;
  const vw = (host?.clientWidth ?? window.innerWidth - 80) - 50;
  const vh = (host?.clientHeight ?? window.innerHeight - 300) - 30;
  setZoom(ctl, Math.floor(Math.min(vw / Math.max(1, w), vh / Math.max(1, h)) * 100));
  s.scroll = { row: r.r1, col: r.c1 };
  ctl.emit();
}

export function setViewMode(ctl: Controller, mode: ViewMode): void {
  ctl.sheet.viewMode = mode;
  ctl.emit();
}

/** Önceki/sonraki notlu hücreye gider (çalışma kitabı boyunca). */
export function gotoNote(ctl: Controller, dir: 1 | -1): void {
  const wb = ctl.wb;
  const n = wb.sheets.length;
  const cur = ctl.sel.active;
  const order = (r: number, c: number) => r * 16384 + c;
  for (let step = 0; step <= n; step++) {
    const idx = (((wb.activeIndex + dir * step) % n) + n) % n;
    const sheet = wb.sheets[idx];
    if (sheet.hidden) continue;
    const keys = [...sheet.notes.keys()].map((k) => ({ row: keyRow(k), col: keyCol(k) })).sort((a, b) => order(a.row, a.col) - order(b.row, b.col));
    if (!keys.length) continue;
    let target;
    if (step === 0) target = dir > 0 ? keys.find((k) => order(k.row, k.col) > order(cur.row, cur.col)) : [...keys].reverse().find((k) => order(k.row, k.col) < order(cur.row, cur.col));
    else target = dir > 0 ? keys[0] : keys[keys.length - 1];
    if (!target) continue;
    if (idx !== wb.activeIndex) ctl.activateSheet(idx);
    ctl.selectCell(target.row, target.col);
    return;
  }
  ctl.alert("Çalışma kitabında başka not yok.");
}

/** Yazım Denetimi (F7): düzenleyicide tarayıcının yazım denetimini açar/kapatır. */
export function toggleSpellcheck(ctl: Controller): void {
  ctl.spellcheck = !ctl.spellcheck;
  ctl.status = ctl.spellcheck ? "Yazım denetimi açık: hücre düzenlenirken hatalı sözcüklerin altı çizilir" : "";
  ctl.emit();
}
