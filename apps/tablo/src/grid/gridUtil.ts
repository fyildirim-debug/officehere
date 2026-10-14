/** Izgara yardımcıları: isabet testi, seçili satır/sütunlar, doldurma ipucu, köprü açma, koyu tema algılama. */
import { useEffect, useState } from "react";
import { formatValue } from "@officehere/formula";
import type { Controller } from "../controller";
import { filterAt, filtersOf } from "../actions/data";
import { MAX_COLS, MAX_ROWS } from "../model/rect";
import type { Rect } from "../model/types";
import type { SheetModel } from "../model/workbook";
import type { GridLayout } from "./layout";
import { checkboxRect } from "./renderCell";

export interface Hit {
  area: "corner" | "colHeader" | "rowHeader" | "cell" | "none";
  row: number;
  col: number;
  colBorder?: number;
  rowBorder?: number;
  fillHandle?: boolean;
  /** Tablonun sağ alt köşesindeki boyutlandırma tutamacı (tablo kimliği). */
  tableHandle?: string;
  selBorder?: boolean;
  filterButton?: boolean;
}

/** Ekran noktasının ızgaradaki karşılığı (başlıklar, kenarlıklar, doldurma tutamacı, filtre düğmesi). */
/** Fare imleci: başlık kenarı, doldurma tutamacı, taşıma kenarı, filtre düğmesi, köprü. */
export function cursorFor(hit: Hit, sheet: SheetModel): string {
  if (hit.area === "colHeader") return hit.colBorder !== undefined ? "col-resize" : "s-resize";
  if (hit.area === "rowHeader") return hit.rowBorder !== undefined ? "row-resize" : "e-resize";
  if (hit.area === "corner") return "se-resize";
  if (hit.fillHandle) return "crosshair";
  if (hit.tableHandle) return "nwse-resize";
  if (hit.selBorder) return "move";
  if (hit.filterButton) return "default";
  return hit.area === "cell" && sheet.linkAt(hit.row, hit.col) ? "pointer" : "cell";
}

/** Nokta hücredeki onay kutusunun üzerinde mi (kutunun çevresinde birkaç piksel pay)? */
export function checkboxHit(ctl: Controller, L: GridLayout, row: number, col: number, x: number, y: number): boolean {
  const st = ctl.wb.styleAt(ctl.sheet, row, col);
  if (!st.checkbox || ctl.sheet.showFormulas) return false;
  const m = ctl.sheet.mergeAt(row, col);
  const rc = m ? L.cellRect(m.r1, m.c1, m.r2, m.c2) : L.cellRect(row, col);
  const b = checkboxRect(rc, st, L.z);
  const pad = 3 * L.z;
  return x >= b.x - pad && x <= b.x + b.s + pad && y >= b.y - pad && y <= b.y + b.s + pad;
}

export function hitTestGrid(L: GridLayout, ctl: Controller, sheet: SheetModel, x: number, y: number): Hit {
  const z = L.z;
  if ((L.outlineW && x < L.outlineW) || (L.outlineH && y < L.outlineH)) return { area: "none", row: -1, col: -1 };
  if (x < L.headerW && y < L.headerH) return { area: "corner", row: -1, col: -1 };
  if (y < L.headerH) {
    const col = L.colAtX(x);
    const cx = L.colX(col);
    const cw = L.colW(col);
    if (Math.abs(x - (cx + cw)) <= 4) return { area: "colHeader", row: -1, col, colBorder: col };
    if (Math.abs(x - cx) <= 4 && col > 0) {
      let prev = col - 1;
      while (prev > 0 && sheet.cols.isHidden(prev) && x < cx + 2) prev--;
      return { area: "colHeader", row: -1, col, colBorder: sheet.cols.isHidden(col - 1) && x >= cx ? col - 1 : prev };
    }
    return { area: "colHeader", row: -1, col };
  }
  if (x < L.headerW) {
    const row = L.rowAtY(y);
    const ry = L.rowY(row);
    const rh = L.rowH(row);
    if (Math.abs(y - (ry + rh)) <= 3) return { area: "rowHeader", row, col: -1, rowBorder: row };
    if (Math.abs(y - ry) <= 3 && row > 0) return { area: "rowHeader", row, col: -1, rowBorder: row - 1 };
    return { area: "rowHeader", row, col: -1 };
  }
  if (x > L.width || y > L.height) return { area: "none", row: -1, col: -1 };
  const row = L.rowAtY(y);
  const col = L.colAtX(x);
  const hit: Hit = { area: "cell", row, col };
  const last = ctl.range;
  const lr = L.cellRect(last.r1, last.c1, Math.min(last.r2, MAX_ROWS - 1), Math.min(last.c2, MAX_COLS - 1));
  if (ctl.sel.ranges.length === 1 && Math.abs(x - (lr.x + lr.w)) <= 5 && Math.abs(y - (lr.y + lr.h)) <= 5) hit.fillHandle = true;
  else if (ctl.sel.ranges.length === 1 && !ctl.edit) {
    const onV = (Math.abs(x - lr.x) <= 3 || Math.abs(x - (lr.x + lr.w)) <= 3) && y >= lr.y - 3 && y <= lr.y + lr.h + 3;
    const onH = (Math.abs(y - lr.y) <= 3 || Math.abs(y - (lr.y + lr.h)) <= 3) && x >= lr.x - 3 && x <= lr.x + lr.w + 3;
    if ((onV || onH) && last.r2 < MAX_ROWS - 1 && last.c2 < MAX_COLS - 1) hit.selBorder = true;
  }
  // Tablo boyutlandırma tutamacı (doldurma tutamacı öncelikli)
  if (!hit.fillHandle && !ctl.edit)
    for (const t of sheet.tables) {
      const tr = L.cellRect(t.ref.r2, t.ref.c2);
      const cx = tr.x + tr.w;
      const cy = tr.y + tr.h;
      if (x <= cx + 1 && x >= cx - 7 * z && y <= cy + 1 && y >= cy - 7 * z) {
        hit.tableHandle = t.id;
        hit.selBorder = false;
        break;
      }
    }
  // Filtre düğmesi
  if (filterAt(sheet, row, col) || (sheet.pivots.length && sheet.pivots.some((p) => p.buttons?.some((b) => b.row === row && b.col === col)))) {
    const rc = L.cellRect(row, col);
    const s = Math.min(rc.h, 18 * z);
    if (x >= rc.x + rc.w - s - 1 && y >= rc.y) hit.filterButton = true;
  }
  return hit;
}

export function selectedLines(ctl: Controller, axis: "row" | "col", fallback: number): number[] {
  const out: number[] = [];
  for (const r of ctl.sel.ranges) {
    const full = axis === "col" ? r.r1 === 0 && r.r2 === MAX_ROWS - 1 : r.c1 === 0 && r.c2 === MAX_COLS - 1;
    if (!full) continue;
    const [a, b] = axis === "col" ? [r.c1, r.c2] : [r.r1, Math.min(r.r2, r.r1 + 50000)];
    if (fallback < a || fallback > b) continue;
    for (let i = a; i <= b; i++) out.push(i);
  }
  return out.length ? out : [fallback];
}

/**
 * Doldurma tutamacı ipucu: bırakınca hedefin en uç hücresine yazılacak değer (doldurmayla aynı
 * seri kuralı). ctrl: Ctrl basılı sürükleme (tek sayı birer artar, diğerleri kopyalanır).
 * Formül kopyalanırken ve içeri sürüklemede ipucu yoktur.
 */
export function fillTipText(ctl: Controller, src: Rect, target: Rect, ctrl = false): string {
  const p = ctl.ops.fillPreview(ctl.sheet, src, target, ctrl ? "alt" : "series");
  if (!p) return "";
  const v = p.input;
  if (v === null || (typeof v === "object" && "formula" in v)) return "";
  return formatValue(v, p.numFmt).text;
}

export function openLink(ctl: Controller, link: string): void {
  if (link.startsWith("#")) {
    ctl.goTo(link.slice(1));
    return;
  }
  window.open(link, "_blank", "noopener");
}

/** Arayüz teması koyu mu (Koyu Gri / Siyah ya da sistem koyu)? */
export function useDarkUi(): boolean {
  const read = () => {
    if (typeof document === "undefined") return false;
    const t = document.documentElement.getAttribute("data-oh-theme");
    if (t === "dark" || t === "black") return true;
    return t === "system" && typeof matchMedia !== "undefined" && matchMedia("(prefers-color-scheme: dark)").matches;
  };
  const [dark, setDark] = useState(read);
  useEffect(() => {
    const update = () => setDark(read());
    const mo = new MutationObserver(update);
    mo.observe(document.documentElement, { attributes: true, attributeFilter: ["data-oh-theme"] });
    const mq = typeof matchMedia !== "undefined" ? matchMedia("(prefers-color-scheme: dark)") : null;
    mq?.addEventListener("change", update);
    return () => {
      mo.disconnect();
      mq?.removeEventListener("change", update);
    };
  }, []);
  return dark;
}

/** Ölçütü olan filtrelerin veri satırları (satır numaraları mavi gösterilir); filtre yoksa undefined. */
export function activeFilterRows(sheet: SheetModel): ((row: number) => boolean) | undefined {
  const fs = filtersOf(sheet).filter((f) => Object.keys(f.columns).length > 0);
  if (!fs.length) return undefined;
  return (row) => fs.some((f) => row > f.range.r1 && row <= f.range.r2);
}
