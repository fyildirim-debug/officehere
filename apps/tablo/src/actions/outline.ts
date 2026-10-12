/** Anahat eylemleri (geri alınabilir): gruplandır, grubu çöz, ayrıntıyı göster/gizle, düzey, alt toplam. */
import type { Controller } from "../controller";
import { isFullCols, isFullRows } from "../model/rect";
import { applySubtotals, groupAt, groupItems, setGroupCollapsed, showLevel, subtotalRows, ungroupItems, type OutlineAxis, type OutlineGroup, type SubtotalSpec } from "../model/outline";
import { hitTestOutline } from "../grid/renderOutline";
import type { GridLayout } from "../grid/layout";

function axisSizes(ctl: Controller, axis: OutlineAxis) {
  return axis === "row" ? ctl.sheet.rows : ctl.sheet.cols;
}

/** Seçimden eksen: tüm sütunlar → sütun, aksi halde satır (iletişim kutusu yoksa). */
export function selectionAxis(ctl: Controller): OutlineAxis | null {
  const r = ctl.range;
  if (isFullCols(r)) return "col";
  if (isFullRows(r)) return "row";
  return null;
}

/** Gruplandır (Alt+Shift+Sağ). Eksen belirsizse iletişim kutusu sorar. */
export function groupSelection(ctl: Controller, axis?: OutlineAxis): void {
  const ax = axis ?? selectionAxis(ctl);
  if (!ax) {
    ctl.openDialog("groupAxis", { mode: "group" });
    return;
  }
  const r = ctl.range;
  const sheet = ctl.sheet;
  ctl.ops.sheetOp(sheet, "Gruplandır", () => {
    const a = axisSizes(ctl, ax);
    if (ax === "row") groupItems(a, r.r1, Math.min(r.r2, r.r1 + 100000));
    else groupItems(a, r.c1, Math.min(r.c2, r.c1 + 16000));
  });
}

/** Grubu Çöz (Alt+Shift+Sol). */
export function ungroupSelection(ctl: Controller, axis?: OutlineAxis): void {
  const ax = axis ?? selectionAxis(ctl);
  if (!ax) {
    ctl.openDialog("groupAxis", { mode: "ungroup" });
    return;
  }
  const r = ctl.range;
  const sheet = ctl.sheet;
  ctl.ops.sheetOp(sheet, "Grubu Çöz", () => {
    const a = axisSizes(ctl, ax);
    if (ax === "row") ungroupItems(a, r.r1, Math.min(r.r2, r.r1 + 100000));
    else ungroupItems(a, r.c1, Math.min(r.c2, r.c1 + 16000));
  });
  ctl.ops.syncHiddenRows(sheet);
}

/** Anahattı Temizle. */
export function clearOutline(ctl: Controller): void {
  const sheet = ctl.sheet;
  if (!sheet.rows.maxLevel() && !sheet.cols.maxLevel()) return;
  ctl.ops.sheetOp(sheet, "Anahattı Temizle", () => {
    for (const a of [sheet.rows, sheet.cols]) {
      const hidden = a.levelItems().map(([i]) => i).filter((i) => a.isHidden(i));
      a.setHiddenMany(hidden, false);
      a.clearOutline();
    }
    ctl.ops.syncHiddenRows(sheet);
  });
}

export function toggleGroup(ctl: Controller, axis: OutlineAxis, g: OutlineGroup, collapse: boolean): void {
  const sheet = ctl.sheet;
  ctl.ops.sheetOp(sheet, collapse ? "Ayrıntıyı Gizle" : "Ayrıntıyı Göster", () => {
    setGroupCollapsed(axisSizes(ctl, axis), g, collapse);
    ctl.ops.syncHiddenRows(sheet);
  });
}

export function showOutlineLevel(ctl: Controller, axis: OutlineAxis, level: number): void {
  const sheet = ctl.sheet;
  ctl.ops.sheetOp(sheet, "Anahat Düzeyi", () => {
    showLevel(axisSizes(ctl, axis), level);
    ctl.ops.syncHiddenRows(sheet);
  });
}

/** Ayrıntıyı Göster/Gizle: etkin hücrenin grubu (önce satırlar, sonra sütunlar). */
export function showDetail(ctl: Controller, show: boolean): void {
  const { row, col } = ctl.sel.active;
  for (const axis of ["row", "col"] as const) {
    const g = groupAt(axisSizes(ctl, axis), axis === "row" ? row : col);
    if (g) {
      if (g.collapsed !== !show) toggleGroup(ctl, axis, g, !show);
      return;
    }
  }
}

/** Izgaradaki anahat şeritlerine tıklama; işlendiyse true. */
export function handleOutlineClick(ctl: Controller, L: GridLayout, x: number, y: number): boolean {
  const hit = hitTestOutline(L, ctl.sheet, x, y);
  if (!hit) return (L.outlineW > 0 && x < L.outlineW) || (L.outlineH > 0 && y < L.outlineH);
  if (hit.kind === "level") showOutlineLevel(ctl, hit.axis, hit.level);
  else toggleGroup(ctl, hit.axis, hit.group, !hit.group.collapsed);
  return true;
}

/** Alt Toplam: satırlar eklenir ve düzeyler kurulur (tek geri alma adımı). */
export function addSubtotals(ctl: Controller, spec: SubtotalSpec, replace: boolean): void {
  const { wb, sheet } = ctl;
  ctl.ops.snapshotOp("Alt Toplam", () => {
    let range = spec.range;
    if (replace) range = removeSubtotalRows(ctl, range);
    const insert = (at: number, count: number) => {
      wb.engine.insertRows(sheet.id, at, count);
      sheet.shiftStructure("row", at, count);
      wb.shiftDataRefs(sheet, "row", at, count);
    };
    const out = applySubtotals(wb, sheet, { ...spec, range }, insert);
    ctl.sheet.selection = { ranges: [out], active: { row: out.r1, col: out.c1 }, anchor: { row: out.r1, col: out.c1 } };
  });
}

/** Mevcut alt toplam satırlarını siler ve satır anahattını temizler; yeni aralığı döndürür. */
function removeSubtotalRows(ctl: Controller, r: { r1: number; c1: number; r2: number; c2: number }) {
  const { wb, sheet } = ctl;
  const rows = subtotalRows(wb, sheet, r);
  for (let i = rows.length - 1; i >= 0; i--) {
    wb.engine.deleteRows(sheet.id, rows[i], 1);
    sheet.shiftStructure("row", rows[i], -1);
    wb.shiftDataRefs(sheet, "row", rows[i], -1);
  }
  const r2 = r.r2 - rows.length;
  const show: number[] = [];
  for (let i = r.r1; i <= r2; i++) {
    if (sheet.rows.isHidden(i)) show.push(i);
    sheet.rows.setLevel(i, 0);
    sheet.rows.setCollapsed(i, false);
  }
  sheet.rows.setHiddenMany(show, false);
  return { ...r, r2 };
}

/** Alt Toplam > Tümünü Kaldır. */
export function removeSubtotals(ctl: Controller, r: { r1: number; c1: number; r2: number; c2: number }): void {
  ctl.ops.snapshotOp("Alt Toplamları Kaldır", () => {
    const out = removeSubtotalRows(ctl, r);
    ctl.sheet.selection = { ranges: [out], active: { row: out.r1, col: out.c1 }, anchor: { row: out.r1, col: out.c1 } };
  });
}
