/** Seçim ve gezinme yardımcıları (birleştirilmiş hücreler, Ctrl+ok, gizli satırlar). */
import { MAX_COLS, MAX_ROWS, contains, rect, union } from "./rect";
import type { CellPos, Rect, Selection } from "./types";
import type { SheetModel, WorkbookModel } from "./workbook";

/** Aralığı içine aldığı birleştirmeleri kapsayacak şekilde genişletir. */
export function expandForMerges(sheet: SheetModel, r: Rect): Rect {
  if (!sheet.merges.length) return r;
  let out = r;
  let changed = true;
  let guard = 0;
  while (changed && guard++ < 20) {
    changed = false;
    for (const m of sheet.merges) {
      if (m.r1 <= out.r2 && out.r1 <= m.r2 && m.c1 <= out.c2 && out.c1 <= m.c2) {
        const u = union(out, m);
        if (u.r1 !== out.r1 || u.c1 !== out.c1 || u.r2 !== out.r2 || u.c2 !== out.c2) {
          out = u;
          changed = true;
        }
      }
    }
  }
  return out;
}

export function singleSelection(sheet: SheetModel, row: number, col: number): Selection {
  const m = sheet.mergeAt(row, col);
  const r = m ?? rect(row, col);
  const active = m ? { row: m.r1, col: m.c1 } : { row, col };
  return { ranges: [r], active, anchor: active };
}

function visibleStep(sheet: SheetModel, axis: "row" | "col", from: number, d: number): number {
  const ax = axis === "row" ? sheet.rows : sheet.cols;
  const max = axis === "row" ? MAX_ROWS - 1 : MAX_COLS - 1;
  let i = from + d;
  while (i >= 0 && i <= max && ax.isHidden(i)) i += d;
  if (i < 0 || i > max) return from;
  return i;
}

/** Ok tuşuyla bir hücre hareket (birleştirmeler atlanır). */
export function stepCell(sheet: SheetModel, pos: CellPos, dr: number, dc: number): CellPos {
  let { row, col } = pos;
  const m = sheet.mergeAt(row, col);
  if (dr > 0) row = visibleStep(sheet, "row", m ? m.r2 : row, 1);
  else if (dr < 0) row = visibleStep(sheet, "row", m ? m.r1 : row, -1);
  if (dc > 0) col = visibleStep(sheet, "col", m ? m.c2 : col, 1);
  else if (dc < 0) col = visibleStep(sheet, "col", m ? m.c1 : col, -1);
  if (m && dr !== 0) col = pos.col;
  if (m && dc !== 0) row = pos.row;
  const m2 = sheet.mergeAt(row, col);
  if (m2) return { row: m2.r1, col: m2.c1 };
  return { row, col };
}

/** Ctrl+ok: veri bölgesinin kenarına atla. */
export function jumpCell(wb: WorkbookModel, sheet: SheetModel, pos: CellPos, dr: number, dc: number): CellPos {
  const ext = wb.dataExtent(sheet);
  const filled = (r: number, c: number) => !wb.isEmpty(sheet, r, c);
  const vertical = dr !== 0;
  const d = vertical ? dr : dc;
  let r = pos.row;
  let c = pos.col;
  const max = vertical ? MAX_ROWS - 1 : MAX_COLS - 1;
  const limit = vertical ? ext.rows : ext.cols;
  const get = (i: number) => (vertical ? filled(i, c) : filled(r, i));
  let i = vertical ? r : c;
  const next = (k: number) => {
    let n = k + d;
    const ax = vertical ? sheet.rows : sheet.cols;
    while (n >= 0 && n <= max && ax.isHidden(n)) n += d;
    return n;
  };
  let n = next(i);
  if (n < 0 || n > max) return pos;
  if (get(i) && get(n)) {
    while (n >= 0 && n <= max && n < limit + 1 && get(n)) {
      i = n;
      n = next(n);
    }
  } else {
    // Sonraki dolu hücreyi bul
    let found = -1;
    if (d > 0) {
      for (let k = n; k < limit && k <= max; k = next(k)) {
        if (get(k)) {
          found = k;
          break;
        }
      }
    } else {
      for (let k = n; k >= 0; k = next(k)) {
        if (get(k)) {
          found = k;
          break;
        }
        if (k === 0) break;
      }
    }
    i = found >= 0 ? found : d > 0 ? max : 0;
  }
  if (vertical) r = i;
  else c = i;
  const m = sheet.mergeAt(r, c);
  return m ? { row: m.r1, col: m.c1 } : { row: r, col: c };
}

/** Seçimin son aralığını etkin hücreden köşeye kadar ayarlar. */
export function extendTo(sheet: SheetModel, sel: Selection, corner: CellPos): Selection {
  const r = expandForMerges(sheet, rect(sel.anchor.row, sel.anchor.col, corner.row, corner.col));
  const ranges = sel.ranges.slice(0, -1).concat([r]);
  return { ...sel, ranges };
}

/** Enter/Tab ile seçim içinde sonraki hücre. */
export function nextInSelection(sheet: SheetModel, sel: Selection, dir: "down" | "up" | "right" | "left"): Selection {
  const r = sel.ranges[sel.ranges.length - 1];
  const single = sel.ranges.length === 1 && r.r1 === r.r2 && r.c1 === r.c2;
  const merged = sheet.mergeAt(sel.active.row, sel.active.col);
  const isMergeOnly = sel.ranges.length === 1 && merged && merged.r1 === r.r1 && merged.r2 === r.r2 && merged.c1 === r.c1 && merged.c2 === r.c2;
  if (single || isMergeOnly) {
    const p = stepCell(sheet, sel.active, dir === "down" ? 1 : dir === "up" ? -1 : 0, dir === "right" ? 1 : dir === "left" ? -1 : 0);
    return singleSelection(sheet, p.row, p.col);
  }
  let { row, col } = sel.active;
  const fwd = dir === "down" || dir === "right";
  for (let guard = 0; guard < 100000; guard++) {
    if (dir === "down" || dir === "up") {
      row += fwd ? 1 : -1;
      if (row > r.r2) {
        row = r.r1;
        col = col + 1 > r.c2 ? r.c1 : col + 1;
      } else if (row < r.r1) {
        row = r.r2;
        col = col - 1 < r.c1 ? r.c2 : col - 1;
      }
    } else {
      col += fwd ? 1 : -1;
      if (col > r.c2) {
        col = r.c1;
        row = row + 1 > r.r2 ? r.r1 : row + 1;
      } else if (col < r.c1) {
        col = r.c2;
        row = row - 1 < r.r1 ? r.r2 : row - 1;
      }
    }
    const m = sheet.mergeAt(row, col);
    if (m && (m.r1 !== row || m.c1 !== col)) continue;
    if (sheet.rows.isHidden(row) || sheet.cols.isHidden(col)) continue;
    break;
  }
  return { ...sel, active: { row, col } };
}

export function selectionContains(sel: Selection, row: number, col: number): boolean {
  return sel.ranges.some((r) => contains(r, row, col));
}

/** Seçim etiketi (Ad Kutusu): "A1", "3S x 2S". */
export function isWholeSheet(r: Rect): boolean {
  return r.r1 === 0 && r.c1 === 0 && r.r2 === MAX_ROWS - 1 && r.c2 === MAX_COLS - 1;
}

/** Otomatik Toplam için toplanacak aralığı tahmin eder (Excel gibi önce yukarı, sonra sol). */
export function guessSumRange(wb: WorkbookModel, sheet: SheetModel, pos: CellPos): Rect | null {
  const isNum = (r: number, c: number) => typeof wb.getValue(sheet, r, c) === "number";
  // Yukarı
  let r = pos.row - 1;
  if (r >= 0 && (isNum(r, pos.col) || !wb.isEmpty(sheet, r, pos.col))) {
    while (r >= 0 && wb.isEmpty(sheet, r, pos.col)) r--;
    const end = r;
    while (r >= 0 && isNum(r, pos.col)) r--;
    if (end > r) return rect(r + 1, pos.col, end, pos.col);
  }
  let c = pos.col - 1;
  if (c >= 0 && isNum(pos.row, c)) {
    const end = c;
    while (c >= 0 && isNum(pos.row, c)) c--;
    return rect(pos.row, c + 1, pos.row, end);
  }
  // Yukarıda sayı yok ama metin başlık varsa altındaki boşluk
  r = pos.row - 1;
  while (r >= 0 && wb.isEmpty(sheet, r, pos.col)) r--;
  if (r >= 0) {
    const end = r;
    while (r >= 0 && isNum(r, pos.col)) r--;
    if (end > r) return rect(r + 1, pos.col, end, pos.col);
  }
  return null;
}

/** Bitişik veri bölgesi (Ctrl+A ve Tablo Ekle için "geçerli bölge"). */
export function currentRegion(wb: WorkbookModel, sheet: SheetModel, pos: CellPos): Rect {
  let r: Rect = rect(pos.row, pos.col);
  const filled = (row: number, col: number) => row >= 0 && col >= 0 && row < MAX_ROWS && col < MAX_COLS && !wb.isEmpty(sheet, row, col);
  let changed = true;
  let guard = 0;
  while (changed && guard++ < 1000) {
    changed = false;
    // Her kenarın bir dışındaki satır/sütunu (çaprazlar dahil) kontrol et
    const check = (row1: number, col1: number, row2: number, col2: number) => {
      for (let i = row1; i <= row2; i++) for (let j = col1; j <= col2; j++) if (filled(i, j)) return true;
      return false;
    };
    if (r.r1 > 0 && check(r.r1 - 1, Math.max(0, r.c1 - 1), r.r1 - 1, r.c2 + 1)) {
      r = { ...r, r1: r.r1 - 1 };
      changed = true;
    }
    if (check(r.r2 + 1, Math.max(0, r.c1 - 1), r.r2 + 1, r.c2 + 1)) {
      r = { ...r, r2: r.r2 + 1 };
      changed = true;
    }
    if (r.c1 > 0 && check(Math.max(0, r.r1 - 1), r.c1 - 1, r.r2 + 1, r.c1 - 1)) {
      r = { ...r, c1: r.c1 - 1 };
      changed = true;
    }
    if (check(Math.max(0, r.r1 - 1), r.c2 + 1, r.r2 + 1, r.c2 + 1)) {
      r = { ...r, c2: r.c2 + 1 };
      changed = true;
    }
  }
  return expandForMerges(sheet, r);
}
