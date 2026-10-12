/** Tablo işlemleri: satır/sütun ekle-sil, birleştir/böl, dağıt. */
import { emptyCell } from "./factory";
import type { Paragraph, TableCell, TableShape } from "./types";

export interface CellRange {
  r1: number;
  c1: number;
  r2: number;
  c2: number;
}

export function normRange(r: CellRange): CellRange {
  return { r1: Math.min(r.r1, r.r2), r2: Math.max(r.r1, r.r2), c1: Math.min(r.c1, r.c2), c2: Math.max(r.c1, r.c2) };
}

function cloneCellStyle(c: TableCell): TableCell {
  const fill = c.fill;
  return { ...emptyCell(), fill, anchor: c.anchor };
}

export function insertRow(t: TableShape, at: number, below: boolean): TableShape {
  const ref = t.rows[Math.max(0, Math.min(t.rows.length - 1, at))];
  const idx = below ? at + 1 : at;
  const row = { h: ref?.h ?? 40, cells: (ref?.cells ?? t.cols.map(emptyCell)).map((c) => ({ ...cloneCellStyle(c), gridSpan: undefined, hMerge: undefined })) };
  const rows = [...t.rows.slice(0, idx), row, ...t.rows.slice(idx)];
  return { ...t, rows, h: rows.reduce((s, r) => s + r.h, 0) };
}

export function insertCol(t: TableShape, at: number, right: boolean): TableShape {
  const idx = right ? at + 1 : at;
  const w = t.cols[Math.max(0, Math.min(t.cols.length - 1, at))] ?? 100;
  // Tablo genişliği korunur: yeni sütun için diğerleri daraltılır (PowerPoint genişletir; burada slayttan taşmayı önlüyoruz).
  const cols = [...t.cols.slice(0, idx), w, ...t.cols.slice(idx)];
  const rows = t.rows.map((r) => ({ ...r, cells: [...r.cells.slice(0, idx), cloneCellStyle(r.cells[Math.min(at, r.cells.length - 1)] ?? emptyCell()), ...r.cells.slice(idx)] }));
  return { ...t, cols, rows, w: cols.reduce((s, c) => s + c, 0) };
}

export function deleteRows(t: TableShape, r1: number, r2: number): TableShape | null {
  const keep = t.rows.filter((_, i) => i < r1 || i > r2);
  if (!keep.length) return null;
  return { ...t, rows: keep, h: keep.reduce((s, r) => s + r.h, 0) };
}

export function deleteCols(t: TableShape, c1: number, c2: number): TableShape | null {
  const cols = t.cols.filter((_, i) => i < c1 || i > c2);
  if (!cols.length) return null;
  const rows = t.rows.map((r) => ({ ...r, cells: r.cells.filter((_, i) => i < c1 || i > c2).map((c) => ({ ...c, gridSpan: undefined, hMerge: undefined })) }));
  return { ...t, cols, rows, w: cols.reduce((s, c) => s + c, 0) };
}

/** Aralıktaki hücreleri birleştirir; metinler sol üst hücrede toplanır. */
export function mergeCells(t: TableShape, range: CellRange): TableShape {
  const r = normRange(range);
  if (r.r1 === r.r2 && r.c1 === r.c2) return t;
  const paras: Paragraph[] = [];
  for (let i = r.r1; i <= r.r2; i++)
    for (let j = r.c1; j <= r.c2; j++) {
      const c = t.rows[i].cells[j];
      if (c.text.paragraphs.some((p) => p.runs.some((x) => x.text))) paras.push(...c.text.paragraphs);
    }
  const rows = t.rows.map((row, i) => ({
    ...row,
    cells: row.cells.map((c, j) => {
      if (i < r.r1 || i > r.r2 || j < r.c1 || j > r.c2) return c;
      if (i === r.r1 && j === r.c1)
        return { ...c, gridSpan: r.c2 - r.c1 + 1 > 1 ? r.c2 - r.c1 + 1 : undefined, rowSpan: r.r2 - r.r1 + 1 > 1 ? r.r2 - r.r1 + 1 : undefined, text: { ...c.text, paragraphs: paras.length ? paras : c.text.paragraphs } };
      return { ...emptyCell(), hMerge: j > r.c1 ? true : undefined, vMerge: i > r.r1 ? true : undefined, fill: c.fill };
    }),
  }));
  return { ...t, rows };
}

/** Birleştirilmiş hücreyi çözer. */
export function splitCell(t: TableShape, row: number, col: number): TableShape {
  const c = t.rows[row]?.cells[col];
  if (!c || (!c.gridSpan && !c.rowSpan)) return t;
  const gs = c.gridSpan ?? 1;
  const rs = c.rowSpan ?? 1;
  const rows = t.rows.map((r, i) => ({
    ...r,
    cells: r.cells.map((x, j) => {
      if (i === row && j === col) return { ...x, gridSpan: undefined, rowSpan: undefined };
      if (i >= row && i < row + rs && j >= col && j < col + gs) return { ...x, hMerge: undefined, vMerge: undefined };
      return x;
    }),
  }));
  return { ...t, rows };
}

export function distributeRows(t: TableShape): TableShape {
  const h = t.rows.reduce((s, r) => s + r.h, 0) / t.rows.length;
  return { ...t, rows: t.rows.map((r) => ({ ...r, h })) };
}

export function distributeCols(t: TableShape): TableShape {
  const w = t.cols.reduce((s, c) => s + c, 0) / t.cols.length;
  return { ...t, cols: t.cols.map(() => w) };
}

export function mapCells(t: TableShape, range: CellRange | null, fn: (c: TableCell, r: number, col: number) => TableCell): TableShape {
  const r = range ? normRange(range) : { r1: 0, c1: 0, r2: t.rows.length - 1, c2: t.cols.length - 1 };
  return {
    ...t,
    rows: t.rows.map((row, i) => ({ ...row, cells: row.cells.map((c, j) => (i >= r.r1 && i <= r.r2 && j >= r.c1 && j <= r.c2 ? fn(c, i, j) : c)) })),
  };
}
