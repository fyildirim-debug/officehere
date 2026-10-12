/** Dikdörtgen (aralık) yardımcıları. */
import { MAX_COLS, MAX_ROWS, formatRange, parseRangeAddress } from "@officehere/formula";
import type { CellPos, Rect } from "./types";

export { MAX_COLS, MAX_ROWS };

export function rect(r1: number, c1: number, r2 = r1, c2 = c1): Rect {
  return { r1: Math.min(r1, r2), c1: Math.min(c1, c2), r2: Math.max(r1, r2), c2: Math.max(c1, c2) };
}

export function rectFromPos(a: CellPos, b: CellPos = a): Rect {
  return rect(a.row, a.col, b.row, b.col);
}

export function contains(r: Rect, row: number, col: number): boolean {
  return row >= r.r1 && row <= r.r2 && col >= r.c1 && col <= r.c2;
}

export function containsRect(outer: Rect, inner: Rect): boolean {
  return inner.r1 >= outer.r1 && inner.r2 <= outer.r2 && inner.c1 >= outer.c1 && inner.c2 <= outer.c2;
}

export function intersects(a: Rect, b: Rect): boolean {
  return a.r1 <= b.r2 && b.r1 <= a.r2 && a.c1 <= b.c2 && b.c1 <= a.c2;
}

export function intersection(a: Rect, b: Rect): Rect | null {
  if (!intersects(a, b)) return null;
  return { r1: Math.max(a.r1, b.r1), c1: Math.max(a.c1, b.c1), r2: Math.min(a.r2, b.r2), c2: Math.min(a.c2, b.c2) };
}

export function union(a: Rect, b: Rect): Rect {
  return { r1: Math.min(a.r1, b.r1), c1: Math.min(a.c1, b.c1), r2: Math.max(a.r2, b.r2), c2: Math.max(a.c2, b.c2) };
}

export function rectEquals(a: Rect, b: Rect): boolean {
  return a.r1 === b.r1 && a.c1 === b.c1 && a.r2 === b.r2 && a.c2 === b.c2;
}

export function rectRows(r: Rect): number {
  return r.r2 - r.r1 + 1;
}

export function rectCols(r: Rect): number {
  return r.c2 - r.c1 + 1;
}

export function area(r: Rect): number {
  return rectRows(r) * rectCols(r);
}

export function isFullCols(r: Rect): boolean {
  return r.r1 === 0 && r.r2 === MAX_ROWS - 1;
}

export function isFullRows(r: Rect): boolean {
  return r.c1 === 0 && r.c2 === MAX_COLS - 1;
}

export function rectToA1(r: Rect): string {
  return formatRange(r);
}

export function a1ToRect(s: string): Rect | null {
  return parseRangeAddress(s);
}

/** Satır/sütun ekleme-silmede dikdörtgeni eşler; tamamen silinirse null. */
export function shiftRect(r: Rect, axis: "row" | "col", at: number, count: number): Rect | null {
  const lo = axis === "row" ? r.r1 : r.c1;
  const hi = axis === "row" ? r.r2 : r.c2;
  const max = axis === "row" ? MAX_ROWS - 1 : MAX_COLS - 1;
  let nlo: number;
  let nhi: number;
  if (count > 0) {
    if (lo === 0 && hi === max) return r;
    nlo = lo >= at ? lo + count : lo;
    nhi = hi >= at ? hi + count : hi;
    nhi = Math.min(nhi, max);
    if (nlo > max) return null;
  } else {
    const n = -count;
    const end = at + n - 1;
    if (lo >= at && hi <= end) return null;
    nlo = lo < at ? lo : lo > end ? lo - n : at;
    nhi = hi < at ? hi : hi > end ? hi - n : at - 1;
    if (nhi < nlo) return null;
  }
  return axis === "row" ? { ...r, r1: nlo, r2: nhi } : { ...r, c1: nlo, c2: nhi };
}

/** Konumu ekleme/silmeye göre eşler; silinirse null. */
export function shiftIndex(i: number, at: number, count: number): number | null {
  if (count > 0) return i >= at ? i + count : i;
  if (i < at) return i;
  if (i < at - count) return null;
  return i + count;
}

/** a'dan b çıkarılınca kalan dikdörtgenler (en çok 4). */
export function subtractRect(a: Rect, b: Rect): Rect[] {
  const i = intersection(a, b);
  if (!i) return [a];
  const out: Rect[] = [];
  if (a.r1 < i.r1) out.push({ r1: a.r1, c1: a.c1, r2: i.r1 - 1, c2: a.c2 });
  if (i.r2 < a.r2) out.push({ r1: i.r2 + 1, c1: a.c1, r2: a.r2, c2: a.c2 });
  if (a.c1 < i.c1) out.push({ r1: i.r1, c1: a.c1, r2: i.r2, c2: i.c1 - 1 });
  if (i.c2 < a.c2) out.push({ r1: i.r1, c1: i.c2 + 1, r2: i.r2, c2: a.c2 });
  return out;
}

/** Listeden kesilecek dikdörtgenler çıkarılır. */
export function subtractRects(list: Rect[], cut: Rect[]): Rect[] {
  let cur = list;
  for (const b of cut) cur = cur.flatMap((a) => subtractRect(a, b));
  return cur;
}
