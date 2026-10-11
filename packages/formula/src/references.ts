/**
 * A1 başvuru yardımcıları: sütun harfleri, hücre/aralık ayrıştırma ve
 * biçimlendirme, sayfa adlarının tırnaklanması.
 */

export const MAX_ROWS = 1048576;
export const MAX_COLS = 16384;

/** 0 tabanlı sütun → "A", "Z", "AA", "XFD". */
export function colToName(col: number): string {
  let n = col + 1;
  let s = "";
  while (n > 0) {
    const m = (n - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

/** "A" → 0; geçersizse -1. */
export function nameToCol(name: string): number {
  let n = 0;
  const up = name.toUpperCase();
  if (!/^[A-Z]{1,3}$/.test(up)) return -1;
  for (let i = 0; i < up.length; i++) n = n * 26 + (up.charCodeAt(i) - 64);
  return n - 1 < MAX_COLS ? n - 1 : -1;
}

export interface CellAddress {
  row: number;
  col: number;
}

export interface CellRef extends CellAddress {
  rowAbs: boolean;
  colAbs: boolean;
}

export interface RangeAddress {
  r1: number;
  c1: number;
  r2: number;
  c2: number;
}

const CELL_RE = /^(\$?)([A-Za-z]{1,3})(\$?)(\d{1,7})$/;

/** "$B$3" → { row: 2, col: 1, rowAbs, colAbs }. */
export function parseCellRef(text: string): CellRef | null {
  const m = CELL_RE.exec(text.trim());
  if (!m) return null;
  const col = nameToCol(m[2]);
  const row = parseInt(m[4], 10) - 1;
  if (col < 0 || row < 0 || row >= MAX_ROWS) return null;
  return { row, col, colAbs: m[1] === "$", rowAbs: m[3] === "$" };
}

export function formatCellRef(ref: CellRef | CellAddress): string {
  const r = ref as CellRef;
  return `${r.colAbs ? "$" : ""}${colToName(ref.col)}${r.rowAbs ? "$" : ""}${ref.row + 1}`;
}

/** Basit adres: (0,0) → "A1". */
export function addr(row: number, col: number): string {
  return `${colToName(col)}${row + 1}`;
}

/** "A1:B5", "A:A", "1:3", "B2" → aralık (sayfa öneki olmadan). */
export function parseRangeAddress(text: string): RangeAddress | null {
  const t = text.trim().replace(/\$/g, "");
  const parts = t.split(":");
  if (parts.length === 1) {
    const c = parseCellRef(parts[0]);
    return c ? { r1: c.row, c1: c.col, r2: c.row, c2: c.col } : null;
  }
  if (parts.length !== 2) return null;
  const [a, b] = parts;
  if (/^\d+$/.test(a) && /^\d+$/.test(b)) {
    const r1 = parseInt(a, 10) - 1;
    const r2 = parseInt(b, 10) - 1;
    if (r1 < 0 || r2 < 0 || r1 >= MAX_ROWS || r2 >= MAX_ROWS) return null;
    return { r1: Math.min(r1, r2), c1: 0, r2: Math.max(r1, r2), c2: MAX_COLS - 1 };
  }
  if (/^[A-Za-z]{1,3}$/.test(a) && /^[A-Za-z]{1,3}$/.test(b)) {
    const c1 = nameToCol(a);
    const c2 = nameToCol(b);
    if (c1 < 0 || c2 < 0) return null;
    return { r1: 0, c1: Math.min(c1, c2), r2: MAX_ROWS - 1, c2: Math.max(c1, c2) };
  }
  const p = parseCellRef(a);
  const q = parseCellRef(b);
  if (!p || !q) return null;
  return { r1: Math.min(p.row, q.row), c1: Math.min(p.col, q.col), r2: Math.max(p.row, q.row), c2: Math.max(p.col, q.col) };
}

/** Aralığı "A1:B5" biçiminde yazar (tüm sütun/satır için "A:B" / "1:3"). */
export function formatRange(r: RangeAddress): string {
  if (r.r1 === 0 && r.r2 === MAX_ROWS - 1) return `${colToName(r.c1)}:${colToName(r.c2)}`;
  if (r.c1 === 0 && r.c2 === MAX_COLS - 1) return `${r.r1 + 1}:${r.r2 + 1}`;
  if (r.r1 === r.r2 && r.c1 === r.c2) return addr(r.r1, r.c1);
  return `${addr(r.r1, r.c1)}:${addr(r.r2, r.c2)}`;
}

/** Sayfa adı tırnak gerektiriyor mu? */
export function needsQuotes(name: string): boolean {
  if (name === "") return true;
  if (!/^[\p{L}_][\p{L}\p{N}_.]*$/u.test(name)) return true;
  // Hücre başvurusuna ya da R1C1'e benzeyen adlar tırnaklanır.
  if (/^[A-Za-z]{1,3}\d+$/.test(name)) return true;
  if (/^(R|C|RC|R\d+C\d+)$/i.test(name)) return true;
  if (/^(TRUE|FALSE)$/i.test(name)) return true;
  return false;
}

export function quoteSheetName(name: string): string {
  return needsQuotes(name) ? `'${name.replace(/'/g, "''")}'` : name;
}

/** "Sayfa1!A1:B2" ya da "'Sayfa 1'!A1" → { sheet, range }. */
export function parseQualifiedRange(text: string): { sheet?: string; range: RangeAddress } | null {
  const t = text.trim();
  let sheet: string | undefined;
  let rest = t;
  const bang = t.lastIndexOf("!");
  if (bang >= 0) {
    let s = t.slice(0, bang);
    if (s.startsWith("'") && s.endsWith("'")) s = s.slice(1, -1).replace(/''/g, "'");
    sheet = s;
    rest = t.slice(bang + 1);
  }
  const range = parseRangeAddress(rest);
  return range ? { sheet, range } : null;
}

export function formatQualifiedRange(sheet: string | undefined, r: RangeAddress): string {
  return sheet ? `${quoteSheetName(sheet)}!${formatRange(r)}` : formatRange(r);
}

export function rangeContains(r: RangeAddress, row: number, col: number): boolean {
  return row >= r.r1 && row <= r.r2 && col >= r.c1 && col <= r.c2;
}

export function rangesIntersect(a: RangeAddress, b: RangeAddress): boolean {
  return a.r1 <= b.r2 && b.r1 <= a.r2 && a.c1 <= b.c2 && b.c1 <= a.c2;
}

export function normalizeRange(r: RangeAddress): RangeAddress {
  return { r1: Math.min(r.r1, r.r2), c1: Math.min(r.c1, r.c2), r2: Math.max(r.r1, r.r2), c2: Math.max(r.c1, r.c2) };
}
