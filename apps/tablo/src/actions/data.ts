/** Veri eylemleri: sıralama, filtreleme, yinelenenleri kaldırma, metni sütunlara dönüştürme, tablolar, koşullu biçim. */
import { compareText, formatValue, isError, numEquals, shiftFormula, type Scalar } from "@officehere/formula";
import type { Controller } from "../controller";
import { displayText } from "../model/display";
import { uniqueTableName } from "../model/ops";
import { rect, rectRows } from "../model/rect";
import { currentRegion } from "../model/selection";
import { DEFAULT_TABLE_STYLE, tableCellStyle } from "../model/tablestyles";
import type { AutoFilter, CondRule, FilterCriteria, Rect, TableDef } from "../model/types";
import type { CellInput, CellState, SheetModel, WorkbookModel } from "../model/workbook";
import { cellKey } from "@officehere/formula";

/* ------------------------------ Sıralama ------------------------------ */

export interface SortKey {
  col: number;
  desc: boolean;
  /** Özel liste (ör. gün adları) sırası. */
  order?: string[];
  by?: "value" | "fill" | "color";
}

function rank(v: Scalar): number {
  if (v === null || v === "") return 5;
  if (typeof v === "number") return 0;
  if (typeof v === "string") return 1;
  if (typeof v === "boolean") return 2;
  return 3;
}

export function compareForSort(a: Scalar, b: Scalar, desc: boolean, order?: string[]): number {
  const ra = rank(a);
  const rb = rank(b);
  // Boşlar her zaman sonda
  if (ra === 5 || rb === 5) return ra === rb ? 0 : ra === 5 ? 1 : -1;
  let c: number;
  if (order && typeof a === "string" && typeof b === "string") {
    const ia = order.findIndex((x) => x.toLocaleUpperCase("tr-TR") === a.toLocaleUpperCase("tr-TR"));
    const ib = order.findIndex((x) => x.toLocaleUpperCase("tr-TR") === b.toLocaleUpperCase("tr-TR"));
    c = (ia < 0 ? 999 : ia) - (ib < 0 ? 999 : ib);
  } else if (ra !== rb) c = ra - rb;
  else if (typeof a === "number") c = numEquals(a, b as number) ? 0 : a < (b as number) ? -1 : 1;
  else if (typeof a === "string") c = compareText(a, b as string);
  else if (typeof a === "boolean") c = a === b ? 0 : a ? 1 : -1;
  else c = 0;
  return desc ? -c : c;
}

/** Aralığın ilk satırı başlık mı (sezgisel)? */
export function guessHeader(wb: WorkbookModel, sheet: SheetModel, r: Rect): boolean {
  if (rectRows(r) < 2) return false;
  let headerText = 0;
  let total = 0;
  let typeDiff = false;
  for (let c = r.c1; c <= r.c2; c++) {
    const a = wb.getValue(sheet, r.r1, c);
    const b = wb.getValue(sheet, r.r1 + 1, c);
    if (a === null) continue;
    total++;
    if (typeof a === "string") headerText++;
    if (b !== null && typeof a !== typeof b) typeDiff = true;
    if (wb.styleAt(sheet, r.r1, c).bold && !wb.styleAt(sheet, r.r1 + 1, c).bold) typeDiff = true;
  }
  return total > 0 && headerText === total && typeDiff;
}

/** Sıralama aralığını belirler: tek hücre seçiliyse geçerli bölge. */
export function sortRegion(ctl: Controller): Rect {
  const r = ctl.range;
  if (r.r1 === r.r2 && r.c1 === r.c2) {
    const t = ctl.sheet.tableAt(r.r1, r.c1);
    if (t) return t.ref;
    return currentRegion(ctl.wb, ctl.sheet, ctl.sel.active);
  }
  const ext = ctl.wb.dataExtent(ctl.sheet);
  return { ...r, r2: Math.min(r.r2, Math.max(r.r1, ext.rows - 1)), c2: Math.min(r.c2, Math.max(r.c1, ext.cols - 1)) };
}

export function sortRange(ctl: Controller, r: Rect, keys: SortKey[], hasHeader: boolean, byColumns = false): void {
  const { wb, sheet } = ctl;
  if (sheet.mergesIn(r).length) {
    ctl.alert("Bunu yapmak için tüm birleştirilmiş hücrelerin aynı boyutta olması gerekir.");
    return;
  }
  const data: Rect = hasHeader ? (byColumns ? { ...r, c1: r.c1 + 1 } : { ...r, r1: r.r1 + 1 }) : r;
  if (data.r1 > data.r2 || data.c1 > data.c2) return;
  ctl.ops.cellsOp(sheet, [data], "Sırala", () => {
    const lines: { idx: number; cells: CellState[]; keyVals: Scalar[] }[] = [];
    const count = byColumns ? data.c2 - data.c1 + 1 : data.r2 - data.r1 + 1;
    const width = byColumns ? data.r2 - data.r1 + 1 : data.c2 - data.c1 + 1;
    for (let i = 0; i < count; i++) {
      const cells: CellState[] = [];
      for (let j = 0; j < width; j++) {
        const row = byColumns ? data.r1 + j : data.r1 + i;
        const col = byColumns ? data.c1 + i : data.c1 + j;
        cells.push(wb.cellState(sheet, cellKey(row, col)));
      }
      const keyVals = keys.map((k) => {
        const row = byColumns ? k.col : data.r1 + i;
        const col = byColumns ? data.c1 + i : k.col;
        if (k.by === "fill") return wb.styleAt(sheet, row, col).fill ?? null;
        if (k.by === "color") return wb.styleAt(sheet, row, col).color ?? null;
        return wb.getValue(sheet, row, col);
      });
      lines.push({ idx: i, cells, keyVals });
    }
    lines.sort((a, b) => {
      for (let k = 0; k < keys.length; k++) {
        const c = compareForSort(a.keyVals[k], b.keyVals[k], keys[k].desc, keys[k].order);
        if (c) return c;
      }
      return a.idx - b.idx;
    });
    lines.forEach((ln, i) => {
      for (let j = 0; j < width; j++) {
        const row = byColumns ? data.r1 + j : data.r1 + i;
        const col = byColumns ? data.c1 + i : data.c1 + j;
        const st = { ...ln.cells[j] };
        const inp = st.input;
        if (inp !== null && typeof inp === "object" && "formula" in inp) {
          const d = i - ln.idx;
          st.input = { formula: shiftFormula(inp.formula, byColumns ? 0 : d, byColumns ? d : 0) };
        }
        wb.applyCellState(sheet, cellKey(row, col), st);
      }
    });
  });
}

/** Hızlı sıralama (A→Z / Z→A) etkin hücrenin sütununa göre. */
export function quickSort(ctl: Controller, desc: boolean): void {
  const r = sortRegion(ctl);
  const header = guessHeader(ctl.wb, ctl.sheet, r) || !!ctl.sheet.tableAt(r.r1, r.c1)?.headerRow || (!!ctl.sheet.autoFilter && ctl.sheet.autoFilter.range.r1 === r.r1);
  sortRange(ctl, r, [{ col: ctl.sel.active.col, desc }], header);
}

/* ------------------------------ Filtre ------------------------------ */

export function filtersOf(sheet: SheetModel): AutoFilter[] {
  const out: AutoFilter[] = [];
  if (sheet.autoFilter) out.push(sheet.autoFilter);
  for (const t of sheet.tables) if (t.filter && t.headerRow && t.filterButton) out.push(t.filter);
  return out;
}

export function filterAt(sheet: SheetModel, row: number, col: number): AutoFilter | null {
  for (const f of filtersOf(sheet)) if (f.range.r1 === row && col >= f.range.c1 && col <= f.range.c2) return f;
  return null;
}

/** Filtre aralığını başlığın altındaki bitişik veriyle genişletir. */
function extendFilterRange(wb: WorkbookModel, sheet: SheetModel, f: AutoFilter): Rect {
  let r2 = f.range.r2;
  const filled = (row: number) => {
    for (let c = f.range.c1; c <= f.range.c2; c++) if (!wb.isEmpty(sheet, row, c)) return true;
    return false;
  };
  while (filled(r2 + 1)) r2++;
  return { ...f.range, r2 };
}

export function criteriaMatch(c: FilterCriteria, v: Scalar, text: string, fill: string | undefined, colNums: number[]): boolean {
  switch (c.kind) {
    case "values": {
      if (v === null || text === "") return c.blanks;
      // Excel değer filtresi büyük/küçük harf duyarsızdır (İzmir = izmir)
      const u = text.toLocaleUpperCase("tr-TR");
      return c.values.some((x) => x === text || x.toLocaleUpperCase("tr-TR") === u);
    }
    case "color":
      return (fill ?? "") === c.color;
    case "top10": {
      if (typeof v !== "number") return false;
      const sorted = [...colNums].sort((a, b) => (c.bottom ? a - b : b - a));
      const n = c.percent ? Math.max(1, Math.floor((sorted.length * c.count) / 100)) : c.count;
      const th = sorted[Math.min(n, sorted.length) - 1];
      return c.bottom ? v <= th : v >= th;
    }
    case "custom": {
      const t1 = testOp(c.op1, c.v1, v, text);
      if (!c.op2) return t1;
      const t2 = testOp(c.op2, c.v2 ?? "", v, text);
      return c.join === "or" ? t1 || t2 : t1 && t2;
    }
  }
}

function testOp(op: string, arg: string, v: Scalar, text: string): boolean {
  const n = Number(arg.replace(/\./g, "").replace(",", "."));
  const isNum = arg.trim() !== "" && Number.isFinite(n);
  const s = text.toLocaleUpperCase("tr-TR");
  const a = arg.toLocaleUpperCase("tr-TR");
  switch (op) {
    case "eq":
      return isNum && typeof v === "number" ? numEquals(v, n) : s === a;
    case "ne":
      return isNum && typeof v === "number" ? !numEquals(v, n) : s !== a;
    case "gt":
      return typeof v === "number" && isNum ? v > n : s > a;
    case "ge":
      return typeof v === "number" && isNum ? v >= n : s >= a;
    case "lt":
      return typeof v === "number" && isNum ? v < n : s < a;
    case "le":
      return typeof v === "number" && isNum ? v <= n : s <= a;
    case "begins":
      return s.startsWith(a);
    case "ends":
      return s.endsWith(a);
    case "contains":
      return s.includes(a);
    case "notContains":
      return !s.includes(a);
    case "aboveAvg":
    case "belowAvg":
      return true;
  }
  return true;
}

/** Tüm filtreleri yeniden uygular ve satırları gizler/gösterir. */
export function recomputeFilters(ctl: Controller, sheet = ctl.sheet): void {
  const { wb } = ctl;
  for (const r of sheet.filterHidden) {
    sheet.rows.setHidden(r, false);
    wb.engine.setRowHidden(sheet.id, r, false, true);
  }
  sheet.filterHidden.clear();
  wb.engine.clearFilteredRows(sheet.id);
  for (const f of filtersOf(sheet)) {
    f.range = extendFilterRange(wb, sheet, f);
    const cols = Object.entries(f.columns);
    if (!cols.length) continue;
    const numsCache = new Map<number, number[]>();
    for (let row = f.range.r1 + 1; row <= f.range.r2; row++) {
      let visible = true;
      for (const [k, crit] of cols) {
        const col = f.range.c1 + Number(k);
        const v = wb.getValue(sheet, row, col);
        let nums = numsCache.get(col);
        if (!nums && crit.kind === "top10") {
          nums = [];
          for (let rr = f.range.r1 + 1; rr <= f.range.r2; rr++) {
            const x = wb.getValue(sheet, rr, col);
            if (typeof x === "number") nums.push(x);
          }
          numsCache.set(col, nums);
        }
        if (!criteriaMatch(crit, v, displayText(wb, sheet, row, col), wb.styleAt(sheet, row, col).fill, nums ?? [])) {
          visible = false;
          break;
        }
      }
      if (!visible) {
        sheet.filterHidden.add(row);
        sheet.rows.setHidden(row, true);
        wb.engine.setRowHidden(sheet.id, row, true, true);
      }
    }
  }
}

export function toggleAutoFilter(ctl: Controller): void {
  const sheet = ctl.sheet;
  const t = sheet.tableAt(ctl.sel.active.row, ctl.sel.active.col);
  if (t) {
    ctl.ops.sheetOp(sheet, "Filtre", () => {
      t.filterButton = !t.filterButton;
      if (!t.filterButton) t.filter = undefined;
      else t.filter = { range: t.ref, columns: {} };
      recomputeFilters(ctl);
    });
    return;
  }
  ctl.ops.sheetOp(sheet, "Filtre", () => {
    if (sheet.autoFilter) {
      sheet.autoFilter = null;
    } else {
      const r = ctl.range;
      const region = r.r1 === r.r2 && r.c1 === r.c2 ? currentRegion(ctl.wb, sheet, ctl.sel.active) : r;
      sheet.autoFilter = { range: region, columns: {} };
    }
    recomputeFilters(ctl);
  });
}

export function setColumnFilter(ctl: Controller, f: AutoFilter, col: number, crit: FilterCriteria | null): void {
  ctl.ops.sheetOp(ctl.sheet, "Filtre", () => {
    const k = col - f.range.c1;
    if (crit) f.columns[k] = crit;
    else delete f.columns[k];
    recomputeFilters(ctl);
  });
  // Durum çubuğu: "5 kayıttan 2 tanesi bulundu"
  const total = Math.max(0, f.range.r2 - f.range.r1);
  if (Object.keys(f.columns).length) {
    let hidden = 0;
    for (let r = f.range.r1 + 1; r <= f.range.r2; r++) if (ctl.sheet.filterHidden.has(r)) hidden++;
    ctl.status = `${total.toLocaleString("tr-TR")} kayıttan ${(total - hidden).toLocaleString("tr-TR")} tanesi bulundu`;
  } else ctl.status = "";
  ctl.emit();
}

export function clearAllFilters(ctl: Controller): void {
  ctl.ops.sheetOp(ctl.sheet, "Filtreyi Temizle", () => {
    for (const f of filtersOf(ctl.sheet)) f.columns = {};
    recomputeFilters(ctl);
  });
}

export function reapplyFilters(ctl: Controller): void {
  ctl.ops.sheetOp(ctl.sheet, "Yeniden Uygula", () => recomputeFilters(ctl));
}

/** Filtre sütunundaki benzersiz görünen metinler. */
export function columnValues(wb: WorkbookModel, sheet: SheetModel, f: AutoFilter, col: number): { values: string[]; hasBlank: boolean } {
  const range = extendFilterRange(wb, sheet, f);
  // Büyük/küçük harf farkı olan metinler tek öğe olarak listelenir (ilk görülen yazım)
  const set = new Map<string, [string, Scalar]>();
  let hasBlank = false;
  for (let row = range.r1 + 1; row <= range.r2; row++) {
    const v = wb.getValue(sheet, row, col);
    const t = displayText(wb, sheet, row, col);
    if (v === null || t === "") hasBlank = true;
    else {
      const k = t.toLocaleUpperCase("tr-TR");
      if (!set.has(k)) set.set(k, [t, v]);
    }
  }
  const values = [...set.values()].sort((a, b) => compareForSort(a[1], b[1], false)).map((e) => e[0]);
  return { values, hasBlank };
}

/* ------------------------- Yinelenenleri Kaldır ------------------------- */

export function removeDuplicates(ctl: Controller, r: Rect, cols: number[], hasHeader: boolean): { removed: number; kept: number } {
  const { wb, sheet } = ctl;
  const data = hasHeader ? { ...r, r1: r.r1 + 1 } : r;
  let removed = 0;
  let kept = 0;
  ctl.ops.cellsOp(sheet, [data], "Yinelenenleri Kaldır", () => {
    const seen = new Set<string>();
    const rows: CellState[][] = [];
    for (let row = data.r1; row <= data.r2; row++) {
      const key = cols.map((c) => keyText(wb.getValue(sheet, row, c))).join("\u0001");
      if (seen.has(key)) {
        removed++;
        continue;
      }
      seen.add(key);
      const cells: CellState[] = [];
      for (let c = data.c1; c <= data.c2; c++) cells.push(wb.cellState(sheet, cellKey(row, c)));
      rows.push(cells);
    }
    kept = rows.length;
    for (let i = 0; i <= data.r2 - data.r1; i++) {
      for (let c = data.c1; c <= data.c2; c++) {
        const k = cellKey(data.r1 + i, c);
        wb.applyCellState(sheet, k, i < rows.length ? rows[i][c - data.c1] : undefined);
      }
    }
  });
  return { removed, kept };
}

function keyText(v: Scalar): string {
  if (v === null) return "";
  if (typeof v === "string") return "s" + v.toLocaleUpperCase("tr-TR");
  if (isError(v)) return "e" + v.code;
  return typeof v + String(v);
}

/* ------------------------- Metni Sütunlara Dönüştür ------------------------- */

export function textToColumns(ctl: Controller, r: Rect, delims: string[], treatConsecutive: boolean, textQualifier = '"'): void {
  const { wb, sheet } = ctl;
  const col = r.c1;
  const parts: string[][] = [];
  let maxCols = 1;
  for (let row = r.r1; row <= r.r2; row++) {
    const v = wb.getValue(sheet, row, col);
    const s = v === null ? "" : typeof v === "string" ? v : displayText(wb, sheet, row, col);
    const out = splitDelimited(s, delims, treatConsecutive, textQualifier);
    parts.push(out);
    maxCols = Math.max(maxCols, out.length);
  }
  const entries: { row: number; col: number; input: CellInput }[] = [];
  parts.forEach((p, i) =>
    p.forEach((t, j) => {
      const n = t.trim().replace(/\./g, "").replace(",", ".");
      const num = t.trim() !== "" && /^-?\d+(\.\d+)?$/.test(n) ? Number(n) : null;
      entries.push({ row: r.r1 + i, col: col + j, input: num ?? (t === "" ? null : t) });
    }),
  );
  ctl.ops.setInputs(sheet, entries, "Metni Sütunlara Dönüştür");
}

export function splitDelimited(s: string, delims: string[], consecutive: boolean, q = '"'): string[] {
  const out: string[] = [];
  let cur = "";
  let inQ = false;
  for (let i = 0; i < s.length; i++) {
    const ch = s[i];
    if (q && ch === q) {
      if (inQ && s[i + 1] === q) {
        cur += q;
        i++;
      } else inQ = !inQ;
      continue;
    }
    if (!inQ && delims.includes(ch)) {
      if (consecutive && cur === "" && out.length && out[out.length - 1] === "") continue;
      out.push(cur);
      cur = "";
      continue;
    }
    cur += ch;
  }
  out.push(cur);
  return consecutive ? out.filter((x, i) => x !== "" || i === 0) : out;
}

/* ------------------------------ Tablolar ------------------------------ */

export function createTable(ctl: Controller, r: Rect, hasHeader: boolean, style = DEFAULT_TABLE_STYLE): void {
  const { wb, sheet } = ctl;
  if (sheet.tables.some((t) => t.ref.r1 <= r.r2 && r.r1 <= t.ref.r2 && t.ref.c1 <= r.c2 && r.c1 <= t.ref.c2)) {
    ctl.alert("Tablo diğer bir tabloyla çakışamaz.");
    return;
  }
  if (sheet.mergesIn(r).length) {
    ctl.alert("Birleştirilmiş hücre içeren bir aralık tabloya dönüştürülemez.");
    return;
  }
  ctl.ops.sheetOp(
    sheet,
    "Tablo Oluştur",
    () => {
      let ref = r;
      if (!hasHeader) {
        // Başlık satırı ekle: veriyi bir satır aşağı kaydır
        wb.engine.shiftCells(sheet.id, r.r1, r.c1, r.r1, r.c2, "down");
        const moved = new Map<number, number>();
        for (const [k, v] of sheet.styles) {
          const row = Math.floor(k / 16384);
          const col = k % 16384;
          if (col >= r.c1 && col <= r.c2 && row >= r.r1) moved.set(cellKey(row + 1, col), v);
          else moved.set(k, v);
        }
        sheet.styles = moved;
        for (let c = r.c1; c <= r.c2; c++) wb.setInputRaw(sheet, r.r1, c, `Sütun${c - r.c1 + 1}`);
        ref = { ...r, r2: r.r2 + 1 };
      }
      const names: string[] = [];
      for (let c = ref.c1; c <= ref.c2; c++) {
        let n = displayText(wb, sheet, ref.r1, c) || `Sütun${c - ref.c1 + 1}`;
        const base = n;
        let k = 2;
        while (names.some((x) => x.toLocaleUpperCase("tr-TR") === n.toLocaleUpperCase("tr-TR"))) n = `${base}${k++}`;
        if (n !== displayText(wb, sheet, ref.r1, c)) wb.setInputRaw(sheet, ref.r1, c, n);
        names.push(n);
      }
      const t: TableDef = {
        id: wb.newId("tbl"),
        name: uniqueTableName(wb),
        ref,
        style,
        headerRow: true,
        totalRow: false,
        bandedRows: true,
        bandedCols: false,
        firstColumn: false,
        lastColumn: false,
        filterButton: true,
        columns: names,
        filter: { range: ref, columns: {} },
      };
      if (sheet.autoFilter && sheet.autoFilter.range.r1 <= ref.r2 && ref.r1 <= sheet.autoFilter.range.r2) sheet.autoFilter = null;
      sheet.tables.push(t);
    },
    [rect(r.r1, r.c1, r.r2 + 1, r.c2)],
  );
}

/** Tablo seçeneği değiştir (Tablo Tasarımı sekmesi). */
export function updateTable(ctl: Controller, t: TableDef, patch: Partial<TableDef>): void {
  const { wb, sheet } = ctl;
  ctl.ops.sheetOp(
    sheet,
    "Tablo Stili Seçenekleri",
    () => {
      const cur = sheet.tables.find((x) => x.id === t.id);
      if (!cur) return;
      if (patch.totalRow !== undefined && patch.totalRow !== cur.totalRow) {
        if (patch.totalRow) {
          const row = cur.ref.r2 + 1;
          cur.ref = { ...cur.ref, r2: row };
          wb.setInputRaw(sheet, row, cur.ref.c1, "Toplam");
          const lastCol = cur.ref.c2;
          const dataRange = `${colName(lastCol)}${cur.ref.r1 + 2}:${colName(lastCol)}${row}`;
          wb.setInputRaw(sheet, row, lastCol, { formula: `=SUBTOTAL(109,${dataRange})` });
        } else {
          const row = cur.ref.r2;
          for (let c = cur.ref.c1; c <= cur.ref.c2; c++) wb.setInputRaw(sheet, row, c, null);
          cur.ref = { ...cur.ref, r2: row - 1 };
        }
      }
      Object.assign(cur, { ...patch, ref: cur.ref });
      if (cur.filter) cur.filter.range = { ...cur.ref, r2: cur.totalRow ? cur.ref.r2 - 1 : cur.ref.r2 };
      if (patch.filterButton === false) cur.filter = undefined;
      if (patch.filterButton === true && !cur.filter) cur.filter = { range: cur.ref, columns: {} };
      recomputeFilters(ctl);
    },
    [rect(t.ref.r1, t.ref.c1, t.ref.r2 + 1, t.ref.c2)],
  );
}

function colName(c: number): string {
  let n = c + 1;
  let s = "";
  while (n > 0) {
    const m = (n - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

export function convertTableToRange(ctl: Controller, t: TableDef): void {
  const { wb, sheet } = ctl;
  ctl.ops.sheetOp(
    sheet,
    "Aralığa Dönüştür",
    () => {
      // Tablo görünümünü kalıcı biçime çevir
      for (let row = t.ref.r1; row <= t.ref.r2; row++)
        for (let c = t.ref.c1; c <= t.ref.c2; c++) {
          const ts = tableCellStyle(t, row, c);
          const own = wb.styleAt(sheet, row, c);
          const merged = { ...ts, ...own, fill: own.fill ?? ts.fill, color: own.color ?? ts.color };
          wb.setStyleRaw(sheet, row, c, wb.styles.intern(merged));
        }
      sheet.tables = sheet.tables.filter((x) => x.id !== t.id);
      recomputeFilters(ctl);
    },
    [t.ref],
  );
}

/* ------------------------- Koşullu Biçimlendirme ------------------------- */

export function addCondRule(ctl: Controller, rule: Omit<CondRule, "id" | "ranges">, ranges?: Rect[]): void {
  const sheet = ctl.sheet;
  ctl.ops.sheetOp(sheet, "Koşullu Biçimlendirme", () => {
    sheet.condFormats.unshift({ ...rule, id: ctl.wb.newId("cf"), ranges: (ranges ?? ctl.sel.ranges).map((r) => ({ ...r })) });
  });
}

export function clearCondRules(ctl: Controller, scope: "selection" | "sheet"): void {
  const sheet = ctl.sheet;
  ctl.ops.sheetOp(sheet, "Kuralları Temizle", () => {
    if (scope === "sheet") sheet.condFormats = [];
    else {
      const sel = ctl.sel.ranges;
      sheet.condFormats = sheet.condFormats
        .map((c) => ({ ...c, ranges: c.ranges.filter((r) => !sel.some((s) => s.r1 <= r.r2 && r.r1 <= s.r2 && s.c1 <= r.c2 && r.c1 <= s.c2)) }))
        .filter((c) => c.ranges.length);
    }
  });
}

export function formatForFilter(v: number, code?: string): string {
  return formatValue(v, code).text;
}
