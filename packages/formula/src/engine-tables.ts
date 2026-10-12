/** Hesap motoru yardımcıları: tablo kaydı (yapısal başvurular) ve satır permütasyonu (sıralama). */
import type { FormulaEngine } from "./engine";
import { SHEET_STRIDE, cellKey, keyCol, keyRow, type CellRecord } from "./engine-types";
import type { TableInfo } from "./structref";
import { renameTableColumnInFormula, renameTableInFormula, shiftAst, shiftFormula } from "./transform";

/** Formüllerin metnini dönüştürür (değişenler yeniden kaydedilir). */
export function rewriteStructRefs(e: FormulaEngine, fn: (text: string) => string): void {
  for (const id of e.order) {
    const sh = e.sheets.get(id)!;
    for (const [k, rec] of sh.cells) {
      if (!rec.formula || !rec.formula.text.includes("[")) continue;
      const t = fn(rec.formula.text);
      if (t !== rec.formula.text) e.setFormulaInternal(id, keyRow(k), keyCol(k), t);
    }
  }
}

export function findTable(tables: TableInfo[], name: string | undefined, sheet: number, row: number, col: number): TableInfo | undefined {
  if (name !== undefined) {
    const u = name.toLocaleUpperCase("tr-TR");
    return tables.find((t) => t.name.toLocaleUpperCase("tr-TR") === u);
  }
  return tables.find((t) => t.sheet === sheet && row >= t.r1 && row <= t.r2 && col >= t.c1 && col <= t.c2);
}

export function renameTableImpl(e: FormulaEngine, oldName: string, newName: string): void {
  rewriteStructRefs(e, (t) => renameTableInFormula(t, oldName, newName));
  for (const t of e.tables) if (t.name.toLocaleUpperCase("tr-TR") === oldName.toLocaleUpperCase("tr-TR")) t.name = newName;
  e.tablesSig = JSON.stringify(e.tables);
}

/** Tablo listesini ayarlar; değiştiyse true (bağımlılıklar yenilenmeli). */
export function setTablesImpl(e: FormulaEngine, list: TableInfo[]): boolean {
  const sig = JSON.stringify(list);
  if (sig === e.tablesSig) return false;
  e.tablesSig = sig;
  e.tables = list.map((t) => ({ ...t, columns: [...t.columns] }));
  return true;
}

export function renameTableColumnImpl(e: FormulaEngine, table: string, oldCol: string, newCol: string): void {
  rewriteStructRefs(e, (t) => renameTableColumnInFormula(t, table, oldCol, newCol));
}

/**
 * Aralıktaki satırları yeniden sıralar: yeni i. satır = eski order[i]. Formüller
 * göreli kaydırılır (metin + ayrıştırılmış ağaç, yeniden ayrıştırmadan).
 * Kaydırma yeni bir #BAŞV! üretecekse hiçbir şey değiştirmeden false döner.
 */
export function permuteRowsImpl(e: FormulaEngine, sheet: number, r1: number, c1: number, r2: number, c2: number, order: number[]): boolean {
  const s = e.sheetOrThrow(sheet);
  const n = r2 - r1 + 1;
  if (order.length !== n) return false;
  // Ön kontrol: yeni #REF! olmamalı
  const moved: (CellRecord | undefined)[][] = [];
  const texts = new Map<CellRecord, string>();
  for (let i = 0; i < n; i++) {
    const src = r1 + order[i];
    const d = r1 + i - src;
    const line: (CellRecord | undefined)[] = [];
    for (let c = c1; c <= c2; c++) {
      const rec = s.cells.get(cellKey(src, c));
      line.push(rec);
      if (rec?.formula && d !== 0) {
        const t = shiftFormula("=" + rec.formula.text, d, 0).slice(1);
        if (t.includes("#REF!") && !rec.formula.text.includes("#REF!")) return false;
        texts.set(rec, t);
      }
    }
    moved.push(line);
  }
  // Eski kayıtları çıkar
  for (let r = r1; r <= r2; r++)
    for (let c = c1; c <= c2; c++) {
      const k = cellKey(r, c);
      const rec = s.cells.get(k);
      if (!rec) continue;
      const g = sheet * SHEET_STRIDE + k;
      if (rec.formula) e.unregister(g);
      s.cells.delete(k);
    }
  // Yeni konumlara yaz
  for (let i = 0; i < n; i++) {
    const row = r1 + i;
    const d = row - (r1 + order[i]);
    for (let c = c1; c <= c2; c++) {
      const rec = moved[i][c - c1];
      const k = cellKey(row, c);
      const g = sheet * SHEET_STRIDE + k;
      if (!rec) {
        e.noteChange(g);
        continue;
      }
      if (rec.formula) {
        const f = rec.formula;
        const text = d !== 0 ? (texts.get(rec) ?? f.text) : f.text;
        const info = { ...f, text, ast: f.ast && d !== 0 ? shiftAst(f.ast, d, 0) : f.ast, cellDeps: [], rangeDeps: [], spill: null, spillTarget: null };
        const nrec: CellRecord = { value: rec.value, formula: info };
        s.cells.set(k, nrec);
        e.registerInfo(g, info, sheet);
        e.dirty.add(g);
      } else s.cells.set(k, { value: rec.value });
      e.noteChange(g);
    }
  }
  s.extentDirty = true;
  return true;
}

export function setRowsHiddenImpl(e: FormulaEngine, sheet: number, rows: Iterable<number>, hidden: boolean, filtered: boolean): void {
  const s = e.sheets.get(sheet);
  if (!s) return;
  const set = filtered ? s.filteredRows : s.hiddenRows;
  if (hidden) for (const r of rows) set.add(r);
  else for (const r of rows) set.delete(r);
  e.subtotalsStale = true;
}

/** hidden null ise elle gizlenenler korunur (yalnız filtre kümesi değişir). */
export function replaceHiddenRowsImpl(e: FormulaEngine, sheet: number, hidden: Iterable<number> | null, filtered: Iterable<number>): void {
  const s = e.sheets.get(sheet);
  if (!s) return;
  if (hidden) s.hiddenRows = new Set(hidden);
  s.filteredRows = new Set(filtered);
  e.subtotalsStale = true;
}
