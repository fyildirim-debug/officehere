/** Hesap motoru okuma yardımcıları: kullanılan alan, taşma bilgisi, değer gezme, değişiklik bildirimi, tek formül değerlendirme. */
import type { FormulaEngine } from "./engine";
import { cellKey, gk, gkLocal, gkSheet, keyCol, keyRow, safeParse, type SheetData, type SpillInfo } from "./engine-types";
import { normalizeResult } from "./engine-util";
import { createContext, evaluateNode } from "./evaluator";
import { Errors, type Scalar } from "./values";

export function recomputeExtentImpl(s: SheetData): void {
  let mr = -1;
  let mc = -1;
  for (const [k, rec] of s.cells) {
    let rr = keyRow(k);
    let cc = keyCol(k);
    if (rec.formula?.spill) {
      rr += rec.formula.spill.rows - 1;
      cc += rec.formula.spill.cols - 1;
    }
    if (rr > mr) mr = rr;
    if (cc > mc) mc = cc;
  }
  s.maxRow = mr;
  s.maxCol = mc;
  s.extentDirty = false;
}

export function getSpillImpl(e: FormulaEngine, sheet: number, row: number, col: number): SpillInfo | undefined {
  const k = gk(sheet, row, col);
  const ownerKey = e.spillOwner.get(k) ?? (e.sheets.get(sheet)?.cells.get(cellKey(row, col))?.formula?.spill ? k : undefined);
  if (ownerKey === undefined) return undefined;
  const sp = e.sheets.get(gkSheet(ownerKey))?.cells.get(gkLocal(ownerKey))?.formula?.spill;
  if (!sp) return undefined;
  return { anchorRow: keyRow(gkLocal(ownerKey)), anchorCol: keyCol(gkLocal(ownerKey)), rows: sp.rows, cols: sp.cols };
}

export function forEachValueImpl(e: FormulaEngine, sheet: number, cb: (row: number, col: number, v: Scalar) => void): void {
  const s = e.sheets.get(sheet);
  if (!s) return;
  for (const [k, rec] of s.cells) {
    const r = keyRow(k);
    const c = keyCol(k);
    const sp = rec.formula?.spill;
    if (sp) {
      for (let i = 0; i < sp.rows; i++) for (let j = 0; j < sp.cols; j++) cb(r + i, c + j, sp.data[i][j]);
    } else cb(r, c, rec.value);
  }
}

/** Değişen konumu bildirir; engellenmiş/taşan dizilerin alanına yazıldıysa çapa yeniden hesaplanır. */
export function noteChangeImpl(e: FormulaEngine, g: number): void {
  e.changed.add(g);
  if (!e.spillAnchors.size) return;
  const sheet = gkSheet(g);
  const local = gkLocal(g);
  const row = keyRow(local);
  const col = keyCol(local);
  for (const a of e.spillAnchors) {
    if (a === g || gkSheet(a) !== sheet) continue;
    const t = e.sheets.get(sheet)?.cells.get(gkLocal(a))?.formula?.spillTarget;
    if (t && row >= t.r1 && row <= t.r2 && col >= t.c1 && col <= t.c2) e.dirty.add(a);
  }
}

/** Formülü verilen hücre bağlamında değerlendirir (normalleştirilmiş sonuç). */
export function evaluateImpl(e: FormulaEngine, sheet: number, row: number, col: number, formula: string): Scalar | Scalar[][] {
  const text = formula.startsWith("=") ? formula.slice(1) : formula;
  const ast = safeParse(text);
  if (!ast) return Errors.NAME;
  const ctx = createContext({ host: e, sheet, row, col, locale: e.locale });
  return normalizeResult(evaluateNode(ast, ctx), e, sheet);
}
