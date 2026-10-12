/**
 * Hedef Ara: formül hücresinin istenen değeri vermesi için değişecek hücreye
 * yazılacak değeri sekant yöntemiyle (gerekirse ikiye bölmeyle) arar.
 * En çok 100 yineleme, 0,001 tolerans (Excel varsayılanları).
 */
import type { SheetModel, WorkbookModel } from "./workbook";
import type { CellPos } from "./types";

export interface GoalSeekResult {
  found: boolean;
  /** Değişecek hücrenin bulunan değeri. */
  value: number;
  /** Hedef hücrenin bu değerdeki sonucu. */
  result: number;
  iterations: number;
}

export function goalSeek(wb: WorkbookModel, sheet: SheetModel, target: CellPos, goal: number, changing: CellPos, maxIter = 100, tol = 0.001): GoalSeekResult {
  const eng = wb.engine;
  const original = eng.hasInput(sheet.id, changing.row, changing.col) ? wb.getInput(sheet, changing.row, changing.col) : null;
  const f = (x: number): number => {
    eng.setValue(sheet.id, changing.row, changing.col, x);
    eng.recalc();
    const v = wb.getValue(sheet, target.row, target.col);
    return typeof v === "number" ? v - goal : NaN;
  };
  let x0 = typeof original === "number" ? original : 0;
  let f0 = f(x0);
  let best = { x: x0, fx: f0 };
  let it = 0;
  let found = Number.isFinite(f0) && Math.abs(f0) <= tol;
  let x1 = x0 === 0 ? 1 : x0 * 1.01;
  let f1 = f(x1);
  // Sekant yöntemi; işaret değişimi görülürse ikiye bölme ile güvenceye alınır
  const br: { lo: { x: number; fx: number } | null; hi: { x: number; fx: number } | null } = { lo: null, hi: null };
  const note = (x: number, fx: number) => {
    if (!Number.isFinite(fx)) return;
    if (Math.abs(fx) < Math.abs(best.fx) || !Number.isFinite(best.fx)) best = { x, fx };
    if (fx < 0 && (!br.lo || Math.abs(fx) < Math.abs(br.lo.fx))) br.lo = { x, fx };
    if (fx > 0 && (!br.hi || Math.abs(fx) < Math.abs(br.hi.fx))) br.hi = { x, fx };
  };
  note(x0, f0);
  note(x1, f1);
  while (!found && it < maxIter) {
    it++;
    let x2: number;
    if (br.lo && br.hi && it % 3 === 0) x2 = (br.lo.x + br.hi.x) / 2;
    else if (Number.isFinite(f0) && Number.isFinite(f1) && f1 !== f0) x2 = x1 - (f1 * (x1 - x0)) / (f1 - f0);
    else x2 = x1 + (Math.abs(x1) + 1) * (it % 2 ? 0.5 : -0.5);
    if (!Number.isFinite(x2)) break;
    const f2 = f(x2);
    note(x2, f2);
    x0 = x1;
    f0 = f1;
    x1 = x2;
    f1 = f2;
    if (Number.isFinite(f2) && Math.abs(f2) <= tol) found = true;
  }
  // Özgün değeri geri yükle (sonucu çağıran uygular)
  wb.setInputRaw(sheet, changing.row, changing.col, original);
  eng.recalc();
  return { found, value: best.x, result: best.fx + goal, iterations: it };
}
