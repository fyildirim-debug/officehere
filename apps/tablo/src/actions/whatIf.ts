/** Durum Çözümlemesi ve veri araçları eylemleri: Hedef Ara, Akıllı Doldurma. */
import type { Controller } from "../controller";
import { displayText } from "../model/display";
import { learnProgram, runProgram } from "../model/flashFill";
import { goalSeek, type GoalSeekResult } from "../model/goalSeek";
import { currentRegion } from "../model/selection";
import type { CellPos } from "../model/types";

export function runGoalSeek(ctl: Controller, target: CellPos, goal: number, changing: CellPos): GoalSeekResult {
  return goalSeek(ctl.wb, ctl.sheet, target, goal, changing);
}

/** Hedef Ara sonucunu uygular (geri alınabilir). */
export function applyGoalSeek(ctl: Controller, changing: CellPos, value: number): void {
  ctl.ops.setInputs(ctl.sheet, [{ row: changing.row, col: changing.col, input: Math.round(value * 1e12) / 1e12 }], "Hedef Ara");
}

/** Akıllı Doldurma (Ctrl+E). Doldurulan hücre sayısını döndürür (bulunamazsa uyarır). */
export function flashFill(ctl: Controller): number {
  const { wb, sheet } = ctl;
  const { row: arow, col } = ctl.sel.active;
  // Kaynak bölge: hedef sütunun solundaki/sağındaki bitişik veri
  const left = col > 0 ? currentRegion(wb, sheet, { row: arow, col: col - 1 }) : null;
  const region = left && !wb.isEmpty(sheet, arow, col - 1) ? left : currentRegion(wb, sheet, ctl.sel.active);
  const c1 = Math.min(region.c1, col);
  const c2 = Math.max(region.c2, col);
  const rowVals = (r: number): string[] => {
    const out: string[] = [];
    for (let c = c1; c <= c2; c++) if (c !== col) out[c] = displayText(wb, sheet, r, c);
    return out;
  };
  const examples: { row: string[]; output: string; r: number }[] = [];
  for (let r = region.r1; r <= region.r2 && examples.length < 50; r++) {
    const t = displayText(wb, sheet, r, col);
    if (t !== "") examples.push({ row: rowVals(r), output: t, r });
  }
  const fail = () => {
    ctl.alert("Akıllı Doldurma bir desen bulamadı. Sütuna birkaç örnek daha yazıp yeniden deneyin.", "OfficeHere Tablo");
    return 0;
  };
  if (!examples.length) return fail();
  let prog = learnProgram(examples);
  // İlk satır başlık olabilir: onsuz yeniden dene
  if (!prog && examples.length > 1) prog = learnProgram(examples.slice(1));
  if (!prog) return fail();
  const entries: { row: number; col: number; input: string }[] = [];
  for (let r = region.r1; r <= region.r2; r++) {
    if (!wb.isEmpty(sheet, r, col)) continue;
    const vals = rowVals(r);
    if (!vals.some((v) => v)) continue;
    const out = runProgram(prog, vals);
    if (out !== null) entries.push({ row: r, col, input: out });
  }
  if (!entries.length) return fail();
  ctl.ops.setInputs(sheet, entries, "Akıllı Doldurma");
  ctl.status = `Akıllı Doldurma: ${entries.length.toLocaleString("tr-TR")} hücre dolduruldu`;
  ctl.emit();
  return entries.length;
}
