/**
 * Akıllı Doldurma canlı önizlemesi: sütuna birkaç örnek yazıldıktan sonra yeni
 * hücreye yazmaya başlayınca desen bulunursa sütunun geri kalanı gri önizleme
 * olarak gösterilir; Enter önerileri kabul eder (tek geri alma adımı).
 */
import type { Controller } from "../controller";
import { displayText } from "../model/display";
import { learnProgram, runProgram } from "../model/flashFill";
import { currentRegion } from "../model/selection";

export interface FlashPreview {
  sheetId: number;
  col: number;
  /** Düzenlenen hücre ve önerilen tam metni. */
  row: number;
  text: string;
  /** Altındaki boş hücrelere önerilenler. */
  fills: { row: number; text: string }[];
}

let cache: { key: string; value: FlashPreview | null } | null = null;

export function flashPreviewFor(ctl: Controller): FlashPreview | null {
  const e = ctl.edit;
  if (!e || e.inBar || !e.text || e.text.startsWith("=") || e.text.startsWith("'")) return null;
  const key = `${ctl.dataVersion}:${e.sheetId}:${e.row}:${e.col}:${e.text}`;
  if (cache?.key === key) return cache.value;
  let value: FlashPreview | null = null;
  try {
    value = compute(ctl, e.sheetId, e.row, e.col, e.text);
  } catch {
    value = null;
  }
  cache = { key, value };
  return value;
}

function compute(ctl: Controller, sheetId: number, row: number, col: number, typed: string): FlashPreview | null {
  const wb = ctl.wb;
  const sheet = wb.sheetById(sheetId);
  if (!sheet || col === 0 || wb.isEmpty(sheet, row, col - 1)) return null;
  const region = currentRegion(wb, sheet, { row, col: col - 1 });
  if (region.c2 < col - 1 || row < region.r1 || row > region.r2 + 0) return null;
  const c1 = region.c1;
  const c2 = Math.max(region.c2, col);
  const rowVals = (r: number): string[] => {
    const out: string[] = [];
    for (let c = c1; c <= c2; c++) if (c !== col) out[c] = displayText(wb, sheet, r, c);
    return out;
  };
  // Örnekler: aynı sütunda yukarıdaki dolu hücreler (en çok 20)
  const examples: { row: string[]; output: string }[] = [];
  for (let r = Math.max(region.r1, row - 200); r < row && examples.length < 20; r++) {
    const t = displayText(wb, sheet, r, col);
    if (t !== "") examples.push({ row: rowVals(r), output: t });
  }
  if (!examples.length) return null;
  let prog = learnProgram(examples);
  if (!prog && examples.length > 1) prog = learnProgram(examples.slice(1));
  if (!prog) return null;
  const full = runProgram(prog, rowVals(row));
  if (full === null || !full.startsWith(typed)) return null;
  const fills: { row: number; text: string }[] = [];
  for (let r = row + 1; r <= region.r2 && fills.length < 5000; r++) {
    if (!wb.isEmpty(sheet, r, col)) continue;
    const vals = rowVals(r);
    if (!vals.some((v) => v)) continue;
    const out = runProgram(prog, vals);
    if (out !== null) fills.push({ row: r, text: out });
  }
  if (!fills.length) return null;
  return { sheetId, col, row, text: full, fills };
}

/** Enter: düzenlenen hücre önerilen metinle onaylanır, aşağıdakiler doldurulur. */
export function acceptFlashPreview(ctl: Controller, p: FlashPreview): void {
  const sheet = ctl.wb.sheetById(p.sheetId);
  if (!sheet) return;
  ctl.edit = null;
  ctl.ops.setInputs(sheet, [{ row: p.row, col: p.col, input: p.text }, ...p.fills.map((f) => ({ row: f.row, col: p.col, input: f.text }))], "Akıllı Doldurma");
  ctl.status = `Akıllı Doldurma: ${(p.fills.length + 1).toLocaleString("tr-TR")} hücre dolduruldu`;
  ctl.selectCell(p.row + 1, p.col);
}
