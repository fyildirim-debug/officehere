import { FormulaEngine, parseCellRef, fromLocalFormula, dateToSerial, type Scalar, isError } from "../src";

export const NOW = dateToSerial(2024, 3, 15) + 0.5;

export function book(cells: Record<string, Scalar | string> = {}, sheetName = "Sayfa1") {
  const e = new FormulaEngine({ now: () => NOW, random: () => 0.25 });
  const s = e.addSheet(sheetName);
  const set = (a: string, v: Scalar | string, sheet = s) => {
    const r = parseCellRef(a)!;
    if (typeof v === "string" && v.startsWith("=")) e.setFormula(sheet, r.row, r.col, v);
    else e.setValue(sheet, r.row, r.col, v as Scalar);
  };
  for (const [a, v] of Object.entries(cells)) set(a, v);
  e.recalc();
  const get = (a: string, sheet = s) => {
    const r = parseCellRef(a)!;
    return e.getValue(sheet, r.row, r.col);
  };
  return { e, s, set, get };
}

/** Kanonik formülü boş bir hücrede (Z1000) değerlendirir. */
export function calc(formula: string, cells: Record<string, Scalar | string> = {}): Scalar {
  const b = book(cells);
  b.set("Z1000", formula.startsWith("=") ? formula : "=" + formula);
  b.e.recalc();
  return b.get("Z1000");
}

/** Türkçe formülü değerlendirir. */
export function calcTr(formula: string, cells: Record<string, Scalar | string> = {}): Scalar {
  return calc(fromLocalFormula(formula, "tr"), cells);
}

export function err(v: Scalar): string | null {
  return isError(v) ? v.code : null;
}
