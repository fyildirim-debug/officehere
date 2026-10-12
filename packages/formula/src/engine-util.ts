/** Hesap motoru yardımcıları: erteleme sinyali, sonuç normalleştirme, skaler karşılaştırma. */
import { Errors, isArray, isError, isLambda, isMultiRef, isRef, type Scalar, type Value } from "./values";
import type { FormulaEngine } from "./engine";

export class DeferSignal {
  constructor(readonly key: number) {}
}

/** Hücre okuma sırasında değer: hata nesnelerini kodlarına göre tekilleştirir. */
export function normalizeResult(v: Value, host: FormulaEngine, sheet: number): Scalar | Scalar[][] {
  if (v === null) return 0;
  if (typeof v !== "object" || isError(v)) return v as Scalar;
  if (isRef(v)) {
    if (v.r1 === v.r2 && v.c1 === v.c2) {
      const s = host.getCellValueInternal(v.sheet, v.r1, v.c1);
      return s === null ? 0 : s;
    }
    const rows = Math.min(v.r2, v.r1 + 100000) - v.r1 + 1;
    const cols = Math.min(v.c2, v.c1 + 16384) - v.c1 + 1;
    if (rows * cols > 2_000_000) return Errors.SPILL;
    const out: Scalar[][] = [];
    for (let i = 0; i < rows; i++) {
      const row: Scalar[] = [];
      for (let j = 0; j < cols; j++) {
        const s = host.getCellValueInternal(v.sheet, v.r1 + i, v.c1 + j);
        row.push(s === null ? 0 : s);
      }
      out.push(row);
    }
    return out;
  }
  if (isArray(v)) {
    if (v.data.length === 1 && v.data[0].length === 1) return v.data[0][0] === null ? 0 : v.data[0][0];
    return v.data.map((r) => r.map((s) => (s === null ? 0 : s)));
  }
  if (isMultiRef(v)) return Errors.VALUE;
  if (isLambda(v)) return Errors.CALC;
  void sheet;
  return Errors.VALUE;
}

export function sameScalar(a: Scalar, b: Scalar): boolean {
  if (a === b) return true;
  if (isError(a) && isError(b)) return a.code === b.code;
  return false;
}

