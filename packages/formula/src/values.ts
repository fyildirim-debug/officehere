/**
 * Formül motorunun değer türleri: skaler değerler, hata değerleri,
 * diziler ve hücre başvuruları.
 */

export type ErrorCode = "#NULL!" | "#DIV/0!" | "#VALUE!" | "#REF!" | "#NAME?" | "#NUM!" | "#N/A" | "#SPILL!" | "#CALC!" | "#GETTING_DATA";

export const ERROR_CODES: ErrorCode[] = ["#NULL!", "#DIV/0!", "#VALUE!", "#REF!", "#NAME?", "#NUM!", "#N/A", "#SPILL!", "#CALC!", "#GETTING_DATA"];

/** Excel hata değeri (#DIV/0!, #N/A …). */
export class FormulaError {
  constructor(
    readonly code: ErrorCode,
    readonly message?: string,
  ) {}
  toString(): string {
    return this.code;
  }
}

export const Errors = {
  NULL: new FormulaError("#NULL!"),
  DIV0: new FormulaError("#DIV/0!"),
  VALUE: new FormulaError("#VALUE!"),
  REF: new FormulaError("#REF!"),
  NAME: new FormulaError("#NAME?"),
  NUM: new FormulaError("#NUM!"),
  NA: new FormulaError("#N/A"),
  SPILL: new FormulaError("#SPILL!"),
  CALC: new FormulaError("#CALC!"),
} as const;

export function errorFromCode(code: string): FormulaError | undefined {
  const c = code.toUpperCase();
  for (const e of Object.values(Errors)) if (e.code === c) return e;
  if (c === "#GETTING_DATA") return new FormulaError("#GETTING_DATA");
  return undefined;
}

/** Skaler hücre değeri; null boş hücredir. */
export type Scalar = number | string | boolean | FormulaError | null;

/** İki boyutlu dizi (satır × sütun). */
export interface ArrayValue {
  kind: "array";
  data: Scalar[][];
}

/** Tek alanlı hücre başvurusu (0 tabanlı, uçlar dahil). */
export interface RefValue {
  kind: "ref";
  sheet: number;
  r1: number;
  c1: number;
  r2: number;
  c2: number;
}

/** Birleşim işleciyle oluşan çok alanlı başvuru. */
export interface MultiRefValue {
  kind: "refs";
  areas: RefValue[];
}

/** LAMBDA değeri (LET/LAMBDA desteği için). */
export interface LambdaValue {
  kind: "lambda";
  params: string[];
  body: unknown;
  scope: Map<string, Value> | undefined;
}

export type Value = Scalar | ArrayValue | RefValue | MultiRefValue | LambdaValue;

export function isError(v: unknown): v is FormulaError {
  return v instanceof FormulaError;
}

export function isArray(v: unknown): v is ArrayValue {
  return typeof v === "object" && v !== null && (v as ArrayValue).kind === "array";
}

export function isRef(v: unknown): v is RefValue {
  return typeof v === "object" && v !== null && (v as RefValue).kind === "ref";
}

export function isMultiRef(v: unknown): v is MultiRefValue {
  return typeof v === "object" && v !== null && (v as MultiRefValue).kind === "refs";
}

export function isLambda(v: unknown): v is LambdaValue {
  return typeof v === "object" && v !== null && (v as LambdaValue).kind === "lambda";
}

export function makeArray(data: Scalar[][]): ArrayValue {
  return { kind: "array", data };
}

export function refRows(r: RefValue): number {
  return r.r2 - r.r1 + 1;
}

export function refCols(r: RefValue): number {
  return r.c2 - r.c1 + 1;
}

/** Boş hücre mi (null)? */
export function isBlank(v: Scalar): v is null {
  return v === null;
}
