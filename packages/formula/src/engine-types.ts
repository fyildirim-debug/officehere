/** Hesap motorunun ortak türleri ve hücre anahtarları. */
import type { Node } from "./ast";
import { parseFormula } from "./locale";
import type { LocaleInfo } from "./coerce";
import { MAX_COLS } from "./references";
import type { Scalar } from "./values";

export const SHEET_STRIDE = 17179869184; // 2^34
export const ROW_STRIDE = MAX_COLS;
export const DEPTH_LIMIT = 300;

export function cellKey(row: number, col: number): number {
  return row * ROW_STRIDE + col;
}

export function keyRow(k: number): number {
  return Math.floor(k / ROW_STRIDE);
}

export function keyCol(k: number): number {
  return k % ROW_STRIDE;
}

export function gk(sheet: number, row: number, col: number): number {
  return sheet * SHEET_STRIDE + row * ROW_STRIDE + col;
}

export function gkSheet(k: number): number {
  return Math.floor(k / SHEET_STRIDE);
}

export function gkLocal(k: number): number {
  return k % SHEET_STRIDE;
}

export interface RangeDep {
  sheet: number;
  r1: number;
  c1: number;
  r2: number;
  c2: number;
  f: number;
}

export interface SpillState {
  rows: number;
  cols: number;
  data: Scalar[][];
}

export interface FormulaInfo {
  text: string;
  ast: Node | null;
  volatile: boolean;
  usesNames: boolean;
  cellDeps: number[];
  rangeDeps: RangeDep[];
  spill: SpillState | null;
  /** Taşmanın hedeflediği alan (engellenmiş olsa bile). */
  spillTarget: { r1: number; c1: number; r2: number; c2: number } | null;
}

export interface CellRecord {
  value: Scalar;
  formula?: FormulaInfo;
}

export interface SheetData {
  id: number;
  name: string;
  cells: Map<number, CellRecord>;
  maxRow: number;
  maxCol: number;
  extentDirty: boolean;
  hiddenRows: Set<number>;
  filteredRows: Set<number>;
}

export interface DefinedName {
  name: string;
  text: string;
  ast: Node | null;
  /** Sayfa kapsamı (yoksa çalışma kitabı). */
  sheet?: number;
}


export interface EngineOptions {
  locale?: LocaleInfo;
  now?: () => number;
  random?: () => number;
}

export interface CellSnapshot {
  row: number;
  col: number;
  value?: Scalar;
  formula?: string;
}

export interface EngineSnapshot {
  sheets: { id: number; name: string; cells: CellSnapshot[] }[];
  names: { name: string; text: string; sheet?: number }[];
}

export interface SpillInfo {
  anchorRow: number;
  anchorCol: number;
  rows: number;
  cols: number;
}

export function safeParse(text: string): Node | null {
  try {
    return parseFormula(text);
  } catch {
    return null;
  }
}
