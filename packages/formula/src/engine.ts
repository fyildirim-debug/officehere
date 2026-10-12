/**
 * Hesap motoru: sayfalar, seyrek hücre deposu, bağımlılık grafiği,
 * artımlı yeniden hesaplama, döngüsel başvuru algılama, dizi taşması
 * (#SPILL!), adlandırılmış aralıklar ve yapısal değişikliklerde
 * başvuru güncelleme.
 */
import "./functions";
import { walk, type Node } from "./ast";
import { TR_LOCALE, type LocaleInfo } from "./coerce";
import type { EvalHost } from "./context";
import { jsDateToSerial } from "./dates";
import { createContext, evaluateNode } from "./evaluator";
import { getFunction } from "./functions/registry";
import { parseFormula } from "./locale";
import { MAX_COLS, MAX_ROWS } from "./references";
import { adjustFormula, moveRefsInFormula, removeSheetFromFormula, renameSheetInFormula, shiftFormula, type MoveSpec } from "./transform";
import { Errors, FormulaError, isArray, isError, isLambda, isMultiRef, isRef, type RefValue, type Scalar, type Value } from "./values";

const SHEET_STRIDE = 17179869184; // 2^34
const ROW_STRIDE = MAX_COLS;
const DEPTH_LIMIT = 300;

export function cellKey(row: number, col: number): number {
  return row * ROW_STRIDE + col;
}

export function keyRow(k: number): number {
  return Math.floor(k / ROW_STRIDE);
}

export function keyCol(k: number): number {
  return k % ROW_STRIDE;
}

function gk(sheet: number, row: number, col: number): number {
  return sheet * SHEET_STRIDE + row * ROW_STRIDE + col;
}

function gkSheet(k: number): number {
  return Math.floor(k / SHEET_STRIDE);
}

function gkLocal(k: number): number {
  return k % SHEET_STRIDE;
}

interface RangeDep {
  sheet: number;
  r1: number;
  c1: number;
  r2: number;
  c2: number;
  f: number;
}

interface SpillState {
  rows: number;
  cols: number;
  data: Scalar[][];
}

interface FormulaInfo {
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

interface SheetData {
  id: number;
  name: string;
  cells: Map<number, CellRecord>;
  maxRow: number;
  maxCol: number;
  extentDirty: boolean;
  hiddenRows: Set<number>;
  filteredRows: Set<number>;
}

interface DefinedName {
  name: string;
  text: string;
  ast: Node | null;
  /** Sayfa kapsamı (yoksa çalışma kitabı). */
  sheet?: number;
}

class DeferSignal {
  constructor(readonly key: number) {}
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

/** Hücre okuma sırasında değer: hata nesnelerini kodlarına göre tekilleştirir. */
function normalizeResult(v: Value, host: FormulaEngine, sheet: number): Scalar | Scalar[][] {
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

export class FormulaEngine implements EvalHost {
  locale: LocaleInfo;
  autoCalc = true;
  private sheets = new Map<number, SheetData>();
  private order: number[] = [];
  private nextSheetId = 1;
  private names = new Map<string, DefinedName>();
  private cellDependents = new Map<number, Set<number>>();
  private rangeIndex = new Map<number, Map<number, Set<RangeDep>>>();
  private wideRanges = new Map<number, Set<RangeDep>>();
  private volatiles = new Set<number>();
  private nameUsers = new Set<number>();
  private spillOwner = new Map<number, number>();
  private spillAnchors = new Set<number>();
  private dirty = new Set<number>();
  private changed = new Set<number>();
  private computing = new Set<number>();
  private computingStack: number[] = [];
  private depth = 0;
  private circular = new Set<number>();
  private nowFn: () => number;
  private randomFn: () => number;
  /** Aynı yeniden hesaplama içinde arama dizinleri. */
  lookupCache = new Map<string, Map<string, number>>();
  /** Son hesaplamada değeri değişen ya da yeniden hesaplanan hücre sayısı. */
  version = 0;

  constructor(opts: EngineOptions = {}) {
    this.locale = opts.locale ?? TR_LOCALE;
    this.nowFn = opts.now ?? (() => jsDateToSerial(new Date()));
    this.randomFn = opts.random ?? Math.random;
  }

  /* ------------------------------ Sayfalar ------------------------------ */

  addSheet(name: string, index?: number): number {
    if (this.getSheetId(name) !== undefined) throw new Error(`"${name}" adlı sayfa zaten var`);
    const id = this.nextSheetId++;
    this.sheets.set(id, { id, name, cells: new Map(), maxRow: -1, maxCol: -1, extentDirty: false, hiddenRows: new Set(), filteredRows: new Set() });
    if (index === undefined || index >= this.order.length) this.order.push(id);
    else this.order.splice(Math.max(0, index), 0, id);
    // Bu ada yapılmış #REF! olmayan başvurular yeniden çözülsün
    this.reregisterAll();
    this.markAllDirty();
    return id;
  }

  removeSheet(id: number): void {
    const s = this.sheets.get(id);
    if (!s) return;
    for (const [k] of s.cells) this.unregister(id * SHEET_STRIDE + k);
    this.sheets.delete(id);
    this.order = this.order.filter((x) => x !== id);
    for (const sid of this.order) {
      const sh = this.sheets.get(sid)!;
      for (const [k, rec] of sh.cells) {
        if (!rec.formula) continue;
        const t = removeSheetFromFormula(rec.formula.text, s.name);
        if (t !== rec.formula.text) this.setFormulaInternal(sid, keyRow(k), keyCol(k), t);
      }
    }
    for (const [key, n] of this.names) {
      if (n.sheet === id) this.names.delete(key);
      else {
        n.text = removeSheetFromFormula(n.text, s.name);
        n.ast = safeParse(n.text);
      }
    }
    for (const k of [...this.spillOwner.keys()]) if (gkSheet(k) === id) this.spillOwner.delete(k);
    for (const k of [...this.spillAnchors]) if (gkSheet(k) === id) this.spillAnchors.delete(k);
    this.reregisterAll();
    this.markAllDirty();
  }

  renameSheet(id: number, name: string): void {
    const s = this.sheets.get(id);
    if (!s || s.name === name) return;
    const other = this.getSheetId(name);
    if (other !== undefined && other !== id) throw new Error(`"${name}" adlı sayfa zaten var`);
    const old = s.name;
    s.name = name;
    for (const sid of this.order) {
      const sh = this.sheets.get(sid)!;
      for (const [, rec] of sh.cells) {
        if (!rec.formula) continue;
        const t = renameSheetInFormula(rec.formula.text, old, name);
        if (t !== rec.formula.text) {
          rec.formula.text = t;
          rec.formula.ast = safeParse(t);
        }
      }
    }
    for (const n of this.names.values()) {
      n.text = renameSheetInFormula(n.text, old, name);
      n.ast = safeParse(n.text);
    }
    this.reregisterAll();
    this.markAllDirty();
  }

  moveSheet(id: number, index: number): void {
    const i = this.order.indexOf(id);
    if (i < 0) return;
    this.order.splice(i, 1);
    this.order.splice(Math.max(0, Math.min(index, this.order.length)), 0, id);
  }

  sheetIds(): number[] {
    return [...this.order];
  }

  getSheetName(id: number): string | undefined {
    return this.sheets.get(id)?.name;
  }

  getSheetId(name: string): number | undefined {
    const u = name.toLocaleUpperCase("tr-TR");
    for (const id of this.order) if (this.sheets.get(id)!.name.toLocaleUpperCase("tr-TR") === u) return id;
    return undefined;
  }

  setRowHidden(sheet: number, row: number, hidden: boolean, filtered = false): void {
    const s = this.sheets.get(sheet);
    if (!s) return;
    const set = filtered ? s.filteredRows : s.hiddenRows;
    if (hidden) set.add(row);
    else set.delete(row);
    // ALTTOPLAM gizli satırlara bağlıdır
    this.markSubtotalsDirty(sheet);
  }

  clearFilteredRows(sheet: number): void {
    const s = this.sheets.get(sheet);
    if (!s) return;
    s.filteredRows.clear();
    this.markSubtotalsDirty(sheet);
  }

  private markSubtotalsDirty(sheet: number): void {
    void sheet;
    for (const id of this.order) {
      const sh = this.sheets.get(id)!;
      for (const [k, rec] of sh.cells) if (rec.formula && /SUBTOTAL|AGGREGATE/i.test(rec.formula.text)) this.dirty.add(id * SHEET_STRIDE + k);
    }
  }

  /* ------------------------------ EvalHost ------------------------------ */

  getCellValue(sheet: number, row: number, col: number): Scalar {
    return this.getCellValueInternal(sheet, row, col);
  }

  getCellValueInternal(sheet: number, row: number, col: number): Scalar {
    const k = gk(sheet, row, col);
    if (this.dirty.has(k)) this.ensureDuringEval(k);
    const s = this.sheets.get(sheet);
    if (!s) return Errors.REF;
    const rec = s.cells.get(cellKey(row, col));
    if (rec) {
      if (rec.formula && this.circularRead(k)) return 0;
      return rec.value;
    }
    const owner = this.spillOwner.get(k);
    if (owner !== undefined) {
      if (this.dirty.has(owner)) this.ensureDuringEval(owner);
      const ownerRec = this.sheets.get(gkSheet(owner))?.cells.get(gkLocal(owner));
      const sp = ownerRec?.formula?.spill;
      if (sp) {
        const or = keyRow(gkLocal(owner));
        const oc = keyCol(gkLocal(owner));
        return sp.data[row - or]?.[col - oc] ?? null;
      }
    }
    return null;
  }

  private circularRead(k: number): boolean {
    if (this.computing.has(k)) {
      this.markCycle(k);
      return true;
    }
    return false;
  }

  /** Hesaplanan yığındaki k'den itibaren tüm hücreler döngünün parçasıdır. */
  private markCycle(k: number): void {
    const i = this.computingStack.lastIndexOf(k);
    for (let j = Math.max(0, i); j < this.computingStack.length; j++) this.circular.add(this.computingStack[j]);
    this.circular.add(k);
  }

  private ensureDuringEval(k: number): void {
    if (this.computing.has(k)) {
      this.markCycle(k);
      return;
    }
    if (this.depth >= DEPTH_LIMIT) throw new DeferSignal(k);
    this.compute(k);
  }

  getUsedExtent(sheet: number): { rows: number; cols: number } {
    const s = this.sheets.get(sheet);
    if (!s) return { rows: 0, cols: 0 };
    if (s.extentDirty) this.recomputeExtent(s);
    return { rows: s.maxRow + 1, cols: s.maxCol + 1 };
  }

  private recomputeExtent(s: SheetData): void {
    let mr = -1;
    let mc = -1;
    for (const [k, rec] of s.cells) {
      const r = keyRow(k);
      const c = keyCol(k);
      let rr = r;
      let cc = c;
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

  private growExtent(s: SheetData, row: number, col: number): void {
    if (row > s.maxRow) s.maxRow = row;
    if (col > s.maxCol) s.maxCol = col;
  }

  sheetIdByName(name: string): number | undefined {
    return this.getSheetId(name);
  }

  sheetName(id: number): string | undefined {
    return this.getSheetName(id);
  }

  sheetIndex(id: number): number {
    return this.order.indexOf(id);
  }

  sheetCount(): number {
    return this.order.length;
  }

  getDefinedName(name: string, sheet: number): Node | undefined {
    const u = name.toLocaleUpperCase("tr-TR");
    const local = this.names.get(`${sheet}!${u}`);
    if (local?.ast) return local.ast;
    return this.names.get(u)?.ast ?? undefined;
  }

  isRowHidden(sheet: number, row: number): boolean {
    const s = this.sheets.get(sheet);
    return !!s && (s.hiddenRows.has(row) || s.filteredRows.has(row));
  }

  isRowFiltered(sheet: number, row: number): boolean {
    return !!this.sheets.get(sheet)?.filteredRows.has(row);
  }

  getFormulaText(sheet: number, row: number, col: number): string | undefined {
    const f = this.sheets.get(sheet)?.cells.get(cellKey(row, col))?.formula;
    return f ? "=" + f.text : undefined;
  }

  isSubtotalCell(sheet: number, row: number, col: number): boolean {
    const f = this.sheets.get(sheet)?.cells.get(cellKey(row, col))?.formula;
    return !!f && /^\s*(SUBTOTAL|AGGREGATE)\s*\(/i.test(f.text);
  }

  now(): number {
    return this.nowFn();
  }

  random(): number {
    return this.randomFn();
  }

  /* ------------------------------ Hücreler ------------------------------ */

  private sheetOrThrow(id: number): SheetData {
    const s = this.sheets.get(id);
    if (!s) throw new Error(`Sayfa yok: ${id}`);
    return s;
  }

  /** Sabit değer yazar (null hücreyi temizler). */
  setValue(sheet: number, row: number, col: number, value: Scalar): void {
    const s = this.sheetOrThrow(sheet);
    const k = cellKey(row, col);
    const g = sheet * SHEET_STRIDE + k;
    const old = s.cells.get(k);
    if (old?.formula) this.unregister(g);
    if (value === null) {
      if (old) {
        s.cells.delete(k);
        s.extentDirty = true;
      }
    } else {
      s.cells.set(k, { value });
      this.growExtent(s, row, col);
    }
    this.dirty.delete(g);
    this.noteChange(g);
  }

  /** Formül yazar ("=" ile ya da olmadan, kanonik İngilizce sözdizimi). */
  setFormula(sheet: number, row: number, col: number, formula: string): void {
    this.setFormulaInternal(sheet, row, col, formula.startsWith("=") ? formula.slice(1) : formula);
  }

  private setFormulaInternal(sheet: number, row: number, col: number, text: string): void {
    const s = this.sheetOrThrow(sheet);
    const k = cellKey(row, col);
    const g = sheet * SHEET_STRIDE + k;
    const old = s.cells.get(k);
    if (old?.formula) this.unregister(g);
    const ast = safeParse(text);
    const info: FormulaInfo = { text, ast, volatile: false, usesNames: false, cellDeps: [], rangeDeps: [], spill: null, spillTarget: null };
    const rec: CellRecord = { value: ast ? 0 : Errors.NAME, formula: info };
    s.cells.set(k, rec);
    this.growExtent(s, row, col);
    this.register(g, info, sheet);
    this.dirty.add(g);
    this.noteChange(g);
  }

  clearCell(sheet: number, row: number, col: number): void {
    this.setValue(sheet, row, col, null);
  }

  /** Hücre içeriğini (değer ya da "=formül") yazar. */
  setInput(sheet: number, row: number, col: number, input: Scalar | { formula: string }): void {
    if (input !== null && typeof input === "object" && !(input instanceof FormulaError)) this.setFormula(sheet, row, col, input.formula);
    else this.setValue(sheet, row, col, input as Scalar);
  }

  getValue(sheet: number, row: number, col: number): Scalar {
    if (this.dirty.size && this.autoCalc && !this.inRecalc) this.recalc();
    const s = this.sheets.get(sheet);
    if (!s) return null;
    const rec = s.cells.get(cellKey(row, col));
    if (rec) return rec.value;
    const owner = this.spillOwner.get(gk(sheet, row, col));
    if (owner !== undefined) {
      const sp = this.sheets.get(gkSheet(owner))?.cells.get(gkLocal(owner))?.formula?.spill;
      if (sp) return sp.data[row - keyRow(gkLocal(owner))]?.[col - keyCol(gkLocal(owner))] ?? null;
    }
    return null;
  }

  /** Formül metni ("=" ile) ya da undefined. */
  getFormula(sheet: number, row: number, col: number): string | undefined {
    const f = this.sheets.get(sheet)?.cells.get(cellKey(row, col))?.formula;
    return f ? "=" + f.text : undefined;
  }

  /** Hücrede kullanıcı girdisi (sabit ya da formül) var mı? */
  hasInput(sheet: number, row: number, col: number): boolean {
    return !!this.sheets.get(sheet)?.cells.has(cellKey(row, col));
  }

  /** Hücre taşan bir dizinin parçasıysa bilgisi. */
  getSpill(sheet: number, row: number, col: number): SpillInfo | undefined {
    const k = gk(sheet, row, col);
    const ownerKey = this.spillOwner.get(k) ?? (this.sheets.get(sheet)?.cells.get(cellKey(row, col))?.formula?.spill ? k : undefined);
    if (ownerKey === undefined) return undefined;
    const sp = this.sheets.get(gkSheet(ownerKey))?.cells.get(gkLocal(ownerKey))?.formula?.spill;
    if (!sp) return undefined;
    return { anchorRow: keyRow(gkLocal(ownerKey)), anchorCol: keyCol(gkLocal(ownerKey)), rows: sp.rows, cols: sp.cols };
  }

  /** Sayfadaki girdili hücreleri gezer. */
  forEachCell(sheet: number, cb: (row: number, col: number, rec: CellRecord) => void): void {
    const s = this.sheets.get(sheet);
    if (!s) return;
    for (const [k, rec] of s.cells) cb(keyRow(k), keyCol(k), rec);
  }

  cellCount(sheet: number): number {
    return this.sheets.get(sheet)?.cells.size ?? 0;
  }

  /** Taşan değerler dahil, değeri olan tüm konumlar. */
  forEachValue(sheet: number, cb: (row: number, col: number, v: Scalar) => void): void {
    const s = this.sheets.get(sheet);
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

  /* ------------------------- Bağımlılık grafiği ------------------------- */

  private register(g: number, info: FormulaInfo, sheet: number): void {
    if (!info.ast) return;
    const cellDeps: number[] = [];
    const rangeDeps: RangeDep[] = [];
    let volatile = false;
    let usesNames = false;
    const seenNames = new Set<string>();
    const visit = (n: Node, depth: number) => {
      if (n.type === "func") {
        const def = getFunction(n.name);
        if (def?.volatile) volatile = true;
      } else if (n.type === "name") {
        usesNames = true;
        const key = n.name.toLocaleUpperCase("tr-TR");
        if (seenNames.has(key) || depth > 8) return;
        seenNames.add(key);
        const ast = this.getDefinedName(n.name, sheet);
        if (ast) walk(ast, (m) => visit(m, depth + 1));
      } else if (n.type === "ref") {
        const t = n.ref;
        if (t.refError) return;
        const sid = t.sheet !== undefined ? this.getSheetId(t.sheet) : sheet;
        if (sid === undefined) return;
        if (t.kind === "cell") cellDeps.push(gk(sid, t.a.row, t.a.col));
        else {
          const b = t.b!;
          let r1: number, r2: number, c1: number, c2: number;
          if (t.kind === "cols") {
            r1 = 0;
            r2 = MAX_ROWS - 1;
            c1 = Math.min(t.a.col, b.col);
            c2 = Math.max(t.a.col, b.col);
          } else if (t.kind === "rows") {
            c1 = 0;
            c2 = MAX_COLS - 1;
            r1 = Math.min(t.a.row, b.row);
            r2 = Math.max(t.a.row, b.row);
          } else {
            r1 = Math.min(t.a.row, b.row);
            r2 = Math.max(t.a.row, b.row);
            c1 = Math.min(t.a.col, b.col);
            c2 = Math.max(t.a.col, b.col);
          }
          if (r1 === r2 && c1 === c2) cellDeps.push(gk(sid, r1, c1));
          else rangeDeps.push({ sheet: sid, r1, c1, r2, c2, f: g });
        }
      } else if (n.type === "binary" && n.op === ":" && n.left.type === "ref" && n.right.type === "ref") {
        // A1:B2 dışındaki ":" kullanımları (ör. Sayfa1!A1:Sayfa1!B2) — sınırlayıcı kutu
        const l = n.left.ref;
        const r = n.right.ref;
        if (!l.refError && !r.refError && l.kind === "cell" && r.kind === "cell") {
          const sid = l.sheet !== undefined ? this.getSheetId(l.sheet) : sheet;
          if (sid !== undefined)
            rangeDeps.push({ sheet: sid, r1: Math.min(l.a.row, r.a.row), c1: Math.min(l.a.col, r.a.col), r2: Math.max(l.a.row, r.a.row), c2: Math.max(l.a.col, r.a.col), f: g });
        }
      }
    };
    walk(info.ast, (n) => visit(n, 0));
    info.cellDeps = cellDeps;
    info.rangeDeps = rangeDeps;
    info.volatile = volatile;
    info.usesNames = usesNames;
    for (const d of cellDeps) {
      let set = this.cellDependents.get(d);
      if (!set) this.cellDependents.set(d, (set = new Set()));
      set.add(g);
    }
    for (const rd of rangeDeps) this.addRangeDep(rd);
    if (volatile) this.volatiles.add(g);
    if (usesNames) this.nameUsers.add(g);
  }

  private addRangeDep(rd: RangeDep): void {
    const b1 = rd.c1 >> 4;
    const b2 = rd.c2 >> 4;
    if (b2 - b1 > 32) {
      let set = this.wideRanges.get(rd.sheet);
      if (!set) this.wideRanges.set(rd.sheet, (set = new Set()));
      set.add(rd);
      return;
    }
    let idx = this.rangeIndex.get(rd.sheet);
    if (!idx) this.rangeIndex.set(rd.sheet, (idx = new Map()));
    for (let b = b1; b <= b2; b++) {
      let set = idx.get(b);
      if (!set) idx.set(b, (set = new Set()));
      set.add(rd);
    }
  }

  private removeRangeDep(rd: RangeDep): void {
    const b1 = rd.c1 >> 4;
    const b2 = rd.c2 >> 4;
    if (b2 - b1 > 32) {
      this.wideRanges.get(rd.sheet)?.delete(rd);
      return;
    }
    const idx = this.rangeIndex.get(rd.sheet);
    if (!idx) return;
    for (let b = b1; b <= b2; b++) idx.get(b)?.delete(rd);
  }

  private unregister(g: number): void {
    const s = this.sheets.get(gkSheet(g));
    const rec = s?.cells.get(gkLocal(g));
    const info = rec?.formula;
    if (!info) return;
    for (const d of info.cellDeps) this.cellDependents.get(d)?.delete(g);
    for (const rd of info.rangeDeps) this.removeRangeDep(rd);
    info.cellDeps = [];
    info.rangeDeps = [];
    this.volatiles.delete(g);
    this.nameUsers.delete(g);
    if (info.spill || info.spillTarget) this.clearSpill(g, info);
    this.spillAnchors.delete(g);
    this.dirty.delete(g);
  }

  private reregisterAll(): void {
    this.cellDependents.clear();
    this.rangeIndex.clear();
    this.wideRanges.clear();
    this.volatiles.clear();
    this.nameUsers.clear();
    for (const id of this.order) {
      const s = this.sheets.get(id)!;
      for (const [k, rec] of s.cells) if (rec.formula) this.register(id * SHEET_STRIDE + k, rec.formula, id);
    }
  }

  /** Konumdaki değişikliğe bağlı formüller. */
  private dependentsOf(g: number, out: (f: number) => void): void {
    const set = this.cellDependents.get(g);
    if (set) for (const f of set) out(f);
    const sheet = gkSheet(g);
    const local = gkLocal(g);
    const row = keyRow(local);
    const col = keyCol(local);
    const idx = this.rangeIndex.get(sheet);
    const bucket = idx?.get(col >> 4);
    if (bucket) for (const rd of bucket) if (row >= rd.r1 && row <= rd.r2 && col >= rd.c1 && col <= rd.c2) out(rd.f);
    const wide = this.wideRanges.get(sheet);
    if (wide) for (const rd of wide) if (row >= rd.r1 && row <= rd.r2 && col >= rd.c1 && col <= rd.c2) out(rd.f);
  }

  private noteChange(g: number): void {
    this.changed.add(g);
    // Engellenmiş/taşan dizilerin alanına yazıldıysa çapa yeniden hesaplanmalı
    if (this.spillAnchors.size) {
      const sheet = gkSheet(g);
      const local = gkLocal(g);
      const row = keyRow(local);
      const col = keyCol(local);
      for (const a of this.spillAnchors) {
        if (a === g || gkSheet(a) !== sheet) continue;
        const t = this.sheets.get(sheet)?.cells.get(gkLocal(a))?.formula?.spillTarget;
        if (t && row >= t.r1 && row <= t.r2 && col >= t.c1 && col <= t.c2) this.dirty.add(a);
      }
    }
  }

  markAllDirty(): void {
    for (const id of this.order) {
      const s = this.sheets.get(id)!;
      for (const [k, rec] of s.cells) if (rec.formula) this.dirty.add(id * SHEET_STRIDE + k);
    }
  }

  /** Yeniden hesaplanmayı bekleyen formül var mı? */
  get pending(): boolean {
    return this.dirty.size > 0 || this.changed.size > 0;
  }

  /* --------------------------- Yeniden hesaplama --------------------------- */

  private inRecalc = false;

  /** Değişikliklerden etkilenen formülleri yeniden hesaplar. full: tümünü. */
  recalc(full = false): void {
    if (this.inRecalc) return;
    this.inRecalc = true;
    try {
      if (full) this.markAllDirty();
      for (const v of this.volatiles) this.dirty.add(v);
      this.lookupCache.clear();
      for (let pass = 0; pass < 8; pass++) {
        // Değişen konumlardan bağımlıları yay
        const queue: number[] = [...this.changed, ...this.dirty];
        this.changed.clear();
        const seen = new Set<number>(queue);
        while (queue.length) {
          const g = queue.pop()!;
          this.dependentsOf(g, (f) => {
            if (!seen.has(f)) {
              seen.add(f);
              this.dirty.add(f);
              queue.push(f);
            }
          });
        }
        if (this.dirty.size === 0) break;
        this.circular.clear();
        const order = [...this.dirty].sort((a, b) => a - b);
        for (const k of order) this.ensure(k);
        if (this.changed.size === 0) break;
      }
      this.version++;
    } finally {
      this.inRecalc = false;
      this.lookupCache.clear();
    }
  }

  /** Döngüsel başvuru içeren hücreler (sayfa, satır, sütun). */
  getCircularRefs(): { sheet: number; row: number; col: number }[] {
    return [...this.circular].map((g) => ({ sheet: gkSheet(g), row: keyRow(gkLocal(g)), col: keyCol(gkLocal(g)) }));
  }

  private ensure(k: number): void {
    if (!this.dirty.has(k)) return;
    const stack = [k];
    let guard = 0;
    while (stack.length) {
      if (++guard > 1_000_000) break;
      const top = stack[stack.length - 1];
      if (!this.dirty.has(top)) {
        stack.pop();
        continue;
      }
      try {
        this.compute(top);
        stack.pop();
      } catch (e) {
        if (e instanceof DeferSignal) {
          if (stack.includes(e.key)) {
            // Yığın üzerinden döngü: değeri 0 kabul et
            this.circular.add(e.key);
            const rec = this.sheets.get(gkSheet(e.key))?.cells.get(gkLocal(e.key));
            if (rec) rec.value = 0;
            this.dirty.delete(e.key);
          } else stack.push(e.key);
          continue;
        }
        throw e;
      }
    }
  }

  private compute(g: number): void {
    const sheet = gkSheet(g);
    const local = gkLocal(g);
    const s = this.sheets.get(sheet);
    const rec = s?.cells.get(local);
    if (!s || !rec?.formula) {
      this.dirty.delete(g);
      return;
    }
    const info = rec.formula;
    const row = keyRow(local);
    const col = keyCol(local);
    this.computing.add(g);
    this.computingStack.push(g);
    this.depth++;
    let result: Value;
    try {
      if (!info.ast) result = Errors.NAME;
      else {
        const ctx = createContext({ host: this, sheet, row, col, locale: this.locale });
        result = evaluateNode(info.ast, ctx);
      }
    } finally {
      this.computing.delete(g);
      this.computingStack.pop();
      this.depth--;
    }
    // Döngüsel başvurudaki hücreler 0 gösterilir (yinelemeli hesaplama kapalı)
    if (this.circular.has(g)) result = 0;
    const norm = normalizeResult(result, this, sheet);
    const oldValue = rec.value;
    if (Array.isArray(norm)) {
      this.applySpill(g, s, info, row, col, norm);
    } else {
      if (info.spill || info.spillTarget) this.clearSpill(g, info);
      rec.value = norm;
    }
    this.dirty.delete(g);
    if (!sameScalar(oldValue, rec.value) && this.inRecalc) {
      // Değer değişti: bu turda henüz hesaplanmamış bağımlılar zaten kirli;
      // hesaplanmış olanlar sıralama nedeniyle kirliyse yeniden işlenir.
    }
  }

  private applySpill(g: number, s: SheetData, info: FormulaInfo, row: number, col: number, data: Scalar[][]): void {
    const rows = data.length;
    const cols = data[0]?.length ?? 0;
    const target = { r1: row, c1: col, r2: row + rows - 1, c2: col + cols - 1 };
    const rec = s.cells.get(cellKey(row, col))!;
    let blocked = target.r2 >= MAX_ROWS || target.c2 >= MAX_COLS;
    if (!blocked) {
      outer: for (let i = 0; i < rows; i++) {
        for (let j = 0; j < cols; j++) {
          if (i === 0 && j === 0) continue;
          const k = cellKey(row + i, col + j);
          if (s.cells.has(k)) {
            blocked = true;
            break outer;
          }
          const owner = this.spillOwner.get(s.id * SHEET_STRIDE + k);
          if (owner !== undefined && owner !== g) {
            blocked = true;
            break outer;
          }
        }
      }
    }
    const old = info.spill;
    const sameShape = old && old.rows === rows && old.cols === cols;
    if (blocked) {
      if (old) this.clearSpill(g, info);
      info.spillTarget = target;
      this.spillAnchors.add(g);
      rec.value = Errors.SPILL;
      return;
    }
    if (!sameShape) {
      if (old) this.clearSpill(g, info);
      for (let i = 0; i < rows; i++)
        for (let j = 0; j < cols; j++) {
          if (i === 0 && j === 0) continue;
          const pk = gk(s.id, row + i, col + j);
          this.spillOwner.set(pk, g);
          this.changed.add(pk);
        }
      this.growExtent(s, target.r2, target.c2);
    } else {
      // Aynı şekil: değişen hücreleri bildir
      for (let i = 0; i < rows; i++)
        for (let j = 0; j < cols; j++) {
          if (i === 0 && j === 0) continue;
          if (!sameScalar(old!.data[i][j], data[i][j])) this.changed.add(gk(s.id, row + i, col + j));
        }
    }
    info.spill = { rows, cols, data };
    info.spillTarget = target;
    this.spillAnchors.add(g);
    rec.value = data[0][0];
  }

  private clearSpill(g: number, info: FormulaInfo): void {
    const sp = info.spill;
    if (sp) {
      const sheet = gkSheet(g);
      const local = gkLocal(g);
      const row = keyRow(local);
      const col = keyCol(local);
      for (let i = 0; i < sp.rows; i++)
        for (let j = 0; j < sp.cols; j++) {
          if (i === 0 && j === 0) continue;
          const pk = gk(sheet, row + i, col + j);
          if (this.spillOwner.get(pk) === g) {
            this.spillOwner.delete(pk);
            this.changed.add(pk);
          }
        }
      const s = this.sheets.get(sheet);
      if (s) s.extentDirty = true;
    }
    info.spill = null;
    info.spillTarget = null;
    this.spillAnchors.delete(g);
  }

  /** Bir formülü verilen hücre bağlamında değerlendirir (koşullu biçim, veri doğrulama). */
  evaluate(sheet: number, row: number, col: number, formula: string): Scalar {
    const text = formula.startsWith("=") ? formula.slice(1) : formula;
    const ast = safeParse(text);
    if (!ast) return Errors.NAME;
    if (this.dirty.size && this.autoCalc && !this.inRecalc) this.recalc();
    const ctx = createContext({ host: this, sheet, row, col, locale: this.locale });
    const v = evaluateNode(ast, ctx);
    const n = normalizeResult(v, this, sheet);
    return Array.isArray(n) ? (n[0]?.[0] ?? null) : n;
  }

  /** Formülün sonucunu dizi olarak döndürür (veri doğrulama listeleri). */
  evaluateArray(sheet: number, row: number, col: number, formula: string): Scalar[][] {
    const text = formula.startsWith("=") ? formula.slice(1) : formula;
    const ast = safeParse(text);
    if (!ast) return [[Errors.NAME]];
    const ctx = createContext({ host: this, sheet, row, col, locale: this.locale });
    const v = evaluateNode(ast, ctx);
    const n = normalizeResult(v, this, sheet);
    return Array.isArray(n) ? n : [[n]];
  }

  /* ----------------------------- Adlar ----------------------------- */

  defineName(name: string, formula: string, sheet?: number): void {
    const text = formula.startsWith("=") ? formula.slice(1) : formula;
    const u = name.toLocaleUpperCase("tr-TR");
    const key = sheet !== undefined ? `${sheet}!${u}` : u;
    this.names.set(key, { name, text, ast: safeParse(text), sheet });
    this.reregisterAll();
    this.markAllDirty();
  }

  removeName(name: string, sheet?: number): void {
    const u = name.toLocaleUpperCase("tr-TR");
    this.names.delete(sheet !== undefined ? `${sheet}!${u}` : u);
    this.reregisterAll();
    this.markAllDirty();
  }

  getNames(): { name: string; formula: string; sheet?: number }[] {
    return [...this.names.values()].map((n) => ({ name: n.name, formula: "=" + n.text, sheet: n.sheet }));
  }

  /* ------------------------- Yapısal değişiklikler ------------------------- */

  /** Satır/sütun ekler (count > 0) ya da siler (count < 0). */
  private structural(sheet: number, axis: "row" | "col", at: number, count: number): void {
    const s = this.sheetOrThrow(sheet);
    const name = s.name;
    // Taşmaları temizle
    for (const a of [...this.spillAnchors]) {
      const rec = this.sheets.get(gkSheet(a))?.cells.get(gkLocal(a));
      if (rec?.formula) this.clearSpill(a, rec.formula);
    }
    this.spillOwner.clear();
    this.spillAnchors.clear();
    // Hücreleri taşı
    const moved = new Map<number, CellRecord>();
    for (const [k, rec] of s.cells) {
      let r = keyRow(k);
      let c = keyCol(k);
      const idx = axis === "row" ? r : c;
      let ni: number | null;
      if (count > 0) ni = idx >= at ? idx + count : idx;
      else ni = idx < at ? idx : idx < at - count ? null : idx + count;
      if (ni === null) continue;
      if (axis === "row") r = ni;
      else c = ni;
      if (r >= MAX_ROWS || c >= MAX_COLS) continue;
      moved.set(cellKey(r, c), rec);
    }
    s.cells = moved;
    s.extentDirty = true;
    // Gizli satırlar
    if (axis === "row") {
      const shiftSet = (set: Set<number>) => {
        const next = new Set<number>();
        for (const r of set) {
          if (count > 0) next.add(r >= at ? r + count : r);
          else if (r < at) next.add(r);
          else if (r >= at - count) next.add(r + count);
        }
        return next;
      };
      s.hiddenRows = shiftSet(s.hiddenRows);
      s.filteredRows = shiftSet(s.filteredRows);
    }
    // Tüm formülleri güncelle
    const ch = { sheet: name, axis, at, count };
    for (const id of this.order) {
      const sh = this.sheets.get(id)!;
      for (const [, rec] of sh.cells) {
        if (!rec.formula) continue;
        const t = adjustFormula(rec.formula.text, sh.name, ch);
        if (t !== rec.formula.text) {
          rec.formula.text = t;
          rec.formula.ast = safeParse(t);
        }
        rec.formula.spill = null;
        rec.formula.spillTarget = null;
      }
    }
    for (const n of this.names.values()) {
      const scopeName = n.sheet !== undefined ? this.getSheetName(n.sheet) ?? name : name;
      const t = adjustFormula(n.text, scopeName, ch);
      if (t !== n.text) {
        n.text = t;
        n.ast = safeParse(t);
      }
    }
    this.reregisterAll();
    this.markAllDirty();
  }

  insertRows(sheet: number, at: number, count = 1): void {
    this.structural(sheet, "row", at, count);
  }

  deleteRows(sheet: number, at: number, count = 1): void {
    this.structural(sheet, "row", at, -count);
  }

  insertCols(sheet: number, at: number, count = 1): void {
    this.structural(sheet, "col", at, count);
  }

  deleteCols(sheet: number, at: number, count = 1): void {
    this.structural(sheet, "col", at, -count);
  }

  /**
   * Hücre aralığını yatay/dikey kaydırır (Hücre Ekle/Sil: "hücreleri sağa/aşağı kaydır").
   * Yalnız verilen satır/sütun bandını etkiler.
   */
  shiftCells(sheet: number, r1: number, c1: number, r2: number, c2: number, direction: "down" | "right" | "up" | "left"): void {
    const s = this.sheetOrThrow(sheet);
    const inputs: { row: number; col: number; input: Scalar | { formula: string } }[] = [];
    const h = r2 - r1 + 1;
    const w = c2 - c1 + 1;
    const inBand = (r: number, c: number) => (direction === "down" || direction === "up" ? c >= c1 && c <= c2 && r >= r1 : r >= r1 && r <= r2 && c >= c1);
    const affected: [number, number][] = [];
    for (const [k] of s.cells) {
      const r = keyRow(k);
      const c = keyCol(k);
      if (inBand(r, c)) affected.push([r, c]);
    }
    for (const [r, c] of affected) {
      const rec = s.cells.get(cellKey(r, c))!;
      let nr = r;
      let nc = c;
      if (direction === "down") nr = r + h;
      else if (direction === "right") nc = c + w;
      else if (direction === "up") {
        if (r <= r2) continue;
        nr = r - h;
      } else {
        if (c <= c2) continue;
        nc = c - w;
      }
      inputs.push({ row: nr, col: nc, input: rec.formula ? { formula: rec.formula.text } : rec.value });
    }
    const mv: MoveSpec = { srcSheet: s.name, r1: 0, c1: 0, r2: 0, c2: 0, dRow: 0, dCol: 0, dstSheet: s.name };
    if (direction === "down") Object.assign(mv, { r1, c1, r2: MAX_ROWS - 1, c2, dRow: h });
    else if (direction === "right") Object.assign(mv, { r1, c1, r2, c2: MAX_COLS - 1, dCol: w });
    else if (direction === "up") Object.assign(mv, { r1: r2 + 1, c1, r2: MAX_ROWS - 1, c2, dRow: -h });
    else Object.assign(mv, { r1, c1: c2 + 1, r2, c2: MAX_COLS - 1, dCol: -w });
    for (const [r, c] of affected) this.setValue(sheet, r, c, null);
    for (const it of inputs) {
      if (it.row < MAX_ROWS && it.col < MAX_COLS) this.setInput(sheet, it.row, it.col, it.input);
    }
    this.rewriteAll((text, sheetName) => moveRefsInFormula(text, sheetName, mv));
  }

  /** Kes-yapıştır: kaynak dikdörtgeni hedefe taşır, başvuruları günceller. */
  moveRange(srcSheet: number, r1: number, c1: number, r2: number, c2: number, dstSheet: number, dRow: number, dCol: number): void {
    const src = this.sheetOrThrow(srcSheet);
    const dst = this.sheetOrThrow(dstSheet);
    const items: { row: number; col: number; input: Scalar | { formula: string } }[] = [];
    for (let r = r1; r <= r2; r++)
      for (let c = c1; c <= c2; c++) {
        const rec = src.cells.get(cellKey(r, c));
        if (rec) items.push({ row: r + dRow, col: c + dCol, input: rec.formula ? { formula: rec.formula.text } : rec.value });
      }
    // Hedefi temizle, kaynağı temizle, yaz
    for (let r = r1; r <= r2; r++) for (let c = c1; c <= c2; c++) this.setValue(srcSheet, r, c, null);
    for (let r = r1 + dRow; r <= r2 + dRow; r++) for (let c = c1 + dCol; c <= c2 + dCol; c++) if (r >= 0 && c >= 0) this.setValue(dstSheet, r, c, null);
    for (const it of items) if (it.row >= 0 && it.col >= 0) this.setInput(dstSheet, it.row, it.col, it.input);
    // Taşınan formüller kendi içlerindeki göreli başvuruları korur (Excel davranışı)
    const mv: MoveSpec = { srcSheet: src.name, r1, c1, r2, c2, dRow, dCol, dstSheet: dst.name };
    this.rewriteAll((text, sheetName) => moveRefsInFormula(text, sheetName, mv));
  }

  private rewriteAll(fn: (text: string, sheetName: string) => string): void {
    for (const id of this.order) {
      const sh = this.sheets.get(id)!;
      for (const [k, rec] of sh.cells) {
        if (!rec.formula) continue;
        const t = fn(rec.formula.text, sh.name);
        if (t !== rec.formula.text) this.setFormulaInternal(id, keyRow(k), keyCol(k), t);
      }
    }
    for (const n of this.names.values()) {
      const t = fn(n.text, n.sheet !== undefined ? this.getSheetName(n.sheet) ?? "" : "");
      if (t !== n.text) {
        n.text = t;
        n.ast = safeParse(t);
      }
    }
  }

  /* ----------------------------- Anlık görüntü ----------------------------- */

  exportInputs(): EngineSnapshot {
    return {
      sheets: this.order.map((id) => {
        const s = this.sheets.get(id)!;
        const cells: CellSnapshot[] = [];
        for (const [k, rec] of s.cells) {
          if (rec.formula) cells.push({ row: keyRow(k), col: keyCol(k), formula: rec.formula.text });
          else cells.push({ row: keyRow(k), col: keyCol(k), value: rec.value });
        }
        return { id, name: s.name, cells };
      }),
      names: [...this.names.values()].map((n) => ({ name: n.name, text: n.text, sheet: n.sheet })),
    };
  }

  /** Bir sayfanın girdilerini ve sayfa listesini anlık görüntüden geri yükler. */
  importInputs(snap: EngineSnapshot): void {
    const keepIds = new Set(snap.sheets.map((s) => s.id));
    for (const id of [...this.order]) if (!keepIds.has(id)) this.sheets.delete(id);
    this.order = snap.sheets.map((s) => s.id);
    for (const ss of snap.sheets) {
      let s = this.sheets.get(ss.id);
      if (!s) {
        s = { id: ss.id, name: ss.name, cells: new Map(), maxRow: -1, maxCol: -1, extentDirty: false, hiddenRows: new Set(), filteredRows: new Set() };
        this.sheets.set(ss.id, s);
      }
      s.name = ss.name;
      s.cells = new Map();
      s.extentDirty = true;
      for (const c of ss.cells) {
        const k = cellKey(c.row, c.col);
        if (c.formula !== undefined) {
          const ast = safeParse(c.formula);
          s.cells.set(k, { value: 0, formula: { text: c.formula, ast, volatile: false, usesNames: false, cellDeps: [], rangeDeps: [], spill: null, spillTarget: null } });
        } else s.cells.set(k, { value: c.value ?? null });
      }
      this.nextSheetId = Math.max(this.nextSheetId, ss.id + 1);
    }
    this.names.clear();
    for (const n of snap.names) {
      const u = n.name.toLocaleUpperCase("tr-TR");
      this.names.set(n.sheet !== undefined ? `${n.sheet}!${u}` : u, { name: n.name, text: n.text, ast: safeParse(n.text), sheet: n.sheet });
    }
    this.spillOwner.clear();
    this.spillAnchors.clear();
    this.reregisterAll();
    this.markAllDirty();
  }
}

function safeParse(text: string): Node | null {
  try {
    return parseFormula(text);
  } catch {
    return null;
  }
}

function sameScalar(a: Scalar, b: Scalar): boolean {
  if (a === b) return true;
  if (isError(a) && isError(b)) return a.code === b.code;
  return false;
}

/** Kopyalanan formülü kaydırır (dışa açık yardımcı). */
export function copyFormula(formula: string, dRow: number, dCol: number): string {
  return shiftFormula(formula, dRow, dCol);
}

export type { RefValue };
