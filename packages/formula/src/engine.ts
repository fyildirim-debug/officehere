/**
 * Hesap motoru: sayfalar, seyrek hücre deposu, bağımlılık grafiği,
 * artımlı yeniden hesaplama, döngüsel başvuru algılama, dizi taşması
 * (#SPILL!), adlandırılmış aralıklar ve yapısal değişikliklerde
 * başvuru güncelleme.
 */
import "./functions";
import type { Node } from "./ast";
import { TR_LOCALE, type LocaleInfo } from "./coerce";
import type { EvalHost } from "./context";
import { jsDateToSerial } from "./dates";
import { createContext, evaluateNode } from "./evaluator";
import { shiftFormula } from "./transform";
import { Errors, FormulaError, type RefValue, type Scalar, type Value } from "./values";
import { DependencyIndex, analyzeDeps } from "./depindex";
import type { TableInfo } from "./structref";
import { findTable, permuteRowsImpl, renameTableColumnImpl, renameTableImpl, replaceHiddenRowsImpl, setRowsHiddenImpl, setTablesImpl, tableRefsToA1Impl } from "./engine-tables";
import { evaluateImpl, forEachValueImpl, getSpillImpl, noteChangeImpl, recomputeExtentImpl } from "./engine-read";
import { DeferSignal, normalizeResult, sameScalar } from "./engine-util";
import { applySpillImpl, clearSpillImpl, exportInputsImpl, importInputsImpl, moveRangeImpl, removeSheetImpl, renameSheetImpl, shiftCellsImpl, structuralChange } from "./engine-structure";
import {
  DEPTH_LIMIT,
  SHEET_STRIDE,
  gk,
  gkLocal,
  gkSheet,
  keyCol,
  keyRow,
  cellKey,
  safeParse,
  type CellRecord,
  type DefinedName,
  type EngineOptions,
  type EngineSnapshot,
  type FormulaInfo,
  type SheetData,
  type SpillInfo,
} from "./engine-types";

export { cellKey, keyCol, keyRow } from "./engine-types";
export type { CellRecord, CellSnapshot, EngineOptions, EngineSnapshot, SpillInfo } from "./engine-types";

export class FormulaEngine implements EvalHost {
  locale: LocaleInfo;
  autoCalc = true;
  /** @internal */
  sheets = new Map<number, SheetData>();
  /** @internal */
  order: number[] = [];
  /** @internal */
  nextSheetId = 1;
  /** @internal */
  names = new Map<string, DefinedName>();
  /** @internal */
  deps = new DependencyIndex();
  private volatiles = new Set<number>();
  private nameUsers = new Set<number>();
  /** @internal */
  spillOwner = new Map<number, number>();
  /** @internal */
  spillAnchors = new Set<number>();
  /** @internal */
  dirty = new Set<number>();
  /** @internal */
  changed = new Set<number>();
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
    removeSheetImpl(this, id);
  }

  renameSheet(id: number, name: string): void {
    renameSheetImpl(this, id, name);
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

  /* Gizli satırlar (ALTTOPLAM bunlara bağlıdır; işaretleme sonraki hesaplamaya ertelenir) */
  setRowHidden(sheet: number, row: number, hidden: boolean, filtered = false): void {
    setRowsHiddenImpl(this, sheet, [row], hidden, filtered);
  }

  /** Toplu gizleme/gösterme (filtre, anahat): tek seferde. */
  setRowsHidden(sheet: number, rows: Iterable<number>, hidden: boolean, filtered = false): void {
    setRowsHiddenImpl(this, sheet, rows, hidden, filtered);
  }

  clearFilteredRows(sheet: number): void {
    replaceHiddenRowsImpl(this, sheet, null, []);
  }

  /** Gizli satır kümesini tümüyle değiştirir (geri al / eşitleme). */
  replaceHiddenRows(sheet: number, hidden: Iterable<number>, filtered: Iterable<number>): void {
    replaceHiddenRowsImpl(this, sheet, hidden, filtered);
  }

  /** @internal */
  subtotalsStale = false;
  /** ALTTOPLAM/TOPLAMA içeren formül hücreleri (gizli satır değişince yeniden hesaplanır). */
  private subtotalCells = new Set<number>();

  private markSubtotalsDirty(): void {
    this.subtotalsStale = false;
    for (const g of this.subtotalCells) this.dirty.add(g);
  }

  /* ------------------------------ Tablolar ------------------------------ */

  /** @internal */
  tables: TableInfo[] = [];
  /** @internal */
  tablesSig = "";

  /** Yapısal başvurular için tablo listesini ayarlar; değiştiyse bağımlılıklar yenilenir. */
  setTables(list: TableInfo[]): void {
    if (!setTablesImpl(this, list)) return;
    this.reregisterAll();
    this.markAllDirty();
  }

  getTables(): TableInfo[] {
    return this.tables;
  }

  getTable(name: string | undefined, sheet: number, row: number, col: number): TableInfo | undefined {
    return findTable(this.tables, name, sheet, row, col);
  }

  /** Tablo yeniden adlandırılınca formüllerdeki yapısal başvuruları günceller. */
  renameTable(oldName: string, newName: string): void {
    renameTableImpl(this, oldName, newName);
    this.reregisterAll();
    this.markAllDirty();
  }

  /** Tablo aralığa dönüştürülürken ona ait yapısal başvuruları A1 başvurusuna çevirir. */
  tableRefsToA1(name: string): void {
    tableRefsToA1Impl(this, name);
    this.reregisterAll();
    this.markAllDirty();
  }

  /** Tablo sütunu yeniden adlandırılınca formülleri günceller. */
  renameTableColumn(table: string, oldCol: string, newCol: string): void {
    renameTableColumnImpl(this, table, oldCol, newCol);
  }

  /** Satırları yeniden sıralar (yeni i. satır = eski order[i]); formüller göreli kaydırılır. */
  permuteRows(sheet: number, r1: number, c1: number, r2: number, c2: number, order: number[]): boolean {
    return permuteRowsImpl(this, sheet, r1, c1, r2, c2, order);
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
    recomputeExtentImpl(s);
  }

  /** @internal */
  growExtent(s: SheetData, row: number, col: number): void {
    if (row > s.maxRow) s.maxRow = row;
    if (col > s.maxCol) s.maxCol = col;
  }

  /* EvalHost sayfa bilgileri */
  sheetIdByName = (name: string): number | undefined => this.getSheetId(name);
  sheetName = (id: number): string | undefined => this.getSheetName(id);
  sheetIndex = (id: number): number => this.order.indexOf(id);
  sheetCount = (): number => this.order.length;

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

  now = (): number => this.nowFn();
  random = (): number => this.randomFn();

  /* ------------------------------ Hücreler ------------------------------ */

  /** @internal */
  sheetOrThrow(id: number): SheetData {
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

  /** @internal */
  setFormulaInternal(sheet: number, row: number, col: number, text: string): void {
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

  /** Hücrenin kaydı (girdi ve önbellekteki değer); toplu okumalar için tek aramalı erişim. */
  getRecord(sheet: number, row: number, col: number): CellRecord | undefined {
    return this.sheets.get(sheet)?.cells.get(cellKey(row, col));
  }

  /** Hücrede kullanıcı girdisi (sabit ya da formül) var mı? */
  hasInput(sheet: number, row: number, col: number): boolean {
    return !!this.sheets.get(sheet)?.cells.has(cellKey(row, col));
  }

  /** Hücre taşan bir dizinin parçasıysa bilgisi. */
  getSpill(sheet: number, row: number, col: number): SpillInfo | undefined {
    return getSpillImpl(this, sheet, row, col);
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
    forEachValueImpl(this, sheet, cb);
  }

  /* ------------------------- Bağımlılık grafiği ------------------------- */

  /** @internal Hazır formül bilgisini bağımlılık dizinine kaydeder. */
  registerInfo(g: number, info: FormulaInfo, sheet: number): void {
    this.register(g, info, sheet);
  }

  private register(g: number, info: FormulaInfo, sheet: number): void {
    if (!info.ast) return;
    const { cellDeps, rangeDeps, volatile, usesNames } = analyzeDeps(info.ast, g, sheet, {
      sheetId: (name) => this.getSheetId(name),
      definedName: (name, sh) => this.getDefinedName(name, sh),
      table: (name, sh, row, col) => this.getTable(name, sh, row, col),
    });
    info.cellDeps = cellDeps;
    info.rangeDeps = rangeDeps;
    info.volatile = volatile;
    info.usesNames = usesNames;
    this.deps.add(g, cellDeps, rangeDeps);
    if (/SUBTOTAL|AGGREGATE/i.test(info.text)) this.subtotalCells.add(g);
    if (volatile) this.volatiles.add(g);
    if (usesNames) this.nameUsers.add(g);
  }

  /** @internal */
  unregister(g: number): void {
    const s = this.sheets.get(gkSheet(g));
    const rec = s?.cells.get(gkLocal(g));
    const info = rec?.formula;
    if (!info) return;
    this.deps.remove(g, info.cellDeps, info.rangeDeps);
    info.cellDeps = [];
    info.rangeDeps = [];
    this.volatiles.delete(g);
    this.nameUsers.delete(g);
    this.subtotalCells.delete(g);
    if (info.spill || info.spillTarget) this.clearSpill(g, info);
    this.spillAnchors.delete(g);
    this.dirty.delete(g);
  }

  /** @internal */
  reregisterAll(): void {
    this.deps.clear();
    this.volatiles.clear();
    this.nameUsers.clear();
    this.subtotalCells.clear();
    for (const id of this.order) {
      const s = this.sheets.get(id)!;
      for (const [k, rec] of s.cells) if (rec.formula) this.register(id * SHEET_STRIDE + k, rec.formula, id);
    }
  }

  /** @internal */
  noteChange(g: number): void {
    noteChangeImpl(this, g);
  }

  markAllDirty(): void {
    for (const id of this.order) {
      const s = this.sheets.get(id)!;
      for (const [k, rec] of s.cells) if (rec.formula) this.dirty.add(id * SHEET_STRIDE + k);
    }
  }

  /** Yeniden hesaplanmayı bekleyen formül var mı? */
  get pending(): boolean {
    return this.dirty.size > 0 || this.changed.size > 0 || this.subtotalsStale;
  }

  /* --------------------------- Yeniden hesaplama --------------------------- */

  private inRecalc = false;

  /** Değişikliklerden etkilenen formülleri yeniden hesaplar. full: tümünü. */
  recalc(full = false): void {
    if (this.inRecalc) return;
    this.inRecalc = true;
    try {
      if (full) this.markAllDirty();
      if (this.subtotalsStale) this.markSubtotalsDirty();
      for (const v of this.volatiles) this.dirty.add(v);
      this.lookupCache.clear();
      for (let pass = 0; pass < 8; pass++) {
        // Değişen konumlardan bağımlıları yay
        const queue: number[] = [...this.changed, ...this.dirty];
        this.changed.clear();
        const seen = new Set<number>(queue);
        while (queue.length) {
          const g = queue.pop()!;
          this.deps.dependentsOf(g, (f) => {
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
    applySpillImpl(this, g, s, info, row, col, data);
  }

  /** @internal */
  clearSpill(g: number, info: FormulaInfo): void {
    clearSpillImpl(this, g, info);
  }

  /** Bir formülü verilen hücre bağlamında değerlendirir (koşullu biçim, veri doğrulama). */
  evaluate(sheet: number, row: number, col: number, formula: string): Scalar {
    if (this.dirty.size && this.autoCalc && !this.inRecalc) this.recalc();
    const n = evaluateImpl(this, sheet, row, col, formula);
    return Array.isArray(n) ? (n[0]?.[0] ?? null) : n;
  }

  /** Formülün sonucunu dizi olarak döndürür (veri doğrulama listeleri). */
  evaluateArray(sheet: number, row: number, col: number, formula: string): Scalar[][] {
    const n = evaluateImpl(this, sheet, row, col, formula);
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

  insertRows(sheet: number, at: number, count = 1): void {
    structuralChange(this, sheet, "row", at, count);
  }

  deleteRows(sheet: number, at: number, count = 1): void {
    structuralChange(this, sheet, "row", at, -count);
  }

  insertCols(sheet: number, at: number, count = 1): void {
    structuralChange(this, sheet, "col", at, count);
  }

  deleteCols(sheet: number, at: number, count = 1): void {
    structuralChange(this, sheet, "col", at, -count);
  }

  /** Hücre aralığını kaydırır (Hücre Ekle/Sil: "hücreleri sağa/aşağı kaydır"). */
  shiftCells(sheet: number, r1: number, c1: number, r2: number, c2: number, direction: "down" | "right" | "up" | "left"): void {
    shiftCellsImpl(this, sheet, r1, c1, r2, c2, direction);
  }

  /** Kes-yapıştır: kaynak dikdörtgeni hedefe taşır, başvuruları günceller. */
  moveRange(srcSheet: number, r1: number, c1: number, r2: number, c2: number, dstSheet: number, dRow: number, dCol: number): void {
    moveRangeImpl(this, srcSheet, r1, c1, r2, c2, dstSheet, dRow, dCol);
  }

  /* ----------------------------- Anlık görüntü ----------------------------- */

  exportInputs(withAst = false): EngineSnapshot {
    return exportInputsImpl(this, withAst);
  }

  /** Sayfa listesini ve girdileri anlık görüntüden geri yükler. */
  importInputs(snap: EngineSnapshot): void {
    importInputsImpl(this, snap);
  }
}

/** Kopyalanan formülü kaydırır (dışa açık yardımcı). */
export function copyFormula(formula: string, dRow: number, dCol: number): string {
  return shiftFormula(formula, dRow, dCol);
}

export type { RefValue };
