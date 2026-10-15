/**
 * Geri alınabilir işlemlerin çekirdeği: geri al yığını ve kayıt türleri
 * (hücre, sayfa, ad, tüm kitap anlık görüntüsü).
 */
import { UndoManager } from "./history";
import { isFullCols, isFullRows } from "./rect";
import type { Rect } from "./types";
import type { CellState, SheetData, SheetModel, WorkbookModel } from "./workbook";

export class OpsCore {
  readonly history = new UndoManager();
  /** Değişiklik sonrası çağrılır (yeniden çizim, kirli işareti). */
  onChange: (label: string) => void = () => {};
  /** Salt okunur açılan dosya (görüntüleyici kipi): hiçbir değişiklik uygulanmaz ve kaydedilmez. */
  readOnly = false;
  /**
   * Hücre içeriği/stili değişince otomatik satır yüksekliklerini ayarlar
   * (arayüz katmanı ölçüm yapar). [satır, eski, yeni] listesi döndürür.
   */
  afterCells: ((sheet: SheetModel, rects: Rect[]) => [number, number | undefined, number | undefined][]) | null = null;

  constructor(public wb: WorkbookModel) {}

  protected recalc(): void {
    this.wb.syncTables();
    if (this.wb.engine.autoCalc) this.wb.engine.recalc();
  }

  protected done(label: string): void {
    this.recalc();
    this.onChange(label);
  }

  undo(): string | undefined {
    const e = this.history.undo();
    if (e) this.done(e.label);
    return e?.label;
  }

  redo(): string | undefined {
    const e = this.history.redo();
    if (e) this.done(e.label);
    return e?.label;
  }

  /* ----------------------------- Genel kayıt ----------------------------- */

  /** Hücre düzeyinde değişiklik: aralıklardaki hücrelerin öncesi/sonrası saklanır. */
  cellsOp(sheet: SheetModel, rects: Rect[], label: string, mutate: () => void, extraKeys?: () => Iterable<number>): void {
    if (this.readOnly) return;
    const wb = this.wb;
    const collect = () => {
      const keys = new Set<number>();
      for (const r of rects) for (const k of wb.cellKeysIn(sheet, r)) keys.add(k);
      if (extraKeys) for (const k of extraKeys()) keys.add(k);
      return keys;
    };
    const lines = rects.some((r) => isFullCols(r) || isFullRows(r));
    const linesBefore = lines ? { c: new Map(sheet.colStyles), r: new Map(sheet.rowStyles) } : null;
    const before = new Map<number, CellState>();
    for (const k of collect()) before.set(k, wb.cellState(sheet, k));
    // Mini grafikler hücreyle birlikte kopyalanabilir (doldurma, yapıştırma)
    const sparkBefore = sheet.sparklines.length ? structuredClone(sheet.sparklines) : null;
    mutate();
    const sparkAfter = sparkBefore || sheet.sparklines.length ? structuredClone(sheet.sparklines) : null;
    const applySpark = (v: typeof sparkBefore) => {
      if (sparkBefore || sparkAfter) sheet.sparklines = structuredClone(v ?? []);
    };
    const rowChanges = this.afterCells ? this.afterCells(sheet, rects) : [];
    const applyRows = (which: 1 | 2) => {
      for (const ch of rowChanges) {
        const v = ch[which];
        if (v === undefined) sheet.rows.reset(ch[0]);
        else sheet.rows.set(ch[0], v, false);
      }
    };
    const after = new Map<number, CellState>();
    for (const k of collect()) after.set(k, wb.cellState(sheet, k));
    for (const k of before.keys()) if (!after.has(k)) after.set(k, wb.cellState(sheet, k));
    const linesAfter = lines ? { c: new Map(sheet.colStyles), r: new Map(sheet.rowStyles) } : null;
    const apply = (states: Map<number, CellState>, other: Map<number, CellState>, ln: typeof linesBefore) => {
      for (const k of other.keys()) if (!states.has(k)) wb.applyCellState(sheet, k, undefined);
      for (const [k, st] of states) wb.applyCellState(sheet, k, st);
      if (ln) {
        sheet.colStyles = new Map(ln.c);
        sheet.rowStyles = new Map(ln.r);
      }
    };
    this.history.push({
      label,
      cost: before.size + after.size,
      undo: () => (apply(before, after, linesBefore), applyRows(1), applySpark(sparkBefore)),
      redo: () => (apply(after, before, linesAfter), applyRows(2), applySpark(sparkAfter)),
    });
    this.done(label);
  }

  /** Sayfa verisinin belirli parçalarında değişiklik (birleştirme, nesneler…). */
  sheetOp(sheet: SheetModel, label: string, mutate: () => void, withCells?: Rect[]): void {
    if (this.readOnly) return;
    const wb = this.wb;
    const before = sheet.cloneData();
    const cellsBefore = withCells ? this.captureCells(sheet, withCells) : null;
    mutate();
    const after = sheet.cloneData();
    const cellsAfter = withCells ? this.captureCells(sheet, withCells) : null;
    const apply = (d: SheetData, cells: Map<number, CellState> | null, other: Map<number, CellState> | null) => {
      sheet.restoreData(d);
      if (cells && other) {
        for (const k of other.keys()) if (!cells.has(k)) wb.applyCellState(sheet, k, undefined);
        for (const [k, st] of cells) wb.applyCellState(sheet, k, st);
      }
      this.syncHiddenRows(sheet);
    };
    this.history.push({ label, cost: sheet.styles.size / 4 + 10, undo: () => apply(before, cellsBefore, cellsAfter), redo: () => apply(after, cellsAfter, cellsBefore) });
    this.done(label);
  }

  /**
   * Satır permütasyonu (sıralama): yeni i. satır = eski order[i]. Hücreler,
   * stiller, notlar ve köprüler taşınır; geri alma ters permütasyondur.
   * Formül kaydırması #BAŞV! üretecekse false döner (çağıran genel yola düşer).
   */
  permuteRowsOp(sheet: SheetModel, r: Rect, order: number[], label: string): boolean {
    if (this.readOnly) return true;
    const wb = this.wb;
    const inv: number[] = new Array(order.length);
    order.forEach((src, i) => (inv[src] = i));
    const moveMaps = (ord: number[]) => {
      const pos: number[] = new Array(ord.length);
      ord.forEach((src, i) => (pos[src] = i));
      for (const m of [sheet.styles, sheet.notes, sheet.links] as Map<number, unknown>[]) {
        if (!m.size) continue;
        const moved: [number, unknown][] = [];
        for (const [k, v] of m) {
          const row = Math.floor(k / 16384);
          const col = k % 16384;
          if (row < r.r1 || row > r.r2 || col < r.c1 || col > r.c2) continue;
          moved.push([(r.r1 + pos[row - r.r1]) * 16384 + col, v]);
          m.delete(k);
        }
        for (const [k, v] of moved) m.set(k, v);
      }
    };
    if (!wb.engine.permuteRows(sheet.id, r.r1, r.c1, r.r2, r.c2, order)) return false;
    moveMaps(order);
    this.history.push({
      label,
      cost: ((r.r2 - r.r1 + 1) * (r.c2 - r.c1 + 1)) / 20 + 10,
      undo: () => {
        wb.engine.permuteRows(sheet.id, r.r1, r.c1, r.r2, r.c2, inv);
        moveMaps(inv);
      },
      redo: () => {
        wb.engine.permuteRows(sheet.id, r.r1, r.c1, r.r2, r.c2, order);
        moveMaps(order);
      },
    });
    this.done(label);
    return true;
  }

  /** Tanımlı ad değişiklikleri (Ad Yöneticisi). */
  namesOp(label: string, mutate: () => void): void {
    if (this.readOnly) return;
    const eng = this.wb.engine;
    const before = eng.getNames();
    mutate();
    const after = eng.getNames();
    const apply = (list: typeof before) => {
      for (const n of eng.getNames()) eng.removeName(n.name, n.sheet);
      for (const n of list) eng.defineName(n.name, n.formula, n.sheet);
    };
    this.history.push({ label, cost: 10, undo: () => apply(before), redo: () => apply(after) });
    this.done(label);
  }

  /** Tüm kitabın anlık görüntüsüyle (yapısal işlemler). */
  snapshotOp(label: string, mutate: () => void): void {
    if (this.readOnly) return;
    const wb = this.wb;
    const before = wb.snapshot();
    mutate();
    // "Sonra" görüntüsü ilk geri almada alınır (doğrusal geçmişte o anki durum budur)
    let after: ReturnType<WorkbookModel["snapshot"]> | null = null;
    const cost = before.engine.sheets.reduce((n, s) => n + s.cells.length, 0) * 2 + 100;
    this.history.push({
      label,
      cost,
      undo: () => {
        after ??= wb.snapshot();
        wb.restore(before);
        wb.sheets.forEach((s) => this.syncHiddenRows(s));
      },
      redo: () => {
        if (after) wb.restore(after);
        wb.sheets.forEach((s) => this.syncHiddenRows(s));
      },
    });
    this.done(label);
  }

  protected captureCells(sheet: SheetModel, rects: Rect[]): Map<number, CellState> {
    const m = new Map<number, CellState>();
    for (const r of rects) for (const k of this.wb.cellKeysIn(sheet, r)) m.set(k, this.wb.cellState(sheet, k));
    return m;
  }

  /** Motorun gizli satır bilgisini sayfa modeliyle eşitler (ALTTOPLAM için). */
  syncHiddenRows(sheet: SheetModel): void {
    const manual: number[] = [];
    for (const r of sheet.rows.hiddenItems()) if (!sheet.filterHidden.has(r)) manual.push(r);
    this.wb.engine.replaceHiddenRows(sheet.id, manual, sheet.filterHidden);
  }

}
