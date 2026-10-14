/**
 * Çalışma kitabı modeli: hesap motoru (@officehere/formula) hücre
 * içeriklerini tutar; bu katman stilleri, boyutları, birleştirmeleri,
 * notları ve sayfa düzeyindeki nesneleri yönetir.
 */
import { FormulaEngine, addr as addrOf, cellKey, keyCol, keyRow, toLocalFormula, numberToText, isDateFormat, isTimeOnlyFormat, formatValue, isError, localErrorName, parseCellInput, unqualifyStructRefs, type Scalar, type TableInfo } from "@officehere/formula";
import { AxisSizes } from "./axis";
import { MAX_COLS, MAX_ROWS, contains, intersects, shiftIndex, shiftRect } from "./rect";
import { StyleTable, type CellStyle } from "./styles";
import { renameRangeText, shiftRangeText } from "./rangeRefs";
import { defaultTheme, type WorkbookTheme } from "./theme";
import type { AutoFilter, ChartDef, CondRule, DataTableDef, ImageDef, Note, PageSetup, PivotDef, Rect, Scenario, Selection, ShapeDef, SlicerDef, SparklineGroup, TableDef, Validation, ViewMode, WorkbookProtection } from "./types";

export type CellInput = Scalar | { formula: string };

export function defaultPage(): PageSetup {
  return {
    orientation: "portrait",
    paper: "A4",
    margins: { top: 0.75, bottom: 0.75, left: 0.7, right: 0.7, header: 0.3, footer: 0.3 },
    printGridlines: false,
    printHeadings: false,
  };
}

export class SheetModel {
  tabColor?: string;
  hidden = false;
  cols = new AxisSizes(64, MAX_COLS);
  rows = new AxisSizes(20, MAX_ROWS);
  styles = new Map<number, number>();
  colStyles = new Map<number, number>();
  rowStyles = new Map<number, number>();
  merges: Rect[] = [];
  notes = new Map<number, Note>();
  links = new Map<number, string>();
  validations: Validation[] = [];
  condFormats: CondRule[] = [];
  autoFilter: AutoFilter | null = null;
  filterHidden = new Set<number>();
  tables: TableDef[] = [];
  charts: ChartDef[] = [];
  images: ImageDef[] = [];
  shapes: ShapeDef[] = [];
  pivots: PivotDef[] = [];
  sparklines: SparklineGroup[] = [];
  /** Senaryo Yöneticisi senaryoları, veri tabloları ve dilimleyiciler. */
  scenarios: Scenario[] = [];
  dataTables: DataTableDef[] = [];
  slicers: SlicerDef[] = [];
  /** Yerinde uygulanmış Gelişmiş Filtre (Temizle ile kalkar). */
  advFilter: { range: Rect; criteria: Rect | null; unique: boolean } | null = null;
  /** Hata Denetimi'nde "Hatayı Yoksay" denen hücreler. */
  ignoredErrors = new Set<number>();
  frozen = { rows: 0, cols: 0 };
  showGridlines = true;
  showHeadings = true;
  showFormulas = false;
  zoom = 100;
  viewMode: ViewMode = "normal";
  page: PageSetup = defaultPage();
  selection: Selection = { ranges: [{ r1: 0, c1: 0, r2: 0, c2: 0 }], active: { row: 0, col: 0 }, anchor: { row: 0, col: 0 } };
  /** Görünümün sol üst hücresi (sayfa değiştirince geri yüklemek için). */
  scroll = { row: 0, col: 0 };
  protected = false;
  /** Koruma parolası (yalnız oturumda tutulur, dosyaya yazılmaz). */
  protectPassword?: string;

  constructor(
    public id: number,
    public name: string,
  ) {}

  /** Birleştirilmiş alan (verilen hücreyi içeren). */
  mergeAt(row: number, col: number): Rect | undefined {
    for (const m of this.merges) if (contains(m, row, col)) return m;
    return undefined;
  }

  mergesIn(r: Rect): Rect[] {
    return this.merges.filter((m) => intersects(m, r));
  }

  noteAt(row: number, col: number): Note | undefined {
    return this.notes.get(cellKey(row, col));
  }

  linkAt(row: number, col: number): string | undefined {
    return this.links.get(cellKey(row, col));
  }

  tableAt(row: number, col: number): TableDef | undefined {
    return this.tables.find((t) => contains(t.ref, row, col));
  }

  /** Derin kopya (yapısal işlemlerin geri alınması için). */
  cloneData(): SheetData {
    return {
      tabColor: this.tabColor,
      hidden: this.hidden,
      cols: this.cols.clone(),
      rows: this.rows.clone(),
      styles: new Map(this.styles),
      colStyles: new Map(this.colStyles),
      rowStyles: new Map(this.rowStyles),
      merges: this.merges.map((m) => ({ ...m })),
      notes: new Map(this.notes),
      links: new Map(this.links),
      validations: structuredClone(this.validations),
      condFormats: structuredClone(this.condFormats),
      autoFilter: structuredClone(this.autoFilter),
      filterHidden: new Set(this.filterHidden),
      tables: structuredClone(this.tables),
      charts: structuredClone(this.charts),
      images: this.images.map((i) => ({ ...i, anchor: { ...i.anchor } })),
      shapes: structuredClone(this.shapes),
      pivots: structuredClone(this.pivots),
      sparklines: structuredClone(this.sparklines),
      scenarios: structuredClone(this.scenarios),
      dataTables: structuredClone(this.dataTables),
      slicers: structuredClone(this.slicers),
      advFilter: structuredClone(this.advFilter),
      ignoredErrors: new Set(this.ignoredErrors),
      frozen: { ...this.frozen },
      page: structuredClone(this.page),
    };
  }

  restoreData(d: SheetData): void {
    this.tabColor = d.tabColor;
    this.hidden = d.hidden;
    this.cols.assign(d.cols);
    this.rows.assign(d.rows);
    this.styles = new Map(d.styles);
    this.colStyles = new Map(d.colStyles);
    this.rowStyles = new Map(d.rowStyles);
    this.merges = d.merges.map((m) => ({ ...m }));
    this.notes = new Map(d.notes);
    this.links = new Map(d.links);
    this.validations = structuredClone(d.validations);
    this.condFormats = structuredClone(d.condFormats);
    this.autoFilter = structuredClone(d.autoFilter);
    this.filterHidden = new Set(d.filterHidden);
    this.tables = structuredClone(d.tables);
    this.charts = structuredClone(d.charts);
    this.images = d.images.map((i) => ({ ...i, anchor: { ...i.anchor } }));
    this.shapes = structuredClone(d.shapes);
    this.pivots = structuredClone(d.pivots);
    this.sparklines = structuredClone(d.sparklines);
    this.scenarios = structuredClone(d.scenarios ?? []);
    this.dataTables = structuredClone(d.dataTables ?? []);
    this.slicers = structuredClone(d.slicers ?? []);
    this.advFilter = structuredClone(d.advFilter ?? null);
    this.ignoredErrors = new Set(d.ignoredErrors);
    this.frozen = { ...d.frozen };
    this.page = structuredClone(d.page);
  }

  /** Satır/sütun ekleme-silmede sayfa nesnelerini kaydırır. */
  shiftStructure(axis: "row" | "col", at: number, count: number): void {
    const moveKeys = <T>(m: Map<number, T>): Map<number, T> => {
      const out = new Map<number, T>();
      for (const [k, v] of m) {
        let r = keyRow(k);
        let c = keyCol(k);
        const ni = shiftIndex(axis === "row" ? r : c, at, count);
        if (ni === null) continue;
        if (axis === "row") r = ni;
        else c = ni;
        if (r >= MAX_ROWS || c >= MAX_COLS) continue;
        out.set(cellKey(r, c), v);
      }
      return out;
    };
    this.styles = moveKeys(this.styles);
    this.notes = moveKeys(this.notes);
    this.links = moveKeys(this.links);
    this.ignoredErrors = new Set(moveKeys(new Map([...this.ignoredErrors].map((k) => [k, true]))).keys());
    const moveLine = (m: Map<number, number>) => {
      const out = new Map<number, number>();
      for (const [k, v] of m) {
        const ni = shiftIndex(k, at, count);
        if (ni !== null) out.set(ni, v);
      }
      return out;
    };
    if (axis === "row") {
      this.rowStyles = moveLine(this.rowStyles);
      this.rows.shift(at, count);
      const fh = new Set<number>();
      for (const r of this.filterHidden) {
        const ni = shiftIndex(r, at, count);
        if (ni !== null) fh.add(ni);
      }
      this.filterHidden = fh;
    } else {
      this.colStyles = moveLine(this.colStyles);
      this.cols.shift(at, count);
    }
    const sr = (r: Rect) => shiftRect(r, axis, at, count);
    this.merges = this.merges.map(sr).filter((m): m is Rect => !!m && (m.r1 !== m.r2 || m.c1 !== m.c2));
    for (const v of this.validations) v.ranges = v.ranges.map(sr).filter((x): x is Rect => !!x);
    this.validations = this.validations.filter((v) => v.ranges.length);
    for (const c of this.condFormats) c.ranges = c.ranges.map(sr).filter((x): x is Rect => !!x);
    this.condFormats = this.condFormats.filter((c) => c.ranges.length);
    if (this.autoFilter) {
      const r = sr(this.autoFilter.range);
      this.autoFilter = r ? { ...this.autoFilter, range: r } : null;
    }
    this.tables = this.tables
      .map((t) => {
        const r = sr(t.ref);
        return r ? { ...t, ref: r } : null;
      })
      .filter((t): t is TableDef => !!t);
    const shiftAnchor = (a: { row: number; col: number }) => {
      if (axis === "row") a.row = Math.max(0, shiftIndex(a.row, at, count) ?? at);
      else a.col = Math.max(0, shiftIndex(a.col, at, count) ?? at);
    };
    for (const o of [...this.charts, ...this.images, ...this.shapes]) shiftAnchor(o.anchor);
    // PivotTable konumu ve çıktısı; mini grafik hücreleri
    for (const p of this.pivots) {
      shiftAnchor(p.anchor);
      if (p.output) p.output = sr(p.output) ?? undefined;
      if (p.buttons)
        p.buttons = p.buttons
          .map((b) => {
            const ni = shiftIndex(axis === "row" ? b.row : b.col, at, count);
            return ni === null ? null : { ...b, row: axis === "row" ? ni : b.row, col: axis === "col" ? ni : b.col };
          })
          .filter((b): b is { row: number; col: number; field: number } => !!b);
    }
    for (const g of this.sparklines) {
      g.items = g.items
        .map((it) => {
          const ni = shiftIndex(axis === "row" ? it.row : it.col, at, count);
          return ni === null ? null : { ...it, row: axis === "row" ? ni : it.row, col: axis === "col" ? ni : it.col };
        })
        .filter((x): x is { row: number; col: number; range: string } => !!x);
    }
    this.sparklines = this.sparklines.filter((g) => g.items.length);
    if (this.page.printArea) this.page.printArea = sr(this.page.printArea) ?? undefined;
    // El ile sayfa sonları satır/sütunlarıyla birlikte kayar
    const key = axis === "row" ? "rowBreaks" : "colBreaks";
    const brk = this.page[key];
    if (brk?.length) this.page[key] = [...new Set(brk.map((i) => shiftIndex(i, at, count)).filter((i): i is number => i !== null && i > 0))];
  }
}

export interface SheetData {
  tabColor?: string;
  hidden: boolean;
  cols: AxisSizes;
  rows: AxisSizes;
  styles: Map<number, number>;
  colStyles: Map<number, number>;
  rowStyles: Map<number, number>;
  merges: Rect[];
  notes: Map<number, Note>;
  links: Map<number, string>;
  validations: Validation[];
  condFormats: CondRule[];
  autoFilter: AutoFilter | null;
  filterHidden: Set<number>;
  tables: TableDef[];
  charts: ChartDef[];
  images: ImageDef[];
  shapes: ShapeDef[];
  pivots: PivotDef[];
  sparklines: SparklineGroup[];
  scenarios?: Scenario[];
  dataTables?: DataTableDef[];
  slicers?: SlicerDef[];
  advFilter?: { range: Rect; criteria: Rect | null; unique: boolean } | null;
  ignoredErrors: Set<number>;
  frozen: { rows: number; cols: number };
  page: PageSetup;
}

export interface CellState {
  input: CellInput;
  style?: number;
  note?: Note;
  link?: string;
}

interface UsedRangeCache {
  inputVersion: number;
  calcVersion: number;
  /** Zayıf başvuru: yenisiyle değiştirilen eski stil haritası önbellek yüzünden bellekte kalmasın. */
  styles: WeakRef<Map<number, number>>;
  styleCount: number;
  styleWrites: number;
  rect: Rect | null;
}

export interface DocProps {
  title?: string;
  subject?: string;
  author?: string;
  company?: string;
  created?: Date;
  modified?: Date;
}

export class WorkbookModel {
  engine: FormulaEngine;
  styles = new StyleTable();
  sheets: SheetModel[] = [];
  activeIndex = 0;
  props: DocProps = {};
  /** Çalışma kitabı teması (Sayfa Düzeni > Temalar). */
  theme: WorkbookTheme = defaultTheme();
  /** İçe aktarmada okunamayıp atlanan öğeler (bilgi çubuğunda gösterilir). */
  loadWarnings: string[] = [];
  /** Çalışma kitabı yapı koruması. */
  protection: WorkbookProtection | null = null;
  private nextObjectId = 1;
  /** usedRange önbelleği (sayfa başına sonuç ve onu üreten damga). */
  private usedCache = new WeakMap<SheetModel, UsedRangeCache>();
  /** setStyleRaw yazma sayacı: boyutu değiştirmeyen stil değişimlerinde önbellek damgası. */
  private styleWrites = 0;

  constructor(empty = false) {
    this.engine = new FormulaEngine();
    if (!empty) this.addSheet("Sayfa1");
  }

  get active(): SheetModel {
    return this.sheets[this.activeIndex];
  }

  newId(prefix: string): string {
    return `${prefix}${this.nextObjectId++}`;
  }

  /** Dışarıdan yüklenen nesnelerin kimlikleriyle çakışmamak için sayacı ilerletir. */
  bumpIds(): void {
    let max = 0;
    for (const s of this.sheets)
      for (const o of [...s.charts, ...s.images, ...s.shapes, ...s.tables, ...s.pivots, ...s.sparklines, ...s.validations, ...s.condFormats, ...s.dataTables, ...s.slicers]) {
        const n = Number(/(\d+)$/.exec(o.id)?.[1] ?? 0);
        if (n > max) max = n;
      }
    this.nextObjectId = Math.max(this.nextObjectId, max + 1);
  }

  /* ------------------------------ Sayfalar ------------------------------ */

  nextSheetName(base = "Sayfa"): string {
    let n = 1;
    const names = new Set(this.sheets.map((s) => s.name.toLocaleUpperCase("tr-TR")));
    while (names.has(`${base}${n}`.toLocaleUpperCase("tr-TR"))) n++;
    return `${base}${n}`;
  }

  addSheet(name = this.nextSheetName(), index = this.sheets.length): SheetModel {
    const id = this.engine.addSheet(name, index);
    const s = new SheetModel(id, name);
    this.sheets.splice(index, 0, s);
    return s;
  }

  removeSheet(index: number): void {
    const s = this.sheets[index];
    if (!s) return;
    this.engine.removeSheet(s.id);
    this.sheets.splice(index, 1);
    if (this.activeIndex >= this.sheets.length) this.activeIndex = this.sheets.length - 1;
  }

  renameSheet(index: number, name: string): void {
    const s = this.sheets[index];
    this.engine.renameSheet(s.id, name);
    // Grafik ve mini grafik veri aralıkları yeni adı göstersin
    for (const o of this.sheets) {
      for (const c of o.charts) {
        c.range = renameRangeText(c.range, s.name, name);
        mapSeriesRanges(c, (t) => renameRangeText(t, s.name, name));
      }
      for (const g of o.sparklines) for (const it of g.items) it.range = renameRangeText(it.range, s.name, name);
    }
    s.name = name;
  }

  moveSheet(from: number, to: number): void {
    const [s] = this.sheets.splice(from, 1);
    this.sheets.splice(to, 0, s);
    this.engine.moveSheet(s.id, to);
  }

  sheetById(id: number): SheetModel | undefined {
    return this.sheets.find((s) => s.id === id);
  }

  sheetByName(name: string): SheetModel | undefined {
    const u = name.toLocaleUpperCase("tr-TR");
    return this.sheets.find((s) => s.name.toLocaleUpperCase("tr-TR") === u);
  }

  isValidSheetName(name: string, except?: SheetModel): string | null {
    const n = name.trim();
    if (!n) return "Sayfa adı boş olamaz.";
    if (n.length > 31) return "Sayfa adı en çok 31 karakter olabilir.";
    if (/[\\/?*[\]:]/.test(n)) return "Sayfa adı şu karakterleri içeremez: \\ / ? * [ ] :";
    if (n.startsWith("'") || n.endsWith("'")) return "Sayfa adı kesme işaretiyle başlayıp bitemez.";
    const other = this.sheetByName(n);
    if (other && other !== except) return "Bu ad zaten kullanılıyor. Başka bir ad deneyin.";
    return null;
  }

  /* ------------------------------ Hücreler ------------------------------ */

  getValue(sheet: SheetModel, row: number, col: number): Scalar {
    return this.engine.getValue(sheet.id, row, col);
  }

  getFormula(sheet: SheetModel, row: number, col: number): string | undefined {
    return this.engine.getFormula(sheet.id, row, col);
  }

  getInput(sheet: SheetModel, row: number, col: number): CellInput {
    // Tek kayıt araması (sıralama gibi toplu işlemlerde sık çağrılır)
    const rec = this.engine.getRecord(sheet.id, row, col);
    if (!rec) return null;
    return rec.formula ? { formula: "=" + rec.formula.text } : rec.value;
  }

  styleIdAt(sheet: SheetModel, row: number, col: number): number {
    const k = cellKey(row, col);
    return sheet.styles.get(k) ?? sheet.rowStyles.get(row) ?? sheet.colStyles.get(col) ?? 0;
  }

  /** Hücrenin kendi stil kaydı (satır/sütun stili hariç). */
  ownStyleId(sheet: SheetModel, row: number, col: number): number | undefined {
    return sheet.styles.get(cellKey(row, col));
  }

  styleAt(sheet: SheetModel, row: number, col: number): CellStyle {
    return this.styles.get(this.styleIdAt(sheet, row, col));
  }

  setInputRaw(sheet: SheetModel, row: number, col: number, input: CellInput): void {
    if (input !== null && typeof input === "object" && "formula" in input) this.engine.setFormula(sheet.id, row, col, input.formula);
    else this.engine.setValue(sheet.id, row, col, input as Scalar);
  }

  setStyleRaw(sheet: SheetModel, row: number, col: number, id: number | undefined): void {
    const k = cellKey(row, col);
    this.styleWrites++;
    if (id === undefined) sheet.styles.delete(k);
    else sheet.styles.set(k, id);
  }

  /** Düzenleyicide gösterilecek metin (Türkçe formül ya da değer). */
  editText(sheet: SheetModel, row: number, col: number): string {
    // Veri tablosu sonuçları: {=TABLO(satır_giriş;sütun_giriş)}
    const dt = sheet.dataTables.length ? sheet.dataTables.find((d) => row >= d.ref.r1 && row <= d.ref.r2 && col >= d.ref.c1 && col <= d.ref.c2) : undefined;
    if (dt) return `{=TABLO(${dt.rowInput ? addrOf(dt.rowInput.row, dt.rowInput.col) : ""};${dt.colInput ? addrOf(dt.colInput.row, dt.colInput.col) : ""})}`;
    const f = this.engine.getFormula(sheet.id, row, col);
    if (f) {
      // Tablonun içindeki formülde kendi tablosuna başvurular adsız gösterilir ([@Adet])
      const t = f.includes("[") ? sheet.tableAt(row, col) : undefined;
      return toLocalFormula(t ? unqualifyStructRefs(f, t.name) : f, "tr");
    }
    if (!this.engine.hasInput(sheet.id, row, col)) return "";
    const v = this.engine.getValue(sheet.id, row, col);
    return valueEditText(v, this.styleAt(sheet, row, col).numFmt);
  }

  /** Düzenleme metni, metin değerlerinde kesme işareti öneki olmadan (arama, metni sütunlara bölme). */
  plainText(sheet: SheetModel, row: number, col: number): string {
    const t = this.editText(sheet, row, col);
    return t.startsWith("'") && typeof this.getValue(sheet, row, col) === "string" ? t.slice(1) : t;
  }

  /** Aralıktaki kayıtlı hücre konumları (içerik, stil, not, bağlantı). */
  cellKeysIn(sheet: SheetModel, r: Rect): Set<number> {
    const keys = new Set<number>();
    const areaSize = (r.r2 - r.r1 + 1) * (r.c2 - r.c1 + 1);
    if (areaSize <= 20000) {
      for (let i = r.r1; i <= r.r2; i++)
        for (let j = r.c1; j <= r.c2; j++) {
          const k = cellKey(i, j);
          if (sheet.styles.has(k) || sheet.notes.has(k) || sheet.links.has(k) || this.engine.hasInput(sheet.id, i, j)) keys.add(k);
        }
      return keys;
    }
    const add = (k: number) => {
      if (contains(r, keyRow(k), keyCol(k))) keys.add(k);
    };
    this.engine.forEachCell(sheet.id, (row, col) => add(cellKey(row, col)));
    for (const k of sheet.styles.keys()) add(k);
    for (const k of sheet.notes.keys()) add(k);
    for (const k of sheet.links.keys()) add(k);
    return keys;
  }

  cellState(sheet: SheetModel, k: number): CellState {
    const row = keyRow(k);
    const col = keyCol(k);
    const st: CellState = { input: this.getInput(sheet, row, col) };
    const style = sheet.styles.get(k);
    if (style !== undefined) st.style = style;
    if (sheet.notes.size) {
      const note = sheet.notes.get(k);
      if (note) st.note = note;
    }
    if (sheet.links.size) {
      const link = sheet.links.get(k);
      if (link) st.link = link;
    }
    return st;
  }

  applyCellState(sheet: SheetModel, k: number, st: CellState | undefined): void {
    const row = keyRow(k);
    const col = keyCol(k);
    this.setInputRaw(sheet, row, col, st?.input ?? null);
    this.setStyleRaw(sheet, row, col, st?.style);
    if (st?.note) sheet.notes.set(k, st.note);
    else sheet.notes.delete(k);
    if (st?.link) sheet.links.set(k, st.link);
    else sheet.links.delete(k);
  }

  /**
   * Sayfanın kullanılan alanı: değeri olan (taşan diziler dahil) ya da kendi stili olan hücrelerin
   * sınırları (Ctrl+End, kaydırma çubukları). Izgara bunu her karede çağırdığı için sonuç sayfa
   * başına önbelleğe alınır. Damga O(1)'dir: motorun girdi ve hesaplama sürümleri (değer/formül,
   * satır-sütun ekle-sil, sıralama, taşıma, geri al/yinele, yükleme) ile stil haritasının kimliği,
   * boyutu ve setStyleRaw sayacı. Stil haritasını yerinde değiştiren kod setStyleRaw'ı kullanmalı ya
   * da haritayı yenisiyle değiştirmeli.
   */
  usedRange(sheet: SheetModel): Rect | null {
    const e = this.engine;
    const c = this.usedCache.get(sheet);
    if (
      !c ||
      c.inputVersion !== e.inputVersion ||
      c.calcVersion !== e.version ||
      c.styles.deref() !== sheet.styles ||
      c.styleCount !== sheet.styles.size ||
      c.styleWrites !== this.styleWrites
    ) {
      const rect = this.scanUsedRange(sheet);
      this.usedCache.set(sheet, { inputVersion: e.inputVersion, calcVersion: e.version, styles: new WeakRef(sheet.styles), styleCount: sheet.styles.size, styleWrites: this.styleWrites, rect });
      return rect && { ...rect };
    }
    // Çağıran sonucu değiştirse de önbellek bozulmasın
    return c.rect && { ...c.rect };
  }

  private scanUsedRange(sheet: SheetModel): Rect | null {
    let r2 = -1;
    let c2 = -1;
    let r1 = Infinity;
    let c1 = Infinity;
    const upd = (r: number, c: number) => {
      if (r > r2) r2 = r;
      if (c > c2) c2 = c;
      if (r < r1) r1 = r;
      if (c < c1) c1 = c;
    };
    this.engine.forEachValue(sheet.id, (r, c, v) => v !== null && upd(r, c));
    for (const k of sheet.styles.keys()) upd(keyRow(k), keyCol(k));
    if (r2 < 0) return null;
    return { r1, c1, r2, c2 };
  }

  /** Değeri olan hücrelerin sınırları (yalnız içerik). */
  dataExtent(sheet: SheetModel): { rows: number; cols: number } {
    return this.engine.getUsedExtent(sheet.id);
  }

  isEmpty(sheet: SheetModel, row: number, col: number): boolean {
    const v = this.engine.getValue(sheet.id, row, col);
    return v === null;
  }

  /** Satır/sütun ekleme-silmede bu sayfayı gösteren PivotTable kaynaklarını, grafik ve mini grafik aralıklarını kaydırır. */
  shiftDataRefs(sheet: SheetModel, axis: "row" | "col", at: number, count: number): void {
    for (const s of this.sheets) {
      for (const p of s.pivots) {
        if (p.sourceSheet !== sheet.id) continue;
        const r = shiftRect(p.source, axis, at, count);
        if (r) p.source = r;
      }
      for (const c of s.charts) {
        c.range = shiftRangeText(c.range, s.name, sheet.name, axis, at, count);
        mapSeriesRanges(c, (t) => shiftRangeText(t, s.name, sheet.name, axis, at, count));
      }
      for (const g of s.sparklines) for (const it of g.items) it.range = shiftRangeText(it.range, s.name, sheet.name, axis, at, count);
    }
  }

  /** Tablo tanımlarını hesap motoruna bildirir (yapısal başvurular); değişmediyse işlem yapılmaz. */
  syncTables(): void {
    const list: TableInfo[] = [];
    for (const s of this.sheets)
      for (const t of s.tables) list.push({ name: t.name, sheet: s.id, r1: t.ref.r1, c1: t.ref.c1, r2: t.ref.r2, c2: t.ref.c2, headerRow: t.headerRow, totalRow: t.totalRow, columns: t.columns });
    this.engine.setTables(list);
  }

  /** Tüm sayfaların anlık görüntüsü (yapısal işlemleri geri almak için). */
  snapshot(): WorkbookSnapshot {
    return {
      engine: this.engine.exportInputs(true),
      sheets: this.sheets.map((s) => ({ id: s.id, name: s.name, data: s.cloneData() })),
      activeIndex: this.activeIndex,
      theme: structuredClone(this.theme),
      protection: structuredClone(this.protection),
    };
  }

  restore(snap: WorkbookSnapshot): void {
    this.engine.importInputs(snap.engine);
    const byId = new Map(this.sheets.map((s) => [s.id, s]));
    this.sheets = snap.sheets.map((ss) => {
      const s = byId.get(ss.id) ?? new SheetModel(ss.id, ss.name);
      s.name = ss.name;
      s.restoreData(ss.data);
      return s;
    });
    this.activeIndex = Math.min(snap.activeIndex, this.sheets.length - 1);
    this.theme = structuredClone(snap.theme);
    if (snap.protection !== undefined) this.protection = structuredClone(snap.protection);
    this.syncTables();
    this.engine.recalc();
  }
}

/** Grafiğin açık seri aralıklarına dönüşüm uygular. */
function mapSeriesRanges(c: ChartDef, fn: (t: string) => string): void {
  if (!c.series) return;
  c.series = c.series.map((s) => ({
    ...s,
    values: fn(s.values),
    ...(s.categories ? { categories: fn(s.categories) } : {}),
    ...(s.xValues ? { xValues: fn(s.xValues) } : {}),
    ...(s.nameRef ? { nameRef: fn(s.nameRef) } : {}),
  }));
}

export interface WorkbookSnapshot {
  engine: ReturnType<FormulaEngine["exportInputs"]>;
  sheets: { id: number; name: string; data: SheetData }[];
  activeIndex: number;
  theme: WorkbookTheme;
  protection?: WorkbookProtection | null;
}

/** Değeri düzenleme metnine çevirir (tarihler gg.aa.yyyy, yüzde %, ondalık virgül). */
export function valueEditText(v: Scalar, numFmt: string | undefined): string {
  if (v === null) return "";
  if (typeof v === "boolean") return v ? "DOĞRU" : "YANLIŞ";
  if (isError(v)) return localErrorName(v.code);
  if (typeof v === "string") {
    // Sayı/tarih/formül gibi okunacak metinler kesme işaretiyle gösterilir ('123): yeniden onaylanınca metin kalır
    if (v === "" || numFmt === "@") return v;
    const p = parseCellInput(v, { numFmt });
    return p.kind === "value" && p.value === v ? v : "'" + v;
  }
  if (numFmt && isDateFormat(numFmt)) {
    if (isTimeOnlyFormat(numFmt)) return formatValue(v, "h:mm:ss").text;
    const frac = v - Math.floor(v);
    return formatValue(v, frac ? "dd.mm.yyyy h:mm:ss" : "dd.mm.yyyy").text;
  }
  if (numFmt && /%/.test(numFmt.replace(/"[^"]*"/g, ""))) return numberToText(Math.round(v * 100 * 1e10) / 1e10) + "%";
  return numberToText(v);
}
