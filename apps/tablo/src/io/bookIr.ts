/**
 * Çalışma kitabının ara gösterimi (IR): dosya G/Ç Worker'ı ile ana iş parçacığı arasında taşınır.
 * Hücreler sayfa başına sütunlu dizilerde (satır, sütun, tür, sayı ya da dize dizini, formül metni
 * dizini) ve tek bir dize tablosunda durur; bu tamponlar aktarılır, kopyalanmaz. Stiller, nesneler,
 * kurallar ve geçiş verisi düz (yapılandırılmış kopyalanabilir) nesnelerdir.
 *
 * Kitap → IR ve IR → kitap dönüşümleri üreteçtir: ana iş parçacığında `runSliced` ile parça parça,
 * Worker'da `runSync` ile tek seferde çalışır. Formüller hesaplanmış değerleriyle taşınır ve
 * yeniden hesaplanmaz (motorun toplu yüklemesi, FormulaEngine.bulkLoad); yalnız taşan diziler ve
 * döngüsel başvurudaki hücreler hesapla yeniden kurulur.
 */
import { Errors, FormulaError, cellKey, isError, keyCol, keyRow, type Scalar } from "@officehere/formula";
import type { AxisSizes } from "../model/axis";
import type { CellStyle } from "../model/styles";
import type { WorkbookTheme } from "../model/theme";
import type { Selection, ViewMode, WorkbookProtection } from "../model/types";
import type { WorkbookPassthrough } from "../model/passthrough";
import { SheetModel, WorkbookModel, type DocProps, type SheetData } from "../model/workbook";

/* Hücre türü (alt 3 bit: değer türü) ve bayraklar. */
const V_NULL = 0;
const V_NUM = 1;
const V_STR = 2;
const V_BOOL = 3;
const V_ERR = 4;
const V_MASK = 7;
/** Formül hücresi: metni `forms` dizisinde. */
const F_FORMULA = 8;
/** Önbellek değeri kullanılmaz, hücre hesapla kurulur (taşan dizi ya da döngüsel başvuru). */
const F_RECALC = 16;

/** Üreteçlerin ara noktalarında bildirdiği ilerleme (0..1); `undefined` yalnız ara noktadır. */
type Step = number | undefined;
/** Hücre döngülerinde bu kadar hücrede bir ara nokta. */
const CHUNK = 2048;

/**
 * Dize tablosu: dizeler sırayla `parts` metinlerine bitiştirilir (her parçada `counts[p]` dize),
 * i. dizenin uzunluğu `lens[i]`. Az sayıda büyük metin hem kopyalamada hem çözmede hızlıdır.
 */
export interface StringTableIr {
  parts: string[];
  counts: Int32Array;
  lens: Int32Array;
}

/** Sayfanın hücreleri (girdi kayıtları; motorun sırasıyla). */
export interface CellsIr {
  rows: Int32Array;
  cols: Int32Array;
  /** Değer türü + bayraklar. */
  kinds: Uint8Array;
  /** Sayı, mantıksal (0/1) ya da dize tablosu dizini (metin, hata). */
  nums: Float64Array;
  /** Formül metninin dize dizini (formülsüz hücrede -1); sayfada formül yoksa null. */
  forms: Int32Array | null;
}

/** Satır/sütun boyutları (AxisSizes). */
export interface AxisIr {
  def: number;
  /** Özel boyutlu öğeler ve boyutları. */
  idx: Int32Array;
  px: Float64Array;
  manual: Int32Array;
  hidden: Int32Array;
  /** Anahat düzeyleri: öğe, düzey çiftleri. */
  levels: Int32Array;
  collapsed: Int32Array;
}

/** Sayfanın hücre, stil ve boyut dışındaki tüm alanları (SheetModel alan adlarıyla). */
export type SheetPropsIr = Omit<SheetData, "rows" | "cols" | "styles"> & {
  showGridlines: boolean;
  showHeadings: boolean;
  showFormulas: boolean;
  zoom: number;
  viewMode: ViewMode;
  selection: Selection;
  scroll: { row: number; col: number };
  protected: boolean;
  protectPassword?: string;
};

export interface SheetIr {
  id: number;
  name: string;
  cells: CellsIr;
  /** Hücre stilleri (anahtar satır*16384+sütun → stil numarası). */
  styleKeys: Float64Array;
  styleIds: Int32Array;
  rows: AxisIr;
  cols: AxisIr;
  props: SheetPropsIr;
}

export interface BookIr {
  v: 1;
  strings: StringTableIr;
  sheets: SheetIr[];
  /** Stil tablosu (0 varsayılan); hücreler numarayla başvurur. */
  styles: CellStyle[];
  /** Tanımlı adlar (kanonik metin, sayfa kapsamı sayfa kimliğiyle). */
  names: { name: string; text: string; sheet?: number }[];
  activeIndex: number;
  props: DocProps;
  theme: WorkbookTheme;
  loadWarnings: string[];
  protection: WorkbookProtection | null;
  passthrough: WorkbookPassthrough | null;
  /** Kaynak paket kaydının kimliği (yazmada Worker kaynağı bu kimlikle kaydeder). */
  kaynak?: string;
}

/* ------------------------------ Dize tablosu ------------------------------ */

/** Yinelenen dizeleri bir kez yazan dize tablosu kurucusu. */
class StringTableBuilder {
  private index = new Map<string, number>();
  private lens: number[] = [];
  private parts: string[] = [];
  private counts: number[] = [];
  private cur: string[] = [];
  private curChars = 0;

  add(s: string): number {
    let i = this.index.get(s);
    if (i !== undefined) return i;
    i = this.lens.length;
    this.index.set(s, i);
    this.lens.push(s.length);
    this.cur.push(s);
    this.curChars += s.length;
    if (this.cur.length >= 32768 || this.curChars >= 1 << 20) this.flush();
    return i;
  }

  private flush(): void {
    if (!this.cur.length) return;
    this.parts.push(this.cur.join(""));
    this.counts.push(this.cur.length);
    this.cur = [];
    this.curChars = 0;
  }

  build(): StringTableIr {
    this.flush();
    return { parts: this.parts, counts: Int32Array.from(this.counts), lens: Int32Array.from(this.lens) };
  }
}

/** Dize tablosunu çözer; ilerleme `total` iş birimine göre (bir dize bir birim). */
function* decodeStrings(t: StringTableIr, total: number): Generator<Step, string[]> {
  const out = new Array<string>(t.lens.length);
  let i = 0;
  for (let p = 0; p < t.parts.length; p++) {
    const part = t.parts[p];
    let pos = 0;
    for (let k = 0; k < t.counts[p]; k++, i++) {
      const end = pos + t.lens[i];
      out[i] = part.slice(pos, end);
      pos = end;
      if ((i & 16383) === 16383) yield i / total;
    }
  }
  return out;
}

/* ------------------------------ Değerler ------------------------------ */

/** Hata değeri dize tablosunda "kod" ya da "kod\0ileti" olarak durur. */
function errorText(e: FormulaError): string {
  return e.message ? `${e.code}\u0000${e.message}` : e.code;
}

const KNOWN_ERRORS = new Map<string, FormulaError>(Object.values(Errors).map((e) => [e.code, e]));

function errorOf(text: string): FormulaError {
  const i = text.indexOf("\u0000");
  if (i < 0) return KNOWN_ERRORS.get(text) ?? new FormulaError(text as FormulaError["code"]);
  return new FormulaError(text.slice(0, i) as FormulaError["code"], text.slice(i + 1));
}

/* ------------------------------ Boyutlar ------------------------------ */

function axisIr(a: AxisSizes): AxisIr {
  const custom = a.customItems();
  const levels = a.levelItems();
  return {
    def: a.defaultSize,
    idx: Int32Array.from(custom, (x) => x[0]),
    px: Float64Array.from(custom, (x) => x[1]),
    manual: Int32Array.from(a.manualItems()),
    hidden: Int32Array.from(a.hiddenItems()),
    levels: Int32Array.from(levels.flat()),
    collapsed: Int32Array.from(a.collapsedItems()),
  };
}

/**
 * Boyutları birebir geri kurar. Varsayılan boyuta eşit özel boyut da korunur: önce varsayılan
 * hiçbir boyuta eşit olamayacak bir değere alınır; boyutsuz "elle" kayıtları gerçek varsayılanla yazılır.
 */
function restoreAxis(a: AxisSizes, p: AxisIr): void {
  const manual = new Set(p.manual);
  a.setDefault(-1);
  for (let i = 0; i < p.idx.length; i++) a.set(p.idx[i], p.px[i], manual.has(p.idx[i]));
  a.setDefault(p.def);
  const sized = new Set(p.idx);
  for (const i of p.manual) if (!sized.has(i)) a.set(i, p.def, true);
  a.setHiddenMany(p.hidden, true);
  for (let i = 0; i + 1 < p.levels.length; i += 2) a.setLevel(p.levels[i], p.levels[i + 1]);
  for (const i of p.collapsed) a.setCollapsed(i, true);
}

/** Sayfanın hücre/stil/boyut dışı alanları; dönen nesne modelin kendi nesnelerini gösterir (kopyalayan ileti gönderimidir). */
function sheetProps(s: SheetModel): SheetPropsIr {
  return {
    tabColor: s.tabColor,
    hidden: s.hidden,
    colStyles: s.colStyles,
    rowStyles: s.rowStyles,
    merges: s.merges,
    notes: s.notes,
    links: s.links,
    validations: s.validations,
    condFormats: s.condFormats,
    autoFilter: s.autoFilter,
    filterHidden: s.filterHidden,
    tables: s.tables,
    charts: s.charts,
    images: s.images,
    shapes: s.shapes,
    pivots: s.pivots,
    sparklines: s.sparklines,
    scenarios: s.scenarios,
    dataTables: s.dataTables,
    slicers: s.slicers,
    advFilter: s.advFilter,
    ignoredErrors: s.ignoredErrors,
    frozen: s.frozen,
    page: s.page,
    structVersion: s.structVersion,
    passthrough: s.passthrough,
    showGridlines: s.showGridlines,
    showHeadings: s.showHeadings,
    showFormulas: s.showFormulas,
    zoom: s.zoom,
    viewMode: s.viewMode,
    selection: s.selection,
    scroll: s.scroll,
    protected: s.protected,
    protectPassword: s.protectPassword,
  };
}

/** IR'nin aktarılacak tamponları. */
export function irTransfer(ir: BookIr): ArrayBuffer[] {
  const out: ArrayBuffer[] = [ir.strings.counts.buffer as ArrayBuffer, ir.strings.lens.buffer as ArrayBuffer];
  for (const s of ir.sheets) {
    const c = s.cells;
    out.push(c.rows.buffer as ArrayBuffer, c.cols.buffer as ArrayBuffer, c.kinds.buffer as ArrayBuffer, c.nums.buffer as ArrayBuffer, s.styleKeys.buffer as ArrayBuffer, s.styleIds.buffer as ArrayBuffer);
    if (c.forms) out.push(c.forms.buffer as ArrayBuffer);
    for (const a of [s.rows, s.cols]) out.push(a.idx.buffer as ArrayBuffer, a.px.buffer as ArrayBuffer, a.manual.buffer as ArrayBuffer, a.hidden.buffer as ArrayBuffer, a.levels.buffer as ArrayBuffer, a.collapsed.buffer as ArrayBuffer);
  }
  return [...new Set(out)];
}

/* ------------------------------ Kitap → IR ------------------------------ */

export interface ToIrOptions {
  /**
   * Ara noktalarda sorulur: kitap bu arada değiştiyse (ana iş parçacığında parça parça çıkarılırken
   * kullanıcı düzenledi) anlık görüntü tutarsızdır, üreteç null döner ve çağıran yeniden dener.
   */
  stale?: () => boolean;
}

/**
 * Kitabın anlık görüntüsünü IR olarak çıkarır. Formül değerleri güncel olmalıdır (çağıran önce
 * hesaplatır). Hücre ve stil haritaları parça parça gezilir; geri kalan alanlar sonda tek adımda alınır.
 */
export function* bookToIr(wb: WorkbookModel, opts: ToIrOptions = {}): Generator<Step, BookIr | null> {
  const strings = new StringTableBuilder();
  const e = wb.engine;
  const total = Math.max(1, wb.sheets.reduce((n, s) => n + e.cellCount(s.id) + s.styles.size, 0));
  let done = 0;
  const circular = new Map<number, Set<number>>();
  for (const c of e.getCircularRefs()) {
    let set = circular.get(c.sheet);
    if (!set) circular.set(c.sheet, (set = new Set()));
    set.add(cellKey(c.row, c.col));
  }
  const sheets: Omit<SheetIr, "rows" | "cols" | "props">[] = [];
  for (const s of wb.sheets) {
    const n = e.cellCount(s.id);
    const rows = new Int32Array(n);
    const cols = new Int32Array(n);
    const kinds = new Uint8Array(n);
    const nums = new Float64Array(n);
    let forms: Int32Array | null = null;
    const circ = circular.get(s.id);
    let i = 0;
    for (const [k, rec] of e.cellEntries(s.id)) {
      // Gezinti sırasında hücre eklendiyse görüntü zaten bayattır
      if (i >= n) return null;
      rows[i] = keyRow(k);
      cols[i] = keyCol(k);
      const v = rec.value;
      let kind = V_NULL;
      if (typeof v === "number") {
        kind = V_NUM;
        nums[i] = v;
      } else if (typeof v === "string") {
        kind = V_STR;
        nums[i] = strings.add(v);
      } else if (typeof v === "boolean") {
        kind = V_BOOL;
        nums[i] = v ? 1 : 0;
      } else if (isError(v)) {
        kind = V_ERR;
        nums[i] = strings.add(errorText(v));
      }
      const f = rec.formula;
      if (f) {
        if (!forms) forms = new Int32Array(n).fill(-1);
        forms[i] = strings.add(f.text);
        kind |= F_FORMULA;
        if (f.spill || f.spillTarget || circ?.has(k)) kind |= F_RECALC;
      }
      kinds[i] = kind;
      i++;
      if ((i & (CHUNK - 1)) === 0) {
        yield (done + i) / total;
        if (opts.stale?.()) return null;
      }
    }
    if (i !== n) return null;
    done += n;
    const styleKeys = new Float64Array(s.styles.size);
    const styleIds = new Int32Array(s.styles.size);
    let j = 0;
    for (const [k, id] of s.styles) {
      if (j >= styleKeys.length) return null;
      styleKeys[j] = k;
      styleIds[j] = id;
      j++;
      if ((j & (CHUNK * 4 - 1)) === 0) {
        yield (done + j) / total;
        if (opts.stale?.()) return null;
      }
    }
    done += j;
    sheets.push({ id: s.id, name: s.name, cells: { rows, cols, kinds, nums, forms }, styleKeys, styleIds });
    yield done / total;
  }
  if (opts.stale?.()) return null;
  // Son adım tek parçadır: hücre dışı alanlar, parçalar arasında değişmediği denetlendikten hemen sonra
  return {
    v: 1,
    strings: strings.build(),
    sheets: sheets.map((sh, i) => {
      const s = wb.sheets[i];
      return { ...sh, rows: axisIr(s.rows), cols: axisIr(s.cols), props: sheetProps(s) };
    }),
    styles: wb.styles.all(),
    names: e.getNames().map((n) => ({ name: n.name, text: n.formula.slice(1), sheet: n.sheet })),
    activeIndex: wb.activeIndex,
    props: wb.props,
    theme: wb.theme,
    loadWarnings: wb.loadWarnings,
    protection: wb.protection,
    passthrough: wb.passthrough,
    kaynak: wb.kaynak,
  };
}

/* ------------------------------ IR → kitap ------------------------------ */

/**
 * IR'den kitabı kurar. IR'nin nesneleri modele kopyalanmadan bağlanır (IR, ileti alımıyla zaten
 * kopyadır). Sayfalar, adlar ve tablolar hücrelerden önce kurulur: formül bağımlılıkları toplu
 * yüklemede çözülür. Kaynak paket kaydı (kaynak) çağıranın işidir.
 */
export function* bookFromIr(ir: BookIr): Generator<Step, WorkbookModel> {
  const total = Math.max(1, ir.strings.lens.length + ir.sheets.reduce((n, s) => n + s.cells.rows.length + s.styleKeys.length, 0));
  const strs = yield* decodeStrings(ir.strings, total);
  const wb = new WorkbookModel(true);
  // Stil tablosu birebir (yeniden normalleştirme numaraları kaydırabilirdi: "font" alanı etkin
  // temanın yazı tipine göre yazılır, bu iş parçacığında başka tema etkin olabilir)
  wb.styles.restore(ir.styles);
  wb.theme = ir.theme;
  wb.props = ir.props;
  wb.loadWarnings = ir.loadWarnings;
  wb.protection = ir.protection;
  wb.passthrough = ir.passthrough;
  // Sayfalar kimlikleriyle (özet tablo kaynakları ve ad kapsamları kimlikle başvurur)
  wb.engine.importInputs({ sheets: ir.sheets.map((s) => ({ id: s.id, name: s.name, cells: [] })), names: ir.names });
  wb.sheets = ir.sheets.map((si) => {
    const s = new SheetModel(si.id, si.name);
    Object.assign(s, si.props);
    restoreAxis(s.rows, si.rows);
    restoreAxis(s.cols, si.cols);
    return s;
  });
  wb.activeIndex = Math.min(Math.max(0, ir.activeIndex), Math.max(0, wb.sheets.length - 1));
  wb.syncTables();
  let done = strs.length;
  yield done / total;
  for (const [si, sheetIr] of ir.sheets.entries()) {
    const sheet = wb.sheets[si];
    const { rows, cols, kinds, nums, forms } = sheetIr.cells;
    const load = wb.engine.bulkLoad(sheet.id);
    const n = rows.length;
    for (let i = 0; i < n; i++) {
      const kind = kinds[i];
      const x = nums[i];
      let v: Scalar;
      switch (kind & V_MASK) {
        case V_NUM:
          v = x;
          break;
        case V_STR:
          v = strs[x];
          break;
        case V_BOOL:
          v = x === 1;
          break;
        case V_ERR:
          v = errorOf(strs[x]);
          break;
        default:
          v = null;
      }
      if (kind & F_FORMULA) load.formula(rows[i], cols[i], strs[forms![i]], kind & F_RECALC ? undefined : v);
      else load.value(rows[i], cols[i], v);
      if ((i & (CHUNK - 1)) === CHUNK - 1) yield (done + i) / total;
    }
    load.finish();
    done += n;
    const styles = sheet.styles;
    const { styleKeys, styleIds } = sheetIr;
    for (let j = 0; j < styleKeys.length; j++) {
      styles.set(styleKeys[j], styleIds[j]);
      if ((j & (CHUNK * 4 - 1)) === CHUNK * 4 - 1) yield (done + j) / total;
    }
    done += styleKeys.length;
  }
  wb.bumpIds();
  return wb;
}

/* ------------------------------ Çalıştırıcılar ------------------------------ */

/** Üreteci tek seferde sonuna kadar çalıştırır (Worker'da ve testlerde). */
export function runSync<T>(gen: Generator<Step, T>): T {
  for (;;) {
    const r = gen.next();
    if (r.done) return r.value;
  }
}

type SchedulerLike = { yield?: () => Promise<void> };

/** Ana iş parçacığını bir kez bırakır: bekleyen giriş olayları ve çizim araya girer. */
export function yieldToMain(): Promise<void> {
  const sch = (globalThis as { scheduler?: SchedulerLike }).scheduler;
  if (typeof sch?.yield === "function") return sch.yield();
  if (typeof MessageChannel === "function" && typeof window !== "undefined") {
    return new Promise((resolve) => {
      const ch = new MessageChannel();
      ch.port1.onmessage = () => resolve();
      ch.port2.postMessage(null);
    });
  }
  return new Promise((resolve) => setTimeout(resolve, 0));
}

export interface SliceOptions {
  signal?: AbortSignal;
  /** Bir dilimin en uzun süresi (ms); sonra iş parçacığı bırakılır. */
  budgetMs?: number;
  onProgress?: (ratio: number) => void;
}

function abortReason(signal: AbortSignal): unknown {
  return signal.reason ?? new DOMException("İşlem iptal edildi.", "AbortError");
}

/**
 * Üreteci dilimler hâlinde çalıştırır: her dilim en çok `budgetMs` sürer, aralarda iş parçacığı
 * bırakılır (scheduler.yield), iptal denetlenir ve ilerleme bildirilir. Uzun görev oluşmaz.
 */
export async function runSliced<T>(gen: Generator<Step, T>, opts: SliceOptions = {}): Promise<T> {
  const budget = opts.budgetMs ?? 12;
  let ratio = 0;
  for (;;) {
    if (opts.signal?.aborted) throw abortReason(opts.signal);
    const t0 = performance.now();
    for (;;) {
      const r = gen.next();
      if (r.done) return r.value;
      if (typeof r.value === "number") ratio = r.value;
      if (performance.now() - t0 >= budget) break;
    }
    opts.onProgress?.(ratio);
    await yieldToMain();
  }
}
