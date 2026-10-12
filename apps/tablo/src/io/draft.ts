/**
 * Otomatik Kurtarma taslağı: çalışma kitabının IndexedDB'ye yapılandırılmış
 * kopya (structured clone) ile yazılabilen sıkı biçimi. Hücreler sütun
 * dizileri (tür, sayı, metin) olarak tutulur; serileştirme üreteç ile
 * parça parça yapılır (büyük kitaplarda arayüz donmaz).
 */
import { errorFromCode, isError } from "@officehere/formula";
import { AxisSizes } from "../model/axis";
import type { CellStyle } from "../model/styles";
import type { WorkbookTheme } from "../model/theme";
import { WorkbookModel, type DocProps, type SheetData } from "../model/workbook";

const K_NUM = 1;
const K_STR = 2;
const K_BOOL = 3;
const K_ERR = 4;
const K_FORMULA = 5;

interface PlainAxis {
  defaultSize: number;
  count: number;
  sizes: [number, number][];
  hidden: number[];
  manual: number[];
}

export interface DraftSheet {
  name: string;
  /** Hücre girdileri: anahtar (satır*16384+sütun), tür, sayı, metin dizileri. */
  keys: Float64Array;
  kinds: Uint8Array;
  nums: Float64Array;
  strs: string[];
  styleKeys: Float64Array;
  styleIds: Int32Array;
  rows: PlainAxis;
  cols: PlainAxis;
  /** Kalan sayfa verisi (nesneler, birleştirmeler, kurallar…): yapılandırılmış kopyalanabilir. */
  rest: Omit<SheetData, "rows" | "cols" | "styles">;
  view: { showGridlines: boolean; showHeadings: boolean; zoom: number; frozen: { rows: number; cols: number }; protected: boolean };
}

export interface WorkbookDraft {
  version: 1;
  sheets: DraftSheet[];
  styles: CellStyle[];
  names: { name: string; formula: string; sheet?: number }[];
  theme: WorkbookTheme;
  props: DocProps;
  activeIndex: number;
  /** Toplam hücre sayısı (bilgi). */
  cellCount: number;
}

function plainAxis(a: AxisSizes): PlainAxis {
  return { defaultSize: a.defaultSize, count: a.count, sizes: a.customItems(), hidden: a.hiddenItems(), manual: a.manualItems() };
}

function restoreAxis(a: AxisSizes, p: PlainAxis): void {
  a.setDefault(p.defaultSize);
  const manual = new Set(p.manual);
  for (const [i, v] of p.sizes) a.set(i, v, manual.has(i));
  a.setHiddenMany(p.hidden, true);
}

/**
 * Kitabı parça parça serileştirir: her `yield` bir ara noktadır (çağıran
 * boşta kalma süresine göre devam eder). Sonuç `return` ile döner.
 */
export function* serializeDraft(wb: WorkbookModel, chunk = 20000): Generator<void, WorkbookDraft, void> {
  const sheets: DraftSheet[] = [];
  let cellCount = 0;
  for (const s of wb.sheets) {
    const n = wb.engine.cellCount(s.id);
    const keys = new Float64Array(n);
    const kinds = new Uint8Array(n);
    const nums = new Float64Array(n);
    const strs: string[] = [];
    let i = 0;
    const strIndex = (t: string) => {
      strs.push(t);
      return strs.length - 1;
    };
    const cells: [number, number, { value: unknown; formula?: { text: string } }][] = [];
    wb.engine.forEachCell(s.id, (row, col, rec) => cells.push([row, col, rec]));
    for (const [row, col, rec] of cells) {
      keys[i] = row * 16384 + col;
      const v = rec.value;
      if (rec.formula) {
        kinds[i] = K_FORMULA;
        nums[i] = strIndex(rec.formula.text);
      } else if (typeof v === "number") {
        kinds[i] = K_NUM;
        nums[i] = v;
      } else if (typeof v === "string") {
        kinds[i] = K_STR;
        nums[i] = strIndex(v);
      } else if (typeof v === "boolean") {
        kinds[i] = K_BOOL;
        nums[i] = v ? 1 : 0;
      } else if (isError(v as never)) {
        kinds[i] = K_ERR;
        nums[i] = strIndex((v as { code: string }).code);
      }
      i++;
      if (i % chunk === 0) yield;
    }
    cellCount += n;
    const styleKeys = new Float64Array(s.styles.size);
    const styleIds = new Int32Array(s.styles.size);
    let j = 0;
    for (const [k, id] of s.styles) {
      styleKeys[j] = k;
      styleIds[j] = id;
      j++;
    }
    const data = s.cloneData();
    const { rows: _r, cols: _c, styles: _s, ...rest } = data;
    void _r;
    void _c;
    void _s;
    sheets.push({
      name: s.name,
      keys,
      kinds,
      nums,
      strs,
      styleKeys,
      styleIds,
      rows: plainAxis(s.rows),
      cols: plainAxis(s.cols),
      rest,
      view: { showGridlines: s.showGridlines, showHeadings: s.showHeadings, zoom: s.zoom, frozen: { ...s.frozen }, protected: s.protected },
    });
    yield;
  }
  const names = wb.engine.getNames().map((n) => ({ name: n.name, formula: n.formula, sheet: n.sheet !== undefined ? wb.sheets.findIndex((s) => s.id === n.sheet) : undefined }));
  return { version: 1, sheets, styles: wb.styles.all().map((x) => ({ ...x })), names, theme: structuredClone(wb.theme), props: { ...wb.props }, activeIndex: wb.activeIndex, cellCount };
}

/** Üreteci sonuna kadar çalıştırır (testler ve sayfa kapanırken). */
export function serializeDraftSync(wb: WorkbookModel): WorkbookDraft {
  const g = serializeDraft(wb, 1e9);
  for (;;) {
    const r = g.next();
    if (r.done) return r.value;
  }
}

/** Taslaktan çalışma kitabını kurar. */
export function workbookFromDraft(d: WorkbookDraft): WorkbookModel {
  const wb = new WorkbookModel(true);
  for (const st of d.styles.slice(1)) wb.styles.intern(st);
  wb.theme = structuredClone(d.theme);
  wb.props = { ...d.props };
  wb.engine.autoCalc = false;
  for (const ds of d.sheets) {
    const s = wb.addSheet(ds.name);
    const n = ds.keys.length;
    for (let i = 0; i < n; i++) {
      const k = ds.keys[i];
      const row = Math.floor(k / 16384);
      const col = k % 16384;
      const kind = ds.kinds[i];
      const x = ds.nums[i];
      if (kind === K_FORMULA) wb.engine.setFormula(s.id, row, col, "=" + ds.strs[x]);
      else if (kind === K_NUM) wb.engine.setValue(s.id, row, col, x);
      else if (kind === K_STR) wb.engine.setValue(s.id, row, col, ds.strs[x]);
      else if (kind === K_BOOL) wb.engine.setValue(s.id, row, col, x === 1);
      else if (kind === K_ERR) wb.engine.setValue(s.id, row, col, errorFromCode(ds.strs[x]) ?? null);
    }
    const rows = new AxisSizes(ds.rows.defaultSize, ds.rows.count);
    const cols = new AxisSizes(ds.cols.defaultSize, ds.cols.count);
    restoreAxis(rows, ds.rows);
    restoreAxis(cols, ds.cols);
    const styles = new Map<number, number>();
    for (let j = 0; j < ds.styleKeys.length; j++) styles.set(ds.styleKeys[j], ds.styleIds[j]);
    s.restoreData({ ...ds.rest, rows, cols, styles });
    s.showGridlines = ds.view.showGridlines;
    s.showHeadings = ds.view.showHeadings;
    s.zoom = ds.view.zoom;
    s.frozen = { ...ds.view.frozen };
    s.protected = ds.view.protected;
  }
  for (const n of d.names) {
    try {
      wb.engine.defineName(n.name, n.formula, n.sheet !== undefined && n.sheet >= 0 ? wb.sheets[n.sheet]?.id : undefined);
    } catch {
      /* geçersiz ad atlanır */
    }
  }
  if (!wb.sheets.length) wb.addSheet("Sayfa1");
  wb.bumpIds();
  wb.activeIndex = Math.min(Math.max(0, d.activeIndex), wb.sheets.length - 1);
  wb.engine.autoCalc = true;
  wb.syncTables();
  wb.engine.recalc(true);
  return wb;
}
