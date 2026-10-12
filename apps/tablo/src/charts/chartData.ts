/** Grafik verisi: aralıktan seri/kategori tahmini (Excel'in sezgisel kuralları). */
import { formatValue, isDateFormat, parseQualifiedRange } from "@officehere/formula";
import { formatScalar } from "../model/display";
import type { ChartDef, Rect } from "../model/types";
import type { SheetModel, WorkbookModel } from "../model/workbook";

export interface ChartSeries {
  name: string;
  values: (number | null)[];
  /** Dağılım grafiği için x değerleri. */
  xs?: (number | null)[];
}

export interface ChartData {
  categories: string[];
  series: ChartSeries[];
  numFmt?: string;
}

export const PALETTES: string[][] = [
  ["#156082", "#E97132", "#196B24", "#0F9ED5", "#A02B93", "#4EA72E", "#4D8FB0", "#EE9767"],
  ["#4472C4", "#ED7D31", "#A5A5A5", "#FFC000", "#5B9BD5", "#70AD47", "#264478", "#9E480E"],
  ["#0E2841", "#156082", "#4D8FB0", "#8AB8D0", "#C0E4F5", "#0F9ED5", "#5BB8E0", "#A7D9EF"],
  ["#9C4A1B", "#E97132", "#EE9767", "#F4BC9C", "#FBE2D5", "#C0504D", "#D99694", "#E6B9B8"],
  ["#196B24", "#4EA72E", "#7DC160", "#A9D69B", "#DAF2D0", "#2E7D32", "#66BB6A", "#A5D6A7"],
  ["#262626", "#595959", "#7F7F7F", "#A6A6A6", "#BFBFBF", "#D9D9D9", "#404040", "#8C8C8C"],
];

export const PALETTE_LABELS = ["Renkli Palet 1", "Renkli Palet 2", "Tek Renkli Palet 1 (Mavi)", "Tek Renkli Palet 2 (Turuncu)", "Tek Renkli Palet 3 (Yeşil)", "Tek Renkli Palet 4 (Gri)"];

export function resolveRange(wb: WorkbookModel, def: ChartDef): { sheet: SheetModel; rect: Rect } | null {
  const q = parseQualifiedRange(def.range);
  if (!q) return null;
  const sheet = q.sheet ? wb.sheetByName(q.sheet) : wb.active;
  if (!sheet) return null;
  const ext = wb.dataExtent(sheet);
  const r = q.range;
  return { sheet, rect: { r1: r.r1, c1: r.c1, r2: Math.min(r.r2, Math.max(r.r1, ext.rows - 1)), c2: Math.min(r.c2, Math.max(r.c1, ext.cols - 1)) } };
}

/** Varsayılan seri yönü: satırlar sütunlardan fazlaysa seriler sütunlarda. */
export function defaultSeriesInRows(r: Rect): boolean {
  return r.c2 - r.c1 > r.r2 - r.r1;
}

export interface ChartLayout {
  sheet: SheetModel;
  rect: Rect;
  headerRow: boolean;
  labelCol: boolean;
  dataR1: number;
  dataC1: number;
}

/** Aralıktaki başlık satırı / etiket sütunu sezgisi. */
export function chartLayout(wb: WorkbookModel, def: ChartDef): ChartLayout | null {
  const res = resolveRange(wb, def);
  if (!res) return null;
  const { sheet, rect } = res;
  const get = (r: number, c: number) => wb.getValue(sheet, r, c);
  const rows = rect.r2 - rect.r1 + 1;
  const cols = rect.c2 - rect.c1 + 1;
  let headerRow = false;
  let labelCol = false;
  const firstRowVals = Array.from({ length: cols }, (_, j) => get(rect.r1, rect.c1 + j));
  const firstColVals = Array.from({ length: rows }, (_, i) => get(rect.r1 + i, rect.c1));
  if (rows > 1 && firstRowVals.slice(cols > 1 ? 1 : 0).some((v) => typeof v === "string") && !firstRowVals.slice(1).every((v) => typeof v === "number")) headerRow = true;
  const firstColFmt = wb.styleAt(sheet, rect.r1 + (headerRow ? 1 : 0), rect.c1).numFmt;
  if (cols > 1 && (firstColVals.slice(headerRow ? 1 : 0).some((v) => typeof v === "string") || isDateFormat(firstColFmt) || (headerRow && (firstRowVals[0] === null || firstRowVals[0] === "")))) labelCol = true;
  if (def.type === "scatter" && cols > 1) labelCol = true;
  return { sheet, rect, headerRow, labelCol, dataR1: rect.r1 + (headerRow ? 1 : 0), dataC1: rect.c1 + (labelCol ? 1 : 0) };
}

export interface SeriesRefs {
  sheetName: string;
  cat: Rect | null;
  series: { name: Rect | null; values: Rect; xs?: Rect | null }[];
}

/** Dosyaya yazmak için seri başvuruları. */
export function seriesRefs(wb: WorkbookModel, def: ChartDef): SeriesRefs | null {
  const L = chartLayout(wb, def);
  if (!L) return null;
  const { rect, headerRow, labelCol, dataR1, dataC1 } = L;
  const out: SeriesRefs = { sheetName: L.sheet.name, cat: null, series: [] };
  if (!def.seriesInRows) {
    if (labelCol) out.cat = { r1: dataR1, c1: rect.c1, r2: rect.r2, c2: rect.c1 };
    for (let c = dataC1; c <= rect.c2; c++)
      out.series.push({ name: headerRow ? { r1: rect.r1, c1: c, r2: rect.r1, c2: c } : null, values: { r1: dataR1, c1: c, r2: rect.r2, c2: c }, xs: def.type === "scatter" && labelCol ? out.cat : undefined });
  } else {
    if (headerRow) out.cat = { r1: rect.r1, c1: dataC1, r2: rect.r1, c2: rect.c2 };
    for (let r = dataR1; r <= rect.r2; r++) out.series.push({ name: labelCol ? { r1: r, c1: rect.c1, r2: r, c2: rect.c1 } : null, values: { r1: r, c1: dataC1, r2: r, c2: rect.c2 } });
  }
  return out;
}

export function extractChartData(wb: WorkbookModel, def: ChartDef): ChartData {
  const L = chartLayout(wb, def);
  if (!L) return { categories: [], series: [] };
  const { sheet, rect, headerRow, labelCol, dataR1, dataC1 } = L;
  const get = (r: number, c: number) => wb.getValue(sheet, r, c);
  const label = (r: number, c: number) => {
    const v = get(r, c);
    if (typeof v === "number") return formatValue(v, wb.styleAt(sheet, r, c).numFmt).text;
    return formatScalar(v);
  };
  const num = (r: number, c: number): number | null => {
    const v = get(r, c);
    return typeof v === "number" ? v : null;
  };
  const numFmt = wb.styleAt(sheet, dataR1, dataC1).numFmt;
  const out: ChartData = { categories: [], series: [], numFmt };
  if (!def.seriesInRows) {
    for (let r = dataR1; r <= rect.r2; r++) out.categories.push(labelCol ? label(r, rect.c1) : String(r - dataR1 + 1));
    for (let c = dataC1; c <= rect.c2; c++) {
      const name = headerRow ? label(rect.r1, c) : `Seri${c - dataC1 + 1}`;
      const values: (number | null)[] = [];
      for (let r = dataR1; r <= rect.r2; r++) values.push(num(r, c));
      const s: ChartSeries = { name, values };
      if (def.type === "scatter" && labelCol) s.xs = Array.from({ length: rect.r2 - dataR1 + 1 }, (_, i) => num(dataR1 + i, rect.c1));
      out.series.push(s);
    }
  } else {
    for (let c = dataC1; c <= rect.c2; c++) out.categories.push(headerRow ? label(rect.r1, c) : String(c - dataC1 + 1));
    for (let r = dataR1; r <= rect.r2; r++) {
      const name = labelCol ? label(r, rect.c1) : `Seri${r - dataR1 + 1}`;
      const values: (number | null)[] = [];
      for (let c = dataC1; c <= rect.c2; c++) values.push(num(r, c));
      out.series.push({ name, values });
    }
  }
  return out;
}

/** Eksen için "güzel" adım ve sınırlar. */
export function niceScale(min: number, max: number, ticks = 5): { min: number; max: number; step: number } {
  if (min === max) {
    if (min === 0) return { min: 0, max: 1, step: 0.2 };
    min = min > 0 ? 0 : min;
    max = max < 0 ? 0 : max;
  }
  const range = max - min;
  const raw = range / ticks;
  const mag = Math.pow(10, Math.floor(Math.log10(raw)));
  const norm = raw / mag;
  const step = (norm <= 1 ? 1 : norm <= 2 ? 2 : norm <= 2.5 ? 2.5 : norm <= 5 ? 5 : 10) * mag;
  return { min: Math.floor(min / step) * step, max: Math.ceil(max / step) * step, step };
}
