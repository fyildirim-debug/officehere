/** Grafik verisi yardımcıları: varsayılan veri, eksen ölçeği, sayı biçimi. */
import { uid } from "./factory";
import type { ChartData, ChartGrouping, ChartShape, ChartType, Color } from "./types";

const ACCENTS = ["accent1", "accent2", "accent3", "accent4", "accent5", "accent6"] as const;
type Accent = (typeof ACCENTS)[number];

/** Renkleri Değiştir paletleri: renkli (vurgu sırası) ve tek renkli (bir vurgunun tonları). */
export const CHART_PALETTES: { name: string; kind: "colorful" | "mono"; order: Accent[] }[] = [
  { name: "Renkli Palet 1", kind: "colorful", order: ["accent1", "accent2", "accent3", "accent4", "accent5", "accent6"] },
  { name: "Renkli Palet 2", kind: "colorful", order: ["accent2", "accent4", "accent6", "accent1", "accent3", "accent5"] },
  { name: "Renkli Palet 3", kind: "colorful", order: ["accent3", "accent5", "accent1", "accent2", "accent4", "accent6"] },
  { name: "Renkli Palet 4", kind: "colorful", order: ["accent4", "accent6", "accent2", "accent3", "accent5", "accent1"] },
  ...ACCENTS.map((a, i) => ({ name: `Tek Renkli Palet ${i + 1}`, kind: "mono" as const, order: [a] })),
];

const MONO_STEPS: Partial<Color>[] = [{}, { lumMod: 0.6 }, { lumMod: 0.6, lumOff: 0.4 }, { lumMod: 0.75 }, { lumMod: 0.8, lumOff: 0.2 }, { lumMod: 0.5 }];

/** i. seri/dilim rengi (6'dan sonrası koyulaştırılmış tonlar). */
export function seriesColor(i: number, explicit?: Color, palette = 0): Color {
  if (explicit) return explicit;
  const pal = CHART_PALETTES[palette] ?? CHART_PALETTES[0];
  if (pal.kind === "mono") return { scheme: pal.order[0], ...MONO_STEPS[i % MONO_STEPS.length] };
  const base: Color = { scheme: pal.order[i % 6] };
  const round = Math.floor(i / 6);
  if (round === 1) return { ...base, lumMod: 0.6 };
  if (round >= 2) return { ...base, lumMod: 0.8, lumOff: 0.2 };
  return base;
}

export const CHART_TYPES: { type: ChartType; label: string }[] = [
  { type: "col", label: "Sütun" },
  { type: "line", label: "Çizgi" },
  { type: "pie", label: "Pasta" },
  { type: "bar", label: "Çubuk" },
  { type: "area", label: "Alan" },
  { type: "scatter", label: "X Y (Dağılım)" },
  { type: "radar", label: "Radar" },
  { type: "combo", label: "Birleşik" },
  { type: "doughnut", label: "Halka" },
];

/** Gruplandırma alt türleri (Grafik Ekle iletişim kutusunun üst satırı). */
export function chartSubtypes(t: ChartType): { grouping: ChartGrouping; label: string }[] {
  const base: Partial<Record<ChartType, string>> = { col: "Sütun", bar: "Çubuk", line: "Çizgi", area: "Alan" };
  const b = base[t];
  if (!b) return [];
  return [
    { grouping: "clustered", label: t === "line" || t === "area" ? b : `Kümelenmiş ${b}` },
    { grouping: "stacked", label: `Yığılmış ${b}` },
    { grouping: "percentStacked", label: `%100 Yığılmış ${b}` },
  ];
}

export function chartTypeLabel(t: ChartType): string {
  return CHART_TYPES.find((c) => c.type === t)?.label ?? "Grafik";
}

export const isPieType = (t: ChartType) => t === "pie" || t === "doughnut";
export const hasAxes = (t: ChartType) => !isPieType(t) && t !== "radar";

/** Satır/Sütun Değiştir: kategoriler ile seriler yer değiştirir. */
export function switchRowColumn(c: ChartData): ChartData {
  const categories = c.series.map((s) => s.name);
  const series = c.categories.map((name, i) => ({ name, values: c.series.map((s) => s.values[i] ?? null) }));
  return { ...c, categories, series: series.length ? series : [{ name: "Seri 1", values: categories.map(() => null) }] };
}

/** Grafik Stilleri (1–8): kılavuz, veri etiketi, arka plan ve çubuk aralığı ön ayarları. */
export interface ChartStylePreset {
  bg: "none" | "light" | "dark";
  outline: boolean;
  gap: number;
  labels: boolean;
  grid: boolean;
}

export const CHART_STYLES: ChartStylePreset[] = [
  { bg: "none", outline: false, gap: 219, labels: false, grid: true },
  { bg: "none", outline: false, gap: 100, labels: true, grid: false },
  { bg: "none", outline: true, gap: 150, labels: false, grid: true },
  { bg: "light", outline: false, gap: 219, labels: false, grid: true },
  { bg: "light", outline: true, gap: 60, labels: true, grid: false },
  { bg: "dark", outline: false, gap: 219, labels: false, grid: true },
  { bg: "dark", outline: false, gap: 80, labels: true, grid: false },
  { bg: "none", outline: true, gap: 40, labels: false, grid: false },
];

export function chartStyle(c: ChartData): ChartStylePreset {
  return CHART_STYLES[Math.max(0, Math.min(CHART_STYLES.length - 1, (c.style ?? 1) - 1))];
}

/** Hızlı Düzen ön ayarları (başlık, gösterge, eksen başlıkları, etiketler, kılavuzlar). */
export const QUICK_LAYOUTS: { name: string; patch: (c: ChartData) => Partial<ChartData> }[] = [
  { name: "Düzen 1", patch: () => ({ showTitle: true, legend: "b", axisTitles: undefined, dataLabels: undefined, gridMajor: undefined }) },
  { name: "Düzen 2", patch: () => ({ showTitle: true, legend: "t", dataLabels: true, gridMajor: false, axisTitles: undefined }) },
  { name: "Düzen 3", patch: () => ({ showTitle: true, legend: "b", axisTitles: undefined, dataLabels: undefined, gridMajor: true, gridMinor: false }) },
  { name: "Düzen 4", patch: () => ({ showTitle: false, legend: "b", dataLabels: true, gridMajor: false }) },
  { name: "Düzen 5", patch: () => ({ showTitle: true, legend: "none", axisTitles: { val: "Eksen Başlığı" }, dataLabels: undefined, gridMajor: true }) },
  { name: "Düzen 6", patch: () => ({ showTitle: true, legend: "r", axisTitles: { cat: "Eksen Başlığı", val: "Eksen Başlığı" }, gridMajor: true }) },
  { name: "Düzen 7", patch: () => ({ showTitle: false, legend: "r", axisTitles: { cat: "Eksen Başlığı", val: "Eksen Başlığı" }, gridMajor: true, gridMinor: true }) },
  { name: "Düzen 8", patch: () => ({ showTitle: true, legend: "none", axisTitles: undefined, dataLabels: true, gridMajor: false }) },
];

/** Yığılmış türlerde her kategorideki serilerin alt/üst değerleri (yüzde ise 0–100). */
export function stackedValues(c: ChartData): { lo: number; hi: number }[][] {
  const pct = c.grouping === "percentStacked";
  const n = c.categories.length;
  const pos = new Array(n).fill(0);
  const neg = new Array(n).fill(0);
  const tot = Array.from({ length: n }, (_, i) => c.series.reduce((t, s) => t + Math.abs(s.values[i] ?? 0), 0) || 1);
  return c.series.map((s) =>
    Array.from({ length: n }, (_, i) => {
      let v = s.values[i] ?? 0;
      if (!Number.isFinite(v)) v = 0;
      if (pct) v = (v / tot[i]) * 100;
      if (v >= 0) {
        const lo = pos[i];
        pos[i] += v;
        return { lo, hi: pos[i] };
      }
      const hi = neg[i];
      neg[i] += v;
      return { lo: neg[i], hi };
    }),
  );
}

/** Dağılım grafiğinde X değerleri (kategoriler sayıysa onlar, değilse 1, 2, 3…). */
export function scatterX(c: ChartData): number[] {
  const nums = c.categories.map((t) => parseChartNumber(t));
  return nums.every((v) => v !== null) ? (nums as number[]) : c.categories.map((_, i) => i + 1);
}

/** Yeni grafiğin örnek verisi (PowerPoint'in varsayılan tablosuna benzer). */
export function defaultChartData(type: ChartType, grouping?: ChartGrouping): ChartData {
  if (type === "scatter") {
    return { type, showTitle: true, categories: ["0,7", "1,8", "2,6"], series: [{ name: "Y Değerleri", values: [2.7, 3.2, 0.8] }], legend: "b" };
  }
  if (type === "pie" || type === "doughnut") {
    return {
      type,
      showTitle: true,
      categories: ["1. Çeyrek", "2. Çeyrek", "3. Çeyrek", "4. Çeyrek"],
      series: [{ name: "Satışlar", values: [8.2, 3.2, 1.4, 1.2] }],
      legend: "b",
    };
  }
  return {
    type,
    showTitle: true,
    categories: ["Kategori 1", "Kategori 2", "Kategori 3", "Kategori 4"],
    series: [
      { name: "Seri 1", values: [4.3, 2.5, 3.5, 4.5] },
      { name: "Seri 2", values: [2.4, 4.4, 1.8, 2.8] },
      { name: "Seri 3", values: [2, 2, 3, 5] },
    ],
    legend: "b",
    grouping: grouping && grouping !== "clustered" ? grouping : undefined,
  };
}

export function makeChart(type: ChartType, x: number, y: number, w: number, h: number, name = "Grafik", grouping?: ChartGrouping): ChartShape {
  return { kind: "chart", id: uid(), name, x, y, w, h, rot: 0, chart: defaultChartData(type, grouping) };
}

/** Görüntülenecek başlık (boşsa tek seride seri adı). */
export function chartTitle(c: ChartData): string {
  if (c.title) return c.title;
  if (c.series.length === 1 || c.type === "pie" || c.type === "doughnut") return c.series[0]?.name ?? "Grafik Başlığı";
  return "Grafik Başlığı";
}

const NF = new Intl.NumberFormat("tr-TR", { maximumFractionDigits: 2 });

export function formatChartNumber(v: number): string {
  return NF.format(Math.abs(v) < 1e-12 ? 0 : v);
}

/** Değer ekseni için "güzel" aralık ve adım. */
export function niceScale(min: number, max: number, maxTicks = 8): { min: number; max: number; step: number } {
  let lo = Math.min(0, min);
  let hi = Math.max(0, max);
  if (lo === hi) hi = lo + 1;
  const range = hi - lo;
  const rough = range / Math.max(2, maxTicks - 1);
  const mag = 10 ** Math.floor(Math.log10(rough));
  const norm = rough / mag;
  const step = (norm <= 1 ? 1 : norm <= 2 ? 2 : norm <= 2.5 ? 2.5 : norm <= 5 ? 5 : 10) * mag;
  lo = Math.floor(lo / step + 1e-9) * step;
  hi = Math.ceil(hi / step - 1e-9) * step;
  if (hi === max && max > 0) hi += step;
  return { min: lo, max: hi, step };
}

/** Veri dizisinin min/max değerleri (boş değerler atlanır). */
export function valueRange(c: ChartData): { min: number; max: number } {
  if (c.grouping && c.grouping !== "clustered" && hasAxes(c.type) && c.type !== "scatter") {
    if (c.grouping === "percentStacked") return { min: 0, max: 100 };
    const st = stackedValues(c);
    let lo = 0;
    let hi = 0;
    for (const s of st) for (const p of s) {
      lo = Math.min(lo, p.lo);
      hi = Math.max(hi, p.hi);
    }
    return { min: lo, max: hi || 1 };
  }
  let min = Infinity;
  let max = -Infinity;
  for (const s of c.series) for (const v of s.values) if (v !== null && Number.isFinite(v)) {
    min = Math.min(min, v);
    max = Math.max(max, v);
  }
  if (!Number.isFinite(min)) return { min: 0, max: 1 };
  return { min, max };
}

/** Tablo düzenleyicisinden gelen ızgarayı (ilk satır seri adları, ilk sütun kategoriler) grafiğe uygular. */
export function chartFromGrid(base: ChartData, grid: string[][]): ChartData {
  const header = grid[0] ?? [];
  const rows = grid.slice(1).filter((r) => r.some((c) => c.trim() !== ""));
  const series = header.slice(1).map((name, j) => ({
    name: name || `Seri ${j + 1}`,
    values: rows.map((r) => parseChartNumber(r[j + 1] ?? "")),
    color: base.series[j]?.color,
  }));
  return { ...base, categories: rows.map((r, i) => r[0] || `Kategori ${i + 1}`), series: series.filter((s, j) => header[j + 1] !== undefined && (s.name || s.values.some((v) => v !== null))) };
}

export function chartToGrid(c: ChartData): string[][] {
  const head = ["", ...c.series.map((s) => s.name)];
  const rows = c.categories.map((cat, i) => [cat, ...c.series.map((s) => (s.values[i] === null || s.values[i] === undefined ? "" : String(s.values[i]).replace(".", ",")))]);
  return [head, ...rows];
}

/** "1.234,5" / "1234.5" → sayı; boşsa null. */
export function parseChartNumber(s: string): number | null {
  const t = s.trim();
  if (!t) return null;
  let norm = t.replace(/\s/g, "");
  if (norm.includes(",")) norm = norm.replace(/\./g, "").replace(",", ".");
  const n = Number(norm);
  return Number.isFinite(n) ? n : null;
}
