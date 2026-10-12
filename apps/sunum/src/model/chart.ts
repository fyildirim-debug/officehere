/** Grafik verisi yardımcıları: varsayılan veri, eksen ölçeği, sayı biçimi. */
import { uid } from "./factory";
import type { ChartData, ChartShape, ChartType, Color } from "./types";

const ACCENTS = ["accent1", "accent2", "accent3", "accent4", "accent5", "accent6"] as const;

/** i. seri/dilim rengi (6'dan sonrası koyulaştırılmış tonlar). */
export function seriesColor(i: number, explicit?: Color): Color {
  if (explicit) return explicit;
  const base: Color = { scheme: ACCENTS[i % 6] };
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
  { type: "doughnut", label: "Halka" },
];

export function chartTypeLabel(t: ChartType): string {
  return CHART_TYPES.find((c) => c.type === t)?.label ?? "Grafik";
}

/** Yeni grafiğin örnek verisi (PowerPoint'in varsayılan tablosuna benzer). */
export function defaultChartData(type: ChartType): ChartData {
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
  };
}

export function makeChart(type: ChartType, x: number, y: number, w: number, h: number, name = "Grafik"): ChartShape {
  return { kind: "chart", id: uid(), name, x, y, w, h, rot: 0, chart: defaultChartData(type) };
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
