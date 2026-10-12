/**
 * Mini grafik (sparkline) verisi ve geometrisi: veri aralığından değerler,
 * eksen sınırları ve çizim ilkelleri. Tuval ve SVG çizicileri aynı
 * ilkelleri kullanır (ekran, yazdırma).
 */
import { parseQualifiedRange } from "@officehere/formula";
import type { Rect, SparklineGroup } from "../model/types";
import type { SheetModel, WorkbookModel } from "../model/workbook";

export type SparkItem = SparklineGroup["items"][number];

/** Varsayılan işaretçi rengi (yüksek/düşük/ilk/son/negatif noktalar). */
export const POINT_RED = "#D00000";

/** Mini grafiğin veri aralığını sayfa + dikdörtgene çevirir (nitelenmemişse sahibi). */
export function sparkSource(wb: WorkbookModel, owner: SheetModel, item: SparkItem): { sheet: SheetModel; rect: Rect } | null {
  const q = parseQualifiedRange(item.range.replace(/^=/, ""));
  if (!q) return null;
  const sheet = q.sheet ? wb.sheetByName(q.sheet) : owner;
  return sheet ? { sheet, rect: q.range } : null;
}

/** Veri değerleri (boş hücreler grup ayarına göre boşluk ya da sıfır). Gizli satır/sütunlar atlanır. */
export function sparkValues(wb: WorkbookModel, owner: SheetModel, g: SparklineGroup, item: SparkItem): (number | null)[] {
  const src = sparkSource(wb, owner, item);
  if (!src) return [];
  const { sheet, rect } = src;
  const out: (number | null)[] = [];
  const push = (r: number, c: number) => {
    if (!g.showHidden && (sheet.rows.size(r) === 0 || sheet.cols.size(c) === 0)) return;
    const v = wb.engine.getValue(sheet.id, r, c);
    if (typeof v === "number") out.push(v);
    else if (v === null || v === "") out.push(g.emptyAs === "zero" ? 0 : null);
    else out.push(null);
  };
  // Tek boyutlu aralık beklenir; iki boyutluysa satır satır okunur (en çok 4096 hücre)
  let n = 0;
  for (let r = rect.r1; r <= rect.r2 && n < 4096; r++) for (let c = rect.c1; c <= rect.c2 && n < 4096; c++, n++) push(r, c);
  return g.rightToLeft ? out.reverse() : out;
}

/** Grup ortak eksen sınırları (Tüm Mini Grafikler İçin Aynı seçiliyse). */
export function groupBounds(wb: WorkbookModel, owner: SheetModel, g: SparklineGroup): { min: number; max: number } | null {
  if (g.minAxis !== "group" && g.maxAxis !== "group") return null;
  let min = Infinity;
  let max = -Infinity;
  for (const it of g.items)
    for (const v of sparkValues(wb, owner, g, it)) {
      if (v === null) continue;
      if (v < min) min = v;
      if (v > max) max = v;
    }
  return Number.isFinite(min) ? { min, max } : null;
}

export type Prim =
  | { k: "path"; pts: [number, number][]; color: string; width: number }
  | { k: "rect"; x: number; y: number; w: number; h: number; color: string }
  | { k: "dot"; x: number; y: number; r: number; color: string }
  | { k: "hline"; x1: number; x2: number; y: number; color: string; width: number };

/** Noktanın vurgu rengi (yüksek/düşük/ilk/son/negatif); yoksa null. */
function pointColor(g: SparklineGroup, v: number, i: number, first: number, last: number, hi: number, lo: number): string | null {
  if (g.lastPoint && i === last) return g.lastColor ?? POINT_RED;
  if (g.firstPoint && i === first) return g.firstColor ?? POINT_RED;
  if (g.highPoint && v === hi) return g.highColor ?? POINT_RED;
  if (g.lowPoint && v === lo) return g.lowColor ?? POINT_RED;
  if (g.negative && v < 0) return g.negativeColor ?? POINT_RED;
  return null;
}

/**
 * Hücre kutusu (x, y, w, h; CSS pikseli) içinde mini grafik ilkelleri.
 * z: yakınlaştırma oranı; bounds: grup ortak eksen sınırları.
 */
export function sparkPrims(g: SparklineGroup, values: (number | null)[], box: { x: number; y: number; w: number; h: number }, z: number, bounds: { min: number; max: number } | null): Prim[] {
  const n = values.length;
  const nums = values.filter((v): v is number => v !== null);
  if (!n || !nums.length) return [];
  let lo = Math.min(...nums);
  let hi = Math.max(...nums);
  const dataLo = lo;
  const dataHi = hi;
  if (bounds && g.minAxis === "group") lo = bounds.min;
  if (bounds && g.maxAxis === "group") hi = bounds.max;
  if (g.minAxis === "custom" && g.minValue !== undefined) lo = g.minValue;
  if (g.maxAxis === "custom" && g.maxValue !== undefined) hi = g.maxValue;
  if (hi < lo) hi = lo;
  const first = values.findIndex((v) => v !== null);
  let last = n - 1;
  while (last > 0 && values[last] === null) last--;
  const out: Prim[] = [];
  const padX = Math.max(2, 3 * z);
  const padY = Math.max(2, 3 * z);
  const x0 = box.x + padX;
  const w = Math.max(1, box.w - 2 * padX);
  const y0 = box.y + padY;
  const h = Math.max(1, box.h - 2 * padY);
  const axisColor = g.axisColor ?? "#000000";

  if (g.type === "line") {
    const lw = Math.max(0.75, ((g.weight ?? 0.75) * 96) / 72) * z;
    const xs = (i: number) => (n === 1 ? x0 + w / 2 : x0 + (i * w) / (n - 1));
    const ys = (v: number) => (hi === lo ? y0 + h / 2 : y0 + ((hi - Math.min(hi, Math.max(lo, v))) / (hi - lo)) * h);
    if (g.showAxis && lo < 0 && hi > 0) out.push({ k: "hline", x1: box.x + 1, x2: box.x + box.w - 1, y: ys(0), color: axisColor, width: Math.max(1, z) });
    let seg: [number, number][] = [];
    const flush = () => {
      if (seg.length > 1) out.push({ k: "path", pts: seg, color: g.color, width: lw });
      seg = [];
    };
    values.forEach((v, i) => {
      if (v === null) {
        if (g.emptyAs !== "span") flush();
        return;
      }
      seg.push([xs(i), ys(v)]);
    });
    flush();
    // İşaretçi çizgiden kalın olsun
    const r = Math.max(1.5, 2 * z, lw * 0.75 + 0.75 * z);
    values.forEach((v, i) => {
      if (v === null) return;
      const c = pointColor(g, v, i, first, last, dataHi, dataLo) ?? (g.markers ? (g.markerColor ?? g.color) : null);
      if (c) out.push({ k: "dot", x: xs(i), y: ys(v), r, color: c });
    });
    // Tek nokta ve işaretçi yoksa görünsün diye küçük nokta
    if (nums.length === 1 && !out.some((p) => p.k === "dot")) out.push({ k: "dot", x: xs(first), y: ys(nums[0]), r: lw, color: g.color });
    return out;
  }

  const slot = w / n;
  const gap = Math.max(1, slot * 0.18);
  const bw = Math.max(1, slot - gap);
  const bx = (i: number) => x0 + i * slot + (slot - bw) / 2;
  const color = (v: number, i: number) => pointColor(g, v, i, first, last, dataHi, dataLo) ?? g.color;

  if (g.type === "winLoss") {
    const mid = y0 + h / 2;
    const half = Math.max(1, h / 2 - Math.max(0.5, z / 2));
    values.forEach((v, i) => {
      if (v === null || v === 0) return;
      const c = v < 0 ? (pointColor(g, v, i, first, last, dataHi, dataLo) ?? g.negativeColor ?? g.color) : color(v, i);
      out.push(v > 0 ? { k: "rect", x: bx(i), y: mid - half, w: bw, h: half, color: c } : { k: "rect", x: bx(i), y: mid + Math.max(0.5, z / 2), w: bw, h: half, color: c });
    });
    if (g.showAxis) out.push({ k: "hline", x1: box.x + 1, x2: box.x + box.w - 1, y: mid, color: axisColor, width: Math.max(1, z) });
    return out;
  }

  // Sütun: karışık işaretlerde sıfır taban; hepsi pozitifse taban en küçük değer (en küçük sütun ince görünür)
  const minH = Math.max(1, z);
  const base = lo >= 0 ? lo : hi <= 0 ? hi : 0;
  const span = hi - lo || 1;
  const yOf = (v: number) => y0 + ((hi - Math.min(hi, Math.max(lo, v))) / span) * h;
  values.forEach((v, i) => {
    if (v === null) return;
    let top = Math.min(yOf(v), yOf(base));
    let bh = Math.abs(yOf(v) - yOf(base));
    if (hi === lo) {
      top = y0;
      bh = h;
    }
    if (bh < minH) {
      bh = minH;
      if (v < base || (base === hi && hi < 0)) top = yOf(base);
      else top = yOf(base) - minH;
    }
    out.push({ k: "rect", x: bx(i), y: top, w: bw, h: bh, color: color(v, i) });
  });
  if (g.showAxis && lo < 0 && hi > 0) out.push({ k: "hline", x1: box.x + 1, x2: box.x + box.w - 1, y: yOf(0), color: axisColor, width: Math.max(1, z) });
  return out;
}

/** İlkelleri tuvale çizer. */
export function drawPrims(ctx: CanvasRenderingContext2D, prims: Prim[]): void {
  for (const p of prims) {
    if (p.k === "path") {
      ctx.strokeStyle = p.color;
      ctx.lineWidth = p.width;
      ctx.lineJoin = "round";
      ctx.lineCap = "round";
      ctx.beginPath();
      p.pts.forEach(([x, y], i) => (i ? ctx.lineTo(x, y) : ctx.moveTo(x, y)));
      ctx.stroke();
    } else if (p.k === "rect") {
      ctx.fillStyle = p.color;
      ctx.fillRect(p.x, p.y, p.w, p.h);
    } else if (p.k === "dot") {
      ctx.fillStyle = p.color;
      ctx.beginPath();
      ctx.arc(p.x, p.y, p.r, 0, Math.PI * 2);
      ctx.fill();
    } else {
      ctx.fillStyle = p.color;
      ctx.fillRect(p.x1, p.y - p.width / 2, p.x2 - p.x1, p.width);
    }
  }
}

const f = (n: number) => String(Math.round(n * 100) / 100);

/** İlkelleri SVG belgesine çevirir (yazdırma ve HTML dışa aktarma). */
export function primsToSvg(prims: Prim[], w: number, h: number): string {
  const body = prims
    .map((p) => {
      if (p.k === "path") return `<polyline points="${p.pts.map(([x, y]) => `${f(x)},${f(y)}`).join(" ")}" fill="none" stroke="${p.color}" stroke-width="${f(p.width)}" stroke-linejoin="round" stroke-linecap="round"/>`;
      if (p.k === "rect") return `<rect x="${f(p.x)}" y="${f(p.y)}" width="${f(p.w)}" height="${f(p.h)}" fill="${p.color}"/>`;
      if (p.k === "dot") return `<circle cx="${f(p.x)}" cy="${f(p.y)}" r="${f(p.r)}" fill="${p.color}"/>`;
      return `<rect x="${f(p.x1)}" y="${f(p.y - p.width / 2)}" width="${f(p.x2 - p.x1)}" height="${f(p.width)}" fill="${p.color}"/>`;
    })
    .join("");
  return `<svg xmlns="http://www.w3.org/2000/svg" width="${f(w)}" height="${f(h)}" viewBox="0 0 ${f(w)} ${f(h)}">${body}</svg>`;
}

/** Hücredeki mini grafik (varsa) ve grubu. */
export function sparkAt(sheet: SheetModel, row: number, col: number): { group: SparklineGroup; item: SparkItem } | null {
  for (const g of sheet.sparklines) {
    const item = g.items.find((it) => it.row === row && it.col === col);
    if (item) return { group: g, item };
  }
  return null;
}

/** Hücre için SVG (yazdırma tablosunda arka plan olarak). */
export function sparkSvgForCell(wb: WorkbookModel, sheet: SheetModel, row: number, col: number, w: number, h: number): string | null {
  const s = sparkAt(sheet, row, col);
  if (!s) return null;
  const prims = sparkPrims(s.group, sparkValues(wb, sheet, s.group, s.item), { x: 0, y: 0, w, h }, 1, groupBounds(wb, sheet, s.group));
  return prims.length ? primsToSvg(prims, w, h) : null;
}
