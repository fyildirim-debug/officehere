/**
 * Hazır şekil geometrileri. Adlar ECMA-376 prstGeom değerleridir; yollar bu
 * proje için yazılmıştır ve ayar (adj) değerleri standarttaki birimleri
 * (çoğunlukla kısa kenarın 100000'de biri) kullanır.
 */
import type { CustomPath } from "../model/types";

export type FillMode = "norm" | "darken" | "darkenLess" | "lighten" | "lightenLess" | "none";

export interface GeomPath {
  d: string;
  fill: FillMode;
  stroke: boolean;
  evenodd?: boolean;
}

export interface GeomResult {
  paths: GeomPath[];
  /** Metin dikdörtgeni (şekil koordinatlarında). */
  text: { x: number; y: number; w: number; h: number };
  /** Açık yol (çizgi) — dolgu yok. */
  open?: boolean;
}

const f = (n: number) => (Math.round(n * 100) / 100).toString();
const pt = (x: number, y: number) => `${f(x)} ${f(y)}`;

function poly(points: [number, number][], close = true): string {
  return points.map((p, i) => `${i ? "L" : "M"}${pt(p[0], p[1])}`).join(" ") + (close ? " Z" : "");
}

function ellipsePath(cx: number, cy: number, rx: number, ry: number, ccw = false): string {
  const s = ccw ? 0 : 1;
  return `M${pt(cx - rx, cy)} A${f(rx)} ${f(ry)} 0 1 ${s} ${pt(cx + rx, cy)} A${f(rx)} ${f(ry)} 0 1 ${s} ${pt(cx - rx, cy)} Z`;
}

function roundRectPath(x: number, y: number, w: number, h: number, r: number, corners = [true, true, true, true]): string {
  r = Math.max(0, Math.min(r, w / 2, h / 2));
  const [tl, tr, br, bl] = corners.map((c) => (c ? r : 0));
  return [
    `M${pt(x + tl, y)}`,
    `L${pt(x + w - tr, y)}`,
    tr ? `A${f(tr)} ${f(tr)} 0 0 1 ${pt(x + w, y + tr)}` : "",
    `L${pt(x + w, y + h - br)}`,
    br ? `A${f(br)} ${f(br)} 0 0 1 ${pt(x + w - br, y + h)}` : "",
    `L${pt(x + bl, y + h)}`,
    bl ? `A${f(bl)} ${f(bl)} 0 0 1 ${pt(x, y + h - bl)}` : "",
    `L${pt(x, y + tl)}`,
    tl ? `A${f(tl)} ${f(tl)} 0 0 1 ${pt(x + tl, y)}` : "",
    "Z",
  ].join(" ");
}

/** Düzgün çokgen/yıldız; dış köşeler kutuyu dolduracak biçimde ölçeklenir. */
function starPoints(n: number, inner: number | null, w: number, h: number, startDeg = -90): [number, number][] {
  const pts: [number, number][] = [];
  const total = inner === null ? n : n * 2;
  for (let i = 0; i < total; i++) {
    const r = inner === null || i % 2 === 0 ? 1 : inner;
    const a = ((startDeg + (360 / total) * i) * Math.PI) / 180;
    pts.push([Math.cos(a) * r, Math.sin(a) * r]);
  }
  const outer = inner === null ? pts : pts.filter((_, i) => i % 2 === 0);
  const minX = Math.min(...outer.map((p) => p[0]));
  const maxX = Math.max(...outer.map((p) => p[0]));
  const minY = Math.min(...outer.map((p) => p[1]));
  const maxY = Math.max(...outer.map((p) => p[1]));
  return pts.map(([x, y]) => [((x - minX) / (maxX - minX)) * w, ((y - minY) / (maxY - minY)) * h]);
}

const full = (w: number, h: number) => ({ x: 0, y: 0, w, h });
const inset = (w: number, h: number, l: number, t: number, r: number, b: number) => ({ x: w * l, y: h * t, w: w * (1 - l - r), h: h * (1 - t - b) });

function one(d: string, w: number, h: number, text = full(w, h), extra: Partial<GeomPath> = {}): GeomResult {
  return { paths: [{ d, fill: "norm", stroke: true, ...extra }], text };
}

function cloudPath(x: number, y: number, w: number, h: number): string {
  const n = 9;
  const cx = x + w / 2;
  const cy = y + h / 2;
  const rx = w * 0.4;
  const ry = h * 0.36;
  const pts: [number, number][] = [];
  for (let i = 0; i < n; i++) {
    const a = (i / n) * Math.PI * 2 - Math.PI / 2 + 0.25;
    pts.push([cx + Math.cos(a) * rx, cy + Math.sin(a) * ry]);
  }
  let d = `M${pt(pts[0][0], pts[0][1])}`;
  for (let i = 0; i < n; i++) {
    const p1 = pts[i];
    const p2 = pts[(i + 1) % n];
    const mx = (p1[0] + p2[0]) / 2;
    const my = (p1[1] + p2[1]) / 2;
    const dx = mx - cx;
    const dy = my - cy;
    const len = Math.hypot(dx / rx, dy / ry) || 1;
    const bulge = 0.55;
    const ox = (dx / len) * bulge + (dx / len) * 0.1;
    const oy = (dy / len) * bulge * (ry / rx) + (dy / len) * 0.1;
    const c1x = p1[0] + (mx - p1[0]) * 0.1 + ox * 0.9;
    const c1y = p1[1] + (my - p1[1]) * 0.1 + oy * 0.9;
    const c2x = p2[0] + (mx - p2[0]) * 0.1 + ox * 0.9;
    const c2y = p2[1] + (my - p2[1]) * 0.1 + oy * 0.9;
    d += ` C${pt(c1x, c1y)} ${pt(c2x, c2y)} ${pt(p2[0], p2[1])}`;
  }
  return d + " Z";
}

function calloutTail(w: number, h: number, tx: number, ty: number): { side: "t" | "b" | "l" | "r"; a: number; b: number } {
  const dx = (tx - w / 2) / (w / 2 || 1);
  const dy = (ty - h / 2) / (h / 2 || 1);
  if (Math.abs(dy) >= Math.abs(dx)) {
    const c = Math.max(w * 0.15, Math.min(w * 0.85, w / 2 + (tx - w / 2) * 0.4));
    const hw = Math.max(4, w * 0.09);
    return { side: dy > 0 ? "b" : "t", a: c - hw, b: c + hw };
  }
  const c = Math.max(h * 0.15, Math.min(h * 0.85, h / 2 + (ty - h / 2) * 0.4));
  const hw = Math.max(4, h * 0.09);
  return { side: dx > 0 ? "r" : "l", a: c - hw, b: c + hw };
}

function rectCallout(w: number, h: number, tx: number, ty: number, r: number): string {
  const inside = tx >= 0 && tx <= w && ty >= 0 && ty <= h;
  if (inside) return roundRectPath(0, 0, w, h, r);
  const t = calloutTail(w, h, tx, ty);
  r = Math.max(0, Math.min(r, w / 2, h / 2));
  const parts: string[] = [`M${pt(r, 0)}`];
  if (t.side === "t") parts.push(`L${pt(t.a, 0)} L${pt(tx, ty)} L${pt(t.b, 0)}`);
  parts.push(`L${pt(w - r, 0)}`);
  if (r) parts.push(`A${f(r)} ${f(r)} 0 0 1 ${pt(w, r)}`);
  if (t.side === "r") parts.push(`L${pt(w, t.a)} L${pt(tx, ty)} L${pt(w, t.b)}`);
  parts.push(`L${pt(w, h - r)}`);
  if (r) parts.push(`A${f(r)} ${f(r)} 0 0 1 ${pt(w - r, h)}`);
  if (t.side === "b") parts.push(`L${pt(t.b, h)} L${pt(tx, ty)} L${pt(t.a, h)}`);
  parts.push(`L${pt(r, h)}`);
  if (r) parts.push(`A${f(r)} ${f(r)} 0 0 1 ${pt(0, h - r)}`);
  if (t.side === "l") parts.push(`L${pt(0, t.b)} L${pt(tx, ty)} L${pt(0, t.a)}`);
  parts.push(`L${pt(0, r)}`);
  if (r) parts.push(`A${f(r)} ${f(r)} 0 0 1 ${pt(r, 0)}`);
  return parts.join(" ") + " Z";
}

/** Özel (custGeom) yolları şekil boyutuna ölçekleyerek SVG yoluna çevirir. */
export function customPathD(p: CustomPath, w: number, h: number): string {
  const sx = p.w ? w / p.w : 1;
  const sy = p.h ? h / p.h : 1;
  return p.cmds
    .map((c) => {
      const xs = c.pts.map((v, i) => f(i % 2 === 0 ? v * sx : v * sy));
      return c.c + (xs.length ? xs.join(" ") : "");
    })
    .join(" ");
}

export const DEFAULT_ADJ: Record<string, Record<string, number>> = {
  roundRect: { adj: 16667 },
  snip1Rect: { adj: 16667 },
  round2SameRect: { adj1: 16667, adj2: 0 },
  triangle: { adj: 50000 },
  parallelogram: { adj: 25000 },
  trapezoid: { adj: 25000 },
  hexagon: { adj: 25000 },
  octagon: { adj: 29289 },
  plus: { adj: 25000 },
  donut: { adj: 25000 },
  noSmoking: { adj: 18750 },
  can: { adj: 25000 },
  cube: { adj: 25000 },
  bevel: { adj: 12500 },
  foldedCorner: { adj: 16667 },
  frame: { adj1: 12500 },
  sun: { adj: 25000 },
  moon: { adj: 50000 },
  rightArrow: { adj1: 50000, adj2: 50000 },
  leftArrow: { adj1: 50000, adj2: 50000 },
  upArrow: { adj1: 50000, adj2: 50000 },
  downArrow: { adj1: 50000, adj2: 50000 },
  leftRightArrow: { adj1: 50000, adj2: 50000 },
  upDownArrow: { adj1: 50000, adj2: 50000 },
  quadArrow: { adj1: 22500, adj2: 22500, adj3: 22500 },
  homePlate: { adj: 50000 },
  chevron: { adj: 50000 },
  wedgeRectCallout: { adj1: -20833, adj2: 62500 },
  wedgeRoundRectCallout: { adj1: -20833, adj2: 62500, adj3: 16667 },
  wedgeEllipseCallout: { adj1: -20833, adj2: 62500 },
  cloudCallout: { adj1: -20833, adj2: 62500 },
  corner: { adj1: 50000, adj2: 50000 },
  pie: { adj1: 0, adj2: 16200000 },
  wave: { adj1: 12500, adj2: 0 },
  mathPlus: { adj1: 23520 },
  mathMinus: { adj1: 23520 },
  mathMultiply: { adj1: 23520 },
  mathEqual: { adj1: 23520, adj2: 11760 },
  bentConnector3: { adj1: 50000 },
  star4: { adj: 12500 },
  star5: { adj: 19098 },
  star6: { adj: 28868 },
  star8: { adj: 38250 },
  star10: { adj: 42533 },
  star12: { adj: 37500 },
};

export function presetGeometry(geom: string, w: number, h: number, adjIn?: Record<string, number>): GeomResult {
  const adj = { ...(DEFAULT_ADJ[geom] ?? {}), ...(adjIn ?? {}) };
  const ss = Math.min(w, h);
  const a = (k = "adj") => (adj[k] ?? 0) / 100000;
  switch (geom) {
    case "rect":
    case "flowChartProcess":
      return one(`M0 0 L${pt(w, 0)} L${pt(w, h)} L${pt(0, h)} Z`, w, h);
    case "roundRect": {
      const r = ss * a();
      return one(roundRectPath(0, 0, w, h, r), w, h, inset(w, h, (r * 0.29) / w, (r * 0.29) / h, (r * 0.29) / w, (r * 0.29) / h));
    }
    case "round2SameRect":
      return one(roundRectPath(0, 0, w, h, ss * a("adj1"), [true, true, false, false]), w, h);
    case "snip1Rect": {
      const c = ss * a();
      return one(poly([[0, 0], [w - c, 0], [w, c], [w, h], [0, h]]), w, h);
    }
    case "ellipse":
    case "flowChartConnector":
      return one(ellipsePath(w / 2, h / 2, w / 2, h / 2), w, h, inset(w, h, 0.146, 0.146, 0.146, 0.146));
    case "triangle": {
      const x = w * a();
      return one(poly([[x, 0], [w, h], [0, h]]), w, h, { x: x / 2, y: h / 2, w: w / 2, h: h / 2 });
    }
    case "rtTriangle":
      return one(poly([[0, 0], [w, h], [0, h]]), w, h, inset(w, h, 0.083, 0.5, 0.5, 0.083));
    case "parallelogram":
    case "flowChartInputOutput": {
      const o = geom === "parallelogram" ? ss * a() : w * 0.2;
      return one(poly([[o, 0], [w, 0], [w - o, h], [0, h]]), w, h, { x: o * 0.85, y: 0, w: w - o * 1.7, h });
    }
    case "trapezoid": {
      const o = ss * a();
      return one(poly([[0, h], [o, 0], [w - o, 0], [w, h]]), w, h, { x: o * 0.66, y: 0, w: w - o * 1.33, h });
    }
    case "diamond":
    case "flowChartDecision":
      return one(poly([[w / 2, 0], [w, h / 2], [w / 2, h], [0, h / 2]]), w, h, inset(w, h, 0.25, 0.25, 0.25, 0.25));
    case "pentagon":
      return one(poly(starPoints(5, null, w, h)), w, h, inset(w, h, 0.19, 0.24, 0.19, 0));
    case "hexagon": {
      const o = ss * a();
      return one(poly([[o, 0], [w - o, 0], [w, h / 2], [w - o, h], [o, h], [0, h / 2]]), w, h, { x: o * 0.6, y: 0, w: w - o * 1.2, h });
    }
    case "heptagon":
      return one(poly(starPoints(7, null, w, h)), w, h, inset(w, h, 0.15, 0.15, 0.15, 0.15));
    case "octagon": {
      const o = ss * a();
      return one(poly([[o, 0], [w - o, 0], [w, o], [w, h - o], [w - o, h], [o, h], [0, h - o], [0, o]]), w, h, inset(w, h, (o * 0.5) / w, (o * 0.5) / h, (o * 0.5) / w, (o * 0.5) / h));
    }
    case "decagon":
      return one(poly(starPoints(10, null, w, h, -90 + 18)), w, h, inset(w, h, 0.1, 0.1, 0.1, 0.1));
    case "dodecagon":
      return one(poly(starPoints(12, null, w, h, -90 + 15)), w, h, inset(w, h, 0.12, 0.12, 0.12, 0.12));
    case "star4":
    case "star5":
    case "star6":
    case "star8":
    case "star10":
    case "star12": {
      const n = parseInt(geom.slice(4), 10);
      const inner = a() * 2;
      const t = n === 4 ? inset(w, h, 0.3, 0.3, 0.3, 0.3) : n === 5 ? inset(w, h, 0.31, 0.38, 0.31, 0.2) : inset(w, h, 0.22, 0.22, 0.22, 0.22);
      return one(poly(starPoints(n, inner, w, h)), w, h, t);
    }
    case "irregularSeal1": {
      const pts: [number, number][] = [
        [0.39, 0.0], [0.49, 0.26], [0.67, 0.05], [0.65, 0.3], [0.97, 0.22], [0.76, 0.42], [1.0, 0.56], [0.72, 0.6], [0.83, 0.86], [0.57, 0.69],
        [0.5, 1.0], [0.41, 0.73], [0.18, 0.92], [0.24, 0.66], [0.0, 0.68], [0.17, 0.5], [0.02, 0.32], [0.25, 0.33], [0.2, 0.08], [0.38, 0.24],
      ];
      return one(poly(pts.map(([x, y]) => [x * w, y * h] as [number, number])), w, h, inset(w, h, 0.25, 0.3, 0.25, 0.3));
    }
    case "plus": {
      const o = ss * a();
      return one(poly([[o, 0], [w - o, 0], [w - o, o], [w, o], [w, h - o], [w - o, h - o], [w - o, h], [o, h], [o, h - o], [0, h - o], [0, o], [o, o]]), w, h, { x: 0, y: o, w, h: h - 2 * o });
    }
    case "donut": {
      const t = ss * a();
      return one(ellipsePath(w / 2, h / 2, w / 2, h / 2) + " " + ellipsePath(w / 2, h / 2, Math.max(0, w / 2 - t), Math.max(0, h / 2 - t), true), w, h, inset(w, h, 0.146, 0.146, 0.146, 0.146), { evenodd: true });
    }
    case "noSmoking": {
      const t = ss * a();
      const rx = w / 2 - t;
      const ry = h / 2 - t;
      const d1 = ellipsePath(w / 2, h / 2, w / 2, h / 2) + " " + ellipsePath(w / 2, h / 2, rx, ry, true);
      const k = 0.7071;
      const bar = poly([
        [w / 2 - rx * k - t * 0.5, h / 2 - ry * k + t * 0.5],
        [w / 2 - rx * k + t * 0.5, h / 2 - ry * k - t * 0.5],
        [w / 2 + rx * k + t * 0.5, h / 2 + ry * k - t * 0.5],
        [w / 2 + rx * k - t * 0.5, h / 2 + ry * k + t * 0.5],
      ]);
      return { paths: [{ d: d1 + " " + bar, fill: "norm", stroke: true }], text: inset(w, h, 0.146, 0.146, 0.146, 0.146) };
    }
    case "can": {
      const ry = (ss * a()) / 2;
      const body = `M0 ${f(ry)} A${f(w / 2)} ${f(ry)} 0 0 0 ${pt(w, ry)} L${pt(w, h - ry)} A${f(w / 2)} ${f(ry)} 0 0 1 ${pt(0, h - ry)} Z`;
      const top = ellipsePath(w / 2, ry, w / 2, ry);
      return { paths: [{ d: body, fill: "norm", stroke: true }, { d: top, fill: "lighten", stroke: true }], text: { x: 0, y: ry * 2, w, h: h - ry * 3 } };
    }
    case "cube": {
      const d = ss * a();
      return {
        paths: [
          { d: poly([[0, d], [w - d, d], [w - d, h], [0, h]]), fill: "norm", stroke: true },
          { d: poly([[0, d], [d, 0], [w, 0], [w - d, d]]), fill: "lightenLess", stroke: true },
          { d: poly([[w - d, d], [w, 0], [w, h - d], [w - d, h]]), fill: "darkenLess", stroke: true },
        ],
        text: { x: 0, y: d, w: w - d, h: h - d },
      };
    }
    case "bevel": {
      const d = ss * a();
      return {
        paths: [
          { d: poly([[d, d], [w - d, d], [w - d, h - d], [d, h - d]]), fill: "norm", stroke: false },
          { d: poly([[0, 0], [w, 0], [w - d, d], [d, d]]), fill: "lightenLess", stroke: false },
          { d: poly([[w, 0], [w, h], [w - d, h - d], [w - d, d]]), fill: "darkenLess", stroke: false },
          { d: poly([[0, h], [d, h - d], [w - d, h - d], [w, h]]), fill: "darken", stroke: false },
          { d: poly([[0, 0], [d, d], [d, h - d], [0, h]]), fill: "lighten", stroke: false },
          { d: poly([[0, 0], [w, 0], [w, h], [0, h]]) + " " + poly([[d, d], [w - d, d], [w - d, h - d], [d, h - d]]), fill: "none", stroke: true },
        ],
        text: { x: d, y: d, w: w - 2 * d, h: h - 2 * d },
      };
    }
    case "foldedCorner": {
      const c = ss * a();
      return {
        paths: [
          { d: poly([[0, 0], [w, 0], [w, h - c], [w - c, h], [0, h]]), fill: "norm", stroke: true },
          { d: poly([[w - c, h], [w - c * 0.8, h - c * 0.8], [w, h - c]]), fill: "darkenLess", stroke: true },
        ],
        text: { x: 0, y: 0, w, h: h - c },
      };
    }
    case "frame": {
      const t = ss * a("adj1");
      return one(poly([[0, 0], [w, 0], [w, h], [0, h]]) + " " + poly([[t, t], [t, h - t], [w - t, h - t], [w - t, t]]), w, h, { x: t, y: t, w: w - 2 * t, h: h - 2 * t }, { evenodd: true });
    }
    case "corner": {
      const t1 = ss * a("adj1");
      const t2 = ss * a("adj2");
      return one(poly([[0, 0], [t2, 0], [t2, h - t1], [w, h - t1], [w, h], [0, h]]), w, h, { x: 0, y: h - t1, w, h: t1 });
    }
    case "heart": {
      const d = `M${pt(w / 2, h * 0.25)} C${pt(w * 0.5, h * 0.12)} ${pt(w * 0.38, 0)} ${pt(w * 0.25, 0)} C${pt(w * 0.08, 0)} ${pt(0, h * 0.14)} ${pt(0, h * 0.3)} C${pt(0, h * 0.58)} ${pt(w * 0.3, h * 0.75)} ${pt(w / 2, h)} C${pt(w * 0.7, h * 0.75)} ${pt(w, h * 0.58)} ${pt(w, h * 0.3)} C${pt(w, h * 0.14)} ${pt(w * 0.92, 0)} ${pt(w * 0.75, 0)} C${pt(w * 0.62, 0)} ${pt(w * 0.5, h * 0.12)} ${pt(w / 2, h * 0.25)} Z`;
      return one(d, w, h, inset(w, h, 0.2, 0.2, 0.2, 0.3));
    }
    case "lightningBolt": {
      const pts: [number, number][] = [
        [0.38, 0], [0.62, 0.28], [0.54, 0.32], [0.78, 0.55], [0.7, 0.59], [1, 1], [0.36, 0.67], [0.46, 0.62], [0.14, 0.44], [0.24, 0.39], [0, 0.15],
      ];
      return one(poly(pts.map(([x, y]) => [x * w, y * h] as [number, number])), w, h, inset(w, h, 0.3, 0.3, 0.3, 0.3));
    }
    case "sun": {
      const r = 0.5 - a() * 0.9;
      const parts: string[] = [ellipsePath(w / 2, h / 2, w * r * 0.92, h * r * 0.92)];
      for (let i = 0; i < 8; i++) {
        const ang = (i * Math.PI) / 4;
        const tip = [w / 2 + Math.cos(ang) * w * 0.5, h / 2 + Math.sin(ang) * h * 0.5];
        const b1 = [w / 2 + Math.cos(ang - 0.17) * w * (r + 0.06), h / 2 + Math.sin(ang - 0.17) * h * (r + 0.06)];
        const b2 = [w / 2 + Math.cos(ang + 0.17) * w * (r + 0.06), h / 2 + Math.sin(ang + 0.17) * h * (r + 0.06)];
        parts.push(poly([tip as [number, number], b1 as [number, number], b2 as [number, number]]));
      }
      return one(parts.join(" "), w, h, inset(w, h, 0.3, 0.3, 0.3, 0.3));
    }
    case "moon": {
      const k = a();
      const d = `M${pt(w, 0)} A${f(w)} ${f(h / 2)} 0 0 0 ${pt(w, h)} A${f(w * (1 - k) + 0.01)} ${f(h / 2)} 0 0 1 ${pt(w, 0)} Z`;
      return one(d, w, h, inset(w, h, 0.1, 0.25, 0.5, 0.25));
    }
    case "smileyFace": {
      const face = ellipsePath(w / 2, h / 2, w / 2, h / 2);
      const eyes = ellipsePath(w * 0.34, h * 0.37, w * 0.055, h * 0.055) + " " + ellipsePath(w * 0.66, h * 0.37, w * 0.055, h * 0.055);
      const smile = `M${pt(w * 0.27, h * 0.65)} Q${pt(w / 2, h * 0.87)} ${pt(w * 0.73, h * 0.65)}`;
      return {
        paths: [
          { d: face, fill: "norm", stroke: true },
          { d: eyes, fill: "darkenLess", stroke: true },
          { d: smile, fill: "none", stroke: true },
        ],
        text: inset(w, h, 0.146, 0.146, 0.146, 0.146),
      };
    }
    case "cloud":
      return one(cloudPath(0, 0, w, h), w, h, inset(w, h, 0.18, 0.22, 0.18, 0.22));
    case "teardrop": {
      const d = `M${pt(w / 2, 0)} L${pt(w, 0)} L${pt(w, h / 2)} A${f(w / 2)} ${f(h / 2)} 0 1 1 ${pt(w / 2, 0)} Z`;
      return one(d, w, h, inset(w, h, 0.146, 0.146, 0.146, 0.146));
    }
    case "pie": {
      const s = (a("adj1") * 100000) / 60000;
      const e = (a("adj2") * 100000) / 60000;
      const sr = (s * Math.PI) / 180;
      const er = (e * Math.PI) / 180;
      const sweep = (((e - s) % 360) + 360) % 360;
      const x1 = w / 2 + (Math.cos(sr) * w) / 2;
      const y1 = h / 2 + (Math.sin(sr) * h) / 2;
      const x2 = w / 2 + (Math.cos(er) * w) / 2;
      const y2 = h / 2 + (Math.sin(er) * h) / 2;
      const d = `M${pt(w / 2, h / 2)} L${pt(x1, y1)} A${f(w / 2)} ${f(h / 2)} 0 ${sweep > 180 ? 1 : 0} 1 ${pt(x2, y2)} Z`;
      return one(d, w, h, inset(w, h, 0.146, 0.146, 0.146, 0.146));
    }
    case "wave": {
      const amp = h * a("adj1");
      const d = `M0 ${f(amp)} C${pt(w / 3, -amp)} ${pt((w * 2) / 3, amp * 3)} ${pt(w, amp)} L${pt(w, h - amp)} C${pt((w * 2) / 3, h + amp)} ${pt(w / 3, h - amp * 3)} ${pt(0, h - amp)} Z`;
      return one(d, w, h, { x: 0, y: amp * 2, w, h: h - amp * 4 });
    }
    case "flowChartTerminator":
      return one(roundRectPath(0, 0, w, h, h / 2), w, h, inset(w, h, 0.1, 0.15, 0.1, 0.15));
    case "flowChartDocument": {
      const d = `M0 0 L${pt(w, 0)} L${pt(w, h * 0.83)} C${pt(w * 0.75, h * 0.7)} ${pt(w * 0.55, h * 0.95)} ${pt(w * 0.35, h * 0.95)} C${pt(w * 0.15, h * 0.95)} ${pt(w * 0.05, h * 0.9)} ${pt(0, h * 0.87)} Z`;
      return one(d, w, h, { x: 0, y: 0, w, h: h * 0.8 });
    }
    case "flowChartPredefinedProcess": {
      const o = w * 0.11;
      return {
        paths: [
          { d: poly([[0, 0], [w, 0], [w, h], [0, h]]), fill: "norm", stroke: true },
          { d: `M${pt(o, 0)} L${pt(o, h)} M${pt(w - o, 0)} L${pt(w - o, h)}`, fill: "none", stroke: true },
        ],
        text: { x: o, y: 0, w: w - 2 * o, h },
      };
    }
    case "rightArrow":
    case "leftArrow": {
      const t = (h * a("adj1")) / 2;
      const hl = Math.min(w, ss * a("adj2"));
      let pts: [number, number][] = [[0, h / 2 - t], [w - hl, h / 2 - t], [w - hl, 0], [w, h / 2], [w - hl, h], [w - hl, h / 2 + t], [0, h / 2 + t]];
      if (geom === "leftArrow") pts = pts.map(([x, y]) => [w - x, y]);
      return one(poly(pts), w, h, { x: geom === "leftArrow" ? hl * 0.5 : 0, y: h / 2 - t, w: w - hl * 0.5, h: 2 * t });
    }
    case "upArrow":
    case "downArrow": {
      const t = (w * a("adj1")) / 2;
      const hl = Math.min(h, ss * a("adj2"));
      let pts: [number, number][] = [[w / 2 - t, h], [w / 2 - t, hl], [0, hl], [w / 2, 0], [w, hl], [w / 2 + t, hl], [w / 2 + t, h]];
      if (geom === "downArrow") pts = pts.map(([x, y]) => [x, h - y]);
      return one(poly(pts), w, h, { x: w / 2 - t, y: geom === "upArrow" ? hl * 0.5 : 0, w: 2 * t, h: h - hl * 0.5 });
    }
    case "leftRightArrow": {
      const t = (h * a("adj1")) / 2;
      const hl = Math.min(w / 2, ss * a("adj2"));
      return one(poly([[0, h / 2], [hl, 0], [hl, h / 2 - t], [w - hl, h / 2 - t], [w - hl, 0], [w, h / 2], [w - hl, h], [w - hl, h / 2 + t], [hl, h / 2 + t], [hl, h]]), w, h, { x: hl * 0.5, y: h / 2 - t, w: w - hl, h: 2 * t });
    }
    case "upDownArrow": {
      const t = (w * a("adj1")) / 2;
      const hl = Math.min(h / 2, ss * a("adj2"));
      return one(poly([[w / 2, 0], [w, hl], [w / 2 + t, hl], [w / 2 + t, h - hl], [w, h - hl], [w / 2, h], [0, h - hl], [w / 2 - t, h - hl], [w / 2 - t, hl], [0, hl]]), w, h, { x: w / 2 - t, y: hl * 0.5, w: 2 * t, h: h - hl });
    }
    case "quadArrow": {
      const t = (ss * a("adj1")) / 2;
      const hw = (ss * a("adj2")) / 2;
      const hl = ss * a("adj3");
      const cx = w / 2;
      const cy = h / 2;
      return one(
        poly([
          [cx, 0], [cx + hw, hl], [cx + t, hl], [cx + t, cy - t], [w - hl, cy - t], [w - hl, cy - hw], [w, cy], [w - hl, cy + hw], [w - hl, cy + t], [cx + t, cy + t], [cx + t, h - hl], [cx + hw, h - hl],
          [cx, h], [cx - hw, h - hl], [cx - t, h - hl], [cx - t, cy + t], [hl, cy + t], [hl, cy + hw], [0, cy], [hl, cy - hw], [hl, cy - t], [cx - t, cy - t], [cx - t, hl], [cx - hw, hl],
        ]),
        w,
        h,
        { x: cx - t, y: cy - t, w: 2 * t, h: 2 * t },
      );
    }
    case "homePlate": {
      const p = Math.min(w, ss * a());
      return one(poly([[0, 0], [w - p, 0], [w, h / 2], [w - p, h], [0, h]]), w, h, { x: 0, y: 0, w: w - p / 2, h });
    }
    case "chevron": {
      const p = Math.min(w, ss * a());
      return one(poly([[0, 0], [w - p, 0], [w, h / 2], [w - p, h], [0, h], [p, h / 2]]), w, h, { x: p, y: 0, w: w - 2 * p, h });
    }
    case "mathPlus": {
      const t = (ss * a("adj1")) / 2;
      const m = Math.min(w, h) * 0.12;
      const cx = w / 2;
      const cy = h / 2;
      return one(poly([[cx - t, m], [cx + t, m], [cx + t, cy - t], [w - m, cy - t], [w - m, cy + t], [cx + t, cy + t], [cx + t, h - m], [cx - t, h - m], [cx - t, cy + t], [m, cy + t], [m, cy - t], [cx - t, cy - t]]), w, h, { x: m, y: cy - t, w: w - 2 * m, h: 2 * t });
    }
    case "mathMinus": {
      const t = (h * a("adj1")) / 2;
      const m = w * 0.12;
      return one(poly([[m, h / 2 - t], [w - m, h / 2 - t], [w - m, h / 2 + t], [m, h / 2 + t]]), w, h, { x: m, y: h / 2 - t, w: w - 2 * m, h: 2 * t });
    }
    case "mathMultiply": {
      const t = (ss * a("adj1")) / 2;
      const cx = w / 2;
      const cy = h / 2;
      const k = 0.7071;
      const r = Math.min(w, h) * 0.42;
      const arm = (ang: number): [number, number][] => {
        const c = Math.cos(ang);
        const s = Math.sin(ang);
        return [
          [cx + c * r - s * t * k * 1.4, cy + s * r + c * t * k * 1.4],
          [cx + c * r + s * t * k * 1.4, cy + s * r - c * t * k * 1.4],
        ];
      };
      void k;
      const pts: [number, number][] = [];
      for (let i = 0; i < 4; i++) {
        const ang = Math.PI / 4 + (i * Math.PI) / 2;
        const [p1, p2] = arm(ang);
        pts.push(p2, p1);
        const mid = ang + Math.PI / 4;
        pts.push([cx + Math.cos(mid) * t * 1.41, cy + Math.sin(mid) * t * 1.41]);
      }
      return one(poly(pts), w, h, inset(w, h, 0.3, 0.3, 0.3, 0.3));
    }
    case "mathEqual": {
      const t = (h * a("adj1")) / 2;
      const g = (h * a("adj2")) / 2;
      const m = w * 0.12;
      const top = poly([[m, h / 2 - g - 2 * t], [w - m, h / 2 - g - 2 * t], [w - m, h / 2 - g], [m, h / 2 - g]]);
      const bot = poly([[m, h / 2 + g], [w - m, h / 2 + g], [w - m, h / 2 + g + 2 * t], [m, h / 2 + g + 2 * t]]);
      return one(top + " " + bot, w, h, { x: m, y: h / 2 - g - 2 * t, w: w - 2 * m, h: 2 * (g + 2 * t) });
    }
    case "wedgeRectCallout":
    case "wedgeRoundRectCallout": {
      const tx = w / 2 + w * a("adj1");
      const ty = h / 2 + h * a("adj2");
      const r = geom === "wedgeRoundRectCallout" ? ss * a("adj3") : 0;
      return one(rectCallout(w, h, tx, ty, r), w, h);
    }
    case "wedgeEllipseCallout": {
      const tx = w / 2 + w * a("adj1");
      const ty = h / 2 + h * a("adj2");
      const ang = Math.atan2((ty - h / 2) / h, (tx - w / 2) / w);
      const d1 = ang - 0.2;
      const d2 = ang + 0.2;
      const p1 = [w / 2 + (Math.cos(d1) * w) / 2, h / 2 + (Math.sin(d1) * h) / 2];
      const p2 = [w / 2 + (Math.cos(d2) * w) / 2, h / 2 + (Math.sin(d2) * h) / 2];
      const d = `M${pt(p2[0], p2[1])} A${f(w / 2)} ${f(h / 2)} 0 1 1 ${pt(p1[0], p1[1])} L${pt(tx, ty)} Z`;
      return one(d, w, h, inset(w, h, 0.146, 0.146, 0.146, 0.146));
    }
    case "cloudCallout": {
      const tx = w / 2 + w * a("adj1");
      const ty = h / 2 + h * a("adj2");
      const parts = [cloudPath(0, 0, w, h)];
      const ex = w / 2 + (tx - w / 2) * 0.62;
      const ey = h / 2 + (ty - h / 2) * 0.62;
      const r1 = ss * 0.08;
      parts.push(ellipsePath(ex, ey, r1, r1 * 0.8));
      parts.push(ellipsePath(ex + (tx - ex) * 0.5, ey + (ty - ey) * 0.5, r1 * 0.6, r1 * 0.5));
      parts.push(ellipsePath(tx, ty, r1 * 0.35, r1 * 0.3));
      return one(parts.join(" "), w, h, inset(w, h, 0.18, 0.22, 0.18, 0.22));
    }
    case "line":
    case "straightConnector1":
      return { paths: [{ d: `M0 0 L${pt(w, h)}`, fill: "none", stroke: true }], text: full(w, h), open: true };
    case "bentConnector3": {
      const x = w * a("adj1");
      return { paths: [{ d: `M0 0 L${pt(x, 0)} L${pt(x, h)} L${pt(w, h)}`, fill: "none", stroke: true }], text: full(w, h), open: true };
    }
    case "curvedConnector3":
      return { paths: [{ d: `M0 0 C${pt(w / 2, 0)} ${pt(w / 2, h)} ${pt(w, h)}`, fill: "none", stroke: true }], text: full(w, h), open: true };
    case "arc":
      return { paths: [{ d: `M${pt(w / 2, 0)} A${f(w / 2)} ${f(h / 2)} 0 0 1 ${pt(w, h / 2)}`, fill: "none", stroke: true }], text: full(w, h), open: true };
    default:
      return one(`M0 0 L${pt(w, 0)} L${pt(w, h)} L${pt(0, h)} Z`, w, h);
  }
}

export function isLineGeom(geom: string): boolean {
  return geom === "line" || geom === "straightConnector1" || geom === "bentConnector3" || geom === "curvedConnector3" || geom === "arc";
}

export interface ShapeCatalogItem {
  geom: string;
  label: string;
  /** Çizgi okları (Ok Çizgisi). */
  tail?: "triangle";
  head?: "triangle";
}

export const SHAPE_CATALOG: { group: string; items: ShapeCatalogItem[] }[] = [
  {
    group: "Çizgiler",
    items: [
      { geom: "line", label: "Çizgi" },
      { geom: "straightConnector1", label: "Ok Çizgisi", tail: "triangle" },
      { geom: "straightConnector1", label: "Çift Ok Çizgisi", tail: "triangle", head: "triangle" },
      { geom: "bentConnector3", label: "Bağlayıcı: Dirsek" },
      { geom: "curvedConnector3", label: "Bağlayıcı: Eğri" },
      { geom: "arc", label: "Yay" },
    ],
  },
  {
    group: "Dikdörtgenler",
    items: [
      { geom: "rect", label: "Dikdörtgen" },
      { geom: "roundRect", label: "Dikdörtgen: Köşeleri Yuvarlatılmış" },
      { geom: "snip1Rect", label: "Dikdörtgen: Tek Köşesi Kesik" },
      { geom: "round2SameRect", label: "Dikdörtgen: Aynı Kenar Köşeleri Yuvarlatılmış" },
    ],
  },
  {
    group: "Temel Şekiller",
    items: [
      { geom: "ellipse", label: "Oval" },
      { geom: "triangle", label: "İkizkenar Üçgen" },
      { geom: "rtTriangle", label: "Dik Üçgen" },
      { geom: "parallelogram", label: "Paralelkenar" },
      { geom: "trapezoid", label: "Yamuk" },
      { geom: "diamond", label: "Elmas" },
      { geom: "pentagon", label: "Normal Beşgen" },
      { geom: "hexagon", label: "Altıgen" },
      { geom: "heptagon", label: "Yedigen" },
      { geom: "octagon", label: "Sekizgen" },
      { geom: "decagon", label: "Ongen" },
      { geom: "dodecagon", label: "Onikigen" },
      { geom: "pie", label: "Pasta" },
      { geom: "teardrop", label: "Gözyaşı" },
      { geom: "frame", label: "Çerçeve" },
      { geom: "corner", label: "L Şekli" },
      { geom: "plus", label: "Artı" },
      { geom: "can", label: "Silindir" },
      { geom: "cube", label: "Küp" },
      { geom: "bevel", label: "Eğim" },
      { geom: "donut", label: "Halka" },
      { geom: "noSmoking", label: "Yasak İşareti" },
      { geom: "foldedCorner", label: "Katlanmış Köşe" },
      { geom: "smileyFace", label: "Gülen Yüz" },
      { geom: "heart", label: "Kalp" },
      { geom: "lightningBolt", label: "Şimşek" },
      { geom: "sun", label: "Güneş" },
      { geom: "moon", label: "Ay" },
      { geom: "cloud", label: "Bulut" },
      { geom: "wave", label: "Dalga" },
    ],
  },
  {
    group: "Blok Oklar",
    items: [
      { geom: "rightArrow", label: "Ok: Sağ" },
      { geom: "leftArrow", label: "Ok: Sol" },
      { geom: "upArrow", label: "Ok: Yukarı" },
      { geom: "downArrow", label: "Ok: Aşağı" },
      { geom: "leftRightArrow", label: "Ok: Sol-Sağ" },
      { geom: "upDownArrow", label: "Ok: Yukarı-Aşağı" },
      { geom: "quadArrow", label: "Ok: Dört Yönlü" },
      { geom: "homePlate", label: "Ok: Beşgen" },
      { geom: "chevron", label: "Ok: Köşeli Çift Ayraç" },
    ],
  },
  {
    group: "Denklem Şekilleri",
    items: [
      { geom: "mathPlus", label: "Artı İşareti" },
      { geom: "mathMinus", label: "Eksi İşareti" },
      { geom: "mathMultiply", label: "Çarpma İşareti" },
      { geom: "mathEqual", label: "Eşittir" },
    ],
  },
  {
    group: "Akış Çizelgesi",
    items: [
      { geom: "flowChartProcess", label: "Akış Çizelgesi: İşlem" },
      { geom: "flowChartDecision", label: "Akış Çizelgesi: Karar" },
      { geom: "flowChartInputOutput", label: "Akış Çizelgesi: Veri" },
      { geom: "flowChartPredefinedProcess", label: "Akış Çizelgesi: Önceden Tanımlı İşlem" },
      { geom: "flowChartDocument", label: "Akış Çizelgesi: Belge" },
      { geom: "flowChartTerminator", label: "Akış Çizelgesi: Sonlandırıcı" },
      { geom: "flowChartConnector", label: "Akış Çizelgesi: Bağlayıcı" },
    ],
  },
  {
    group: "Yıldızlar ve Başlık Sancakları",
    items: [
      { geom: "irregularSeal1", label: "Patlama: 8 Köşeli" },
      { geom: "star4", label: "Yıldız: 4 Köşeli" },
      { geom: "star5", label: "Yıldız: 5 Köşeli" },
      { geom: "star6", label: "Yıldız: 6 Köşeli" },
      { geom: "star8", label: "Yıldız: 8 Köşeli" },
      { geom: "star10", label: "Yıldız: 10 Köşeli" },
      { geom: "star12", label: "Yıldız: 12 Köşeli" },
    ],
  },
  {
    group: "Belirtme Çizgileri",
    items: [
      { geom: "wedgeRectCallout", label: "Konuşma Balonu: Dikdörtgen" },
      { geom: "wedgeRoundRectCallout", label: "Konuşma Balonu: Köşeleri Yuvarlatılmış Dikdörtgen" },
      { geom: "wedgeEllipseCallout", label: "Konuşma Balonu: Oval" },
      { geom: "cloudCallout", label: "Düşünce Balonu: Bulut" },
    ],
  },
];

export const SHAPE_LABELS: Record<string, string> = Object.fromEntries(SHAPE_CATALOG.flatMap((g) => g.items.map((i) => [i.geom, i.label])));
