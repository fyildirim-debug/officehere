/** Boyutlandırma, döndürme ve çizgi uç noktası geometrisi. */
import type { Shape } from "../model/types";

export interface Geom {
  x: number;
  y: number;
  w: number;
  h: number;
  rot: number;
  flipH?: boolean;
  flipV?: boolean;
}

export const geomOf = (s: Shape): Geom => ({ x: s.x, y: s.y, w: s.w, h: s.h, rot: s.rot, flipH: s.flipH, flipV: s.flipV });

export function rotateVec(x: number, y: number, deg: number): [number, number] {
  const r = (deg * Math.PI) / 180;
  const c = Math.cos(r);
  const s = Math.sin(r);
  return [x * c - y * s, x * s + y * c];
}

export interface ResizeOpts {
  keepAspect: boolean;
  fromCenter: boolean;
  min?: number;
}

/** (hx, hy) ∈ {-1,0,1}: tutamacın yerel konumu; (lx, ly): yerel sürükleme farkı. */
export function resizeGeom(g: Geom, hx: number, hy: number, lx: number, ly: number, o: ResizeOpts): Geom {
  const k = o.fromCenter ? 2 : 1;
  let nw = hx ? g.w + hx * lx * k : g.w;
  let nh = hy ? g.h + hy * ly * k : g.h;
  if (o.keepAspect && g.w > 0 && g.h > 0) {
    if (hx && hy) {
      const s = Math.max(Math.abs(nw) / g.w, Math.abs(nh) / g.h);
      nw = Math.sign(nw || 1) * g.w * s;
      nh = Math.sign(nh || 1) * g.h * s;
    } else if (hx) nh = (g.h * Math.abs(nw)) / g.w;
    else if (hy) nw = (g.w * Math.abs(nh)) / g.h;
  }
  const min = o.min ?? 0;
  const flipX = hx !== 0 && nw < 0;
  const flipY = hy !== 0 && nh < 0;
  let aw = Math.max(min, Math.abs(nw));
  let ah = Math.max(min, Math.abs(nh));
  if (!isFinite(aw)) aw = g.w;
  if (!isFinite(ah)) ah = g.h;
  let cx = 0;
  let cy = 0;
  if (!o.fromCenter) {
    if (hx) cx = -hx * (g.w / 2) + (hx * (flipX ? -aw : aw)) / 2;
    if (hy) cy = -hy * (g.h / 2) + (hy * (flipY ? -ah : ah)) / 2;
  }
  const [wx, wy] = rotateVec(cx, cy, g.rot);
  const ocx = g.x + g.w / 2 + wx;
  const ocy = g.y + g.h / 2 + wy;
  return {
    x: ocx - aw / 2,
    y: ocy - ah / 2,
    w: aw,
    h: ah,
    rot: g.rot,
    flipH: flipX ? !g.flipH : g.flipH,
    flipV: flipY ? !g.flipV : g.flipV,
  };
}

export function angleFromCenter(g: Geom, px: number, py: number): number {
  const cx = g.x + g.w / 2;
  const cy = g.y + g.h / 2;
  return (Math.atan2(py - cy, px - cx) * 180) / Math.PI + 90;
}

export function normAngle(a: number): number {
  return ((a % 360) + 360) % 360;
}

/** Çizginin dünya koordinatlarındaki uç noktaları. */
export function lineEnds(g: Geom): [[number, number], [number, number]] {
  const sx = g.flipH ? g.x + g.w : g.x;
  const sy = g.flipV ? g.y + g.h : g.y;
  const ex = g.flipH ? g.x : g.x + g.w;
  const ey = g.flipV ? g.y : g.y + g.h;
  if (!g.rot) return [
    [sx, sy],
    [ex, ey],
  ];
  const cx = g.x + g.w / 2;
  const cy = g.y + g.h / 2;
  const a = rotateVec(sx - cx, sy - cy, g.rot);
  const b = rotateVec(ex - cx, ey - cy, g.rot);
  return [
    [cx + a[0], cy + a[1]],
    [cx + b[0], cy + b[1]],
  ];
}

export function geomFromEnds(a: [number, number], b: [number, number]): Geom {
  return {
    x: Math.min(a[0], b[0]),
    y: Math.min(a[1], b[1]),
    w: Math.abs(b[0] - a[0]),
    h: Math.abs(b[1] - a[1]),
    rot: 0,
    flipH: b[0] < a[0],
    flipV: b[1] < a[1],
  };
}

/** Shift ile çizgi açısını 45°'ye yuvarlar. */
export function constrainLine(a: [number, number], b: [number, number]): [number, number] {
  const dx = b[0] - a[0];
  const dy = b[1] - a[1];
  const len = Math.hypot(dx, dy);
  const ang = Math.round(Math.atan2(dy, dx) / (Math.PI / 4)) * (Math.PI / 4);
  return [a[0] + Math.cos(ang) * len, a[1] + Math.sin(ang) * len];
}

export const HANDLES: { id: string; hx: number; hy: number; cursor: string }[] = [
  { id: "nw", hx: -1, hy: -1, cursor: "nwse-resize" },
  { id: "n", hx: 0, hy: -1, cursor: "ns-resize" },
  { id: "ne", hx: 1, hy: -1, cursor: "nesw-resize" },
  { id: "e", hx: 1, hy: 0, cursor: "ew-resize" },
  { id: "se", hx: 1, hy: 1, cursor: "nwse-resize" },
  { id: "s", hx: 0, hy: 1, cursor: "ns-resize" },
  { id: "sw", hx: -1, hy: 1, cursor: "nesw-resize" },
  { id: "w", hx: -1, hy: 0, cursor: "ew-resize" },
];

const CURSORS = ["ns-resize", "nesw-resize", "ew-resize", "nwse-resize"];

/** Döndürülmüş tutamaç için uygun imleç. */
export function cursorFor(hx: number, hy: number, rot: number): string {
  const base = (Math.atan2(hy, hx) * 180) / Math.PI + 90 + rot;
  const idx = Math.round(normAngle(base) / 45) % 4;
  return CURSORS[idx];
}
