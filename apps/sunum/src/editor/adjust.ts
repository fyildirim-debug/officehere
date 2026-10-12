/** Şekil ayar tutamaçları (sarı): konumları ve sürüklemeyle ayar değeri hesaplama. */
import { DEFAULT_ADJ } from "../render/geometry";

export interface AdjustHandle {
  /** Yerel koordinatta tutamaç konumu. */
  pos: (w: number, h: number, adj: Record<string, number>) => [number, number];
  /** Yerel noktadan yeni ayar değerleri. */
  set: (x: number, y: number, w: number, h: number, adj: Record<string, number>) => Record<string, number>;
}

const clamp = (v: number, a: number, b: number) => Math.max(a, Math.min(b, v));
const ss = (w: number, h: number) => Math.min(w, h) || 1;

function edgeX(key: string, max: number): AdjustHandle {
  return {
    pos: (w, h, a) => [(ss(w, h) * a[key]) / 100000, 0],
    set: (x, _y, w, h) => ({ [key]: Math.round(clamp((x / ss(w, h)) * 100000, 0, max)) }),
  };
}

const ARROW_H = (vertical: boolean, reverse: boolean): AdjustHandle[] => {
  if (!vertical)
    return [
      {
        pos: (w, h, a) => {
          const hl = Math.min(w, (ss(w, h) * a.adj2) / 100000);
          return [reverse ? hl : w - hl, h / 2 - (h * a.adj1) / 200000];
        },
        set: (x, y, w, h) => ({
          adj2: Math.round(clamp(((reverse ? x : w - x) / ss(w, h)) * 100000, 0, (w / ss(w, h)) * 100000)),
          adj1: Math.round(clamp(((h / 2 - y) * 200000) / h, 0, 100000)),
        }),
      },
    ];
  return [
    {
      pos: (w, h, a) => {
        const hl = Math.min(h, (ss(w, h) * a.adj2) / 100000);
        return [w / 2 - (w * a.adj1) / 200000, reverse ? h - hl : hl];
      },
      set: (x, y, w, h) => ({
        adj2: Math.round(clamp(((reverse ? h - y : y) / ss(w, h)) * 100000, 0, (h / ss(w, h)) * 100000)),
        adj1: Math.round(clamp(((w / 2 - x) * 200000) / w, 0, 100000)),
      }),
    },
  ];
};

const CALLOUT: AdjustHandle = {
  pos: (w, h, a) => [w / 2 + (w * a.adj1) / 100000, h / 2 + (h * a.adj2) / 100000],
  set: (x, y, w, h) => ({ adj1: Math.round(((x - w / 2) / (w || 1)) * 100000), adj2: Math.round(((y - h / 2) / (h || 1)) * 100000) }),
};

function star(): AdjustHandle {
  return {
    pos: (w, h, a) => [w / 2, h / 2 - (h / 2) * ((a.adj * 2) / 100000)],
    set: (_x, y, _w, h) => ({ adj: Math.round(clamp(((h / 2 - y) / (h / 2 || 1)) * 50000, 1000, 50000)) }),
  };
}

export const ADJUST_HANDLES: Record<string, AdjustHandle[]> = {
  roundRect: [edgeX("adj", 50000)],
  round2SameRect: [edgeX("adj1", 50000)],
  snip1Rect: [
    {
      pos: (w, h, a) => [w - (ss(w, h) * a.adj) / 100000, 0],
      set: (x, _y, w, h) => ({ adj: Math.round(clamp(((w - x) / ss(w, h)) * 100000, 0, 50000)) }),
    },
  ],
  triangle: [{ pos: (w, _h, a) => [(w * a.adj) / 100000, 0], set: (x, _y, w) => ({ adj: Math.round(clamp((x / (w || 1)) * 100000, 0, 100000)) }) }],
  parallelogram: [edgeX("adj", 100000)],
  trapezoid: [edgeX("adj", 50000)],
  hexagon: [edgeX("adj", 50000)],
  octagon: [edgeX("adj", 50000)],
  plus: [edgeX("adj", 50000)],
  frame: [edgeX("adj1", 50000)],
  bevel: [edgeX("adj", 50000)],
  cube: [
    {
      pos: (w, h, a) => [0, (ss(w, h) * a.adj) / 100000],
      set: (_x, y, w, h) => ({ adj: Math.round(clamp((y / ss(w, h)) * 100000, 0, 100000)) }),
    },
  ],
  can: [
    {
      pos: (w, h, a) => [w / 2, (ss(w, h) * a.adj) / 100000],
      set: (_x, y, w, h) => ({ adj: Math.round(clamp((y / ss(w, h)) * 100000, 0, 50000)) }),
    },
  ],
  donut: [edgeX("adj", 50000)],
  foldedCorner: [
    {
      pos: (w, h, a) => [w - (ss(w, h) * a.adj) / 100000, h],
      set: (x, _y, w, h) => ({ adj: Math.round(clamp(((w - x) / ss(w, h)) * 100000, 0, 50000)) }),
    },
  ],
  rightArrow: ARROW_H(false, false),
  leftArrow: ARROW_H(false, true),
  downArrow: ARROW_H(true, true),
  upArrow: ARROW_H(true, false),
  homePlate: [
    {
      pos: (w, h, a) => [w - Math.min(w, (ss(w, h) * a.adj) / 100000), 0],
      set: (x, _y, w, h) => ({ adj: Math.round(clamp(((w - x) / ss(w, h)) * 100000, 0, (w / ss(w, h)) * 100000)) }),
    },
  ],
  chevron: [
    {
      pos: (w, h, a) => [w - Math.min(w, (ss(w, h) * a.adj) / 100000), 0],
      set: (x, _y, w, h) => ({ adj: Math.round(clamp(((w - x) / ss(w, h)) * 100000, 0, (w / ss(w, h)) * 100000)) }),
    },
  ],
  star4: [star()],
  star5: [star()],
  star6: [star()],
  star8: [star()],
  star10: [star()],
  star12: [star()],
  wedgeRectCallout: [CALLOUT],
  wedgeRoundRectCallout: [CALLOUT],
  wedgeEllipseCallout: [CALLOUT],
  cloudCallout: [CALLOUT],
};

export function effectiveAdj(geom: string, adj?: Record<string, number>): Record<string, number> {
  return { ...(DEFAULT_ADJ[geom] ?? {}), ...(adj ?? {}) };
}
