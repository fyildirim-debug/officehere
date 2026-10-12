/** Slayt alanı yardımcıları: sürükleme türleri, grup gezinme, ızgara ve önizleme çizimleri. */
import type { Box } from "../model/ops";
import { findShape } from "../model/ops";
import type { GroupShape, Shape, Slide, TableShape } from "../model/types";
import type { GeomOverride } from "../render/SlideView";
import type { AppState, InkTool } from "../state/store";
import type { Geom } from "./interact";

export type Pt = [number, number];

export type Drag =
  | { kind: "pending" | "move"; id: string; start: Pt; startClient: Pt; wasSelected: boolean; textHit: boolean; cell?: [number, number]; ids: string[]; geoms: Map<string, Geom>; ctrl: boolean; point: { x: number; y: number } }
  | { kind: "resize"; hx: number; hy: number; start: Pt; geoms: Map<string, Geom>; kinds: Map<string, string>; crop?: { id: string; c: { l: number; t: number; r: number; b: number } } }
  | { kind: "rotate"; start: Pt; startAngle: number; geoms: Map<string, Geom> }
  | { kind: "end"; id: string; which: "start" | "end"; geom: Geom }
  | { kind: "adj"; id: string; index: number; geom: Geom; adj: Record<string, number>; geomName: string }
  | { kind: "lasso"; start: Pt; additive: boolean }
  | { kind: "insert"; start: Pt }
  | { kind: "cells"; id: string; start: [number, number] }
  | { kind: "guide"; dir: "h" | "v"; index: number }
  | { kind: "cropPan"; id: string; start: Pt; geom: Geom; c: { l: number; t: number; r: number; b: number } }
  | { kind: "ink"; pts: Pt[] }
  | { kind: "erase"; key: string };

export const MARGIN = 40;

export function topLevelOf(shapes: Shape[], id: string): Shape | undefined {
  for (const s of shapes) {
    if (s.id === id) return s;
    if (s.kind === "group" && findShape(s.children, id)) return s;
  }
  return undefined;
}

export function childOfGroupContaining(g: GroupShape, id: string): Shape | undefined {
  return g.children.find((c) => c.id === id || (c.kind === "group" && findShape(c.children, id)));
}

export function animationLabels(slide: Slide): Map<string, string[]> {
  const out = new Map<string, string[]>();
  let click = 0;
  for (const a of slide.animations) {
    if (a.trigger === "click") click += 1;
    const list = out.get(a.shapeId) ?? [];
    list.push(String(click));
    out.set(a.shapeId, list);
  }
  return out;
}


/** Silgi: istemci noktasının çevresindeki mürekkep şekilleri. */
export function inkIdsAt(slide: Slide, cx: number, cy: number): string[] {
  const ids = new Set<string>();
  for (const [dx, dy] of [[0, 0], [-5, 0], [5, 0], [0, -5], [0, 5]]) {
    for (const el of document.elementsFromPoint(cx + dx, cy + dy)) {
      const id = (el.closest("[data-shape-id]") as HTMLElement | null)?.dataset.shapeId;
      if (!id) continue;
      const sh = findShape(slide.shapes, id);
      if (sh && sh.kind === "sp" && sh.ink) ids.add(id);
    }
  }
  return [...ids];
}

/** Seçili tablo hücre aralığının slayt koordinatlarındaki dikdörtgeni. */
export function cellSelRects(t: TableShape, cs: NonNullable<AppState["cellSel"]>): { shape: TableShape; rects: Box[] } {
  const r1 = Math.min(cs.r1, cs.r2);
  const r2 = Math.max(cs.r1, cs.r2);
  const c1 = Math.min(cs.c1, cs.c2);
  const c2 = Math.max(cs.c1, cs.c2);
  let y = t.y;
  for (let r = 0; r < r1; r++) y += t.rows[r].h;
  let x = t.x;
  for (let c = 0; c < c1; c++) x += t.cols[c];
  let h = 0;
  for (let r = r1; r <= r2; r++) h += t.rows[r]?.h ?? 0;
  let w = 0;
  for (let c = c1; c <= c2; c++) w += t.cols[c] ?? 0;
  return { shape: t, rects: [{ x, y, w, h }] };
}

/** Kırpma modunda resmin tamamının (soluk) konumu. */
export function cropGhostOf(slide: Slide, cropId: string, overrides?: Map<string, GeomOverride>): { src: string; x: number; y: number; w: number; h: number } | null {
  const pic = findShape(slide.shapes, cropId);
  if (!pic || pic.kind !== "pic") return null;
  const g = overrides?.get(pic.id) ?? pic;
  const c = (overrides?.get(pic.id)?.crop ?? pic.crop) ?? { l: 0, t: 0, r: 0, b: 0 };
  const fw = g.w / Math.max(0.01, 1 - c.l - c.r);
  const fh = g.h / Math.max(0.01, 1 - c.t - c.b);
  return { src: pic.src, x: g.x - c.l * fw, y: g.y - c.t * fh, w: fw, h: fh };
}

/** Çizilmekte olan mürekkebin canlı önizlemesi. */
export function InkPreview({ pts, ink, z, w, h }: { pts: Pt[]; ink: InkTool; z: number; w: number; h: number }) {
  return (
    <svg className="ink-preview" width={w * z} height={h * z} overflow="visible">
      <path
        d={pts.map((p, i) => `${i ? "L" : "M"}${(p[0] * z).toFixed(1)} ${(p[1] * z).toFixed(1)}`).join(" ") + (pts.length === 1 ? " l0.01 0.01" : "")}
        fill="none"
        stroke={ink.color}
        strokeOpacity={ink.tool === "highlighter" ? 0.5 : ink.tool === "pencil" ? 0.85 : 1}
        strokeWidth={Math.max(1, ((ink.width * 96) / 72) * z)}
        strokeLinecap="round"
        strokeLinejoin="round"
      />
    </svg>
  );
}

export function GridOverlay({ w, h, z }: { w: number; h: number; z: number }) {
  const step = (96 / 2.54) * z;
  return (
    <svg className="grid-overlay" width={w * z} height={h * z}>
      <defs>
        <pattern id="oh-grid" width={step} height={step} patternUnits="userSpaceOnUse" x={((w * z) / 2) % step} y={((h * z) / 2) % step}>
          <circle cx={0.5} cy={0.5} r={0.7} fill="rgba(0,0,0,0.35)" />
        </pattern>
      </defs>
      <rect width="100%" height="100%" fill="url(#oh-grid)" />
    </svg>
  );
}
