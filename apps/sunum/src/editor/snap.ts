/** Akıllı kılavuzlar: slayt kenarı/ortası, diğer nesneler ve kılavuz çizgilerine yapışma. */
import type { Box } from "../model/ops";

export interface GuideLine {
  /** "v": dikey çizgi (x sabit), "h": yatay çizgi (y sabit). */
  dir: "v" | "h";
  pos: number;
  from: number;
  to: number;
}

export interface SnapResult {
  dx: number;
  dy: number;
  lines: GuideLine[];
}

interface Candidate {
  pos: number;
  from: number;
  to: number;
}

export function snapMove(moving: Box, others: Box[], slideW: number, slideH: number, threshold: number, guides?: { h: number[]; v: number[] }): SnapResult {
  const xs: Candidate[] = [
    { pos: 0, from: 0, to: slideH },
    { pos: slideW / 2, from: 0, to: slideH },
    { pos: slideW, from: 0, to: slideH },
  ];
  const ys: Candidate[] = [
    { pos: 0, from: 0, to: slideW },
    { pos: slideH / 2, from: 0, to: slideW },
    { pos: slideH, from: 0, to: slideW },
  ];
  for (const g of guides?.v ?? []) xs.push({ pos: g, from: 0, to: slideH });
  for (const g of guides?.h ?? []) ys.push({ pos: g, from: 0, to: slideW });
  for (const o of others) {
    for (const x of [o.x, o.x + o.w / 2, o.x + o.w]) xs.push({ pos: x, from: o.y, to: o.y + o.h });
    for (const y of [o.y, o.y + o.h / 2, o.y + o.h]) ys.push({ pos: y, from: o.x, to: o.x + o.w });
  }
  const mx = [moving.x, moving.x + moving.w / 2, moving.x + moving.w];
  const my = [moving.y, moving.y + moving.h / 2, moving.y + moving.h];
  let best: { d: number; c: Candidate } | null = null;
  for (const c of xs)
    for (const m of mx) {
      const d = c.pos - m;
      if (Math.abs(d) <= threshold && (!best || Math.abs(d) < Math.abs(best.d))) best = { d, c };
    }
  let bestY: { d: number; c: Candidate } | null = null;
  for (const c of ys)
    for (const m of my) {
      const d = c.pos - m;
      if (Math.abs(d) <= threshold && (!bestY || Math.abs(d) < Math.abs(bestY.d))) bestY = { d, c };
    }
  const dx = best?.d ?? 0;
  const dy = bestY?.d ?? 0;
  const moved = { x: moving.x + dx, y: moving.y + dy, w: moving.w, h: moving.h };
  const lines: GuideLine[] = [];
  if (best) {
    // Aynı konuma denk gelen tüm adaylar için çizgi.
    for (const c of xs) {
      if (Math.abs(c.pos - (best.c.pos)) > 0.01) continue;
      if (![moved.x, moved.x + moved.w / 2, moved.x + moved.w].some((m) => Math.abs(m - c.pos) < 0.5)) continue;
      lines.push({ dir: "v", pos: c.pos, from: Math.min(c.from, moved.y), to: Math.max(c.to, moved.y + moved.h) });
    }
  }
  if (bestY) {
    for (const c of ys) {
      if (Math.abs(c.pos - bestY.c.pos) > 0.01) continue;
      if (![moved.y, moved.y + moved.h / 2, moved.y + moved.h].some((m) => Math.abs(m - c.pos) < 0.5)) continue;
      lines.push({ dir: "h", pos: c.pos, from: Math.min(c.from, moved.x), to: Math.max(c.to, moved.x + moved.w) });
    }
  }
  return { dx, dy, lines };
}

/** Tek bir değeri (kenar) adaylara yapıştırır. */
export function snapValue(v: number, candidates: number[], threshold: number): { v: number; hit?: number } {
  let best: number | undefined;
  for (const c of candidates) if (Math.abs(c - v) <= threshold && (best === undefined || Math.abs(c - v) < Math.abs(best - v))) best = c;
  return best === undefined ? { v } : { v: best, hit: best };
}

export function gridSnap(v: number, step: number): number {
  return Math.round(v / step) * step;
}
