/** Akıllı kılavuzlar: slayt kenarı/ortası, diğer nesneler ve kılavuz çizgilerine yapışma. */
import type { Box } from "../model/ops";

export interface GuideLine {
  /** "v": dikey çizgi (x sabit), "h": yatay çizgi (y sabit). */
  dir: "v" | "h";
  pos: number;
  from: number;
  to: number;
  /** Eşit aralık işareti (iki nesne arasındaki boşluğu gösteren uçlu çizgi). */
  gap?: boolean;
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

/* ------------------------- Eşit aralık ------------------------- */

interface Spacing {
  /** Taşınan nesnenin yapışacağı konum (x ya da y). */
  pos: number;
  /** Eşit boşluk çiftleri (önce/sonra nesneleri; taşınan nesne null). */
  pairs: [Box | null, Box | null][];
}

/** Eksen soyutlaması: "x" yatay aralık (satırdaki nesneler), "y" dikey aralık (sütundaki nesneler). */
const AX = {
  x: { s: (b: Box) => b.x, l: (b: Box) => b.w, cs: (b: Box) => b.y, cl: (b: Box) => b.h },
  y: { s: (b: Box) => b.y, l: (b: Box) => b.h, cs: (b: Box) => b.x, cl: (b: Box) => b.w },
};

/** Aynı satırda/sütunda (diğer eksende örtüşen) nesneler. */
function lane(ax: keyof typeof AX, ref: Box, all: Box[]): Box[] {
  const A = AX[ax];
  return all.filter((o) => o !== ref && A.cs(o) < A.cs(ref) + A.cl(ref) && A.cs(o) + A.cl(o) > A.cs(ref));
}

/**
 * Eşit aralık adayları: iki komşunun tam ortası ya da komşu çiftinin boşluğu kadar
 * uzaklık (sola/sağa ya da yukarı/aşağı). En yakın aday eşik içindeyse döner.
 */
function spacingSnap(ax: keyof typeof AX, m: Box, others: Box[], threshold: number): Spacing | null {
  const A = AX[ax];
  const end = (b: Box) => A.s(b) + A.l(b);
  const row = lane(ax, m, others).filter((o) => A.l(o) > 0);
  const before = row.filter((o) => end(o) <= A.s(m) + threshold).sort((a, b) => end(b) - end(a));
  const after = row.filter((o) => A.s(o) >= end(m) - threshold).sort((a, b) => A.s(a) - A.s(b));
  const L = before[0];
  const R = after[0];
  const cands: Spacing[] = [];
  if (L && R && A.s(R) - end(L) > A.l(m)) cands.push({ pos: (end(L) + A.s(R) - A.l(m)) / 2, pairs: [[L, null], [null, R]] });
  if (L) {
    const LL = lane(ax, L, others)
      .filter((o) => end(o) <= A.s(L))
      .sort((a, b) => end(b) - end(a))[0];
    if (LL) cands.push({ pos: end(L) + (A.s(L) - end(LL)), pairs: [[LL, L], [L, null]] });
  }
  if (R) {
    const RR = lane(ax, R, others)
      .filter((o) => A.s(o) >= end(R))
      .sort((a, b) => A.s(a) - A.s(b))[0];
    if (RR) cands.push({ pos: A.s(R) - (A.s(RR) - end(R)) - A.l(m), pairs: [[null, R], [R, RR]] });
  }
  let best: Spacing | null = null;
  for (const c of cands) if (Math.abs(c.pos - A.s(m)) <= threshold && (!best || Math.abs(c.pos - A.s(m)) < Math.abs(best.pos - A.s(m)))) best = c;
  return best;
}

/** Eşit aralık çiftlerinin işaret çizgileri (taşınan nesnenin son konumuyla). */
function spacingMarks(ax: keyof typeof AX, sp: Spacing, moved: Box): GuideLine[] {
  const A = AX[ax];
  return sp.pairs.map(([a0, b0]) => {
    const a = a0 ?? moved;
    const b = b0 ?? moved;
    const lo = Math.max(A.cs(a), A.cs(b));
    const hi = Math.min(A.cs(a) + A.cl(a), A.cs(b) + A.cl(b));
    const mid = hi > lo ? (lo + hi) / 2 : (A.cs(a) + A.cl(a) / 2 + A.cs(b) + A.cl(b) / 2) / 2;
    return { dir: ax === "x" ? "h" : "v", pos: mid, from: A.s(a) + A.l(a), to: A.s(b), gap: true };
  });
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
  // Eşit aralık hizalamadan daha yakınsa (ya da eşitse) o eksende o kullanılır.
  const spX = spacingSnap("x", moving, others, threshold);
  const spY = spacingSnap("y", moving, others, threshold);
  const useSpX = !!spX && (!best || Math.abs(spX.pos - moving.x) <= Math.abs(best.d) + 0.01);
  const useSpY = !!spY && (!bestY || Math.abs(spY.pos - moving.y) <= Math.abs(bestY.d) + 0.01);
  if (useSpX) best = null;
  if (useSpY) bestY = null;
  const dx = useSpX ? spX!.pos - moving.x : (best?.d ?? 0);
  const dy = useSpY ? spY!.pos - moving.y : (bestY?.d ?? 0);
  const moved = { x: moving.x + dx, y: moving.y + dy, w: moving.w, h: moving.h };
  const lines: GuideLine[] = [];
  if (useSpX) lines.push(...spacingMarks("x", spX!, moved));
  if (useSpY) lines.push(...spacingMarks("y", spY!, moved));
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
