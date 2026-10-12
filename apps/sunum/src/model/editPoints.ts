/**
 * Noktaları Düzenle: hazır şekli özel yola (custGeom) çevirir, köşe
 * noktalarını listeler, bir noktayı (ve bağlı Bézier denetim noktalarını)
 * taşır ve sınır kutusunu yeniden hesaplar.
 */
import { presetGeometry } from "../render/geometry";
import { svgPathToCmds } from "../render/svgPath";
import type { CustomPath, PathCmd, SpShape } from "./types";

/** Yol üzerindeki bir köşe noktasının konumu: yol, komut. */
export interface VertexRef {
  path: number;
  cmd: number;
}

/** Hazır şekli (ayarlarıyla) şekil boyutunda özel yola çevirir. */
export function toCustomShape(s: SpShape): SpShape {
  if (s.geom === "custom" && s.paths?.length) return s;
  const g = presetGeometry(s.geom, s.w, s.h, s.adj);
  const paths: CustomPath[] = g.paths.map((p) => ({ w: s.w, h: s.h, cmds: svgPathToCmds(p.d), fill: p.fill !== "none" && !g.open, stroke: p.stroke }));
  return { ...s, geom: "custom", adj: undefined, paths };
}

/** Köşe (bitiş) noktaları — Z dışındaki komutların son koordinat çifti. */
export function vertices(s: SpShape): (VertexRef & { x: number; y: number })[] {
  const out: (VertexRef & { x: number; y: number })[] = [];
  (s.paths ?? []).forEach((p, pi) => {
    const sx = p.w ? s.w / p.w : 1;
    const sy = p.h ? s.h / p.h : 1;
    let start: [number, number] | null = null;
    p.cmds.forEach((c, ci) => {
      if (c.c === "Z" || c.pts.length < 2) return;
      const x = c.pts[c.pts.length - 2];
      const y = c.pts[c.pts.length - 1];
      if (c.c === "M") start = [x, y];
      // Kapanan yolun son L komutu başlangıçla çakışıyorsa tekrar gösterme.
      else if (c.c === "L" && p.cmds[ci + 1]?.c === "Z" && start && Math.abs(start[0] - x) < 1e-6 && Math.abs(start[1] - y) < 1e-6) return;
      out.push({ path: pi, cmd: ci, x: x * sx, y: y * sy });
    });
  });
  return out;
}

/** Noktayı (şekil koordinatlarında dx, dy) taşır; bağlı denetim noktaları da kayar. */
export function moveVertex(s: SpShape, ref: VertexRef, dx: number, dy: number): SpShape {
  const paths = (s.paths ?? []).map((p, pi) => {
    if (pi !== ref.path) return p;
    const kx = p.w ? p.w / s.w : 1;
    const ky = p.h ? p.h / s.h : 1;
    const ddx = dx * kx;
    const ddy = dy * ky;
    const cmds: PathCmd[] = p.cmds.map((c, ci) => {
      const pts = c.pts.slice();
      if (ci === ref.cmd) {
        // Bitiş noktası ve (eğriyse) ona giren denetim noktası.
        pts[pts.length - 2] += ddx;
        pts[pts.length - 1] += ddy;
        if (c.c === "C") {
          pts[2] += ddx;
          pts[3] += ddy;
        }
      } else if (ci === ref.cmd + 1 && c.c === "C") {
        // Noktadan çıkan ilk denetim noktası.
        pts[0] += ddx;
        pts[1] += ddy;
      }
      return pts === c.pts ? c : { ...c, pts };
    });
    return { ...p, cmds };
  });
  return { ...s, paths };
}

/** Yol noktaları kutunun dışına çıktıysa sınır kutusunu yeniden hesaplar (döndürülmemiş şekiller). */
export function normalizeBounds(s: SpShape): SpShape {
  if (!s.paths?.length || s.rot) return s;
  let minX = Infinity;
  let minY = Infinity;
  let maxX = -Infinity;
  let maxY = -Infinity;
  for (const p of s.paths) {
    const sx = p.w ? s.w / p.w : 1;
    const sy = p.h ? s.h / p.h : 1;
    for (const c of p.cmds)
      for (let i = 0; i + 1 < c.pts.length; i += 2) {
        minX = Math.min(minX, c.pts[i] * sx);
        maxX = Math.max(maxX, c.pts[i] * sx);
        minY = Math.min(minY, c.pts[i + 1] * sy);
        maxY = Math.max(maxY, c.pts[i + 1] * sy);
      }
  }
  if (!Number.isFinite(minX)) return s;
  const w = Math.max(1, maxX - minX);
  const h = Math.max(1, maxY - minY);
  if (Math.abs(minX) < 0.01 && Math.abs(minY) < 0.01 && Math.abs(w - s.w) < 0.01 && Math.abs(h - s.h) < 0.01) return s;
  const paths = s.paths.map((p) => {
    const sx = p.w ? s.w / p.w : 1;
    const sy = p.h ? s.h / p.h : 1;
    return { ...p, w, h, cmds: p.cmds.map((c) => ({ ...c, pts: c.pts.map((v, i) => (i % 2 === 0 ? v * sx - minX : v * sy - minY)) })) };
  });
  return { ...s, x: s.x + minX, y: s.y + minY, w, h, paths };
}

/* ---------------- Nokta bağlam menüsü işlemleri ---------------- */

type P2 = [number, number];
const endOf = (c: PathCmd): P2 => [c.pts[c.pts.length - 2], c.pts[c.pts.length - 1]];

function mapPath(s: SpShape, pi: number, fn: (p: CustomPath) => CustomPath): SpShape {
  return { ...s, paths: (s.paths ?? []).map((p, i) => (i === pi ? fn(p) : p)) };
}

/** Komutun başlangıç noktası (önceki komutun bitişi; Z'den sonra alt yolun başı). */
function startOf(cmds: PathCmd[], ci: number): P2 | null {
  for (let i = ci - 1; i >= 0; i--) if (cmds[i].c !== "Z" && cmds[i].pts.length >= 2) return endOf(cmds[i]);
  return null;
}

/** Yol kapalı mı (Z içeriyor mu)? */
export function isPathClosed(s: SpShape, pi: number): boolean {
  return !!s.paths?.[pi]?.cmds.some((c) => c.c === "Z");
}

/** Noktadan sonraki kesimin ortasına yeni nokta ekler (eğri ise de Casteljau ile bölünür). */
export function addVertexAfter(s: SpShape, ref: VertexRef): SpShape {
  return mapPath(s, ref.path, (p) => {
    const cmds = p.cmds.slice();
    const ni = ref.cmd + 1;
    // Sonraki kesim yoksa (kapalı yolda) kapanış kesimi: başlangıç noktasına dönen çizgi.
    let next = cmds[ni];
    const a = endOf(cmds[ref.cmd]);
    if (!next || next.c === "Z" || next.c === "M") {
      const m = cmds.slice(0, ref.cmd + 1).reverse().find((c) => c.c === "M");
      if (!m || !next || next.c !== "Z") return p;
      const b = endOf(m);
      cmds.splice(ni, 0, { c: "L", pts: [(a[0] + b[0]) / 2, (a[1] + b[1]) / 2] });
      return { ...p, cmds };
    }
    if (next.c === "Q") {
      // İkinci dereceden eğri kübiğe çevrilir.
      const [qx, qy, ex, ey] = next.pts;
      next = { c: "C", pts: [a[0] + ((qx - a[0]) * 2) / 3, a[1] + ((qy - a[1]) * 2) / 3, ex + ((qx - ex) * 2) / 3, ey + ((qy - ey) * 2) / 3, ex, ey] };
    }
    if (next.c === "C") {
      const [x1, y1, x2, y2, x3, y3] = next.pts;
      const m = (u: number, v: number) => (u + v) / 2;
      const p01 = [m(a[0], x1), m(a[1], y1)];
      const p12 = [m(x1, x2), m(y1, y2)];
      const p23 = [m(x2, x3), m(y2, y3)];
      const q0 = [m(p01[0], p12[0]), m(p01[1], p12[1])];
      const q1 = [m(p12[0], p23[0]), m(p12[1], p23[1])];
      const mid = [m(q0[0], q1[0]), m(q0[1], q1[1])];
      cmds.splice(ni, 1, { c: "C", pts: [...p01, ...q0, ...mid] }, { c: "C", pts: [...q1, ...p23, x3, y3] });
      return { ...p, cmds };
    }
    const b = endOf(next);
    cmds.splice(ni, 0, { c: "L", pts: [(a[0] + b[0]) / 2, (a[1] + b[1]) / 2] });
    return { ...p, cmds };
  });
}

/** Noktayı siler (yolda en az iki nokta kalır; ilk nokta silinirse sonraki nokta başlangıç olur). */
export function deleteVertex(s: SpShape, ref: VertexRef): SpShape {
  const p = s.paths?.[ref.path];
  if (!p) return s;
  const count = p.cmds.filter((c) => c.c !== "Z" && c.pts.length >= 2).length;
  if (count <= 2) return s;
  return mapPath(s, ref.path, (pp) => {
    const cmds = pp.cmds.slice();
    const cur = cmds[ref.cmd];
    if (cur.c === "M") {
      const nx = cmds[ref.cmd + 1];
      if (!nx || nx.c === "Z") return pp;
      cmds.splice(ref.cmd, 2, { c: "M", pts: endOf(nx) });
    } else cmds.splice(ref.cmd, 1);
    return { ...pp, cmds };
  });
}

/** Yolu Aç / Yolu Kapat: kapanış (Z) komutu kaldırılır ya da eklenir; açık yol dolgusuz çizilir. */
export function togglePathClosed(s: SpShape, pi: number): SpShape {
  const closed = isPathClosed(s, pi);
  return mapPath(s, pi, (p) => (closed ? { ...p, cmds: p.cmds.filter((c) => c.c !== "Z"), fill: false } : { ...p, cmds: [...p.cmds, { c: "Z", pts: [] }], fill: true }));
}

/** Kesimi (noktaya gelen) düz ya da eğri yapar. */
export function setSegmentKind(s: SpShape, ref: VertexRef, kind: "straight" | "curved"): SpShape {
  return mapPath(s, ref.path, (p) => {
    const cmds = p.cmds.slice();
    const c = cmds[ref.cmd];
    const a = startOf(cmds, ref.cmd);
    if (!c || !a || c.c === "M" || c.c === "Z") return p;
    const b = endOf(c);
    if (kind === "straight") cmds[ref.cmd] = { c: "L", pts: b };
    else if (c.c === "L") {
      // Hafif bükülmüş eğri: denetim noktaları kesime dik yönde kaydırılır.
      const dx = b[0] - a[0];
      const dy = b[1] - a[1];
      const k = 0.2;
      cmds[ref.cmd] = { c: "C", pts: [a[0] + dx / 3 - dy * k, a[1] + dy / 3 + dx * k, a[0] + (2 * dx) / 3 - dy * k, a[1] + (2 * dy) / 3 + dx * k, b[0], b[1]] };
    }
    return { ...p, cmds };
  });
}

/** Nokta türü: düzgün (eşit olmayan, aynı doğrultuda tutamaçlar), düz (eşit tutamaçlar) ya da köşe. */
export function setVertexKind(s: SpShape, ref: VertexRef, kind: "smooth" | "straight" | "corner"): SpShape {
  return mapPath(s, ref.path, (p) => {
    const cmds = p.cmds.slice();
    const inC = cmds[ref.cmd];
    const outC = cmds[ref.cmd + 1];
    const v = endOf(inC);
    const prev = startOf(cmds, ref.cmd);
    const next = outC && outC.c !== "Z" && outC.c !== "M" ? endOf(outC) : null;
    const toC = (i: number, from: P2, to: P2): void => {
      const c = cmds[i];
      if (c.c === "C") return;
      cmds[i] = { c: "C", pts: [from[0] + (to[0] - from[0]) / 3, from[1] + (to[1] - from[1]) / 3, from[0] + ((to[0] - from[0]) * 2) / 3, from[1] + ((to[1] - from[1]) * 2) / 3, to[0], to[1]] };
    };
    if (inC.c !== "M" && prev) toC(ref.cmd, prev, v);
    if (next) toC(ref.cmd + 1, v, next);
    const hasIn = cmds[ref.cmd].c === "C" && !!prev;
    const hasOut = !!next && cmds[ref.cmd + 1].c === "C";
    if (kind === "corner") {
      // Tutamaçlar komşu noktalara yönelir: köşe keskinleşir.
      if (hasIn && prev) cmds[ref.cmd] = { c: "C", pts: [cmds[ref.cmd].pts[0], cmds[ref.cmd].pts[1], v[0] + (prev[0] - v[0]) / 3, v[1] + (prev[1] - v[1]) / 3, v[0], v[1]] };
      if (hasOut && next) {
        const o = cmds[ref.cmd + 1].pts;
        cmds[ref.cmd + 1] = { c: "C", pts: [v[0] + (next[0] - v[0]) / 3, v[1] + (next[1] - v[1]) / 3, o[2], o[3], o[4], o[5]] };
      }
      return { ...p, cmds };
    }
    if (!prev || !next || !hasIn || !hasOut) return { ...p, cmds };
    let tx = next[0] - prev[0];
    let ty = next[1] - prev[1];
    const tl = Math.hypot(tx, ty) || 1;
    tx /= tl;
    ty /= tl;
    let lin = Math.hypot(v[0] - prev[0], v[1] - prev[1]) / 3;
    let lout = Math.hypot(next[0] - v[0], next[1] - v[1]) / 3;
    if (kind === "straight") lin = lout = (lin + lout) / 2;
    const ip = cmds[ref.cmd].pts;
    cmds[ref.cmd] = { c: "C", pts: [ip[0], ip[1], v[0] - tx * lin, v[1] - ty * lin, v[0], v[1]] };
    const op = cmds[ref.cmd + 1].pts;
    cmds[ref.cmd + 1] = { c: "C", pts: [v[0] + tx * lout, v[1] + ty * lout, op[2], op[3], op[4], op[5]] };
    return { ...p, cmds };
  });
}
