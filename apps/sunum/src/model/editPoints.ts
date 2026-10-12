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
    p.cmds.forEach((c, ci) => {
      if (c.c === "Z" || c.pts.length < 2) return;
      // Kapanan yolun son L komutu başlangıçla çakışıyorsa tekrar gösterme.
      const x = c.pts[c.pts.length - 2];
      const y = c.pts[c.pts.length - 1];
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
