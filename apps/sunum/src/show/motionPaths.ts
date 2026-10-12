/**
 * Hareket Yolları: yollar ECMA-376 p:animMotion biçiminde (slayt genişlik ve
 * yüksekliğine göre kesirli, "M 0 0 L 0 0.25 E") tutulur. Gösteride yol
 * örneklenip şeklin öteleme anahtar karelerine çevrilir; düzenleyicide yol
 * kesik çizgiyle gösterilir.
 */
import type { Animation } from "../model/types";

const LINE: Record<string, [number, number]> = { down: [0, 0.25], up: [0, -0.25], left: [-0.25, 0], right: [0.25, 0] };

/** Hazır yol (yön seçeneğiyle) ya da animasyonun kendi yolu. */
export function motionPathOf(a: Pick<Animation, "effect" | "option" | "path">, W = 1280, H = 720): string {
  if (a.path) return a.path;
  const o = a.option ?? "down";
  const ar = W / H;
  switch (a.effect) {
    case "pathLine": {
      const [x, y] = LINE[o] ?? LINE.down;
      return `M 0 0 L ${x} ${y} E`;
    }
    case "pathArc": {
      // Yön boyunca giden ve yan tarafa kavis yapan yay.
      const [x, y] = LINE[o] ?? LINE.down;
      const k = 0.55;
      const c1 = x ? [x * 0.1, (Math.abs(x) * k * ar) / 2] : [(Math.abs(y) * k) / ar / 2, y * 0.1];
      const c2 = x ? [x * 0.9, (Math.abs(x) * k * ar) / 2] : [(Math.abs(y) * k) / ar / 2, y * 0.9];
      return `M 0 0 C ${f(c1[0])} ${f(c1[1])} ${f(c2[0])} ${f(c2[1])} ${f(x)} ${f(y)} E`;
    }
    case "pathCircle": {
      // Saat yönünde daire: başlangıç noktası üstte, merkez aşağıda.
      const ry = 0.15;
      const rx = ry / ar;
      const c = 0.5523;
      return `M 0 0 C ${f(rx * c)} 0 ${f(rx)} ${f(ry * (1 - c))} ${f(rx)} ${f(ry)} C ${f(rx)} ${f(ry * (1 + c))} ${f(rx * c)} ${f(2 * ry)} 0 ${f(2 * ry)} C ${f(-rx * c)} ${f(2 * ry)} ${f(-rx)} ${f(ry * (1 + c))} ${f(-rx)} ${f(ry)} C ${f(-rx)} ${f(ry * (1 - c))} ${f(-rx * c)} 0 0 0 Z`;
    }
    default:
      return "M 0 0 L 0 0.25 E";
  }
}

const f = (n: number) => String(Math.round(n * 10000) / 10000);

/** Yolu piksel noktalarına örnekler (W × H slayt). */
export function samplePath(path: string, W: number, H: number, steps = 16): [number, number][] {
  const tok = path.replace(/([MLCZE])/gi, " $1 ").trim().split(/[\s,]+/).filter(Boolean);
  const pts: [number, number][] = [];
  let cur: [number, number] = [0, 0];
  let start: [number, number] = [0, 0];
  let i = 0;
  const num = () => Number(tok[i++]);
  while (i < tok.length) {
    const c = tok[i++].toUpperCase();
    if (c === "M") {
      cur = [num() * W, num() * H];
      start = cur;
      pts.push(cur);
    } else if (c === "L") {
      cur = [num() * W, num() * H];
      pts.push(cur);
    } else if (c === "C") {
      const p1: [number, number] = [num() * W, num() * H];
      const p2: [number, number] = [num() * W, num() * H];
      const p3: [number, number] = [num() * W, num() * H];
      const p0 = cur;
      for (let k = 1; k <= steps; k++) {
        const t = k / steps;
        const u = 1 - t;
        pts.push([u * u * u * p0[0] + 3 * u * u * t * p1[0] + 3 * u * t * t * p2[0] + t * t * t * p3[0], u * u * u * p0[1] + 3 * u * u * t * p1[1] + 3 * u * t * t * p2[1] + t * t * t * p3[1]]);
      }
      cur = p3;
    } else if (c === "Z") {
      cur = start;
      pts.push(cur);
    } else if (c === "E") break;
    else if (!Number.isNaN(Number(c))) i--; // yinelenen koordinat (komutsuz)
  }
  return pts.length ? pts : [[0, 0]];
}

/** Gösteride öteleme anahtar kareleri: noktalar yay uzunluğuna göre zamana yayılır. */
export function motionKeyframes(a: Animation, W: number, H: number): Keyframe[] {
  const pts = samplePath(motionPathOf(a, W, H), W, H);
  const len = [0];
  for (let i = 1; i < pts.length; i++) len.push(len[i - 1] + Math.hypot(pts[i][0] - pts[i - 1][0], pts[i][1] - pts[i - 1][1]));
  const total = len[len.length - 1] || 1;
  return pts.map((p, i) => ({ translate: `${p[0].toFixed(2)}px ${p[1].toFixed(2)}px`, offset: len[i] / total }));
}

/** İçe aktarılan yolu hazır yola eşler (eşleşmezse özel yol olarak tutulur). */
export function classifyPath(path: string): { effect: string; option?: string; path?: string } {
  const p = path.trim();
  for (const eff of ["pathLine", "pathArc", "pathCircle"])
    for (const o of eff === "pathCircle" ? ["circle"] : ["down", "up", "left", "right"]) if (norm(motionPathOf({ effect: eff, option: o })) === norm(p)) return { effect: eff, option: eff === "pathCircle" ? undefined : o };
  return { effect: "pathCustom", path: p };
}

const norm = (s: string) => s.replace(/\s+/g, " ").trim().toUpperCase();

/** Düzenleyicide yolun SVG çizimi (şekil merkezinden başlayarak). */
export function motionPathSvg(a: Animation, cx: number, cy: number, W: number, H: number): { d: string; end: [number, number] } {
  const pts = samplePath(motionPathOf(a, W, H), W, H);
  const d = pts.map((p, i) => `${i ? "L" : "M"}${(cx + p[0]).toFixed(1)} ${(cy + p[1]).toFixed(1)}`).join(" ");
  const e = pts[pts.length - 1];
  return { d, end: [cx + e[0], cy + e[1]] };
}
