/** SVG yol verisini (d) mutlak M/L/C/Q/Z komutlarına çevirir (yay → Bézier). */
import type { PathCmd } from "../model/types";

function tokenize(d: string): (string | number)[] {
  const out: (string | number)[] = [];
  const re = /([MmLlHhVvCcSsQqTtAaZz])|(-?(?:\d+\.?\d*|\.\d+)(?:[eE][-+]?\d+)?)/g;
  let m: RegExpExecArray | null;
  while ((m = re.exec(d))) out.push(m[1] ?? Number(m[2]));
  return out;
}

function arcToCubic(x1: number, y1: number, rx: number, ry: number, phi: number, fa: number, fs: number, x2: number, y2: number): number[][] {
  if (rx === 0 || ry === 0) return [[x1, y1, x2, y2, x2, y2]];
  const rad = (phi * Math.PI) / 180;
  const cos = Math.cos(rad);
  const sin = Math.sin(rad);
  const dx = (x1 - x2) / 2;
  const dy = (y1 - y2) / 2;
  const x1p = cos * dx + sin * dy;
  const y1p = -sin * dx + cos * dy;
  rx = Math.abs(rx);
  ry = Math.abs(ry);
  const lam = (x1p * x1p) / (rx * rx) + (y1p * y1p) / (ry * ry);
  if (lam > 1) {
    rx *= Math.sqrt(lam);
    ry *= Math.sqrt(lam);
  }
  const sign = fa === fs ? -1 : 1;
  const num = rx * rx * ry * ry - rx * rx * y1p * y1p - ry * ry * x1p * x1p;
  const den = rx * rx * y1p * y1p + ry * ry * x1p * x1p;
  const coef = sign * Math.sqrt(Math.max(0, num / (den || 1)));
  const cxp = (coef * rx * y1p) / ry;
  const cyp = (-coef * ry * x1p) / rx;
  const cx = cos * cxp - sin * cyp + (x1 + x2) / 2;
  const cy = sin * cxp + cos * cyp + (y1 + y2) / 2;
  const ang = (ux: number, uy: number, vx: number, vy: number) => {
    const a = Math.atan2(ux * vy - uy * vx, ux * vx + uy * vy);
    return a;
  };
  const th1 = ang(1, 0, (x1p - cxp) / rx, (y1p - cyp) / ry);
  let dth = ang((x1p - cxp) / rx, (y1p - cyp) / ry, (-x1p - cxp) / rx, (-y1p - cyp) / ry);
  if (!fs && dth > 0) dth -= 2 * Math.PI;
  if (fs && dth < 0) dth += 2 * Math.PI;
  const segs = Math.max(1, Math.ceil(Math.abs(dth) / (Math.PI / 2)));
  const d = dth / segs;
  const k = (4 / 3) * Math.tan(d / 4);
  const out: number[][] = [];
  let t = th1;
  const pt = (a: number): [number, number] => [cx + rx * Math.cos(a) * cos - ry * Math.sin(a) * sin, cy + rx * Math.cos(a) * sin + ry * Math.sin(a) * cos];
  const dpt = (a: number): [number, number] => [-rx * Math.sin(a) * cos - ry * Math.cos(a) * sin, -rx * Math.sin(a) * sin + ry * Math.cos(a) * cos];
  for (let i = 0; i < segs; i++) {
    const t2 = t + d;
    const p1 = pt(t);
    const p2 = pt(t2);
    const d1 = dpt(t);
    const d2 = dpt(t2);
    out.push([p1[0] + k * d1[0], p1[1] + k * d1[1], p2[0] - k * d2[0], p2[1] - k * d2[1], p2[0], p2[1]]);
    t = t2;
  }
  return out;
}

export function svgPathToCmds(d: string): PathCmd[] {
  const tk = tokenize(d);
  const out: PathCmd[] = [];
  let i = 0;
  let cmd = "";
  let x = 0;
  let y = 0;
  let sx = 0;
  let sy = 0;
  let lastC: [number, number] | null = null;
  let lastQ: [number, number] | null = null;
  const num = () => {
    const v = tk[i++];
    return typeof v === "number" ? v : 0;
  };
  while (i < tk.length) {
    if (typeof tk[i] === "string") cmd = tk[i++] as string;
    const rel = cmd === cmd.toLowerCase();
    const C = cmd.toUpperCase();
    const ox = rel ? x : 0;
    const oy = rel ? y : 0;
    switch (C) {
      case "M": {
        x = ox + num();
        y = oy + num();
        sx = x;
        sy = y;
        out.push({ c: "M", pts: [x, y] });
        cmd = rel ? "l" : "L";
        lastC = lastQ = null;
        break;
      }
      case "L":
        x = ox + num();
        y = oy + num();
        out.push({ c: "L", pts: [x, y] });
        lastC = lastQ = null;
        break;
      case "H":
        x = (rel ? x : 0) + num();
        out.push({ c: "L", pts: [x, y] });
        lastC = lastQ = null;
        break;
      case "V":
        y = (rel ? y : 0) + num();
        out.push({ c: "L", pts: [x, y] });
        lastC = lastQ = null;
        break;
      case "C": {
        const p = [ox + num(), oy + num(), ox + num(), oy + num(), ox + num(), oy + num()];
        out.push({ c: "C", pts: p });
        lastC = [p[2], p[3]];
        lastQ = null;
        x = p[4];
        y = p[5];
        break;
      }
      case "S": {
        const c1: number[] = lastC ? [2 * x - lastC[0], 2 * y - lastC[1]] : [x, y];
        const p: number[] = [c1[0], c1[1], ox + num(), oy + num(), ox + num(), oy + num()];
        out.push({ c: "C", pts: p });
        lastC = [p[2], p[3]];
        lastQ = null;
        x = p[4];
        y = p[5];
        break;
      }
      case "Q": {
        const p = [ox + num(), oy + num(), ox + num(), oy + num()];
        out.push({ c: "Q", pts: p });
        lastQ = [p[0], p[1]];
        lastC = null;
        x = p[2];
        y = p[3];
        break;
      }
      case "T": {
        const c1: number[] = lastQ ? [2 * x - lastQ[0], 2 * y - lastQ[1]] : [x, y];
        const p: number[] = [c1[0], c1[1], ox + num(), oy + num()];
        out.push({ c: "Q", pts: p });
        lastQ = [p[0], p[1]];
        lastC = null;
        x = p[2];
        y = p[3];
        break;
      }
      case "A": {
        const rx = num();
        const ry = num();
        const phi = num();
        const fa = num();
        const fs = num();
        const nx = ox + num();
        const ny = oy + num();
        for (const b of arcToCubic(x, y, rx, ry, phi, fa, fs, nx, ny)) out.push({ c: "C", pts: b });
        x = nx;
        y = ny;
        lastC = lastQ = null;
        break;
      }
      case "Z":
        out.push({ c: "Z", pts: [] });
        x = sx;
        y = sy;
        lastC = lastQ = null;
        break;
      default:
        i++;
    }
  }
  return out;
}
