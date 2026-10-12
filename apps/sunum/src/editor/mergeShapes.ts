/**
 * Şekilleri Birleştir: Birleşim, Birleştir (tek-çift), Parçala, Kesişim, Çıkar.
 * Şekiller slayt koordinatlarında çokgenlere düzleştirilir (eğriler
 * örneklenir); iki çokgen arasında Greiner–Hormann kırpma uygulanır.
 * Sonuç, ilk seçilen şeklin biçimini alan tek bir özel şekildir.
 */
import { uid } from "../model/factory";
import type { CustomPath, PathCmd, SpShape } from "../model/types";
import { shapeGeometry } from "../render/SlideView";
import { svgPathToCmds } from "../render/svgPath";
import type { Actions } from "../state/actions";

export type Pt = [number, number];
export type Ring = Pt[];
export type MergeOp = "union" | "combine" | "intersect" | "subtract" | "fragment";

/* ------------------------- Düzleştirme ------------------------- */

function sampleCubic(a: Pt, c1: Pt, c2: Pt, b: Pt, n: number, out: Ring) {
  for (let i = 1; i <= n; i++) {
    const t = i / n;
    const u = 1 - t;
    out.push([u * u * u * a[0] + 3 * u * u * t * c1[0] + 3 * u * t * t * c2[0] + t * t * t * b[0], u * u * u * a[1] + 3 * u * u * t * c1[1] + 3 * u * t * t * c2[1] + t * t * t * b[1]]);
  }
}

/** Komut listesini (şekil koordinatlarında) kapalı halkalara çevirir. */
export function cmdsToRings(cmds: PathCmd[], steps = 16): Ring[] {
  const rings: Ring[] = [];
  let cur: Ring = [];
  const flush = () => {
    if (cur.length >= 3) rings.push(cur);
    cur = [];
  };
  for (const c of cmds) {
    const last = cur[cur.length - 1];
    if (c.c === "M") {
      flush();
      cur.push([c.pts[0], c.pts[1]]);
    } else if (c.c === "L") cur.push([c.pts[0], c.pts[1]]);
    else if (c.c === "C" && last) sampleCubic(last, [c.pts[0], c.pts[1]], [c.pts[2], c.pts[3]], [c.pts[4], c.pts[5]], steps, cur);
    else if (c.c === "Q" && last) {
      const q: Pt = [c.pts[0], c.pts[1]];
      const b: Pt = [c.pts[2], c.pts[3]];
      sampleCubic(last, [last[0] + (2 / 3) * (q[0] - last[0]), last[1] + (2 / 3) * (q[1] - last[1])], [b[0] + (2 / 3) * (q[0] - b[0]), b[1] + (2 / 3) * (q[1] - b[1])], b, steps, cur);
    } else if (c.c === "Z") flush();
  }
  flush();
  // Ardışık yinelenen noktaları ve kapanış tekrarını temizle.
  return rings
    .map((r) => r.filter((p, i) => i === 0 || Math.hypot(p[0] - r[i - 1][0], p[1] - r[i - 1][1]) > 1e-6))
    .map((r) => (r.length > 1 && Math.hypot(r[0][0] - r[r.length - 1][0], r[0][1] - r[r.length - 1][1]) < 1e-6 ? r.slice(0, -1) : r))
    .filter((r) => r.length >= 3);
}

/** Şeklin dolgulu bölgesini slayt koordinatlarında halkalara çevirir (döndürme/çevirme uygulanır). */
export function shapeRings(s: SpShape): Ring[] {
  const geo = shapeGeometry(s, s.w, s.h);
  const cx = s.w / 2;
  const cy = s.h / 2;
  const rad = (s.rot * Math.PI) / 180;
  const cos = Math.cos(rad);
  const sin = Math.sin(rad);
  const tf = ([x, y]: Pt): Pt => {
    let dx = x - cx;
    let dy = y - cy;
    if (s.flipH) dx = -dx;
    if (s.flipV) dy = -dy;
    return [s.x + cx + dx * cos - dy * sin, s.y + cy + dx * sin + dy * cos];
  };
  const out: Ring[] = [];
  for (const p of geo.paths) {
    if (p.fill === "none" && geo.paths.length > 1) continue;
    for (const r of cmdsToRings(svgPathToCmds(p.d))) out.push(r.map(tf));
  }
  return out;
}

/* ------------------------- Geometri yardımcıları ------------------------- */

export function area(r: Ring): number {
  let a = 0;
  for (let i = 0; i < r.length; i++) {
    const [x1, y1] = r[i];
    const [x2, y2] = r[(i + 1) % r.length];
    a += x1 * y2 - x2 * y1;
  }
  return a / 2;
}

export function pointInRing(p: Pt, r: Ring): boolean {
  let inside = false;
  for (let i = 0, j = r.length - 1; i < r.length; j = i++) {
    const [xi, yi] = r[i];
    const [xj, yj] = r[j];
    if (yi > p[1] !== yj > p[1] && p[0] < ((xj - xi) * (p[1] - yi)) / (yj - yi) + xi) inside = !inside;
  }
  return inside;
}

/* ------------------------- Greiner–Hormann ------------------------- */

interface Node {
  p: Pt;
  next: Node;
  prev: Node;
  isect: boolean;
  entry: boolean;
  neighbor: Node | null;
  alpha: number;
  visited: boolean;
}

function makeList(r: Ring): Node {
  let first: Node | null = null;
  let prev: Node | null = null;
  for (const p of r) {
    const n = { p, isect: false, entry: false, neighbor: null, alpha: 0, visited: false } as unknown as Node;
    if (!first) first = n;
    if (prev) {
      prev.next = n;
      n.prev = prev;
    }
    prev = n;
  }
  prev!.next = first!;
  first!.prev = prev!;
  return first!;
}

function nextNonIsect(n: Node): Node {
  let c = n;
  while (c.isect) c = c.next;
  return c;
}

function insertBetween(n: Node, a: Node, b: Node) {
  let c = a.next;
  while (c !== b && c.alpha < n.alpha) c = c.next;
  n.next = c;
  n.prev = c.prev;
  c.prev.next = n;
  c.prev = n;
}

function segIntersect(a1: Pt, a2: Pt, b1: Pt, b2: Pt): { ta: number; tb: number; p: Pt } | null {
  const d = (b2[1] - b1[1]) * (a2[0] - a1[0]) - (b2[0] - b1[0]) * (a2[1] - a1[1]);
  if (Math.abs(d) < 1e-12) return null;
  const ta = ((b2[0] - b1[0]) * (a1[1] - b1[1]) - (b2[1] - b1[1]) * (a1[0] - b1[0])) / d;
  const tb = ((a2[0] - a1[0]) * (a1[1] - b1[1]) - (a2[1] - a1[1]) * (a1[0] - b1[0])) / d;
  if (ta <= 0 || ta >= 1 || tb <= 0 || tb >= 1) return null;
  return { ta, tb, p: [a1[0] + ta * (a2[0] - a1[0]), a1[1] + ta * (a2[1] - a1[1])] };
}

/** İki basit halka arasında kırpma; kesişme yoksa null (çağıran kapsama durumunu çözer). */
function clipRings(subject: Ring, clip: Ring, op: "union" | "intersect" | "subtract"): Ring[] | null {
  const s = makeList(subject);
  const c = makeList(clip);
  let found = false;
  // Kesişimleri iki listeye ekle.
  for (let a = s; ; ) {
    if (!a.isect) {
      const a2 = nextNonIsect(a.next);
      for (let b = c; ; ) {
        if (!b.isect) {
          const b2 = nextNonIsect(b.next);
          const hit = segIntersect(a.p, a2.p, b.p, b2.p);
          if (hit) {
            found = true;
            const n1 = { p: hit.p, isect: true, entry: false, alpha: hit.ta, visited: false } as unknown as Node;
            const n2 = { p: hit.p, isect: true, entry: false, alpha: hit.tb, visited: false } as unknown as Node;
            n1.neighbor = n2;
            n2.neighbor = n1;
            insertBetween(n1, a, a2);
            insertBetween(n2, b, b2);
          }
        }
        b = b.next;
        if (b === c) break;
      }
    }
    a = a.next;
    if (a === s) break;
  }
  if (!found) return null;
  // Giriş/çıkış işaretleri.
  const mark = (list: Node, other: Ring, invert: boolean) => {
    let inside = pointInRing(list.p, other);
    if (invert) inside = !inside;
    for (let n = list; ; ) {
      if (n.isect) {
        n.entry = !inside;
        inside = !inside;
      }
      n = n.next;
      if (n === list) break;
    }
  };
  // Birleşim: ikisinin de dışından; kesişim: içinden; fark: öznede dışarı, kırpmada içeri.
  mark(s, clip, op === "union" || op === "subtract");
  mark(c, subject, op === "union");
  const out: Ring[] = [];
  for (let n = s; ; ) {
    if (n.isect && !n.visited) {
      const ring: Ring = [];
      let cur: Node = n;
      ring.push(cur.p);
      let guard = 0;
      do {
        cur.visited = true;
        if (cur.neighbor) cur.neighbor.visited = true;
        if (cur.entry) {
          do {
            cur = cur.next;
            ring.push(cur.p);
          } while (!cur.isect && guard++ < 100000);
        } else {
          do {
            cur = cur.prev;
            ring.push(cur.p);
          } while (!cur.isect && guard++ < 100000);
        }
        cur.visited = true;
        cur = cur.neighbor!;
      } while (!cur.visited && guard++ < 100000);
      if (ring.length >= 3) out.push(ring);
    }
    n = n.next;
    if (n === s) break;
  }
  return out;
}

/** Dejenere durumları (çakışan kenarlar/köşeler) önlemek için çok küçük kaydırma. */
function jitter(r: Ring, k: number): Ring {
  return r.map(([x, y], i) => [x + 1e-7 * k * (1 + (i % 3)), y - 1e-7 * k * (1 + (i % 5))]);
}

const contains = (a: Ring, b: Ring) => b.every((p) => pointInRing(p, a));

/** İki halka üzerinde işlem (kapsama ve ayrıklık durumları dahil). */
export function ringOp(a: Ring, b: Ring, op: "union" | "intersect" | "subtract"): Ring[] {
  const res = clipRings(a, jitter(b, 1), op);
  if (res) return res;
  const aInB = contains(b, a);
  const bInA = contains(a, b);
  if (op === "union") return aInB ? [b] : bInA ? [a] : [a, b];
  if (op === "intersect") return aInB ? [a] : bInA ? [b] : [];
  // Fark: B, A'nın içindeyse delik olur.
  return aInB ? [] : bInA ? [a, b.slice().reverse()] : [a];
}

/** Halka listeleri (her biri bir şekil) üzerinde birleştirme işlemi. */
export function mergeRings(shapes: Ring[][], op: MergeOp): Ring[] {
  if (!shapes.length) return [];
  if (op === "combine") return shapes.flat();
  let acc = shapes[0];
  for (const next of shapes.slice(1)) {
    if (op === "union") {
      let pool = [...acc, ...next];
      // Çakışan halkaları birleşene kadar birleştir.
      let merged = true;
      let guard = 0;
      while (merged && guard++ < 200) {
        merged = false;
        outer: for (let i = 0; i < pool.length; i++)
          for (let j = i + 1; j < pool.length; j++) {
            const r = ringOp(pool[i], pool[j], "union");
            if (r.length === 1) {
              pool = [...pool.filter((_, k) => k !== i && k !== j), r[0]];
              merged = true;
              break outer;
            }
          }
      }
      acc = pool;
    } else if (op === "intersect") {
      const out: Ring[] = [];
      for (const a of acc) for (const b of next) out.push(...ringOp(a, b, "intersect"));
      acc = out;
    } else {
      let cur = acc;
      for (const b of next) cur = cur.flatMap((a) => ringOp(a, b, "subtract"));
      acc = cur;
    }
  }
  return acc.filter((r) => Math.abs(area(r)) > 0.5);
}

/** Parçala: şekillerin örtüşmesiyle oluşan her bölge ayrı bir halka kümesi olur. */
export function fragmentRings(shapes: Ring[][]): Ring[][] {
  let pieces: Ring[][] = [];
  const done: Ring[] = [];
  const big = (rs: Ring[]) => rs.filter((r) => Math.abs(area(r)) > 0.5);
  for (const b of shapes) {
    const next: Ring[][] = [];
    for (const p of pieces) {
      const inter = big(mergeRings([p, b], "intersect"));
      const diff = big(mergeRings([p, b], "subtract"));
      if (inter.length) next.push(inter);
      if (diff.length) next.push(diff);
    }
    const rest = done.length ? big(mergeRings([b, done], "subtract")) : b;
    if (rest.length) next.push(rest);
    done.push(...b);
    pieces = next;
  }
  // Ayrık bölgeler ayrı parçalara bölünür; delikler (başka halkanın içindeki halkalar) kapsayanına eklenir.
  const out: Ring[][] = [];
  for (const p of pieces) {
    const holes = new Set<Ring>();
    for (const r of p) if (p.some((o) => o !== r && Math.abs(area(o)) > Math.abs(area(r)) && pointInRing(r[0], o))) holes.add(r);
    for (const r of p) if (!holes.has(r)) out.push([r, ...[...holes].filter((h) => pointInRing(h[0], r))]);
  }
  return out;
}

function ringsToShape(rings: Ring[], first: SpShape): SpShape | null {
  if (!rings.length) return null;
  let minX = Infinity;
  let minY = Infinity;
  let maxX = -Infinity;
  let maxY = -Infinity;
  for (const r of rings)
    for (const [x, y] of r) {
      minX = Math.min(minX, x);
      minY = Math.min(minY, y);
      maxX = Math.max(maxX, x);
      maxY = Math.max(maxY, y);
    }
  const w = Math.max(1, maxX - minX);
  const h = Math.max(1, maxY - minY);
  const cmds: PathCmd[] = [];
  const r2 = (v: number) => Math.round(v * 100) / 100;
  for (const r of rings) {
    r.forEach(([x, y], i) => cmds.push({ c: i ? "L" : "M", pts: [r2(x - minX), r2(y - minY)] }));
    cmds.push({ c: "Z", pts: [] });
  }
  const path: CustomPath = { w, h, cmds };
  return { ...first, id: uid(), x: minX, y: minY, w, h, rot: 0, flipH: undefined, flipV: undefined, geom: "custom", adj: undefined, paths: [path], text: undefined, ph: undefined, txBox: undefined, name: `Serbest Biçim ${first.name.replace(/\D+/g, "") || ""}`.trim() };
}

/** Seçili şekilleri birleştirip yeni özel şekil(ler) döndürür (Parçala birden çok şekil verir). */
export function mergeShapes(list: SpShape[], op: MergeOp): SpShape[] {
  if (list.length < 2) return [];
  const sets = op === "fragment" ? fragmentRings(list.map(shapeRings)) : [mergeRings(list.map(shapeRings), op)];
  return sets.map((rs) => ringsToShape(rs, list[0])).filter((x): x is SpShape => !!x);
}

/** Seçili şekilleri birleştirir; sonuç ilk seçilen şeklin yerine ve biçimiyle eklenir. */
export function mergeSelection(a: Actions, op: MergeOp): boolean {
  const sl = a.slide();
  if (!sl) return false;
  const sel = a.s.sel;
  const list = sel.map((id) => sl.shapes.find((s) => s.id === id)).filter((s): s is SpShape => !!s && s.kind === "sp" && !s.ink);
  if (list.length < 2) return false;
  const merged = mergeShapes(list, op);
  const ids = new Set(list.map((s) => s.id));
  const at = Math.min(...list.map((s) => sl.shapes.indexOf(s)));
  const i = a.s.slideIdx;
  a.store.commit(
    (p) => {
      const slides = p.slides.slice();
      const cur = slides[i];
      const rest = cur.shapes.filter((s) => !ids.has(s.id));
      const shapes = [...rest.slice(0, at), ...merged, ...rest.slice(at)];
      slides[i] = { ...cur, shapes, animations: cur.animations.filter((an) => !ids.has(an.shapeId)) };
      return { ...p, slides };
    },
    { sel: merged.map((m) => m.id), editing: null },
  );
  return true;
}
