/**
 * Gösteri efektleri: Web Animations API anahtar kareleri. Bireysel
 * dönüşüm özellikleri (translate/scale/rotate) kullanılır; böylece şeklin
 * kendi döndürmesi (transform) korunur. Desenli geçişler clip-path çokgen
 * kareleriyle yapılır.
 */
import type { Animation, Slide, Transition } from "../model/types";

type KF = Keyframe[];

const pct = (v: number) => `${(v * 100).toFixed(3)}%`;

/** Rastgele ama tekrarlanabilir sayılar. */
function rng(seed: number) {
  let s = seed >>> 0 || 1;
  return () => {
    s ^= s << 13;
    s ^= s >>> 17;
    s ^= s << 5;
    return ((s >>> 0) % 10000) / 10000;
  };
}

function poly(points: [number, number][]): string {
  return `polygon(${points.map(([x, y]) => `${pct(x)} ${pct(y)}`).join(", ")})`;
}

/** n şerit; her şerit kendi içinde [a,b] aralığını kaplar. */
function barsPolygon(bars: [number, number][], horizontal: boolean): string {
  const pts: [number, number][] = [[0, 0]];
  for (const [a, b] of bars) {
    if (horizontal) pts.push([0, a], [1, a], [1, b], [0, b], [0, a]);
    else pts.push([a, 0], [a, 1], [b, 1], [b, 0], [a, 0]);
  }
  pts.push([0, 0]);
  return poly(pts);
}

function frames(n: number, fn: (t: number) => string): KF {
  const out: KF = [];
  for (let i = 0; i <= n; i++) out.push({ clipPath: fn(i / n), offset: i / n });
  return out;
}

export function blindsFrames(horizontal: boolean, count = 6): KF {
  return frames(10, (t) =>
    barsPolygon(
      Array.from({ length: count }, (_, i) => [i / count, i / count + t / count] as [number, number]),
      horizontal,
    ),
  );
}

export function randomBarFrames(horizontal: boolean, seed = 7): KF {
  const n = 40;
  const r = rng(seed);
  const order = Array.from({ length: n }, (_, i) => ({ i, at: r() }));
  return frames(12, (t) => barsPolygon(order.map((o) => [o.i / n, o.at < t ? (o.i + 1) / n : o.i / n] as [number, number]), horizontal));
}

export function checkerFrames(vertical: boolean): KF {
  const cols = 10;
  const rows = 6;
  return frames(10, (t) => {
    const pts: [number, number][] = [[0, 0]];
    for (let rr = 0; rr < rows; rr++) {
      for (let c = 0; c < cols; c++) {
        const off = (rr + c) % 2 === 0 ? 0 : 0.5;
        const p = Math.max(0, Math.min(1, (t - off * 0.5) * 2 - off));
        const prog = Math.max(0, Math.min(1, t * 2 - off));
        void p;
        if (!vertical) {
          const x0 = c / cols;
          const x1 = x0 + prog / cols;
          const y0 = rr / rows;
          const y1 = (rr + 1) / rows;
          pts.push([x0, y0], [x1, y0], [x1, y1], [x0, y1], [x0, y0], [0, 0]);
        } else {
          const y0 = rr / rows;
          const y1 = y0 + prog / rows;
          const x0 = c / cols;
          const x1 = (c + 1) / cols;
          pts.push([x0, y0], [x1, y0], [x1, y1], [x0, y1], [x0, y0], [0, 0]);
        }
      }
    }
    return poly(pts);
  });
}

export function dissolveFrames(seed = 3): KF {
  const cols = 24;
  const rows = 14;
  const r = rng(seed);
  const at = Array.from({ length: cols * rows }, () => r());
  return frames(12, (t) => {
    const pts: [number, number][] = [[0, 0]];
    for (let i = 0; i < cols * rows; i++) {
      const c = i % cols;
      const rr = Math.floor(i / cols);
      const on = at[i] < t;
      const x0 = c / cols;
      const y0 = rr / rows;
      const x1 = on ? (c + 1) / cols : x0;
      const y1 = on ? (rr + 1) / rows : y0;
      pts.push([x0, y0], [x1, y0], [x1, y1], [x0, y1], [x0, y0], [0, 0]);
    }
    return poly(pts);
  });
}

function sweepPolygon(t: number, segments = 36, twoSided = false): string {
  const pts: [number, number][] = [[0.5, 0.5]];
  const R = 1.5;
  for (let i = 0; i <= segments; i++) {
    const a = -Math.PI / 2 + (twoSided ? -1 : 1) * Math.PI * (twoSided ? 1 : 2) * t * (i / segments);
    pts.push([0.5 + Math.cos(a) * R, 0.5 + Math.sin(a) * R]);
  }
  if (twoSided) {
    for (let i = segments; i >= 0; i--) {
      const a = -Math.PI / 2 + Math.PI * t * (i / segments);
      pts.push([0.5 + Math.cos(a) * R, 0.5 + Math.sin(a) * R]);
    }
  }
  return poly(pts);
}

export function wheelFrames(): KF {
  return frames(16, (t) => sweepPolygon(t));
}

export function wedgeFrames(): KF {
  return frames(16, (t) => sweepPolygon(t, 24, true));
}

function splitFrames(variant: string): KF {
  const horizontal = variant.startsWith("horz");
  const out = variant.endsWith("Out");
  return frames(8, (t) => {
    if (out) {
      const a = 0.5 - t / 2;
      const b = 0.5 + t / 2;
      return barsPolygon([[a, b]], horizontal);
    }
    return barsPolygon(
      [
        [0, t / 2],
        [1 - t / 2, 1],
      ],
      horizontal,
    );
  });
}

function shapeFrames(kind: string): KF {
  if (kind === "diamond")
    return [
      { clipPath: poly([[0.5, 0.5], [0.5, 0.5], [0.5, 0.5], [0.5, 0.5]]) },
      { clipPath: poly([[0.5, -0.55], [1.55, 0.5], [0.5, 1.55], [-0.55, 0.5]]) },
    ];
  if (kind === "plus") {
    const plus = (s: number): string => {
      const a = 0.5 - s;
      const b = 0.5 + s;
      const c = 0.5 - s * 2.2;
      const d = 0.5 + s * 2.2;
      return poly([[a, c], [b, c], [b, a], [d, a], [d, b], [b, b], [b, d], [a, d], [a, b], [c, b], [c, a], [a, a]]);
    };
    return [{ clipPath: plus(0) }, { clipPath: plus(0.5) }];
  }
  return [{ clipPath: "circle(0% at 50% 50%)" }, { clipPath: "circle(75% at 50% 50%)" }];
}

function wipeInset(dir: string, reveal: boolean): KF {
  // dir: kaynağın geldiği taraf ("l": soldan sağa açılır)
  const from = dir === "l" ? "inset(0 100% 0 0)" : dir === "r" ? "inset(0 0 0 100%)" : dir === "t" ? "inset(0 0 100% 0)" : "inset(100% 0 0 0)";
  return reveal ? [{ clipPath: from }, { clipPath: "inset(0 0 0 0)" }] : [{ clipPath: "inset(0 0 0 0)" }, { clipPath: from }];
}

/* ------------------------------------------------------------------ */
/* Nesne animasyonları                                                  */
/* ------------------------------------------------------------------ */

export interface ShapeGeomInfo {
  x: number;
  y: number;
  w: number;
  h: number;
}

function flyOffset(dir: string, g: ShapeGeomInfo, W: number, H: number): [number, number] {
  let dx = 0;
  let dy = 0;
  if (dir.includes("b")) dy = H - g.y + 10;
  if (dir.includes("t")) dy = -(g.y + g.h + 10);
  if (dir.includes("l")) dx = -(g.x + g.w + 10);
  if (dir.includes("r")) dx = W - g.x + 10;
  return [dx, dy];
}

const reverse = (k: KF): KF => k.slice().reverse().map((f, i, arr) => ({ ...f, offset: f.offset !== undefined ? 1 - (f.offset as number) : i / Math.max(1, arr.length - 1) }));

/** Animasyonun anahtar kareleri. Giriş efektleri görünmezden başlar. */
export function animationKeyframes(a: Animation, g: ShapeGeomInfo, W: number, H: number): { kf: KF; easing: string } {
  const opt = a.option ?? "";
  switch (a.effect) {
    case "appear":
      return { kf: [{ opacity: 0 }, { opacity: 1 }], easing: "step-end" };
    case "fade":
      return { kf: [{ opacity: 0 }, { opacity: 1 }], easing: "linear" };
    case "flyIn": {
      const [dx, dy] = flyOffset(opt || "b", g, W, H);
      return { kf: [{ translate: `${dx}px ${dy}px` }, { translate: "0 0" }], easing: "cubic-bezier(.2,.7,.3,1)" };
    }
    case "floatIn":
      return { kf: [{ opacity: 0, translate: `0 ${opt === "down" ? -60 : 60}px` }, { opacity: 1, translate: "0 0" }], easing: "ease-out" };
    case "split":
      return { kf: splitFrames(opt || "vertOut"), easing: "linear" };
    case "wipe":
      return { kf: wipeInset(opt || "b", true), easing: "linear" };
    case "shape":
      return { kf: shapeFrames(opt || "circle"), easing: "linear" };
    case "wheel":
      return { kf: wheelFrames(), easing: "linear" };
    case "randomBars":
      return { kf: randomBarFrames(opt !== "vert"), easing: "linear" };
    case "growTurn":
      return { kf: [{ opacity: 0, scale: "0", rotate: "-90deg" }, { opacity: 1, scale: "1", rotate: "0deg" }], easing: "ease-out" };
    case "zoom":
      return { kf: [{ opacity: 0, scale: "0.1" }, { opacity: 1, scale: "1" }], easing: "ease-out" };
    case "swivel":
      return { kf: [{ scale: "0 1" }, { scale: "1 1", offset: 0.3 }, { scale: "-1 1", offset: 0.5 }, { scale: "1 1" }], easing: "ease-in-out" };
    case "bounce":
      return {
        kf: [
          { translate: `0 ${-(g.y + g.h)}px`, opacity: 0 },
          { translate: "0 0", opacity: 1, offset: 0.45 },
          { translate: "0 -40px", offset: 0.62 },
          { translate: "0 0", offset: 0.78 },
          { translate: "0 -12px", offset: 0.89 },
          { translate: "0 0" },
        ],
        easing: "ease-in",
      };
    case "pulse":
      return { kf: [{ scale: "1" }, { scale: "1.06" }, { scale: "1" }], easing: "ease-in-out" };
    case "colorPulse":
      return { kf: [{ filter: "brightness(1)" }, { filter: "brightness(1.6) saturate(1.4)" }, { filter: "brightness(1)" }], easing: "ease-in-out" };
    case "teeter":
      return { kf: [{ rotate: "0deg" }, { rotate: "5deg" }, { rotate: "-5deg" }, { rotate: "4deg" }, { rotate: "-4deg" }, { rotate: "0deg" }], easing: "ease-in-out" };
    case "spin":
      return { kf: [{ rotate: "0deg" }, { rotate: "360deg" }], easing: "ease-in-out" };
    case "growShrink":
      return { kf: [{ scale: "1" }, { scale: "1.5" }], easing: "ease-in-out" };
    case "transparency":
      return { kf: [{ opacity: 1 }, { opacity: 0.5 }], easing: "linear" };
    case "disappear":
      return { kf: [{ opacity: 1 }, { opacity: 0 }], easing: "step-end" };
    case "fadeOut":
      return { kf: [{ opacity: 1 }, { opacity: 0 }], easing: "linear" };
    case "flyOut": {
      const [dx, dy] = flyOffset(opt || "b", g, W, H);
      return { kf: [{ translate: "0 0" }, { translate: `${dx}px ${dy}px` }], easing: "cubic-bezier(.7,0,.8,.3)" };
    }
    case "floatOut":
      return { kf: [{ opacity: 1, translate: "0 0" }, { opacity: 0, translate: `0 ${opt === "up" ? -60 : 60}px` }], easing: "ease-in" };
    case "splitOut":
      return { kf: reverse(splitFrames(opt === "vertIn" ? "vertOut" : opt === "horzIn" ? "horzOut" : opt === "vertOut" ? "vertIn" : "horzIn")), easing: "linear" };
    case "wipeOut":
      return { kf: wipeInset(opt || "b", false), easing: "linear" };
    case "zoomOut":
      return { kf: [{ opacity: 1, scale: "1" }, { opacity: 0, scale: "0.1" }], easing: "ease-in" };
    case "shrinkTurn":
      return { kf: [{ opacity: 1, scale: "1", rotate: "0deg" }, { opacity: 0, scale: "0", rotate: "90deg" }], easing: "ease-in" };
    default:
      return { kf: [{ opacity: 0 }, { opacity: 1 }], easing: "linear" };
  }
}

export interface StepEffect {
  anim: Animation;
  start: number; // saniye (adım içinde)
}

/** Animasyonları tıklama adımlarına böler. İlk adım tıklamasız (otomatik) olabilir. */
export function buildSteps(slide: Slide): { auto: StepEffect[]; clicks: StepEffect[][] } {
  const auto: StepEffect[] = [];
  const clicks: StepEffect[][] = [];
  let cur: StepEffect[] = auto;
  let parentStart = 0;
  let prevEnd = 0;
  for (const a of slide.animations) {
    if (a.trigger === "click") {
      cur = [];
      clicks.push(cur);
      parentStart = 0;
      prevEnd = 0;
    } else if (a.trigger === "after") {
      parentStart = prevEnd;
    }
    const start = parentStart + Math.max(0, a.delay);
    cur.push({ anim: a, start });
    prevEnd = Math.max(prevEnd, start + Math.max(0.01, a.duration));
  }
  return { auto, clicks };
}

/** Gösteri başında gizlenecek şekiller (ilk animasyonu giriş olanlar). */
export function initiallyHidden(slide: Slide): Set<string> {
  const seen = new Set<string>();
  const hidden = new Set<string>();
  for (const a of slide.animations) {
    if (seen.has(a.shapeId)) continue;
    seen.add(a.shapeId);
    if (a.cls === "entr") hidden.add(a.shapeId);
  }
  return hidden;
}

/* ------------------------------------------------------------------ */
/* Slayt geçişleri                                                      */
/* ------------------------------------------------------------------ */

export interface TransitionPlan {
  incoming?: KF;
  outgoing?: KF;
  easing?: string;
  /** Giden katman yeni katmanın üstünde mi? */
  outgoingOnTop?: boolean;
  /** Siyah zemin üzerinden. */
  thruBlack?: boolean;
}

const DIR_VEC: Record<string, [number, number]> = { u: [0, 1], d: [0, -1], l: [1, 0], r: [-1, 0] };

export function transitionPlan(t: Transition | undefined): TransitionPlan | null {
  if (!t || t.type === "none" || t.type === "cut") return null;
  const o = t.option ?? "";
  switch (t.type) {
    case "fade":
      if (o === "black") return { thruBlack: true, outgoing: [{ opacity: 1 }, { opacity: 0, offset: 0.5 }, { opacity: 0 }], incoming: [{ opacity: 0 }, { opacity: 0, offset: 0.5 }, { opacity: 1 }], easing: "linear" };
      return { incoming: [{ opacity: 0 }, { opacity: 1 }], easing: "linear" };
    case "reveal":
      return { thruBlack: true, outgoing: [{ opacity: 1 }, { opacity: 0, offset: 0.45 }, { opacity: 0 }], incoming: [{ opacity: 0 }, { opacity: 0, offset: 0.55 }, { opacity: 1 }], easing: "ease-in-out" };
    case "push": {
      const [vx, vy] = DIR_VEC[o || "u"];
      return {
        incoming: [{ translate: `${vx * 100}% ${vy * 100}%` }, { translate: "0 0" }],
        outgoing: [{ translate: "0 0" }, { translate: `${-vx * 100}% ${-vy * 100}%` }],
        easing: "cubic-bezier(.4,0,.2,1)",
      };
    }
    case "cover": {
      const [vx, vy] = DIR_VEC[o || "l"];
      return { incoming: [{ translate: `${vx * 100}% ${vy * 100}%` }, { translate: "0 0" }], easing: "cubic-bezier(.4,0,.2,1)" };
    }
    case "uncover": {
      const [vx, vy] = DIR_VEC[o || "l"];
      return { outgoing: [{ translate: "0 0" }, { translate: `${-vx * 100}% ${-vy * 100}%` }], outgoingOnTop: true, easing: "cubic-bezier(.4,0,.2,1)" };
    }
    case "wipe": {
      const side = o === "l" ? "r" : o === "r" ? "l" : o === "d" ? "t" : "b";
      return { incoming: wipeInset(side, true), easing: "linear" };
    }
    case "split":
      return { incoming: splitFrames(o || "vertOut"), easing: "ease-in-out" };
    case "randomBar":
      return { incoming: randomBarFrames(o === "horz"), easing: "linear" };
    case "shape":
      return { incoming: shapeFrames(o || "circle"), easing: "ease-in" };
    case "zoom":
      if (o === "out") return { outgoing: [{ scale: "1", opacity: 1 }, { scale: "1.8", opacity: 0 }], outgoingOnTop: true, easing: "ease-in" };
      return { incoming: [{ scale: "0.3", opacity: 0 }, { scale: "1", opacity: 1 }], easing: "ease-out" };
    case "dissolve":
      return { incoming: dissolveFrames(), easing: "linear" };
    case "checker":
      return { incoming: checkerFrames(o === "vert"), easing: "linear" };
    case "blinds":
      return { incoming: blindsFrames(o === "horz"), easing: "linear" };
    case "wheel":
      return { incoming: wheelFrames(), easing: "linear" };
    case "wedge":
      return { incoming: wedgeFrames(), easing: "linear" };
    case "newsflash":
      return { incoming: [{ scale: "0", rotate: "-720deg" }, { scale: "1", rotate: "0deg" }], easing: "ease-out" };
    default:
      return { incoming: [{ opacity: 0 }, { opacity: 1 }], easing: "linear" };
  }
}
