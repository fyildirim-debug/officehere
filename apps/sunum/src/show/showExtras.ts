/**
 * Gösteride özel geçişler: Dönüştür (şekiller konum/boyut/renk değiştirerek
 * geçer) ve Yakınlaştırma (slayt yakınlaştırmasına tıklanınca hedef slayda
 * yakınlaşma; dönüşte uzaklaşma). Web Animations ile katmanlar üzerinde.
 */
import { matchShapes } from "../model/morph";
import type { Presentation, Shape, Slide } from "../model/types";

const EASE = "cubic-bezier(.45,0,.25,1)";

function elOf(layer: HTMLElement, id: string): HTMLElement | null {
  return layer.querySelector<HTMLElement>(`.sv > [data-shape-id="${CSS.escape(id)}"]`);
}

const tf = (dx: number, dy: number, rot: number, sx: number, sy: number) => `translate(${dx.toFixed(2)}px, ${dy.toFixed(2)}px) rotate(${rot}deg) scale(${sx.toFixed(4)}, ${sy.toFixed(4)})`;

/** Dönüştür geçişinin animasyonları (gelen katman üstte). */
export function morphAnimations(inEl: HTMLElement, outEl: HTMLElement, from: Slide, to: Slide, dur: number): globalThis.Animation[] {
  const anims: globalThis.Animation[] = [];
  const opts: KeyframeAnimationOptions = { duration: dur, easing: EASE, fill: "both" };
  // Arka plan ve süsler yumuşakça geçer; eski katman altta kalır.
  inEl.querySelectorAll<HTMLElement>(".sv > .sv-bg, .sv > .sv-deco").forEach((e) => anims.push(e.animate([{ opacity: 0 }, { opacity: 1 }], opts)));
  const pairs = matchShapes(from.shapes, to.shapes);
  const matchedNew = new Set(pairs.map((p) => p[1].id));
  const matchedOld = new Set(pairs.map((p) => p[0].id));
  for (const [o, n] of pairs) {
    const ne = elOf(inEl, n.id);
    const oe = elOf(outEl, o.id);
    if (!ne) continue;
    const dx = o.x + o.w / 2 - (n.x + n.w / 2);
    const dy = o.y + o.h / 2 - (n.y + n.h / 2);
    const sx = n.w ? o.w / n.w : 1;
    const sy = n.h ? o.h / n.h : 1;
    const same = JSON.stringify({ ...o, id: "", x: 0, y: 0, w: 0, h: 0, rot: 0 }) === JSON.stringify({ ...n, id: "", x: 0, y: 0, w: 0, h: 0, rot: 0 });
    anims.push(ne.animate([{ transform: tf(dx, dy, o.rot, sx, sy), opacity: same ? 1 : 0 }, { transform: tf(0, 0, n.rot, 1, 1), opacity: 1 }], opts));
    if (oe) anims.push(oe.animate([{ transform: tf(0, 0, o.rot, 1, 1), opacity: 1 }, { transform: tf(-dx, -dy, n.rot, 1 / sx, 1 / sy), opacity: same ? 1 : 0 }], opts));
  }
  // Eşleşmeyen şekiller: yeniler belirir, eskiler kaybolur.
  for (const s of to.shapes) {
    const e = matchedNew.has(s.id) ? null : elOf(inEl, s.id);
    if (e) anims.push(e.animate([{ opacity: 0 }, { opacity: 1 }], opts));
  }
  for (const s of from.shapes) {
    const e = matchedOld.has(s.id) ? null : elOf(outEl, s.id);
    if (e) anims.push(e.animate([{ opacity: 1 }, { opacity: 0 }], { ...opts, duration: dur * 0.6 }));
  }
  return anims;
}

export interface ZoomRect {
  x: number;
  y: number;
  w: number;
  h: number;
  /** true: hedeften yakınlaştırma slaydına geri uzaklaşma. */
  out?: boolean;
}

/** Yakınlaştırma: eski katman dikdörtgene doğru büyür, yeni katman dikdörtgenden tam ekrana açılır. */
export function zoomAnimations(inEl: HTMLElement, outEl: HTMLElement, r: ZoomRect, W: number, H: number, dur: number): globalThis.Animation[] {
  const k = Math.max(0.01, r.w / W);
  const big = `translate(${(-r.x / k).toFixed(2)}px, ${(-r.y / k).toFixed(2)}px) scale(${(1 / k).toFixed(4)})`;
  const small = `translate(${r.x.toFixed(2)}px, ${r.y.toFixed(2)}px) scale(${k.toFixed(4)}, ${(r.h / H).toFixed(4)})`;
  const opts: KeyframeAnimationOptions = { duration: dur, easing: EASE, fill: "both" };
  inEl.style.transformOrigin = outEl.style.transformOrigin = "0 0";
  if (!r.out) {
    return [outEl.animate([{ transform: "none", opacity: 1 }, { transform: big, opacity: 0 }], opts), inEl.animate([{ transform: small, opacity: 0.6 }, { transform: "none", opacity: 1 }], opts)];
  }
  // Geri dönüş: hedef küçülerek yakınlaştırma çerçevesine iner.
  outEl.style.zIndex = "2";
  return [outEl.animate([{ transform: "none", opacity: 1 }, { transform: small, opacity: 0.2 }], opts), inEl.animate([{ transform: big, opacity: 0.4 }, { transform: "none", opacity: 1 }], opts)];
}

/** Gösteride yakınlaştırma şeklinin bilgisi: hedef dizin, dönüş noktası ve kapsamın son slaydı. */
export function zoomJump(pres: Presentation, slideIdx: number, shapeId: string): { target: number; rect: ZoomRect; ret: boolean; last: number } | null {
  const sl = pres.slides[slideIdx];
  const sh = sl?.shapes.find((s): s is Shape & { kind: "pic" } => s.id === shapeId && s.kind === "pic");
  if (!sh?.zoom) return null;
  const target = pres.slides.findIndex((s) => s.id === sh.zoom!.slideId);
  if (target < 0) return null;
  let last = target;
  if (sh.zoom.summary) {
    // Bölüm: sonraki bölüm başlangıcına (ya da sona) kadar.
    while (last + 1 < pres.slides.length && !pres.slides[last + 1].section && last + 1 !== slideIdx) last++;
  }
  return { target, rect: { x: sh.x, y: sh.y, w: sh.w, h: sh.h }, ret: !!sh.zoom.ret, last };
}
