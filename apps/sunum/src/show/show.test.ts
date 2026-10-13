/** Gösteri motoru birimleri: Belir/Kaybol kareleri, gösteri sırası ve özel gösteri yedeği. */
import { describe, expect, it } from "vitest";
import { newPresentation, newSlide } from "../model/presentation";
import type { Animation } from "../model/types";
import { animationKeyframes } from "./effects";
import { DEFAULT_SHOW_SETTINGS, showOrder } from "./settings";

const anim = (effect: string, cls: Animation["cls"]): Animation => ({ id: "a", shapeId: "s", cls, effect, trigger: "after", duration: 0, delay: 0 });
const g = { x: 0, y: 0, w: 100, h: 100 };

/** Kare listesinin p ilerlemedeki opaklığı (doğrusal ara değer). */
function opacityAt(kf: Keyframe[], p: number): number {
  const pts = kf.map((k, i) => ({ o: (k.offset as number | undefined) ?? i / (kf.length - 1), v: Number(k.opacity) }));
  let v = pts[0].v;
  for (let i = 0; i < pts.length - 1; i++) {
    const a = pts[i];
    const b = pts[i + 1];
    if (p >= a.o && p <= b.o) v = b.o === a.o ? b.v : a.v + ((b.v - a.v) * (p - a.o)) / (b.o - a.o);
  }
  return v;
}

describe("Belir / Kaybol anlık efektleri", () => {
  it("adım yumuşatması kullanılmaz; bitişe çok yakın ilerlemede (kayan nokta) sonuç değeri kesindir", () => {
    const appear = animationKeyframes(anim("appear", "entr"), g, 1280, 720);
    expect(appear.easing).not.toMatch(/step/);
    expect(opacityAt(appear.kf, 0)).toBe(0);
    // Krom'da gecikmeli 1 ms'lik efektin bitişi 0,9999999… ilerleme olarak hesaplanabilir.
    expect(opacityAt(appear.kf, 0.9999999)).toBeGreaterThan(0.99);
    const gone = animationKeyframes(anim("disappear", "exit"), g, 1280, 720);
    expect(gone.easing).not.toMatch(/step/);
    expect(opacityAt(gone.kf, 0)).toBe(1);
    expect(opacityAt(gone.kf, 0.9999999)).toBeLessThan(0.01);
  });
});

describe("gösteri sırası", () => {
  it("silinmiş özel gösteriye bağlı ayar tüm slaytlara döner (boş gösteri açılmaz)", () => {
    const p = newPresentation("varsayilan", 0);
    p.slides.push(newSlide(p, "obj"), newSlide(p, "obj"));
    expect(showOrder(p, { kind: "custom", id: "yok" })).toEqual([0, 1, 2]);
    p.slides[1].hidden = true;
    expect(showOrder(p, DEFAULT_SHOW_SETTINGS.range)).toEqual([0, 2]);
  });
});
