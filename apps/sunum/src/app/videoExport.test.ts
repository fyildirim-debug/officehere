/** Videoya aktarma zaman çizelgesi: gizli slaytlar, zamanlamalar, geçişler ve animasyon adımları. */
import { describe, expect, it } from "vitest";
import { makeSp } from "../model/factory";
import { newPresentation, newSlide } from "../model/presentation";
import type { Animation, Presentation } from "../model/types";
import { slideStates, videoSize, videoTimeline, type VideoOptions } from "./videoExport";

const OPTS: VideoOptions = { height: 720, secondsPerSlide: 4, useTimings: true };

function deck(n: number): Presentation {
  const p = newPresentation("varsayilan", 0);
  while (p.slides.length < n) p.slides.push(newSlide(p, "obj"));
  return p;
}
const anim = (shapeId: string, cls: Animation["cls"], trigger: Animation["trigger"]): Animation => ({ id: `${shapeId}-${cls}`, shapeId, cls, effect: "fade", trigger, duration: 0.5, delay: 0 });

describe("animasyon durumları", () => {
  it("giriş efektli şekil önce gizli, tıklamadan sonra görünür; çıkış efektiyle gizlenir", () => {
    const p = deck(1);
    const sl = p.slides[0];
    const a = makeSp("rect", 0, 0, 10, 10);
    const b = makeSp("rect", 20, 0, 10, 10);
    sl.shapes.push(a, b);
    sl.animations = [anim(a.id, "entr", "click"), anim(b.id, "entr", "with"), anim(a.id, "exit", "click")];
    const st = slideStates(sl);
    expect(st).toHaveLength(3);
    expect([...st[0]].sort()).toEqual([a.id, b.id].sort());
    expect([...st[1]]).toEqual([]);
    expect([...st[2]]).toEqual([a.id]);
  });
});

describe("zaman çizelgesi", () => {
  it("gizli slayt atlanır; zamanlaması olan slayt kendi süresini, diğerleri varsayılanı kullanır", () => {
    const p = deck(3);
    p.slides[1].hidden = true;
    p.slides[2].transition = { type: "fade", duration: 0.8, advanceOnClick: true, advanceAfter: 2 };
    const { segs, total } = videoTimeline(p, OPTS);
    expect(segs.map((s) => s.kind)).toEqual(["hold", "trans", "hold"]);
    expect(segs.every((s) => s.slide !== 1)).toBe(true);
    expect(total).toBe(4000 + 800 + 2000);
    const tr = segs[1];
    expect(tr.kind === "trans" && tr.prev).toEqual({ slide: 0, state: 0 });
    // Zamanlamalar kullanılmazsa her slayt varsayılan süredir.
    expect(videoTimeline(p, { ...OPTS, useTimings: false }).total).toBe(4000 + 800 + 4000);
  });

  it("animasyon adımları slayt süresini paylaşır; aralar arka arkaya, boşluksuz", () => {
    const p = deck(1);
    const sl = p.slides[0];
    const a = makeSp("rect", 0, 0, 10, 10);
    sl.shapes.push(a);
    sl.animations = [anim(a.id, "entr", "click")];
    const { segs, total } = videoTimeline(p, OPTS);
    expect(segs.map((s) => s.kind)).toEqual(["hold", "step", "hold"]);
    expect(total).toBe(4000);
    for (let i = 1; i < segs.length; i++) expect(segs[i].start).toBe(segs[i - 1].start + segs[i - 1].dur);
  });

  it("video genişliği slayt oranından ve çift sayıdır", () => {
    const p = deck(1);
    expect(videoSize(p, 1080)).toEqual({ w: 1920, h: 1080 });
    expect(videoSize(p, 480).w % 2).toBe(0);
    p.width = 960;
    expect(videoSize(p, 720)).toEqual({ w: 960, h: 720 });
  });
});
