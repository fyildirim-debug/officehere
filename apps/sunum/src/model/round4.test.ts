/** 4. tur model birimleri: Dönüştür eşleştirmesi, hareket yolları, erişilebilirlik, sıkıştırma, eylem düğmeleri. */
import { describe, expect, it } from "vitest";
import { actionButtonGeometry, defaultActionFor } from "../render/actionButtons";
import { presetGeometry } from "../render/geometry";
import { classifyPath, motionKeyframes, motionPathOf, samplePath } from "../show/motionPaths";
import { targetSize } from "../state/compress";
import { checkAccessibility, contrastRatio } from "./accessibility";
import { makePicture, makeSp, textBody } from "./factory";
import { matchShapes } from "./morph";
import { newPresentation, newSlide } from "./presentation";
import type { Animation, SpShape } from "./types";

describe("Dönüştür: şekil eşleştirme", () => {
  it("aynı adlı, aynı metinli ve aynı geometrili şekiller eşlenir", () => {
    const a1 = makeSp("rect", 0, 0, 10, 10, { name: "Kutu" });
    const a2 = makeSp("ellipse", 0, 0, 10, 10, { name: "Daire 3" });
    const a3 = makeSp("rect", 0, 0, 10, 10, { name: "Metin 1", text: textBody("Merhaba") });
    const b1 = makeSp("rect", 50, 50, 20, 20, { name: "Kutu" });
    const b2 = makeSp("ellipse", 80, 0, 30, 30, { name: "Daire 9" });
    const b3 = makeSp("rect", 100, 100, 10, 10, { name: "Metin 7", text: textBody("Merhaba") });
    const pairs = matchShapes([a1, a2, a3], [b1, b2, b3]);
    const ids = new Map(pairs.map(([o, n]) => [n.id, o.id]));
    expect(ids.get(b1.id)).toBe(a1.id);
    expect(ids.get(b3.id)).toBe(a3.id);
    expect(ids.get(b2.id)).toBe(a2.id);
  });
  it("eşi olmayan şekil eşlenmez; '!!' önekli adlar her zaman eşlenir", () => {
    const o = makeSp("rect", 0, 0, 1, 1, { name: "!!Logo" });
    const n = makeSp("star5", 0, 0, 1, 1, { name: "!!Logo" });
    const lone = makeSp("triangle", 0, 0, 1, 1, { name: "Üçgen" });
    expect(matchShapes([o], [n, lone])).toHaveLength(1);
  });
});

describe("Hareket yolları", () => {
  it("çizgi yolu slayt kesirleriyle yazılır ve piksel noktalarına örneklenir", () => {
    const a = { effect: "pathLine", option: "right" } as Animation;
    expect(motionPathOf(a)).toBe("M 0 0 L 0.25 0 E");
    const pts = samplePath(motionPathOf(a), 1280, 720);
    expect(pts[pts.length - 1]).toEqual([320, 0]);
  });
  it("daire yolu başlangıca döner, anahtar kareler 0'dan 1'e", () => {
    const a = { effect: "pathCircle", duration: 2, delay: 0 } as Animation;
    const pts = samplePath(motionPathOf(a), 1280, 720);
    expect(Math.hypot(pts[0][0] - pts[pts.length - 1][0], pts[0][1] - pts[pts.length - 1][1])).toBeLessThan(0.01);
    const kf = motionKeyframes(a, 1280, 720);
    expect(kf[0].offset).toBe(0);
    expect(kf[kf.length - 1].offset).toBe(1);
  });
  it("içe aktarılan yollar hazır yollara eşlenir, diğerleri özel yol kalır", () => {
    expect(classifyPath("M 0 0 L 0 -0.25 E")).toEqual({ effect: "pathLine", option: "up" });
    expect(classifyPath("M 0 0 L 0.1 0.3 E")).toEqual({ effect: "pathCustom", path: "M 0 0 L 0.1 0.3 E" });
  });
});

describe("Erişilebilirlik Denetleyicisi", () => {
  const pres = newPresentation("varsayilan", 0);
  pres.slides.push(newSlide(pres, "obj"));
  it("alternatif metin, başlık ve kontrast sorunlarını bulur", () => {
    const p = structuredClone(pres);
    const s = p.slides[1];
    s.shapes.push(makePicture("data:image/png;base64,AAAA", 10, 10, 50, 50, "Resim 4"));
    const pale = makeSp("rect", 100, 300, 300, 80, { name: "Soluk Metin", text: textBody("Açık gri yazı", {}, { color: { rgb: "#EEEEEE" }, size: 12 }), fill: { type: "none" } });
    s.shapes.push(pale);
    const issues = checkAccessibility(p);
    expect(issues.some((i) => i.kind === "altText" && i.label === "Resim 4")).toBe(true);
    expect(issues.filter((i) => i.kind === "noTitle").length).toBeGreaterThanOrEqual(1);
    expect(issues.some((i) => i.kind === "contrast" && i.shapeId === pale.id)).toBe(true);
    // Hatalar uyarılardan önce sıralanır.
    expect(issues[0].level).toBe("error");
  });
  it("başlık ve alternatif metin verilince sorun kalmaz; WCAG oranı doğru", () => {
    const p = structuredClone(pres);
    p.slides.forEach((s, i) => (s.shapes = s.shapes.map((sh) => (sh.kind === "sp" && (sh.ph?.type === "title" || sh.ph?.type === "ctrTitle") ? ({ ...sh, text: textBody(`Başlık ${i}`) } as SpShape) : sh))));
    expect(checkAccessibility(p).filter((i) => i.kind === "noTitle" || i.kind === "altText")).toEqual([]);
    expect(contrastRatio("#000000", "#FFFFFF")).toBeCloseTo(21, 0);
    expect(contrastRatio("#777777", "#FFFFFF")).toBeGreaterThan(4.4);
  });
});

describe("Resimleri Sıkıştır: hedef boyut", () => {
  it("görüntülenme boyutu ve ppi'ye göre küçültür, özgünden büyütmez", () => {
    // 4000×3000 resim slaytta 480×360 px (5×3,75 inç) görünüyor: 150 ppi → 750×563.
    const t = targetSize(4000, 3000, 480, 360, undefined, 150, true);
    expect(t.w).toBe(750);
    expect(t.h).toBe(563);
    expect(targetSize(400, 300, 480, 360, undefined, 220, true).w).toBe(400);
  });
  it("kırpılan alan silinince kaynak dikdörtgen kırpılır", () => {
    const t = targetSize(1000, 1000, 96, 192, { l: 0.25, t: 0, r: 0.25, b: 0 }, 96, true);
    expect([t.sx, t.sw]).toEqual([250, 500]);
    expect(t.w).toBe(96);
  });
});

describe("Eylem düğmeleri", () => {
  it("gövde ve koyulaştırılmış simge çizilir; varsayılan eylem türe göre", () => {
    const g = actionButtonGeometry("actionButtonForwardNext", 100, 60);
    expect(g.paths).toHaveLength(2);
    expect(g.paths[1].fill).toBe("darken");
    expect(presetGeometry("actionButtonHome", 80, 80).paths.length).toBeGreaterThan(1);
    expect(defaultActionFor("actionButtonForwardNext")).toBe("#next");
    expect(defaultActionFor("actionButtonBeginning")).toBe("#first");
    expect(defaultActionFor("actionButtonBlank")).toBeUndefined();
  });
});
