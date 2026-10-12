/** 3. tur: kararlı küçültme, taslak seçimi, gösteri sırası, komut araması, noktalar, ana hat, yeniden kullanma, grafik. */
import { describe, expect, it } from "vitest";
import { CommitGuard, currentStep, pickShrinkStep, SHRINK_STEPS } from "../editor/autofitFit";
import { draftsToPrune, pickDraft } from "../app/draftStore";
import { foldText as fold, rankCommands, type CommandEntry } from "@officehere/ui";
import { showOrder, stepInOrder, uniqueShowName } from "../show/settings";
import { SLIDE_FONT_SIZES, stepFontSize } from "../state/textActions";
import { makeChart, seriesColor, stackedValues, switchRowColumn, valueRange } from "./chart";
import { moveVertex, normalizeBounds, toCustomShape, vertices } from "./editPoints";
import { makeSp, textBody } from "./factory";
import { parseTextOutline, slidesFromOutline } from "./outlineImport";
import { newPresentation, newSlide } from "./presentation";
import { freezeColors, matchLayout, reuseSlide } from "./reuse";
import type { SpShape } from "./types";

describe("otomatik sığdırma (kararlı küçültme)", () => {
  it("sığan en büyük adımı bulur ve sonuç o anki ölçekten bağımsızdır", () => {
    // Yükseklik adımla azalır; kutu 70 birim.
    const heights = SHRINK_STEPS.map((_, i) => 100 - i * 8);
    const h = (i: number) => heights[i];
    const step = pickShrinkStep(h, 70);
    expect(heights[step]).toBeLessThanOrEqual(71);
    expect(heights[step - 1]).toBeGreaterThan(71);
    // Aynı girdiyle tekrar: aynı sonuç (salınım yok).
    expect(pickShrinkStep(h, 70)).toBe(step);
  });

  it("sığıyorsa küçültmez, hiçbiri sığmıyorsa en küçük adımı verir", () => {
    expect(pickShrinkStep(() => 10, 50)).toBe(0);
    expect(pickShrinkStep(() => 999, 50)).toBe(SHRINK_STEPS.length - 1);
  });

  it("satır kırılmasıyla yükseklik sıçrasa da (eski salınım nedeni) ölçüm sayısı sınırlıdır", () => {
    let calls = 0;
    // 1. adımda iki satır (taşar), sonrakilerde tek satır: eski doğrusal tahmin burada salınıyordu.
    const h = (i: number) => {
      calls++;
      return i === 0 ? 120 : 60;
    };
    expect(pickShrinkStep(h, 80)).toBe(1);
    expect(calls).toBeLessThanOrEqual(Math.ceil(Math.log2(SHRINK_STEPS.length)) + 2);
  });

  it("gövdedeki ölçek/aralık değeri adıma eşlenir", () => {
    expect(currentStep(undefined, undefined)).toBe(0);
    expect(currentStep(0.925, 0.1)).toBe(2);
    expect(currentStep(0.33, 0.2)).toBe(-1);
  });

  it("koruma, kısa sürede çok sayıda güncellemeyi keser", () => {
    const g = new CommitGuard(3, 1000);
    expect([g.allow("a", 0), g.allow("a", 10), g.allow("a", 20), g.allow("a", 30)]).toEqual([true, true, true, false]);
    expect(g.allow("b", 30)).toBe(true);
    expect(g.allow("a", 2000)).toBe(true);
  });
});

describe("Otomatik Kurtarma taslak seçimi", () => {
  const d = (session: string, savedAt: number) => ({ session, savedAt });
  it("kendi ve canlı pencerelerin taslaklarını atlayıp en yeniyi seçer", () => {
    const list = [d("eski", 1), d("canli", 9), d("ben", 10), d("yeni", 5)];
    expect(pickDraft(list, new Set(["canli"]), "ben")?.session).toBe("yeni");
    expect(pickDraft([d("ben", 1)], new Set(), "ben")).toBeNull();
  });
  it("en eski taslaklar budanır", () => {
    const list = [1, 2, 3, 4, 5, 6, 7].map((t) => d(`s${t}`, t));
    expect(draftsToPrune(list, 5).sort()).toEqual(["s1", "s2"]);
  });
});

describe("gösteri sırası", () => {
  const p = newPresentation("varsayilan", 0);
  for (let i = 0; i < 4; i++) p.slides.push(newSlide(p, "obj"));
  p.slides[2].hidden = true;
  p.customShows = [{ id: "c1", name: "Kısa", slideIds: [p.slides[4].id, p.slides[0].id, "silinmis"] }];

  it("tümü / aralık / özel gösteri; gizli slaytlar atlanır", () => {
    expect(showOrder(p, { kind: "all" })).toEqual([0, 1, 3, 4]);
    expect(showOrder(p, { kind: "range", from: 2, to: 4 })).toEqual([1, 3]);
    expect(showOrder(p, { kind: "range", from: 9, to: 1 })).toEqual([4]);
    expect(showOrder(p, { kind: "custom", id: "c1" })).toEqual([4, 0]);
  });

  it("sonraki/önceki; sıra dışı slayttan en yakına", () => {
    const o = [0, 1, 3, 4];
    expect(stepInOrder(o, 1, 1)).toBe(3);
    expect(stepInOrder(o, 4, 1)).toBeNull();
    expect(stepInOrder(o, 0, -1)).toBeNull();
    expect(stepInOrder(o, 2, 1)).toBe(3);
    expect(stepInOrder(o, 2, -1)).toBe(1);
  });

  it("özel gösteri adı benzersizleştirilir", () => {
    expect(uniqueShowName([{ id: "a", name: "Özel Gösteri", slideIds: [] }], "Özel Gösteri")).toBe("Özel Gösteri (2)");
  });
});

describe("yazı tipi boyutu adımları", () => {
  it("sunu listesini izler", () => {
    expect(SLIDE_FONT_SIZES).toContain(32);
    const up: number[] = [];
    let s = 28;
    for (let i = 0; i < 8; i++) up.push((s = stepFontSize(s, 1)));
    expect(up).toEqual([32, 36, 40, 44, 48, 54, 60, 66]);
    expect(stepFontSize(96, 1)).toBe(100);
    expect(stepFontSize(100, -1)).toBe(96);
    expect(stepFontSize(54, -1)).toBe(48);
  });
});

describe("komut araması", () => {
  const e = (title: string, tabLabel = "Giriş", desc = ""): CommandEntry => ({ id: title, title, desc, tabId: "t", tabLabel, group: "", nth: 0 });
  const list = [e("Slaytı Gizle", "Slayt Gösterisi"), e("Yeni Slayt"), e("Gri Tonlama", "Görünüm"), e("Bul", "Giriş", "Sunudaki metni bulun.")];
  it("Türkçe büyük/küçük harf ve şapka farkını yok sayar", () => {
    expect(fold("SLAYTI GİZLE")).toBe(fold("slaytı gizle"));
    expect(rankCommands(list, "slayti gizle")[0].title).toBe("Slaytı Gizle");
    expect(rankCommands(list, "gri ton").map((c) => c.title)).toEqual(["Gri Tonlama"]);
    expect(rankCommands(list, "metni")[0].title).toBe("Bul");
    expect(rankCommands(list, "yok böyle")).toEqual([]);
  });
});

describe("Noktaları Düzenle", () => {
  it("dikdörtgen özel yola çevrilir, nokta taşınır ve kutu yeniden hesaplanır", () => {
    const r = makeSp("rect", 100, 100, 200, 100);
    const c = toCustomShape(r);
    expect(c.geom).toBe("custom");
    const vs = vertices(c);
    expect(vs.length).toBeGreaterThanOrEqual(4);
    const far = vs.reduce((m, v) => (v.x + v.y > m.x + m.y ? v : m), vs[0]);
    const moved = normalizeBounds(moveVertex(c, far, 50, 30));
    expect(moved.w).toBeCloseTo(250, 0);
    expect(moved.h).toBeCloseTo(130, 0);
    expect(moved.x).toBeCloseTo(100, 0);
  });
});

describe("Anahattan Slaytlar", () => {
  it("girintiye göre başlık ve madde düzeyleri", () => {
    const o = parseTextOutline("Giriş\n\tBirinci\n\t\tAlt\n# İkinci Bölüm\n- Madde\n");
    expect(o).toEqual([
      { title: "Giriş", lines: ["Birinci", "\tAlt"] },
      { title: "İkinci Bölüm", lines: ["Madde"] },
    ]);
    const p = newPresentation("varsayilan", 0);
    const slides = slidesFromOutline(p, o);
    expect(slides).toHaveLength(2);
    const texts = slides[0].shapes.filter((s): s is SpShape => s.kind === "sp" && !!s.text).map((s) => s.text!.paragraphs.map((x) => `${x.level}:${x.runs.map((r) => r.text).join("")}`).join("|"));
    expect(texts).toContain("0:Giriş");
    expect(texts).toContain("0:Birinci|1:Alt");
  });
});

describe("Slaytları Yeniden Kullan", () => {
  it("tema renkleri sabitlenir, düzen türü eşleşir, boyut ölçeklenir", () => {
    const src = newPresentation("gunbatimi", 0);
    const dst = newPresentation("varsayilan", 0, { width: 960, height: 720 });
    const sl = newSlide(src, "obj");
    sl.shapes.push(makeSp("rect", 0, 0, 1280, 720, { fill: { type: "solid", color: { scheme: "accent1" } } }));
    const keep = reuseSlide(dst, src, sl, true);
    expect(dst.layouts.find((l) => l.id === keep.layoutId)?.type).toBe("obj");
    const rect = keep.shapes[keep.shapes.length - 1];
    expect(rect.w).toBeCloseTo(960, 0);
    expect(rect.kind === "sp" && rect.fill.type === "solid" && rect.fill.color.rgb?.toUpperCase()).toBe(src.theme.colors.accent1.toUpperCase());
    expect(JSON.stringify(keep)).not.toContain('"scheme"');
    const plain = reuseSlide(dst, src, sl, false);
    expect(JSON.stringify(plain.shapes[plain.shapes.length - 1])).toContain('"scheme":"accent1"');
    expect(matchLayout(dst, undefined).type).toBe("obj");
    expect(freezeColors({ c: { scheme: "dk1" }, font: "+mj" }, src.theme)).toEqual({ c: { rgb: expect.stringMatching(/^#/) }, font: src.theme.fonts.major });
  });
});

describe("grafik yardımcıları", () => {
  it("Satır/Sütun Değiştir, yığılmış değerler ve paletler", () => {
    const c = makeChart("col", 0, 0, 100, 100).chart;
    const t = switchRowColumn(c);
    expect(t.categories).toEqual(c.series.map((s) => s.name));
    expect(t.series.map((s) => s.name)).toEqual(c.categories);
    expect(switchRowColumn(t)).toEqual({ ...c, series: c.series.map((s) => ({ name: s.name, values: s.values })) });
    const st = stackedValues({ ...c, grouping: "stacked" });
    expect(st[1][0].lo).toBeCloseTo(c.series[0].values[0]!, 5);
    expect(valueRange({ ...c, grouping: "percentStacked" })).toEqual({ min: 0, max: 100 });
    const pct = stackedValues({ ...c, grouping: "percentStacked" });
    expect(pct[pct.length - 1][0].hi).toBeCloseTo(100, 5);
    expect(seriesColor(0, undefined, 1)).toEqual({ scheme: "accent2" });
    expect(seriesColor(1, undefined, 4).scheme).toBe("accent1");
  });
  it("metin gövdesi yardımcısı (bağımlılık sınaması)", () => {
    expect(textBody("x").paragraphs).toHaveLength(1);
  });
});
