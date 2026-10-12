/** Bölüm, asıl görünüm, grafik ve köprü model işlemleri. */
import { describe, expect, it } from "vitest";
import { chartFromGrid, chartToGrid, defaultChartData, niceScale, parseChartNumber } from "./chart";
import { textBody } from "./factory";
import { linkSlideIndex, mapSlideLinks, normalizeUrl } from "./links";
import { fromMasterView, insertLayout, MASTER_SLIDE_ID, syncMasterView, toggleLayoutPlaceholders, toMasterView } from "./masterView";
import { newPresentation, newSlide } from "./presentation";
import { addSectionAt, moveSection, moveSlidesKeepSections, removeSection, removeSlidesKeepSections, sectionsOf } from "./sections";
import type { Presentation, SpShape } from "./types";

function deck(n = 5): Presentation {
  const p = newPresentation("varsayilan", 0);
  for (let i = 1; i < n; i++) p.slides.push(newSlide(p, "obj"));
  return p;
}

describe("bölümler", () => {
  it("bölüm eklenince ilk slayt Varsayılan Bölüm olur", () => {
    const p = deck();
    const r = addSectionAt(p.slides, 2, "B")!;
    const secs = sectionsOf(r.slides);
    expect(secs.map((s) => [s.name, s.slideIds.length])).toEqual([
      ["Varsayılan Bölüm", 2],
      ["B", 3],
    ]);
  });

  it("bölümün ilk slaydı silinince bölüm sonraki slayttan sürer", () => {
    const p = deck();
    const slides = addSectionAt(p.slides, 2, "B")!.slides;
    const out = removeSlidesKeepSections(slides, new Set([slides[2].id]));
    expect(sectionsOf(out).map((s) => [s.name, s.slideIds.length])).toEqual([
      ["Varsayılan Bölüm", 2],
      ["B", 2],
    ]);
  });

  it("taşınan slayt bırakıldığı bölüme katılır; ilk slayt hep bölüm başlatır", () => {
    const p = deck();
    const slides = addSectionAt(p.slides, 2, "B")!.slides;
    const moved = moveSlidesKeepSections(slides, [slides[4].id], 0);
    expect(moved[0].id).toBe(slides[4].id);
    expect(moved[0].section?.name).toBe("Varsayılan Bölüm");
    expect(sectionsOf(moved).map((s) => s.slideIds.length)).toEqual([3, 2]);
  });

  it("bölüm taşıma ve kaldırma", () => {
    const p = deck();
    const slides = addSectionAt(p.slides, 3, "B")!.slides;
    const b = sectionsOf(slides)[1];
    const up = moveSection(slides, b.id, -1);
    expect(up[0].section?.name).toBe("B");
    expect(sectionsOf(up).map((s) => s.slideIds.length)).toEqual([2, 3]);
    const removed = removeSection(slides, b.id);
    expect(sectionsOf(removed)).toHaveLength(1);
    expect(removed).toHaveLength(5);
  });
});

describe("asıl slayt görünümü", () => {
  it("düzen yer tutucusu taşınınca konumu düzenden gelen slayt yer tutucuları da taşınır", () => {
    const p = deck(2);
    const slide = p.slides[1];
    const layoutId = slide.layoutId;
    let mv = toMasterView(p);
    expect(mv.slides[0].id).toBe(MASTER_SLIDE_ID);
    const li = mv.slides.findIndex((s) => s.id === layoutId);
    mv = {
      ...mv,
      slides: mv.slides.map((s, i) => (i === li ? { ...s, shapes: s.shapes.map((sh) => (sh.kind === "sp" && sh.ph?.type === "title" ? { ...sh, y: sh.y + 100 } : sh)) } : s)),
    };
    const out = fromMasterView(p, mv);
    const oldTitle = slide.shapes.find((s) => s.kind === "sp" && s.ph?.type === "title")!;
    const newTitle = out.slides[1].shapes.find((s) => s.kind === "sp" && s.ph?.type === "title")!;
    expect(newTitle.y).toBeCloseTo(oldTitle.y + 100, 3);
    expect(out.masterEdited).toBe(true);
    // Örnek metin düzene katlanmaz.
    const lt = out.layouts.find((l) => l.id === layoutId)!.shapes.find((s) => s.kind === "sp" && s.ph?.type === "title") as SpShape;
    expect(lt.text?.paragraphs.every((pp) => pp.runs.length === 0)).toBe(true);
  });

  it("asıl başlık biçimi düzey stiline katlanır ve slaytlara kalıtılır", () => {
    const p = deck(2);
    let mv = toMasterView(p);
    mv = {
      ...mv,
      slides: mv.slides.map((s) =>
        s.id === MASTER_SLIDE_ID
          ? { ...s, shapes: s.shapes.map((sh) => (sh.kind === "sp" && sh.ph?.type === "title" ? { ...sh, text: { ...sh.text!, paragraphs: sh.text!.paragraphs.map((pp) => ({ ...pp, runs: pp.runs.map((r) => ({ ...r, props: { bold: true, color: { rgb: "#FF0000" } } })) })) } } : sh)) }
          : s,
      ),
    };
    mv = syncMasterView(mv);
    // Asıl görünüm açıkken de düzenler biçimi miras alsın diye örnek metin biçimi düzey stiline katlanır.
    const live = mv.master.shapes.find((s) => s.kind === "sp" && s.ph?.type === "title") as SpShape;
    expect(live.text?.levels?.[0]?.run?.bold).toBe(true);
    expect(syncMasterView(mv)).toBe(mv);
    const out = fromMasterView(p, mv);
    const mt = out.master.shapes.find((s) => s.kind === "sp" && s.ph?.type === "title") as SpShape;
    expect(mt.text?.levels?.[0]?.run?.bold).toBe(true);
    expect(mt.text?.levels?.[0]?.run?.color).toEqual({ rgb: "#FF0000" });
  });

  it("yeni düzen eklenir, başlık/alt bilgi yer tutucuları açılıp kapatılır", () => {
    const p = deck(1);
    const mv0 = toMasterView(p);
    const { pres: mv, index } = insertLayout(mv0, 1, "Yeni");
    const id = mv.slides[index].id;
    expect(mv.slides[index].shapes.some((s) => s.kind === "sp" && s.ph?.type === "title")).toBe(true);
    const noTitle = toggleLayoutPlaceholders(mv, id, "title", false);
    expect(noTitle.slides[index].shapes.some((s) => s.kind === "sp" && s.ph?.type === "title")).toBe(false);
    const out = fromMasterView(p, noTitle);
    expect(out.layouts).toHaveLength(p.layouts.length + 1);
    expect(out.layouts.find((l) => l.id === id)?.name).toBe("Yeni");
  });
});

describe("grafik verisi", () => {
  it("ızgara ↔ veri dönüşümü (Türkçe ondalık)", () => {
    const c = defaultChartData("col");
    const g = chartToGrid(c);
    expect(g[1][1]).toBe("4,3");
    g[1][1] = "1.234,5";
    const c2 = chartFromGrid(c, g);
    expect(c2.series[0].values[0]).toBe(1234.5);
    expect(parseChartNumber("")).toBeNull();
    expect(parseChartNumber("2.5")).toBe(2.5);
  });

  it("eksen ölçeği güzel adımlar üretir", () => {
    expect(niceScale(0, 4.5)).toEqual({ min: 0, max: 5, step: 1 });
    expect(niceScale(0, 5).max).toBe(6);
    const s = niceScale(-3, 97, 6);
    expect(s.min).toBeLessThanOrEqual(-3);
    expect(s.max).toBeGreaterThanOrEqual(97);
  });
});

describe("köprü yardımcıları", () => {
  it("adres tamamlama ve slayt hedefi", () => {
    expect(normalizeUrl("www.ornek.com")).toBe("https://www.ornek.com");
    expect(normalizeUrl("ad@ornek.com")).toBe("mailto:ad@ornek.com");
    const p = deck(3);
    expect(linkSlideIndex("#next", p, 0)).toBe(1);
    expect(linkSlideIndex("#prev", p, 0)).toBe(-1);
    expect(linkSlideIndex(`#slide:${p.slides[2].id}`, p, 0)).toBe(2);
    expect(linkSlideIndex("https://x", p, 0)).toBeNull();
    const s = { ...p.slides[1], shapes: [...p.slides[1].shapes] };
    s.shapes = s.shapes.map((x) => (x.kind === "sp" && x.ph ? { ...x, text: { ...textBody("a"), paragraphs: [{ runs: [{ text: "a", props: { link: "#next" } }], level: 0, props: {} }] } } : x));
    const m = mapSlideLinks(s, () => undefined);
    expect(JSON.stringify(m)).not.toContain("#next");
  });
});
