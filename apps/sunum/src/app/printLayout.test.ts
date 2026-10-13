/** Yazdırma düzeni: slayt aralığı, gizli slaytlar, dinleyici notu yerleşimi ve ana hat sayfaları. */
import { describe, expect, it } from "vitest";
import { DEFAULT_PRINT_SETTINGS } from "../io/presProps";
import { textBody } from "../model/factory";
import { newPresentation, newSlide } from "../model/presentation";
import type { Presentation, PrintSettings, SpShape } from "../model/types";
import { A4, buildPrintPages, handoutSlots, paperSize, parseSlideRange, printSlideIndices } from "./printLayout";

function deck(n: number): Presentation {
  const p = newPresentation("varsayilan", 0);
  while (p.slides.length < n) p.slides.push(newSlide(p, "obj"));
  return p;
}
const S = (patch: Partial<PrintSettings> = {}): PrintSettings => ({ ...DEFAULT_PRINT_SETTINGS, ...patch });

describe("slayt aralığı", () => {
  it("virgül/noktalı virgül, ters aralık ve sınır dışı numaralar", () => {
    expect(parseSlideRange("1-3, 5", 10)).toEqual([0, 1, 2, 4]);
    expect(parseSlideRange("5-3; 3", 10)).toEqual([2, 3, 4]);
    expect(parseSlideRange("8-12", 9)).toEqual([7, 8]);
    expect(parseSlideRange("1–2", 9)).toEqual([0, 1]);
    expect(parseSlideRange("a-3", 9)).toBeNull();
    expect(parseSlideRange("", 9)).toBeNull();
  });
});

describe("yazdırılacak slaytlar", () => {
  it("gizli slaytlar ayar açık değilse atlanır; geçerli slayt her zaman basılır", () => {
    const p = deck(5);
    p.slides[1].hidden = true;
    expect(printSlideIndices(p, S(), { kind: "all" }, 0, [])).toEqual([0, 2, 3, 4]);
    expect(printSlideIndices(p, S({ hiddenSlides: true }), { kind: "all" }, 0, [])).toEqual([0, 1, 2, 3, 4]);
    expect(printSlideIndices(p, S(), { kind: "current" }, 1, [])).toEqual([1]);
    expect(printSlideIndices(p, S(), { kind: "selection" }, 0, [4, 2, 2])).toEqual([2, 4]);
    expect(printSlideIndices(p, S(), { kind: "custom", text: "1-3" }, 0, [])).toEqual([0, 2]);
  });
});

describe("sayfalar", () => {
  it("tam sayfa slaytlarda kâğıt slayt boyutundadır; dinleyici notları A4", () => {
    const p = deck(2);
    expect(paperSize(p, "slides", "portrait")).toEqual({ w: 1280, h: 720 });
    expect(paperSize(p, "handouts6", "portrait").w).toBeCloseTo(A4.w);
    expect(paperSize(p, "notes", "landscape").w).toBeCloseTo(A4.h);
  });

  it("13 slayt 6'lı dinleyici notunda 3 sayfa; yatay sıra satır satır, dikey sıra sütun sütun", () => {
    const p = deck(13);
    const pages = buildPrintPages(p, S({ what: "handouts6" }), p.slides.map((_, i) => i));
    expect(pages).toHaveLength(3);
    expect(pages[2].kind === "handout" && pages[2].cells.map((c) => c.index)).toEqual([12]);
    const horz = handoutSlots(6, A4, 16 / 9, "horz");
    const vert = handoutSlots(6, A4, 16 / 9, "vert");
    // Yatay: 2. hücre birinciyle aynı satırda (sağında); dikey: altında.
    expect(horz[1].slide.y).toBeCloseTo(horz[0].slide.y);
    expect(horz[1].slide.x).toBeGreaterThan(horz[0].slide.x);
    expect(vert[1].slide.x).toBeCloseTo(vert[0].slide.x);
    expect(vert[1].slide.y).toBeGreaterThan(vert[0].slide.y);
    // Hücreler sayfa içinde ve slayt oranındadır.
    for (const c of horz) {
      expect(c.slide.x + c.slide.w).toBeLessThanOrEqual(A4.w);
      expect(c.slide.y + c.slide.h).toBeLessThanOrEqual(A4.h);
      expect(c.slide.w / c.slide.h).toBeCloseTo(16 / 9);
    }
  });

  it("3'lü düzende her slaydın yanında not satırları vardır", () => {
    const slots = handoutSlots(3, A4, 16 / 9);
    expect(slots).toHaveLength(3);
    for (const s of slots) {
      expect(s.lines).toBeDefined();
      expect(s.lines!.x).toBeGreaterThan(s.slide.x + s.slide.w);
    }
  });

  it("ana hat başlık ve gövde metniyle sayfalanır; not sayfası slayt başına bir sayfa", () => {
    const p = deck(40);
    p.slides.forEach((s, i) => {
      const t = s.shapes.find((x): x is SpShape => x.kind === "sp" && x.ph?.type === "title");
      if (t) t.text = textBody(`Başlık ${i + 1}`);
      const b = s.shapes.find((x): x is SpShape => x.kind === "sp" && x.ph?.type === "obj");
      if (b) b.text = textBody(Array.from({ length: 6 }, (_, k) => `Madde ${k + 1}`).join("\n"));
    });
    const idx = p.slides.map((_, i) => i);
    const outline = buildPrintPages(p, S({ what: "outline" }), idx);
    expect(outline.length).toBeGreaterThan(1);
    const all = outline.flatMap((pg) => (pg.kind === "outline" ? pg.blocks : []));
    expect(all.map((b) => b.index)).toEqual(idx);
    expect(all[1].title).toBe("Başlık 2");
    expect(all[1].lines).toHaveLength(6);
    expect(buildPrintPages(p, S({ what: "notes" }), [3, 7])).toEqual([
      { kind: "notes", index: 3 },
      { kind: "notes", index: 7 },
    ]);
  });
});
