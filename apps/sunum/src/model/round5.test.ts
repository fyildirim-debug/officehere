/** 5. tur model birimleri: paragraf taşıma (Alt+Shift+Yukarı/Aşağı), ana hat indirme/yükseltme. */
import { describe, expect, it } from "vitest";
import { textBody } from "./factory";
import { bodyShape, demoteSlide, promoteLine, titleShape } from "./outline";
import { newPresentation, newSlide } from "./presentation";
import { moveParagraphs, paraText } from "./text";
import type { TextBody } from "./types";

const texts = (b: { paragraphs: Parameters<typeof paraText>[0][] }) => b.paragraphs.map(paraText);

describe("paragrafları taşı", () => {
  it("tek paragraf yukarı/aşağı taşınır; düzeyi korunur", () => {
    const b = textBody("Bir\nİki\nÜç");
    b.paragraphs[2].level = 1;
    const up = moveParagraphs(b, 2, 2, -1)!;
    expect(texts(up)).toEqual(["Bir", "Üç", "İki"]);
    expect(up.paragraphs[1].level).toBe(1);
    expect(texts(moveParagraphs(b, 0, 0, 1)!)).toEqual(["İki", "Bir", "Üç"]);
  });
  it("seçili blok birlikte taşınır; sınırda değişiklik yoktur", () => {
    const b = textBody("A\nB\nC\nD");
    expect(texts(moveParagraphs(b, 1, 2, 1)!)).toEqual(["A", "D", "B", "C"]);
    expect(texts(moveParagraphs(b, 2, 1, -1)!)).toEqual(["B", "C", "A", "D"]);
    expect(moveParagraphs(b, 0, 1, -1)).toBeNull();
    expect(moveParagraphs(b, 2, 3, 1)).toBeNull();
  });
});

describe("Ana hat: slayt indirme ve satır yükseltme", () => {
  const deck = () => {
    const p = newPresentation("varsayilan", 0);
    p.slides.push(newSlide(p, "obj"), newSlide(p, "obj"));
    const set = (i: number, title: string, body?: string) => {
      const s = p.slides[i];
      const t = titleShape(s)!;
      t.text = textBody(title);
      const b = bodyShape(s);
      if (b && body !== undefined) b.text = textBody(body);
    };
    set(1, "Gündem", "Madde A");
    set(2, "Sonuçlar", "Bulgu 1\nBulgu 2");
    return p;
  };
  const lines = (b: TextBody | undefined) => b!.paragraphs.map((x) => "\t".repeat(x.level) + paraText(x));

  it("başlıkta Tab: slayt önceki slaydın gövdesine (başlık ilk düzey, maddeler bir alt düzey) katılır", () => {
    const p = deck();
    const r = demoteSlide(p, 2)!;
    expect(r.slides).toHaveLength(2);
    expect(lines(bodyShape(r.slides[1])!.text)).toEqual(["Madde A", "Sonuçlar", "\tBulgu 1", "\tBulgu 2"]);
    // İlk slayt indirilemez; notlu ya da başka içerikli slayt indirilmez.
    expect(demoteSlide(p, 0)).toBeNull();
    const q = deck();
    q.slides[2].notes = "Not";
    expect(demoteSlide(q, 2)).toBeNull();
  });

  it("ilk düzeydeki satırda Shift+Tab: satır yeni slaydın başlığı olur, sonraki satırlar taşınır", () => {
    const p = deck();
    const r = promoteLine(p, 2, 0, () => newSlide(p, "obj"))!;
    expect(r.slides).toHaveLength(4);
    expect(titleShape(r.slides[3])!.text!.paragraphs.map(paraText)).toEqual(["Bulgu 1"]);
    expect(lines(bodyShape(r.slides[3])!.text)).toEqual(["Bulgu 2"]);
    expect(lines(bodyShape(r.slides[2])!.text)).toEqual([""]);
    // Alt düzeydeki satır yükseltilmez (önce düzeyi azaltılır).
    const q = deck();
    bodyShape(q.slides[2])!.text!.paragraphs[1].level = 1;
    expect(promoteLine(q, 2, 1, () => newSlide(q, "obj"))).toBeNull();
  });
});
