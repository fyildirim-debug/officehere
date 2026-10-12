import { describe, expect, it } from "vitest";
import { createTypeahead, firstEnabled, foldLabel, lastEnabled, matchByPrefix, nextEnabled, scrollToReveal, spatialNext, type Box } from "./menuNav";

describe("sonraki etkin öğe", () => {
  const disabled = new Set([1, 3]);
  const en = (i: number) => !disabled.has(i);

  it("devre dışı öğeleri atlar ve uçlarda döner", () => {
    expect(nextEnabled(5, 0, 1, en)).toBe(2);
    expect(nextEnabled(5, 2, 1, en)).toBe(4);
    expect(nextEnabled(5, 4, 1, en)).toBe(0);
    expect(nextEnabled(5, 0, -1, en)).toBe(4);
    expect(nextEnabled(5, 4, -1, en)).toBe(2);
  });

  it("seçili öğe yokken ↓ ilk, ↑ son etkin öğeye gider", () => {
    expect(nextEnabled(5, -1, 1, en)).toBe(0);
    expect(nextEnabled(5, -1, -1, (i) => i !== 4)).toBe(3);
    expect(firstEnabled(4, (i) => i > 1)).toBe(2);
    expect(lastEnabled(4, (i) => i < 2)).toBe(1);
  });

  it("dönmeden uca gelince -1; hiç etkin öğe yoksa -1", () => {
    expect(nextEnabled(5, 4, 1, en, false)).toBe(-1);
    expect(nextEnabled(5, 0, -1, en, false)).toBe(-1);
    expect(nextEnabled(3, 0, 1, () => false)).toBe(-1);
    expect(nextEnabled(0, 0, 1)).toBe(-1);
  });
});

describe("harfle atlama (Türkçe)", () => {
  const labels = ["Kes", "Kopyala", "Yapıştır", "İçindekiler", "Işık", "ızgara", "(Tanımlı ad yok)", "Ölçü"];

  it("ilk harf eşleşmesi geçerli öğeden sonra aranır ve döner", () => {
    expect(matchByPrefix(labels, -1, "k")).toBe(0);
    expect(matchByPrefix(labels, 0, "k")).toBe(1);
    expect(matchByPrefix(labels, 1, "k")).toBe(0);
    expect(matchByPrefix(labels, 0, "kk")).toBe(1);
  });

  it("Türkçe büyük/küçük harf: i ↔ İ, ı ↔ I", () => {
    expect(matchByPrefix(labels, -1, "i")).toBe(3);
    expect(matchByPrefix(labels, -1, "İ")).toBe(3);
    expect(matchByPrefix(labels, -1, "ı")).toBe(4);
    expect(matchByPrefix(labels, 4, "I")).toBe(5);
    expect(matchByPrefix(labels, -1, "ö")).toBe(7);
  });

  it("birden çok harf geçerli öğeden başlar; baştaki imler yok sayılır", () => {
    expect(matchByPrefix(labels, 1, "ko")).toBe(1);
    expect(matchByPrefix(labels, 0, "kop")).toBe(1);
    expect(matchByPrefix(labels, -1, "tan")).toBe(6);
    expect(matchByPrefix(labels, -1, "zz")).toBe(-1);
    expect(matchByPrefix(labels, -1, " ")).toBe(-1);
  });

  it("devre dışı öğeler atlanır", () => {
    expect(matchByPrefix(labels, -1, "k", (i) => i !== 0)).toBe(1);
  });

  it("etiket katlama", () => {
    expect(foldLabel("İNDİR")).toBe("indir");
    expect(foldLabel("…Diğer")).toBe("diğer");
  });

  it("art arda basılan harfler birikir, ara verilince sıfırlanır", () => {
    const t = createTypeahead(500);
    expect(t("k", 0)).toBe("k");
    expect(t("o", 100)).toBe("ko");
    expect(t("p", 700)).toBe("p");
  });
});

describe("ızgara/liste ok tuşu hedefi", () => {
  // 3×2 ızgara (20 px kare, 4 px aralık), altında tam genişlikte bir öğe.
  const cell = (c: number, r: number): Box => ({ left: c * 24, top: r * 24, right: c * 24 + 20, bottom: r * 24 + 20 });
  const grid: Box[] = [cell(0, 0), cell(1, 0), cell(2, 0), cell(0, 1), cell(1, 1), cell(2, 1), { left: 0, top: 52, right: 68, bottom: 76 }];

  it("aynı sütunda satır değiştirir; alttaki tam genişlik öğeye iner", () => {
    expect(spatialNext(grid, 1, "down")).toBe(4);
    expect(spatialNext(grid, 4, "down")).toBe(6);
    expect(spatialNext(grid, 4, "up")).toBe(1);
    expect(spatialNext(grid, 6, "up")).toBe(3);
  });

  it("uçlarda döner", () => {
    expect(spatialNext(grid, 6, "down")).toBe(0);
    expect(spatialNext(grid, 2, "up")).toBe(6);
  });

  it("sol/sağ aynı satırda; satır ucunda -1", () => {
    expect(spatialNext(grid, 0, "right")).toBe(1);
    expect(spatialNext(grid, 5, "left")).toBe(4);
    expect(spatialNext(grid, 2, "right")).toBe(-1);
    expect(spatialNext(grid, 6, "left")).toBe(-1);
  });

  it("tek sütunlu liste düz sırayla gezilir, devre dışı atlanır", () => {
    const list: Box[] = [0, 1, 2, 3].map((i) => ({ left: 0, top: i * 28, right: 160, bottom: i * 28 + 28 }));
    expect(spatialNext(list, 0, "down")).toBe(1);
    expect(spatialNext(list, 0, "down", (i) => i !== 1)).toBe(2);
    expect(spatialNext(list, 3, "down")).toBe(0);
    expect(spatialNext(list, 0, "up")).toBe(3);
  });

  it("listede kaydırma: görünür öğe için değişmez, dışındaki için kayar", () => {
    expect(scrollToReveal(0, 100, 20, 20)).toBe(0);
    expect(scrollToReveal(0, 100, 120, 20)).toBe(42);
    expect(scrollToReveal(80, 100, 40, 20)).toBe(38);
  });
});
