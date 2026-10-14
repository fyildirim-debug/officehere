// Birleştirme süpürmesi (bookUtil): çakışma ayıklama ve hücre sorgusu, doğrusala yakın süre.
import { describe, expect, it } from "vitest";
import { disjointMerges, MergeLookup } from "../bookUtil";
import type { Rect } from "../../model/types";

const R = (r1: number, c1: number, r2: number, c2: number): Rect => ({ r1, c1, r2, c2 });
const overlaps = (a: Rect, b: Rect) => a.r1 <= b.r2 && b.r1 <= a.r2 && a.c1 <= b.c2 && b.c1 <= a.c2;

/** Tohumlu sözde rastgele sayılar (deterministik). */
function rng(seed: number) {
  return (n: number) => {
    seed = (seed * 1103515245 + 12345) & 0x7fffffff;
    return seed % n;
  };
}

describe("disjointMerges", () => {
  it("geçerli olanlar özgün sırasıyla kalır; tek hücrelik, ters ve sayfa dışı aralıklar atılır", () => {
    const list = [R(4, 0, 5, 1), R(0, 0, 0, 0), R(0, 0, 1, 1), R(3, 3, 2, 2), R(-1, 0, 1, 1), R(1048575, 0, 1048576, 1), R(0, 16383, 1, 16384), R(0.5, 0, 1, 1)];
    const { merges, invalid, overlapping } = disjointMerges(list);
    expect(merges).toEqual([R(4, 0, 5, 1), R(0, 0, 1, 1)]);
    expect(invalid).toBe(6);
    expect(overlapping).toBe(0);
  });

  it("sayfa sınırına dayanan birleştirmeler geçerlidir", () => {
    const list = [R(1048574, 0, 1048575, 1), R(0, 16382, 1, 16383), R(1048574, 16382, 1048575, 16383)];
    expect(disjointMerges(list)).toEqual({ merges: list, invalid: 0, overlapping: 0, bigDropped: 0 });
  });

  it("dev birleştirme çakıştığı küçük birleştirmelere yenilir", () => {
    const giant = R(0, 0, 1048575, 16383);
    const a = R(1, 2, 3, 4);
    const b = R(5, 2, 8, 2);
    expect(disjointMerges([giant, a, b])).toEqual({ merges: [a, b], invalid: 0, overlapping: 0, bigDropped: 1 });
    expect(disjointMerges([a, giant, b]).merges).toEqual([a, b]);
    // Çakışmayan dev birleştirme kalır; iki dev birleştirmeden sol üstü önde olan kalır
    const far = R(1000, 0, 1_000_000, 10);
    const far2 = R(2000, 5, 1_040_000, 20);
    expect(disjointMerges([a, far, far2])).toEqual({ merges: [a, far], invalid: 0, overlapping: 0, bigDropped: 1 });
  });

  it("çakışmada sol üstü önde olan kalır (sıradan bağımsız)", () => {
    const a = R(0, 0, 1, 1);
    const b = R(1, 1, 2, 2);
    const c = R(0, 3, 0, 4);
    expect(disjointMerges([b, c, a]).merges).toEqual([c, a]);
    expect(disjointMerges([a, b, c]).overlapping).toBe(1);
    // Aynı aralığın kopyaları: biri kalır
    const dup = Array.from({ length: 1000 }, () => R(2, 2, 3, 3));
    const d = disjointMerges(dup);
    expect(d.merges).toHaveLength(1);
    expect(d.overlapping).toBe(999);
  });

  it("rastgele girdide sonuç çakışmasızdır ve kaba kuvvetle aynıdır", () => {
    const ri = rng(7);
    const list: Rect[] = [];
    for (let i = 0; i < 400; i++) {
      const r1 = ri(60);
      const c1 = ri(30);
      list.push(R(r1, c1, r1 + ri(6), c1 + ri(4)));
    }
    const { merges } = disjointMerges(list);
    for (let i = 0; i < merges.length; i++) for (let j = i + 1; j < merges.length; j++) expect(overlaps(merges[i], merges[j])).toBe(false);
    // Kaba kuvvet: sol üst sırasıyla, daha önce kabul edilenlerle çakışmayan kabul edilir
    const order = list
      .map((m, i) => ({ m, i }))
      .filter(({ m }) => m.r1 !== m.r2 || m.c1 !== m.c2)
      .sort((a, b) => a.m.r1 - b.m.r1 || a.m.c1 - b.m.c1 || a.i - b.i);
    const kept: { m: Rect; i: number }[] = [];
    for (const x of order) if (!kept.some((k) => overlaps(k.m, x.m))) kept.push(x);
    expect(merges).toEqual(kept.sort((a, b) => a.i - b.i).map((k) => k.m));
  });

  it("50.000 küçük ve 10.000 uzun birleştirmede hızlıdır (O(n²) değil)", () => {
    const list: Rect[] = [];
    for (let r = 0; r < 50_000; r++) list.push(r % 2 ? R(r, 1, r, 2) : R(r, 0, r, 1));
    // Karışık sütunlarda başlayıp farklı satırlarda biten uzun sütun birleştirmeleri (etkin küme büyük)
    for (let c = 0; c < 10_000; c++) {
      const col = 10 + ((c * 7919) % 10_000);
      list.push(R(60_000 + c, col, 60_001 + c + ((c * 104_729) % 900_000), col));
    }
    const t = performance.now();
    const { merges, overlapping } = disjointMerges(list);
    expect(performance.now() - t).toBeLessThan(2_000);
    expect(merges).toHaveLength(60_000);
    expect(overlapping).toBe(0);
  });
});

describe("MergeLookup", () => {
  // Çakışmasız: 3×4'lük bloklar, bazı bloklar boş
  const merges: Rect[] = [];
  for (let r = 0; r < 60; r += 3) for (let c = 0; c < 40; c += 4) if ((r + c) % 7) merges.push(R(r, c, r + 1 + ((r + c) % 2), c + 2));
  merges.push(R(70, 0, 900, 0), R(70, 1, 75, 3));
  const brute = (r: number, c: number) => merges.find((m) => r >= m.r1 && r <= m.r2 && c >= m.c1 && c <= m.c2);

  it("satır sırasıyla sorguda kaba kuvvetle aynıdır", () => {
    const q = new MergeLookup(merges);
    for (let r = 0; r < 100; r++) for (let c = 0; c < 45; c++) expect(q.at(r, c)).toBe(brute(r, c));
    expect(q.at(900, 0)).toBe(merges[merges.length - 2]);
    expect(q.at(901, 0)).toBeUndefined();
  });

  it("geriye dönen ve satır atlayan sorgularda da doğrudur", () => {
    const q = new MergeLookup(merges);
    const ri = rng(11);
    for (let i = 0; i < 3000; i++) {
      const r = ri(100);
      const c = ri(45);
      expect(q.at(r, c)).toBe(brute(r, c));
    }
  });

  it("boş listede hiçbir hücre birleşik değildir", () => {
    expect(new MergeLookup([]).at(0, 0)).toBeUndefined();
  });
});
