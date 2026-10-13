/** Akıllı kılavuzlar: hizalama ve eşit aralık yapışması. */
import { describe, expect, it } from "vitest";
import { snapMove } from "./snap";

const W = 1280;
const H = 720;

describe("eşit aralık kılavuzları", () => {
  it("iki nesnenin arasına taşınan nesne ortaya yapışır; iki boşluk işareti çizilir", () => {
    const a = { x: 100, y: 300, w: 100, h: 80 };
    const b = { x: 500, y: 320, w: 100, h: 80 };
    // Ortası: (200 + 500 - 80) / 2 = 310
    const r = snapMove({ x: 306, y: 200, w: 80, h: 160 }, [a, b], W, H, 6);
    expect(r.dx).toBeCloseTo(4);
    const gaps = r.lines.filter((l) => l.gap);
    expect(gaps).toHaveLength(2);
    expect(gaps.map((g) => [g.from, g.to])).toEqual([
      [200, 310],
      [390, 500],
    ]);
    expect(gaps.every((g) => g.dir === "h")).toBe(true);
  });

  it("komşu çiftin boşluğu kadar sağa (ve aşağı) yerleştirmede yapışır", () => {
    const a = { x: 100, y: 100, w: 100, h: 100 };
    const b = { x: 260, y: 100, w: 100, h: 100 };
    // a–b boşluğu 60: c'nin x'i 360 + 60 = 420 olmalı
    const r = snapMove({ x: 423, y: 120, w: 100, h: 50 }, [a, b], W, H, 6);
    expect(r.dx).toBeCloseTo(-3);
    expect(r.lines.filter((l) => l.gap).map((g) => g.to - g.from)).toEqual([60, 60]);
    // Dikey: a üstte, d altta, e daha aşağıda; boşluk 40
    const top = { x: 600, y: 50, w: 100, h: 100 };
    const mid = { x: 600, y: 190, w: 100, h: 100 };
    const v = snapMove({ x: 610, y: 333, w: 80, h: 60 }, [top, mid], W, H, 6);
    expect(v.dy).toBeCloseTo(330 - 333);
    expect(v.lines.filter((l) => l.gap && l.dir === "v")).toHaveLength(2);
  });

  it("satırda örtüşmeyen nesneler eşit aralığa katılmaz; eşik dışında yapışma olmaz", () => {
    const a = { x: 100, y: 600, w: 100, h: 50 };
    const b = { x: 500, y: 600, w: 100, h: 50 };
    const far = snapMove({ x: 306, y: 100, w: 80, h: 80 }, [a, b], W, H, 6);
    expect(far.lines.some((l) => l.gap)).toBe(false);
    const out = snapMove({ x: 330, y: 590, w: 80, h: 80 }, [a, b], W, H, 6);
    expect(out.lines.some((l) => l.gap)).toBe(false);
  });

  it("kenar hizalaması korunur", () => {
    const r = snapMove({ x: 103, y: 400, w: 50, h: 50 }, [{ x: 100, y: 100, w: 200, h: 50 }], W, H, 6);
    expect(r.dx).toBeCloseTo(-3);
    expect(r.lines.some((l) => l.dir === "v" && !l.gap && l.pos === 100)).toBe(true);
  });
});
