/** Şekilleri Birleştir: çokgen işlemleri. */
import { describe, expect, it } from "vitest";
import { makeSp } from "../model/factory";
import { area, cmdsToRings, fragmentRings, mergeRings, mergeShapes, ringOp, type Ring } from "./mergeShapes";

const sq = (x: number, y: number, s: number): Ring => [
  [x, y],
  [x + s, y],
  [x + s, y + s],
  [x, y + s],
];
const total = (rs: Ring[]) => rs.reduce((t, r) => t + Math.abs(area(r)), 0);

describe("çokgen işlemleri", () => {
  const a = sq(0, 0, 10);
  const b = sq(5, 5, 10);
  it("kesişen kareler: birleşim, kesişim, fark", () => {
    expect(total(ringOp(a, b, "union"))).toBeCloseTo(175, 3);
    expect(total(ringOp(a, b, "intersect"))).toBeCloseTo(25, 3);
    expect(total(ringOp(a, b, "subtract"))).toBeCloseTo(75, 3);
  });
  it("ayrık ve iç içe durumlar", () => {
    const far = sq(50, 50, 5);
    expect(ringOp(a, far, "union")).toHaveLength(2);
    expect(ringOp(a, far, "intersect")).toHaveLength(0);
    const inner = sq(2, 2, 3);
    expect(total(ringOp(a, inner, "union"))).toBeCloseTo(100, 3);
    expect(total(ringOp(a, inner, "intersect"))).toBeCloseTo(9, 3);
    // Delik: dış halka + ters yönlü iç halka.
    const hole = ringOp(a, inner, "subtract");
    expect(hole).toHaveLength(2);
    expect(Math.sign(area(hole[0]))).toBe(-Math.sign(area(hole[1])));
  });
  it("üç şekil birleşimi ve tek-çift birleştirme", () => {
    const c = sq(12, 0, 6);
    expect(total(mergeRings([[a], [b], [c]], "union"))).toBeCloseTo(175 + 36 - 3, 2);
    expect(mergeRings([[a], [b]], "combine")).toHaveLength(2);
  });
  it("eğriler örneklenir", () => {
    const rings = cmdsToRings([{ c: "M", pts: [0, 0] }, { c: "C", pts: [0, 10, 10, 10, 10, 0] }, { c: "Z", pts: [] }]);
    expect(rings[0].length).toBeGreaterThan(10);
  });
  it("şekiller tek özel şekle dönüşür", () => {
    const r1 = makeSp("rect", 0, 0, 100, 100);
    const r2 = makeSp("ellipse", 50, 50, 100, 100);
    const [u] = mergeShapes([r1, r2], "union");
    expect(u.geom).toBe("custom");
    expect(u.w).toBeCloseTo(150, 0);
    expect(mergeShapes([r1, makeSp("rect", 300, 300, 10, 10)], "intersect")).toEqual([]);
  });
  it("Parçala: örtüşen iki kare üç parçaya ayrılır, alanlar korunur", () => {
    const parts = fragmentRings([[a], [b]]);
    expect(parts).toHaveLength(3);
    expect(parts.map((p) => total(p)).sort((x, y) => x - y)).toEqual([expect.closeTo(25, 2), expect.closeTo(75, 2), expect.closeTo(75, 2)]);
    const shapes = mergeShapes([makeSp("rect", 0, 0, 100, 100), makeSp("rect", 50, 0, 100, 100)], "fragment");
    expect(shapes).toHaveLength(3);
    expect(shapes.every((x) => x.geom === "custom")).toBe(true);
  });
});
