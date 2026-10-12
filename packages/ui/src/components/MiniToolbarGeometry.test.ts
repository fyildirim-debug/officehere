import { describe, expect, it } from "vitest";
import { miniBarPosition, nearestRect } from "./MiniToolbarGeometry";

const view = { w: 1000, h: 800 };
const size = { w: 300, h: 70 };

describe("mini araç çubuğu yerleşimi", () => {
  it("seçim kipinde seçimin üstünde ve biraz solunda", () => {
    expect(miniBarPosition({ x: 200, y: 300, context: false }, size, view)).toEqual({ left: 180, top: 220 });
  });

  it("üstte yer yoksa seçimin altına iner", () => {
    expect(miniBarPosition({ x: 200, y: 40, context: false }, size, view)).toEqual({ left: 180, top: 68 });
    expect(miniBarPosition({ x: 200, y: 40, bottom: 90, context: false }, size, view)).toEqual({ left: 180, top: 98 });
  });

  it("pencerenin sağından taşmaz", () => {
    expect(miniBarPosition({ x: 900, y: 300, context: false }, size, view).left).toBe(696);
  });

  it("bağlam kipinde tıklama noktasının ya da menünün üstünde", () => {
    expect(miniBarPosition({ x: 200, y: 300, context: true }, size, view)).toEqual({ left: 200, top: 224 });
    // Alta sığmayan menü yukarı açılmıştır: çubuk menünün üstüne çıkar.
    const menu = { left: 200, top: 400, right: 430, bottom: 700 };
    expect(miniBarPosition({ x: 200, y: 700, context: true }, size, view, menu)).toEqual({ left: 200, top: 324 });
    // Menünün üstünde yer yoksa altına.
    expect(miniBarPosition({ x: 200, y: 20, context: true }, size, view, { left: 200, top: 20, right: 430, bottom: 320 })).toEqual({ left: 200, top: 326 });
  });

  it("noktaya bitişik menü bulunur, uzaktaki yok sayılır", () => {
    const a = { left: 100, top: 100, right: 200, bottom: 300 };
    const b = { left: 500, top: 500, right: 600, bottom: 600 };
    expect(nearestRect(100, 100, [b, a])).toBe(a);
    expect(nearestRect(300, 100, [a, b])).toBeNull();
  });
});
