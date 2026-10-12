import { describe, expect, it } from "vitest";
import { alignDrawing, layerDrawing, positionDrawing, wrapDrawing, type Drawing, type Prim, type WrapKind } from "./arrangeGeometry";

type Box = Extract<Prim, { k: "rect" }>;
const rects = (d: Drawing, ink?: string) => d.prims.filter((p): p is Box => p.k === "rect" && (!ink || p.ink === ink));
const overlaps = (a: Box, b: Box) => a.x < b.x + b.w && b.x < a.x + a.w && a.y < b.y + b.h && b.y < a.y + a.h;
const rowsBeside = (d: Drawing, pic: Box) => rects(d, "line").filter((l) => l.y < pic.y + pic.h && pic.y < l.y + l.h);

const WRAPS: WrapKind[] = ["inline", "square", "tight", "topBottom", "behind", "front"];
const SIZES = [16, 20, 24, 32];

describe("yerleşim çizimleri: piksel ızgarası", () => {
  it("düz kenarlar her boyutta tamsayı koordinatta ve kutunun içinde", () => {
    const all: Drawing[] = [];
    for (const s of SIZES) {
      for (const m of WRAPS) all.push(wrapDrawing(m, s));
      for (const k of ["forward", "front", "backward", "back"] as const) all.push(layerDrawing(k, s));
      for (const k of ["left", "center", "right", "top", "middle", "bottom", "distH", "distV"] as const) all.push(alignDrawing(k, s));
    }
    all.push(positionDrawing({ v: "middle", h: "left" }, 40, 52), positionDrawing("inline", 40, 52), positionDrawing({ v: "top", h: "right" }, 24, 30));
    for (const d of all) {
      for (const p of d.prims) {
        if (p.k === "path") continue;
        for (const v of [p.x, p.y, p.w, p.h]) expect(Number.isInteger(v)).toBe(true);
        expect(p.w).toBeGreaterThan(0);
        expect(p.h).toBeGreaterThan(0);
        expect(p.x).toBeGreaterThanOrEqual(0);
        expect(p.y).toBeGreaterThanOrEqual(0);
        expect(p.x + p.w).toBeLessThanOrEqual(d.w);
        expect(p.y + p.h).toBeLessThanOrEqual(d.h);
      }
    }
  });

  it("satır kalınlığı 16-24 px'te 1, 32 px'te 2 piksel", () => {
    expect(new Set(rects(wrapDrawing("front", 16), "line").map((l) => l.h))).toEqual(new Set([1]));
    expect(new Set(rects(wrapDrawing("front", 24), "line").map((l) => l.h))).toEqual(new Set([1]));
    expect(new Set(rects(wrapDrawing("front", 32), "line").map((l) => l.h))).toEqual(new Set([2]));
  });
});

describe("Metni Kaydır çizimleri", () => {
  for (const s of SIZES) {
    it(`${s} px: satırlar seçeneğe göre resmin çevresinden akar`, () => {
      const sq = wrapDrawing("square", s);
      const pic = rects(sq, "pic")[0];
      // Kare: satırlar resme değmez, resmin yanında kısalır.
      expect(rects(sq, "line").some((l) => overlaps(l, pic))).toBe(false);
      expect(rowsBeside(sq, pic).length).toBe(3);
      // Üst ve Alt: resmin yanında satır yok.
      const tb = wrapDrawing("topBottom", s);
      expect(rowsBeside(tb, rects(tb, "pic")[0])).toHaveLength(0);
      // Satır içi: yalnız taban satırında resmin iki yanında parça.
      const il = wrapDrawing("inline", s);
      const ip = rects(il, "pic")[0];
      const side = rowsBeside(il, ip);
      expect(side).toHaveLength(2);
      expect(side[0].y + side[0].h).toBe(ip.y + ip.h);
      // Arkasında: soluk resim önce çizilir, satırlar üstünden geçer; önünde: resim satırlardan sonra.
      const bh = wrapDrawing("behind", s);
      expect(bh.prims[0]).toMatchObject({ ink: "ghost" });
      expect(rects(bh, "line").some((l) => overlaps(l, rects(bh, "ghost")[0]))).toBe(true);
      const fr = wrapDrawing("front", s);
      const fp = fr.prims.findIndex((p) => p.ink === "pic");
      expect(fr.prims.slice(fp).some((p) => p.ink === "line")).toBe(false);
      expect(rowsBeside(fr, rects(fr, "pic")[0]).length).toBe(3);
    });
  }

  it("Sıkı: satırlar resmin daraldığı yerde karedekinden daha yakın başlar", () => {
    for (const s of SIZES) {
      const sqStart = Math.min(...rects(wrapDrawing("square", s), "line").filter((l) => l.x > 2).map((l) => l.x));
      const tightStarts = rects(wrapDrawing("tight", s), "line").filter((l) => l.x > 2).map((l) => l.x);
      expect(Math.min(...tightStarts)).toBeLessThan(sqStart);
      expect(new Set(tightStarts).size).toBeGreaterThan(1);
    }
  });
});

describe("Konum sayfa önizlemesi", () => {
  it("sol/sağda satırlar resmin öbür yanında, ortada resmin yanı boş", () => {
    const left = positionDrawing({ v: "top", h: "left" }, 40, 52);
    const lp = rects(left, "pic")[0];
    expect(lp.y).toBe(rects(left, "line")[0].y);
    expect(rowsBeside(left, lp).every((l) => l.x >= lp.x + lp.w)).toBe(true);
    const right = positionDrawing({ v: "bottom", h: "right" }, 40, 52);
    const rp = rects(right, "pic")[0];
    expect(rp.x + rp.w).toBe(35);
    expect(rowsBeside(right, rp).every((l) => l.x + l.w <= rp.x)).toBe(true);
    const mid = positionDrawing({ v: "middle", h: "center" }, 40, 52);
    const mp = rects(mid, "pic")[0];
    expect(rowsBeside(mid, mp)).toHaveLength(0);
    // Tam ortada: sayfanın yatay ve dikey ortası resmin ortası.
    expect(mp.x + mp.w / 2).toBe(20);
    expect(mp.y + mp.h / 2).toBe(26);
  });

  it("dokuz konum birbirinden farklı", () => {
    const keys = new Set<string>();
    for (const v of ["top", "middle", "bottom"] as const)
      for (const h of ["left", "center", "right"] as const) {
        const p = rects(positionDrawing({ v, h }, 40, 52), "pic")[0];
        keys.add(`${p.x},${p.y}`);
      }
    expect(keys.size).toBe(9);
  });
});

describe("Öne Getir / Arkaya Gönder ve Hizala çizimleri", () => {
  it("vurgu kartı öne getirince en son, arkaya gönderince en önce çizilir", () => {
    for (const s of SIZES) {
      const cards = (k: "forward" | "front" | "backward" | "back") => layerDrawing(k, s).prims.filter((p) => p.ink === "pic" || p.ink === "page");
      expect(cards("forward")).toHaveLength(2);
      expect(cards("front")).toHaveLength(3);
      expect(cards("forward").at(-1)!.ink).toBe("pic");
      expect(cards("front").at(-1)!.ink).toBe("pic");
      expect(cards("backward")[0].ink).toBe("pic");
      expect(cards("back")[0].ink).toBe("pic");
    }
  });

  it("çubuklar kılavuza göre hizalı; dağıtta aralıklar eşit", () => {
    for (const s of SIZES) {
      const [a, b] = rects(alignDrawing("left", s), "pic");
      expect(a.x).toBe(b.x);
      const [c, d] = rects(alignDrawing("right", s), "pic");
      expect(c.x + c.w).toBe(d.x + d.w);
      const [e, f] = rects(alignDrawing("center", s), "pic");
      const g = rects(alignDrawing("center", s), "ink")[0];
      expect(e.x * 2 + e.w).toBe(g.x * 2 + g.w);
      expect(f.x * 2 + f.w).toBe(g.x * 2 + g.w);
      const [t1, t2] = rects(alignDrawing("top", s), "pic");
      expect(t1.y).toBe(t2.y);
      const bars = rects(alignDrawing("distH", s), "pic");
      expect(bars).toHaveLength(3);
      expect(bars[1].x - (bars[0].x + bars[0].w)).toBe(bars[2].x - (bars[1].x + bars[1].w));
    }
  });
});
