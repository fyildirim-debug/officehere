import { describe, expect, it } from "vitest";
import { cascadeFrom, centerIn, isReachable, planWindow, sanitizeSaved } from "../src/main/windowBounds";

const primary = { x: 0, y: 0, width: 1920, height: 1040 };
const second = { x: 1920, y: 0, width: 1280, height: 1024 };
const defaults = { width: 1360, height: 860 };

describe("pencere yerleşimi", () => {
  it("kayıt yoksa varsayılan boyut, konum Electron'a bırakılır", () => {
    expect(planWindow({ defaults, areas: [primary], primary })).toEqual({ width: 1360, height: 860, maximized: false });
  });

  it("ekrana düşen kayıt olduğu gibi geri yüklenir", () => {
    const saved = { x: 2000, y: 40, width: 900, height: 600, maximized: true };
    expect(planWindow({ defaults, saved, areas: [primary, second], primary })).toEqual({ x: 2000, y: 40, width: 900, height: 600, maximized: true });
  });

  it("hiçbir ekrana düşmeyen konum birincil ekranda ortalanır", () => {
    const saved = { x: 2000, y: 40, width: 900, height: 600 };
    expect(planWindow({ defaults, saved, areas: [primary], primary })).toEqual({ x: 510, y: 220, width: 900, height: 600, maximized: false });
    // Başlık çubuğu ekranın üstünde ya da sağında kalmışsa erişilemez sayılır.
    expect(isReachable({ x: 100, y: -500, width: 900, height: 600 }, [primary])).toBe(false);
    expect(isReachable({ x: 1880, y: 10, width: 900, height: 600 }, [primary])).toBe(false);
    expect(isReachable({ x: 1800, y: 10, width: 900, height: 600 }, [primary])).toBe(true);
  });

  it("ekran küçülmüşse boyut çalışma alanına sığdırılır", () => {
    const saved = { x: 0, y: 0, width: 2400, height: 1400 };
    expect(planWindow({ defaults, saved, areas: [primary], primary })).toMatchObject({ width: 1920, height: 1040 });
    expect(centerIn({ width: 3000, height: 500 }, primary)).toEqual({ x: 0, y: 270, width: 1920, height: 500 });
  });

  it("aynı türden ikinci pencere 24 px kaydırılır, alandan taşacaksa köşeye döner", () => {
    const previous = { x: 100, y: 80, width: 1000, height: 700 };
    expect(planWindow({ defaults, previous, areas: [primary], primary })).toEqual({ x: 124, y: 104, width: 1000, height: 700, maximized: false });
    expect(cascadeFrom({ x: 900, y: 330, width: 1000, height: 700 }, primary)).toEqual({ x: 0, y: 0, width: 1000, height: 700 });
    // İkinci ekrandaki pencere o ekranın içinde kayar.
    expect(cascadeFrom({ x: 1970, y: 50, width: 800, height: 600 }, second)).toEqual({ x: 1994, y: 74, width: 800, height: 600 });
    // Alandan geniş pencere yalnız başlık çubuğu erişilebilir kaldıkça kayar.
    expect(cascadeFrom({ x: -40, y: 82, width: 1360, height: 860 }, { x: 0, y: 0, width: 1280, height: 1024 })).toEqual({ x: -16, y: 106, width: 1360, height: 860 });
  });

  it("bozuk kayıt yok sayılır", () => {
    expect(sanitizeSaved(null)).toBeUndefined();
    expect(sanitizeSaved({ width: "a", height: 10 })).toBeUndefined();
    expect(sanitizeSaved({ width: 800.4, height: 600, x: "?", y: 3 })).toEqual({ width: 800, height: 600, maximized: false });
  });
});
