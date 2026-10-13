import { readFileSync } from "node:fs";
import { resolve } from "node:path";
import { describe, expect, it } from "vitest";
import { isCfb, readCfb } from "../src/cfb";

const fixture = (name: string) => new Uint8Array(readFileSync(resolve(__dirname, "fixtures", name)));

describe("CFB okuyucusu", () => {
  it("imzayı tanır", () => {
    expect(isCfb(fixture("ornek.doc"))).toBe(true);
    expect(isCfb(new Uint8Array(600))).toBe(false);
    expect(isCfb(new Uint8Array([0xd0, 0xcf]))).toBe(false);
  });

  it(".xls: Workbook akışı ve küçük akışlar okunur, boyutlar dizinle aynı", () => {
    const cfb = readCfb(fixture("ornek.xls"));
    const names = cfb.entries.map((e) => e.path);
    expect(names).toContain("Workbook");
    for (const e of cfb.entries.filter((x) => x.type === "stream")) {
      expect(cfb.read(e.path)?.length).toBe(e.size);
    }
    const wb = cfb.read("workbook")!; // büyük/küçük harfe duyarsız
    // BIFF8 akışı BOF kaydıyla (0x0809) başlar.
    expect(wb[0] | (wb[1] << 8)).toBe(0x0809);
    // 4096 baytın altındaki akış küçük akıştan (mini stream) gelir.
    const small = cfb.entries.find((e) => e.type === "stream" && e.size > 0 && e.size < 4096);
    expect(small).toBeTruthy();
    expect(cfb.read(small!.path)!.some((x) => x !== 0)).toBe(true);
  });

  it(".doc: WordDocument akışı FIB imzasıyla başlar, tablo akışı vardır", () => {
    const cfb = readCfb(fixture("ornek.doc"));
    const wd = cfb.read("WordDocument")!;
    expect(wd[0] | (wd[1] << 8)).toBe(0xa5ec);
    expect(cfb.has("1Table") || cfb.has("0Table")).toBe(true);
    expect(cfb.read("Olmayan")).toBeNull();
  });

  it("bozuk dosyada sonsuz döngüye girmez, hata verir", () => {
    const b = fixture("ornek.xls").slice();
    // İlk dizin kesimini kendine bağla: FAT'te döngü.
    const dv = new DataView(b.buffer);
    const firstDir = dv.getUint32(0x30, true);
    const fatSector = dv.getUint32(0x4c, true);
    dv.setUint32((fatSector + 1) * 512 + firstDir * 4, firstDir, true);
    expect(() => readCfb(b)).toThrow();
    expect(() => readCfb(new Uint8Array(512).fill(0xd0))).toThrow();
  });
});
