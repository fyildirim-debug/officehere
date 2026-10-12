// @vitest-environment jsdom
/** Başka bir uygulamanın (LibreOffice Impress) ürettiği PPTX dosyalarını okuma. */
import { describe, expect, it } from "vitest";
import { flatShapes } from "../model/ops";
import type { Presentation } from "../model/types";
import { plainText } from "../model/textStyle";
import { readPptx } from "./pptxReader";
import { writePptx } from "./pptxWriter";

async function load(name: string): Promise<Presentation> {
  const fs = (await import(/* @vite-ignore */ "node:" + "fs")) as { readFileSync: (p: string) => Uint8Array };
  const proc = (globalThis as unknown as { process: { cwd: () => string } }).process;
  return readPptx(fs.readFileSync(`${proc.cwd()}/src/io/__fixtures__/${name}`));
}

function allText(p: Presentation): string {
  return p.slides.map((s) => flatShapes(s.shapes).map((sh) => (sh.kind === "sp" ? plainText(sh.text) : sh.kind === "table" ? sh.rows.map((r) => r.cells.map((c) => plainText(c.text)).join(" ")).join(" ") : "")).join(" ")).join(" ");
}

describe("LibreOffice PPTX içe aktarma", () => {
  it("proje sunusu: slaytlar, metinler, tablo ve notlar", async () => {
    const p = await load("lo-proje.pptx");
    expect(p.slides.length).toBe(7);
    const text = allText(p);
    for (const w of ["Müşteri Portalı Yenileme Projesi", "Gündem", "Hedefler", "Zaman Çizelgesi", "Teşekkürler", "Kalite Güvence"]) expect(text).toContain(w);
    const s4 = p.slides[3];
    expect(flatShapes(s4.shapes).some((s) => s.kind === "table")).toBe(true);
    expect(p.slides[4].notes).toContain("Ekip üyelerini");
    // Tema renkleri okunur.
    expect(p.theme.colors.accent1).toMatch(/^#[0-9A-F]{6}$/);
    // Her slaydın geçerli bir düzeni vardır.
    for (const s of p.slides) expect(p.layouts.some((l) => l.id === s.layoutId)).toBe(true);
    // Birden çok asıl slayt: içerik slaytları kendi asıl slaytlarının görünümünü (beyaz zemin + dalgalar) alır.
    const lay = (i: number) => p.layouts.find((l) => l.id === p.slides[i].layoutId)!;
    expect(lay(1).id).not.toBe(lay(0).id);
    expect(lay(1).name).toBe("Başlık ve İçerik");
    expect(lay(1).background).toMatchObject({ type: "solid", color: { scheme: "lt1" } });
    expect(lay(1).shapes.filter((x) => !(x.kind === "sp" && x.ph)).length).toBeGreaterThan(0);
  });

  it("açıklamalar: yazar ve metinler okunur", async () => {
    const p = await load("lo-yorum.pptx");
    expect(p.slides.length).toBe(1);
    const c = p.slides[0].comments ?? [];
    expect(c.map((x) => x.text)).toEqual(["İlk açıklama: başlık rengi koyulaşsın.", "İkinci açıklama"]);
    expect(c[0].author).toBe("Kullanıcı");
    expect(Number.isFinite(c[0].date)).toBe(true);
    expect(allText(p)).toContain("Açıklamalı Sunu");
  });

  it("ürün sunusu: şekiller ve yeniden yazma", async () => {
    const p = await load("lo-urun.pptx");
    expect(p.slides.length).toBe(5);
    expect(allText(p)).toContain("Nova X1 Akıllı Saat");
    const geoms = new Set(p.slides.flatMap((s) => flatShapes(s.shapes)).flatMap((s) => (s.kind === "sp" ? [s.geom] : [])));
    expect(geoms.has("heart") || geoms.has("custom")).toBe(true);
    // Okunan dosya yeniden yazılıp okunabilmeli.
    const again = await readPptx(await writePptx(p));
    expect(again.slides.length).toBe(5);
    expect(allText(again)).toContain("Nova X1 Akıllı Saat");
  });
});
