// @vitest-environment jsdom
/**
 * Dosya G/Ç sağlamlığı: eksik/bozuk parçalar uyarıyla (yerine boş slayt ya da
 * paket içinden kurtarma) açılır, sessiz kayıp olmaz; zip bombası ve dev
 * parçalar sınırla reddedilir; tüm hatalar Türkçe ve sınıflandırılmıştır.
 */
import { readFileSync } from "node:fs";
import JSZip from "jszip";
import { afterEach, describe, expect, it, vi } from "vitest";
import { flatShapes } from "../model/ops";
import { plainText } from "../model/textStyle";
import type { Presentation } from "../model/types";
import { FormatError, readPresentationFile, toFormatError } from "./formats";
import { LoadWarnings, ordinalList } from "./loadWarnings";
import { readTransition } from "./readTiming";
import { parseXml } from "./xml";
import { ZIP_LIMITS } from "./zipSafe";

// Paket ayrıştırma ağır: yüklü makinede (paralel CI işleri) 5 sn varsayılan yetmeyebilir.
vi.setConfig({ testTimeout: 30_000 });

const FIX = `${process.cwd()}/src/io/__fixtures__`;
const load = (n: string) => new Uint8Array(readFileSync(`${FIX}/${n}`));
const layoutName = (p: Presentation, i: number) => p.layouts.find((l) => l.id === p.slides[i].layoutId)?.name;
const texts = (p: Presentation) => p.slides.map((s) => flatShapes(s.shapes).map((sh) => (sh.kind === "sp" ? plainText(sh.text) : "")).join(" "));

async function mutate(name: string, fn: (z: JSZip) => void | Promise<void>): Promise<Uint8Array> {
  const z = await JSZip.loadAsync(load(name));
  await fn(z);
  return z.generateAsync({ type: "uint8array", compression: "DEFLATE" });
}

const LIMITS = { ...ZIP_LIMITS };
afterEach(() => Object.assign(ZIP_LIMITS, LIMITS));

describe("PPTX: eksik ve bozuk parçalar sessizce kaybolmaz", () => {
  it("silinen slayt parçasının yerine boş slayt konur ve uyarı verilir (slayt sayısı korunur)", async () => {
    const orig = await readPresentationFile(load("lo-proje.pptx"), "a.pptx");
    expect(orig.warnings).toEqual([]);
    const bytes = await mutate("lo-proje.pptx", (z) => void z.remove("ppt/slides/slide3.xml"));
    const r = await readPresentationFile(bytes, "a.pptx");
    expect(r.pres.slides).toHaveLength(orig.pres.slides.length);
    expect(r.pres.slides[2].shapes).toEqual([]);
    expect(texts(r.pres)[1]).toBe(texts(orig.pres)[1]);
    expect(texts(r.pres)[3]).toBe(texts(orig.pres)[3]);
    expect(r.warnings).toEqual(["3. slayt dosyada bulunamadı; yerine boş slayt kondu."]);
    expect(r.details.join("\n")).toContain("ppt/slides/slide3.xml");
  });

  it("sunu ilişkileri bozuk (hedef yok / döngü) ise slaytlar, asıl slayt ve tema paket içinden kurtarılır", async () => {
    const orig = await readPresentationFile(load("lo-proje.pptx"), "a.pptx");
    for (const target of ['Target="../../yok/olmayan.xml"', 'Target="presentation.xml"']) {
      const bytes = await mutate("lo-proje.pptx", async (z) => {
        const rels = await z.file("ppt/_rels/presentation.xml.rels")!.async("string");
        z.file("ppt/_rels/presentation.xml.rels", rels.replace(/Target="[^"]+"/g, target));
      });
      const r = await readPresentationFile(bytes, "a.pptx");
      expect(r.pres.slides).toHaveLength(orig.pres.slides.length);
      expect(texts(r.pres)).toEqual(texts(orig.pres));
      expect(r.pres.theme.colors).toEqual(orig.pres.theme.colors);
      expect(r.pres.slides.map((_, i) => layoutName(r.pres, i))).toEqual(orig.pres.slides.map((_, i) => layoutName(orig.pres, i)));
      expect(r.warnings.join(" ")).toMatch(/paket içinden bulunarak açıldı/);
    }
  });

  it("ilişki bölümü tamamen silinse de sunu kurtarılır", async () => {
    const orig = await readPresentationFile(load("lo-proje.pptx"), "a.pptx");
    const r = await readPresentationFile(await mutate("lo-proje.pptx", (z) => void z.remove("ppt/_rels/presentation.xml.rels")), "a.pptx");
    expect(texts(r.pres)).toEqual(texts(orig.pres));
    expect(r.warnings.length).toBeGreaterThan(0);
  });

  it("hiçbir slayt okunamıyorsa (kurtarılacak parça yok) anlaşılır hata", async () => {
    const bytes = await mutate("lo-proje.pptx", (z) => {
      for (const p of Object.keys(z.files)) if (/^ppt\/slides\/slide\d+\.xml$/.test(p)) z.remove(p);
    });
    const e = await readPresentationFile(bytes, "a.pptx").catch((x) => x);
    expect(e).toBeInstanceOf(FormatError);
    expect(e.message).toMatch(/slaytların hiçbiri okunamadı/);
  });

  it("bozuk ya da boş slayt XML'i: yerine boş slayt + uyarı; asıl slayt/tema/düzen eksikliği bildirilir", async () => {
    const bytes = await mutate("lo-proje.pptx", (z) => {
      z.file("ppt/slides/slide2.xml", "<p:sld><yarım");
      z.file("ppt/slides/slide4.xml", "");
      z.remove("ppt/theme/theme1.xml");
      z.remove("ppt/slideLayouts/slideLayout2.xml");
    });
    const r = await readPresentationFile(bytes, "a.pptx");
    expect(r.pres.slides).toHaveLength(7);
    const w = r.warnings.join("\n");
    expect(w).toMatch(/2\. ve 4\. slaytlar bozuk olduğu için okunamadı; yerlerine boş slayt kondu\./);
    expect(w).toMatch(/Tema okunamadı/);
    expect(w).toMatch(/slayt düzenleri okunamadı/);
  });

  it("silinen asıl slayt: düzenler paket içinden bulunur, slaytlar düzenlerini korur", async () => {
    const orig = await readPresentationFile(load("lo-proje.pptx"), "a.pptx");
    const r = await readPresentationFile(await mutate("lo-proje.pptx", (z) => void z.remove("ppt/slideMasters/slideMaster1.xml")), "a.pptx");
    expect(texts(r.pres)).toEqual(texts(orig.pres));
    for (let i = 0; i < orig.pres.slides.length; i++) expect(layoutName(r.pres, i)).toBe(layoutName(orig.pres, i));
    expect(r.warnings.join(" ")).toMatch(/Asıl slayt okunamadı/);
  });

  it("silinen slayt ilişkileri ve notlar bildirilir", async () => {
    const bytes = await mutate("lo-proje.pptx", (z) => {
      z.remove("ppt/slides/_rels/slide6.xml.rels");
      z.file("ppt/notesSlides/notesSlide5.xml", "bozuk");
    });
    const r = await readPresentationFile(bytes, "a.pptx");
    const w = r.warnings.join("\n");
    expect(w).toMatch(/6\. slaydın ilişki bölümü eksik/);
    expect(w).toMatch(/5\. slaydın konuşmacı notları okunamadı/);
  });

  it("[Content_Types].xml silinse de sunu tanınıp açılır (uyarıyla)", async () => {
    const r = await readPresentationFile(await mutate("lo-proje.pptx", (z) => void z.remove("[Content_Types].xml")), "a.pptx");
    expect(r.format).toBe("pptx");
    expect(r.pres.slides).toHaveLength(7);
    expect(r.warnings.join(" ")).toMatch(/Content_Types/);
  });
});

/** ZIP'te verilen girdinin sıkıştırılmış verisini (yerel başlıktan sonra) bozar. */
function corruptEntry(zip: Uint8Array, name: string): Uint8Array {
  const out = new Uint8Array(zip);
  const dv = new DataView(out.buffer);
  const enc = new TextEncoder().encode(name);
  for (let i = 0; i + 30 < out.length; i++) {
    if (dv.getUint32(i, true) !== 0x04034b50) continue;
    const nameLen = dv.getUint16(i + 26, true);
    const extraLen = dv.getUint16(i + 28, true);
    const compSize = dv.getUint32(i + 18, true);
    const nm = out.subarray(i + 30, i + 30 + nameLen);
    if (nm.length !== enc.length || !nm.every((b, k) => b === enc[k])) continue;
    const start = i + 30 + nameLen + extraLen;
    for (let k = 2; k < Math.min(compSize, 600); k++) out[start + k] = 0xff;
    return out;
  }
  throw new Error(`girdi yok: ${name}`);
}

/** Yerel başlık ve merkezi dizindeki açılmış boyutu sahte (küçük) bir değerle değiştirir. */
function fakeSize(zip: Uint8Array, name: string, size: number): Uint8Array {
  const out = new Uint8Array(zip);
  const dv = new DataView(out.buffer);
  const enc = new TextEncoder().encode(name);
  const match = (at: number, len: number) => len === enc.length && out.subarray(at, at + len).every((b, k) => b === enc[k]);
  for (let i = 0; i + 46 < out.length; i++) {
    const sig = dv.getUint32(i, true);
    if (sig === 0x04034b50 && match(i + 30, dv.getUint16(i + 26, true))) dv.setUint32(i + 22, size, true);
    if (sig === 0x02014b50 && match(i + 46, dv.getUint16(i + 28, true))) dv.setUint32(i + 24, size, true);
  }
  return out;
}

describe("Sıkıştırılmış veri bozulması ve zip bombası", () => {
  it("bozuk sıkıştırılmış slayt verisi ham hata değil, uyarı + boş slayt verir", async () => {
    const z = await JSZip.loadAsync(load("lo-proje.pptx"));
    const bytes = corruptEntry(await z.generateAsync({ type: "uint8array", compression: "DEFLATE" }), "ppt/slides/slide4.xml");
    const r = await readPresentationFile(bytes, "a.pptx");
    expect(r.pres.slides).toHaveLength(7);
    expect(r.warnings.join(" ")).toMatch(/4\. slayt bozuk/);
  });

  it("sınırı aşan (bildirilen boyutu büyük) slayt parçası okunmadan atlanır", async () => {
    ZIP_LIMITS.xml = 64 * 1024;
    const bytes = await mutate("lo-proje.pptx", async (z) => {
      const x = await z.file("ppt/slides/slide2.xml")!.async("string");
      z.file("ppt/slides/slide2.xml", x + " ".repeat(200 * 1024));
    });
    const t = performance.now();
    const r = await readPresentationFile(bytes, "a.pptx");
    expect(performance.now() - t).toBeLessThan(5000);
    expect(r.pres.slides).toHaveLength(7);
    expect(r.warnings).toContain("2. slayt olağandışı büyük olduğu için okunmadı; yerine boş slayt kondu.");
  });

  it("boyutunu küçük bildirip büyük açılan sahte parça akış sırasında durdurulur", async () => {
    ZIP_LIMITS.xml = 64 * 1024;
    const z = await JSZip.loadAsync(load("lo-proje.pptx"));
    const x = await z.file("ppt/slides/slide2.xml")!.async("string");
    z.file("ppt/slides/slide2.xml", x + " ".repeat(300 * 1024));
    const bytes = fakeSize(await z.generateAsync({ type: "uint8array", compression: "DEFLATE" }), "ppt/slides/slide2.xml", 1000);
    const r = await readPresentationFile(bytes, "a.pptx");
    expect(r.warnings.join(" ")).toMatch(/2\. slayt olağandışı büyük/);
  });

  it("ODP içeriği sınırı aşarsa anlaşılır hatayla reddedilir", async () => {
    ZIP_LIMITS.xml = 16 * 1024;
    const e = await readPresentationFile(load("lo-bicim.odp"), "a.odp").catch((x) => x);
    expect(e).toBeInstanceOf(FormatError);
    expect(e.message).toMatch(/content\.xml\) olağandışı büyük/);
  });

  it("paketin toplam açılmış boyutu sınırı aşarsa anlaşılır hata", async () => {
    ZIP_LIMITS.total = 100 * 1024;
    const e = await readPresentationFile(load("lo-proje.pptx"), "a.pptx").catch((x) => x);
    expect(e).toBeInstanceOf(FormatError);
    expect(e.message).toMatch(/açılmış boyutu çok büyük/);
  });
});

describe("ODP: eksik parçalar bildirilir", () => {
  it("silinen resim ve stil bölümü uyarı üretir; içerik açılır", async () => {
    const bytes = await mutate("lo-bicim.odp", (z) => {
      for (const p of Object.keys(z.files)) if (p.startsWith("Pictures/")) z.remove(p);
      z.remove("styles.xml");
    });
    const r = await readPresentationFile(bytes, "a.odp");
    expect(r.pres.slides).toHaveLength(4);
    const w = r.warnings.join("\n");
    expect(w).toMatch(/resim pakette bulunamadı/);
    expect(w).toMatch(/styles\.xml\) dosyada bulunamadı/);
  });

  it("bozuk sıkıştırılmış content.xml Türkçe hatayla reddedilir", async () => {
    const z = await JSZip.loadAsync(load("lo-bicim.odp"));
    const bytes = corruptEntry(await z.generateAsync({ type: "uint8array", compression: "DEFLATE" }), "content.xml");
    const e = await readPresentationFile(bytes, "a.odp").catch((x) => x);
    expect(e).toBeInstanceOf(FormatError);
    expect(e.message).toMatch(/[çğıöşü]/);
    expect(e.message).not.toMatch(/Bug|mismatch|invalid/i);
  });
});

describe("Hata sınıflandırması ve uyarı metinleri", () => {
  it("ham teknik hatalar Türkçe sınıflara çevrilir; ayrıntı korunur", () => {
    const zipErr = toFormatError(new Error("Bug : uncompressed data size mismatch"));
    expect(zipErr.message).toBe("Dosya bozuk; sunu paketi okunamadı.");
    expect(zipErr.detail).toMatch(/size mismatch/);
    expect(toFormatError(new RangeError("Maximum call stack size exceeded")).message).toMatch(/iç içe/);
    expect(toFormatError(new RangeError("Invalid string length")).message).toMatch(/büyük/);
    expect(toFormatError(new TypeError("Cannot read properties of undefined (reading 'x')")).message).toMatch(/^Dosya bozuk/);
    expect(toFormatError(new Error("Sunu bölümü bulunamadı.")).message).toBe("Sunu bölümü bulunamadı.");
  });

  it("aynı türden slayt sorunları tek cümlede toplanır", () => {
    expect(ordinalList([5, 1, 3])).toBe("1., 3. ve 5.");
    const w = new LoadWarnings();
    w.slide("missing", 1);
    w.slide("missing", 3);
    w.slide("notes", 2);
    w.add("Tema okunamadı; varsayılan tema renkleri ve yazı tipleri kullanıldı.");
    expect(w.list()).toEqual(["1. ve 3. slaytlar dosyada bulunamadı; yerlerine boş slayt kondu.", "2. slaydın konuşmacı notları okunamadı.", "Tema okunamadı; varsayılan tema renkleri ve yazı tipleri kullanıldı."]);
  });
});

describe("Geçişler: etkisiz geçiş 'yok' sayılır (gidiş-dönüş gürültüsü)", () => {
  const sld = (t: string) => parseXml(`<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">${t}</p:sld>`).documentElement;
  it("efektsiz, tıklamayla ilerleyen ve sessiz geçiş okunmaz; anlamlı olanlar korunur", () => {
    expect(readTransition(sld('<p:transition spd="slow"/>'))).toBeUndefined();
    expect(readTransition(sld('<p:transition spd="med" advTm="3000"/>'))).toMatchObject({ type: "none", advanceAfter: 3 });
    expect(readTransition(sld('<p:transition advClick="0"/>'))).toMatchObject({ type: "none", advanceOnClick: false });
    expect(readTransition(sld('<p:transition spd="slow"><p:fade/></p:transition>'))).toMatchObject({ type: "fade" });
  });
});
