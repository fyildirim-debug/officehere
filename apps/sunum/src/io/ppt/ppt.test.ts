// @vitest-environment jsdom
/** 97-2003 sunu (.ppt) okuyucu testleri: LibreOffice örnekleri, bozuk ve şifreli dosyalar. */
import { readFileSync } from "node:fs";
import { describe, expect, it } from "vitest";
import { readCfb } from "@officehere/core";
import { flatShapes } from "../../model/ops";
import { plainText } from "../../model/textStyle";
import type { PicShape, Presentation, SpShape } from "../../model/types";
import { readPresentationFile, sniffFormat } from "../formats";
import { dibToBmp } from "./blips";
import { readPpt } from "./reader";
import { writeCfb } from "./testCfb";

const FIX = `${process.cwd()}/src/io/__fixtures__`;
const load = (n: string) => new Uint8Array(readFileSync(`${FIX}/${n}`));

function texts(p: Presentation): string[] {
  return p.slides.map((s) =>
    flatShapes(s.shapes)
      .map((sh) => (sh.kind === "sp" ? plainText(sh.text) : ""))
      .join(" "),
  );
}

/** CFB'yi akışları değiştirerek yeniden yazar. */
function rewrite(bytes: Uint8Array, change: (name: string, data: Uint8Array) => Uint8Array | null): Uint8Array {
  const cfb = readCfb(bytes);
  const streams: { name: string; data: Uint8Array }[] = [];
  for (const e of cfb.entries) {
    if (e.type !== "stream") continue;
    const d = change(e.name, cfb.read(e.path)!);
    if (d) streams.push({ name: e.name, data: d });
  }
  return writeCfb(streams);
}

describe("PPT okuma", () => {
  it("slayt sırası, Türkçe metin, yer tutucular ve slayt boyutu", async () => {
    const p = await readPpt(load("lo-bicim.ppt"));
    expect(p.slides).toHaveLength(4);
    expect(Math.round((p.width * 2.54) / 96)).toBe(28);
    const t = texts(p);
    expect(t[0]).toContain("Çeyrek Sonuçları — Iğdır Şubesi");
    expect(t[0]).toContain("Satış ve Pazarlama Ekibi · Şükrü Öztürk");
    expect(t[1]).toContain("Gündem");
    expect(t[2]).toContain("Görsel ve Şekiller");
    expect(t[3]).toContain("Teşekkürler!");
    const s1 = p.slides[0].shapes.filter((s): s is SpShape => s.kind === "sp");
    expect(s1.some((s) => s.ph?.type === "ctrTitle")).toBe(true);
    expect(s1.some((s) => s.ph?.type === "subTitle")).toBe(true);
    const title = s1.find((s) => s.ph?.type === "ctrTitle")!;
    expect(title.text!.paragraphs[0].runs[0].props).toMatchObject({ bold: true, size: 32, color: { rgb: "#1F4E79" } });
    // Konum: 2 cm, 4 cm.
    expect((title.x * 2.54) / 96).toBeCloseTo(2, 1);
    expect((title.y * 2.54) / 96).toBeCloseTo(4, 1);
  });

  it("madde işaretleri ve düzeyler, notlar", async () => {
    const p = await readPpt(load("lo-bicim.ppt"));
    const body = p.slides[1].shapes.find((s): s is SpShape => s.kind === "sp" && plainText(s.text).includes("Yeni ürünler"))!;
    const paras = body.text!.paragraphs;
    expect(paras.map((x) => [x.runs.map((r) => r.text).join(""), x.level])).toEqual([
      ["Geçen çeyreğin özeti", 0],
      ["Yeni ürünler", 0],
      ["Akıllı saat çığır açıyor", 1],
      ["Kablosuz kulaklık", 1],
      ["Soru ve cevap", 0],
    ]);
    expect(paras[0].props.bullet).toMatchObject({ type: "char", char: "•" });
    expect(paras[0].runs[0].props.size).toBe(32);
    expect(p.slides[0].notes).toBe("Açılış: herkese hoş geldiniz deyin.");
    expect(p.slides[1].notes).toContain("Ayşe Hanım");
  });

  it("resim (Pictures akışı), şekiller, döndürme, biçimli metin ve arka plan", async () => {
    const p = await readPpt(load("lo-bicim.ppt"));
    const s3 = p.slides[2].shapes;
    const pic = s3.find((s): s is PicShape => s.kind === "pic")!;
    expect(pic.src).toMatch(/^data:image\/png;base64,iVBOR/);
    expect(Math.round((pic.w * 2.54) / 96)).toBe(6);
    const box = s3.find((s): s is SpShape => s.kind === "sp" && plainText(s.text) === "Kırmızı Kutu")!;
    expect(box.fill).toEqual({ type: "solid", color: { rgb: "#C00000" } });
    expect(box.line.fill).toEqual({ type: "solid", color: { rgb: "#800000" } });
    const ell = s3.find((s): s is SpShape => s.kind === "sp" && s.geom === "ellipse")!;
    expect(ell.rot).toBeCloseTo(330, 0);
    expect((ell.w * 2.54) / 96).toBeCloseTo(5, 1);
    expect(s3.some((s) => s.kind === "sp" && s.geom === "line")).toBe(true);
    const fmt = s3.find((s): s is SpShape => s.kind === "sp" && plainText(s.text).startsWith("Düz metin"))!;
    const runs = fmt.text!.paragraphs[0].runs;
    expect(runs.find((r) => r.text === "eğik ve altı çizili")!.props).toMatchObject({ italic: true, underline: true });
    expect(runs.find((r) => r.text === "serif turuncu")!.props).toMatchObject({ size: 24, color: { rgb: "#C55A11" } });
    expect(p.slides[3].background).toEqual({ type: "solid", color: { rgb: "#FFF2CC" } });
  });

  it("daha büyük sunu: 7 slayt, notlar ve asıl slayt görünümü", async () => {
    const p = await readPpt(load("lo-proje.ppt"));
    expect(p.slides).toHaveLength(7);
    const t = texts(p);
    expect(t[0]).toContain("Müşteri Portalı Yenileme Projesi");
    expect(t[1]).toContain("Projenin amacı ve kapsamı");
    expect(t[3]).toContain("Zaman Çizelgesi");
    expect(t[3]).toContain("Kalite Güvence");
    expect(t[6]).toContain("Teşekkürler");
    expect(p.slides[4].notes).toContain("Ekip üyelerini");
    // Başlık yer tutucuları tanınır.
    expect(p.slides[1].shapes.some((s) => s.kind === "sp" && s.ph?.type === "title")).toBe(true);
    // Asıl slaytların arka plan/süsleri düzenlere taşınır.
    expect(p.master.shapes.some((s) => !(s.kind === "sp" && s.ph))).toBe(true);
  });

  it("biçim tanıma: içerik imzası uzantıdan önce gelir", async () => {
    const data = load("lo-bicim.ppt");
    expect(await sniffFormat(data, "yanlis-uzanti.pptx")).toBe("ppt");
    const r = await readPresentationFile(data, "sunu.ppt");
    expect(r.format).toBe("ppt");
    expect(r.pres.slides).toHaveLength(4);
  });
});

describe("PPT hata durumları", () => {
  it("bileşik dosya olmayan ve bozuk dosyalar anlaşılır hatayla reddedilir", async () => {
    await expect(readPpt(new Uint8Array(1024))).rejects.toThrow(/97-2003/);
    const bytes = load("lo-bicim.ppt");
    await expect(readPpt(bytes.slice(0, 1536))).rejects.toThrow();
    // Sunu akışı kesilmiş: düzenleme kaydı bulunamaz.
    const cut = rewrite(bytes, (n, d) => (n === "PowerPoint Document" ? d.slice(0, 200) : d));
    await expect(readPpt(cut)).rejects.toThrow(/bozuk|bulunamadı/);
    // Sunu akışı yok, kelime işlemci akışı var.
    const word = writeCfb([{ name: "WordDocument", data: new Uint8Array(5000) }]);
    await expect(readPpt(word)).rejects.toThrow(/kelime işlemci/);
  });

  it("şifreli sunu (CurrentUserAtom imzası) reddedilir", async () => {
    const enc = rewrite(load("lo-bicim.ppt"), (n, d) => {
      if (n !== "Current User") return d;
      const c = d.slice();
      new DataView(c.buffer).setUint32(12, 0xf3d1c4df, true);
      return c;
    });
    await expect(readPpt(enc)).rejects.toThrow(/parola/);
  });

  it("şifreli sunu (CryptSession10Container) reddedilir", async () => {
    const enc = rewrite(load("lo-bicim.ppt"), (n, d) => {
      if (n !== "PowerPoint Document") return d;
      // Akışın sonuna şifreleme oturumu kaydı eklenir.
      const extra = new Uint8Array(16);
      const dv = new DataView(extra.buffer);
      dv.setUint16(0, 0x000f, true);
      dv.setUint16(2, 0x2f14, true);
      dv.setUint32(4, 8, true);
      const out = new Uint8Array(d.length + extra.length);
      out.set(d);
      out.set(extra, d.length);
      return out;
    });
    await expect(readPpt(enc)).rejects.toThrow(/parola/);
  });

  it("DIB resim BMP'ye çevrilir", () => {
    const dib = new Uint8Array(40 + 4);
    const dv = new DataView(dib.buffer);
    dv.setUint32(0, 40, true);
    dv.setInt32(4, 1, true);
    dv.setInt32(8, 1, true);
    dv.setUint16(12, 1, true);
    dv.setUint16(14, 24, true);
    const bmp = dibToBmp(dib);
    expect(String.fromCharCode(bmp[0], bmp[1])).toBe("BM");
    expect(new DataView(bmp.buffer).getUint32(10, true)).toBe(54);
  });
});
