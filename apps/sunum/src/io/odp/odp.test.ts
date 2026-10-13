// @vitest-environment jsdom
/** OpenDocument sunusu (.odp/.otp) okuma, yazma ve gidiş-dönüş testleri. */
import { execFileSync } from "node:child_process";
import { existsSync, mkdtempSync, readFileSync, rmSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import JSZip from "jszip";
import { describe, expect, it } from "vitest";
import { flatShapes } from "../../model/ops";
import { plainText } from "../../model/textStyle";
import type { PicShape, Presentation, SpShape, TableShape } from "../../model/types";
import { TEMPLATES } from "../../templates/templates";
import { readPresentationFile, sniffFormat } from "../formats";
import { readPptx } from "../pptxReader";
import { readOdp } from "./reader";
import { writeOdp } from "./writer";

const FIX = `${process.cwd()}/src/io/__fixtures__`;
const load = (n: string) => new Uint8Array(readFileSync(`${FIX}/${n}`));

function texts(p: Presentation): string[] {
  return p.slides.map((s) =>
    flatShapes(s.shapes)
      .map((sh) => (sh.kind === "sp" ? plainText(sh.text) : sh.kind === "table" ? sh.rows.map((r) => r.cells.map((c) => plainText(c.text)).join(" ")).join(" ") : ""))
      .join(" "),
  );
}

const sps = (p: Presentation, i: number) => p.slides[i].shapes.filter((s): s is SpShape => s.kind === "sp");

describe("ODP okuma (LibreOffice örneği)", () => {
  it("sayfa boyutu, slaytlar, yer tutucular ve Türkçe metin", async () => {
    const p = await readOdp(load("lo-bicim.odp"));
    expect(Math.round((p.width * 2.54) / 96)).toBe(28);
    expect(Math.round(((p.height * 2.54) / 96) * 100) / 100).toBeCloseTo(15.75, 1);
    expect(p.slides).toHaveLength(4);
    expect(p.slides.map((s) => s.name)).toEqual(["Kapak", "Gündem", "Görsel ve Şekil", "Kapanış"]);
    const t = texts(p);
    expect(t[0]).toContain("Çeyrek Sonuçları — Iğdır Şubesi");
    expect(t[0]).toContain("Şükrü Öztürk");
    const title = sps(p, 0).find((s) => s.ph?.type === "ctrTitle")!;
    expect(title.text!.paragraphs[0].runs[0].props).toMatchObject({ bold: true, size: 32, color: { rgb: "#1F4E79" } });
    expect(title.text!.paragraphs[0].props.align).toBe("ctr");
    expect(sps(p, 0).some((s) => s.ph?.type === "subTitle")).toBe(true);
    expect(p.layouts.find((l) => l.id === p.slides[0].layoutId)?.type).toBe("title");
  });

  it("madde işaretli liste düzeyleri ve notlar", async () => {
    const p = await readOdp(load("lo-bicim.odp"));
    const body = sps(p, 1).find((s) => s.ph?.type === "obj")!;
    const paras = body.text!.paragraphs;
    expect(paras.map((x) => [x.runs.map((r) => r.text).join(""), x.level])).toEqual([
      ["Geçen çeyreğin özeti", 0],
      ["Yeni ürünler", 0],
      ["Akıllı saat çığır açıyor", 1],
      ["Kablosuz kulaklık", 1],
      ["Soru ve cevap", 0],
    ]);
    expect(paras[0].props.bullet?.type).toBe("char");
    expect(p.slides[0].notes).toBe("Açılış: herkese hoş geldiniz deyin.");
    expect(p.slides[1].notes).toContain("Ayşe Hanım");
  });

  it("resim, şekiller (dolgu, çizgi, döndürme), biçimli metin ve arka plan", async () => {
    const p = await readOdp(load("lo-bicim.odp"));
    const s3 = p.slides[2].shapes;
    const pic = s3.find((s): s is PicShape => s.kind === "pic")!;
    expect(pic.src).toMatch(/^data:image\/png;base64,/);
    expect(pic.descr).toBe("Mavi kare");
    const box = s3.find((s): s is SpShape => s.kind === "sp" && s.name === "Kırmızı Kutu")!;
    expect(box.fill).toEqual({ type: "solid", color: { rgb: "#C00000" } });
    expect(box.line.fill).toEqual({ type: "solid", color: { rgb: "#800000" } });
    expect(plainText(box.text)).toBe("Kırmızı Kutu");
    const ell = s3.find((s): s is SpShape => s.kind === "sp" && s.geom === "ellipse")!;
    expect(ell.rot).toBeCloseTo(330, 0);
    expect(Math.round(ell.w)).toBe(189);
    const line = s3.find((s): s is SpShape => s.kind === "sp" && s.geom === "line")!;
    expect(line.line.width).toBeCloseTo(2.27, 1);
    const fmt = s3.find((s): s is SpShape => s.kind === "sp" && s.name === "Biçimli Metin")!;
    const runs = fmt.text!.paragraphs[0].runs;
    expect(runs.find((r) => r.text === "eğik ve altı çizili")!.props).toMatchObject({ italic: true, underline: true });
    expect(runs.find((r) => r.text === "serif turuncu")!.props).toMatchObject({ font: "Liberation Serif", size: 24, color: { rgb: "#C55A11" } });
    expect(p.slides[3].background).toEqual({ type: "solid", color: { rgb: "#FFF2CC" } });
  });

  it("çok ana sayfalı sunu: tablo, notlar ve ikincil ana sayfa görünümü", async () => {
    const p = await readOdp(load("lo-proje.odp"));
    expect(p.slides).toHaveLength(7);
    const t = texts(p);
    for (const w of ["Müşteri Portalı Yenileme Projesi", "Gündem", "Hedefler", "Zaman Çizelgesi", "Teşekkürler", "Kalite Güvence"]) expect(t.join(" ")).toContain(w);
    const table = p.slides[3].shapes.find((s): s is TableShape => s.kind === "table")!;
    expect(table.rows.length).toBeGreaterThanOrEqual(5);
    expect(table.cols).toHaveLength(4);
    expect(plainText(table.rows[0].cells[0].text)).toBe("Aşama");
    expect(p.slides[4].notes).toContain("Ekip üyelerini");
    // İkincil ana sayfalar düzen kopyası olarak gelir.
    expect(new Set(p.slides.map((s) => s.layoutId)).size).toBeGreaterThan(1);
    expect(p.layouts.some((l) => l.background && !l.showMasterShapes)).toBe(true);
  });

  it("slayt geçişleri okunur (soluklaştır, it)", async () => {
    const p = await readOdp(load("lo-proje.odp"));
    expect(p.slides[0].transition).toMatchObject({ type: "fade", duration: 1 });
    expect(p.slides[1].transition).toMatchObject({ type: "push", option: "u", duration: 0.75 });
  });

  it("şablon (.otp) tanınır ve okunur", async () => {
    const data = load("lo-bicim.otp");
    expect(await sniffFormat(data, "x.otp")).toBe("otp");
    const r = await readPresentationFile(data, "sablon.otp");
    expect(r.format).toBe("otp");
    expect(r.pres.slides).toHaveLength(4);
  });
});

describe("ODP yazma", () => {
  it("paket: mimetype ilk ve sıkıştırılmamış, bildirim ve parçalar", async () => {
    const bytes = await writeOdp(await readPptx(load("lo-proje.pptx")));
    // Yerel dosya başlığı: ilk girdi "mimetype", yöntem 0 (STORE).
    expect(Buffer.from(bytes.subarray(30, 38)).toString()).toBe("mimetype");
    expect(bytes[8] | (bytes[9] << 8)).toBe(0);
    expect(Buffer.from(bytes.subarray(38, 38 + 47)).toString()).toBe("application/vnd.oasis.opendocument.presentation");
    const zip = await JSZip.loadAsync(bytes);
    for (const f of ["META-INF/manifest.xml", "content.xml", "styles.xml", "meta.xml"]) expect(zip.file(f)).toBeTruthy();
    const manifest = await zip.file("META-INF/manifest.xml")!.async("string");
    expect(manifest).toContain('manifest:full-path="/"');
    for (const f of ["content.xml", "styles.xml", "meta.xml"]) {
      const doc = new DOMParser().parseFromString(await zip.file(f)!.async("string"), "application/xml");
      expect(doc.getElementsByTagName("parsererror")).toHaveLength(0);
    }
  });

  it("gidiş-dönüş (PPTX → ODP → model): slaytlar, metin, tablo, resim ve notlar", async () => {
    const src = await readPptx(load("lo-proje.pptx"));
    const back = await readOdp(await writeOdp(src));
    expect(back.slides).toHaveLength(src.slides.length);
    const a = texts(src);
    const b = texts(back);
    a.forEach((t, i) => {
      for (const w of t.split(/\s+/).filter((x) => x.length > 2)) expect(b[i]).toContain(w);
    });
    expect(back.slides[4].notes).toBe(src.slides[4].notes);
    expect(back.slides[3].shapes.some((s) => s.kind === "table")).toBe(true);
    expect(Math.round(back.width)).toBe(Math.round(src.width));
  });

  it("gidiş-dönüş (ODP → ODP): biçimler ve konumlar korunur", async () => {
    const src = await readOdp(load("lo-bicim.odp"));
    const back = await readOdp(await writeOdp(src));
    expect(back.slides.map((s) => s.name)).toEqual(src.slides.map((s) => s.name));
    expect(texts(back)).toEqual(texts(src));
    const e1 = src.slides[2].shapes.find((s) => s.kind === "sp" && s.geom === "ellipse")!;
    const e2 = back.slides[2].shapes.find((s) => s.kind === "sp" && s.geom === "ellipse")!;
    for (const k of ["x", "y", "w", "h", "rot"] as const) expect(e2[k]).toBeCloseTo(e1[k], 0);
    const box = back.slides[2].shapes.find((s): s is SpShape => s.kind === "sp" && s.name === "Kırmızı Kutu")!;
    expect(box.fill).toEqual({ type: "solid", color: { rgb: "#C00000" } });
    expect(box.text!.paragraphs[0].runs[0].props).toMatchObject({ bold: true, color: { rgb: "#FFFFFF" } });
    const body = back.slides[1].shapes.find((s): s is SpShape => s.kind === "sp" && s.ph?.type === "obj")!;
    expect(body.text!.paragraphs.map((p) => p.level)).toEqual([0, 0, 1, 1, 0]);
    expect(back.slides[0].notes).toBe(src.slides[0].notes);
    expect(back.slides[3].background).toEqual(src.slides[3].background);
    expect(back.slides[2].shapes.some((s) => s.kind === "pic")).toBe(true);
  });

  it("gizli şekil ve slayt, belge özellikleri korunur", async () => {
    const src = await readOdp(load("lo-bicim.odp"));
    src.slides[2] = { ...src.slides[2], hidden: true, shapes: src.slides[2].shapes.map((s) => (s.name === "Kırmızı Kutu" ? { ...s, hidden: true } : s)) };
    src.props = { ...src.props, title: "Biçim Örneği", creator: "Ayşe Yılmaz" };
    const back = await readOdp(await writeOdp(src));
    expect(back.slides[2].hidden).toBe(true);
    expect(back.slides[2].shapes.find((s) => s.name === "Kırmızı Kutu")?.hidden).toBe(true);
    expect(back.props).toMatchObject({ title: "Biçim Örneği", creator: "Ayşe Yılmaz" });
    expect(Number.isFinite(Date.parse(back.props.created!))).toBe(true);
  });

  it("geçişler ve otomatik ilerleme gidiş-dönüşte korunur", async () => {
    const src = TEMPLATES.find((t) => t.id === "blank")!.build();
    const s0 = src.slides[0];
    src.slides = [
      { ...s0, transition: { type: "push", option: "l", duration: 0.5, advanceOnClick: true } },
      { ...s0, id: "ikinci", transition: { type: "split", option: "horzIn", duration: 1.25, advanceOnClick: true, advanceAfter: 3 } },
      { ...s0, id: "ucuncu", transition: { type: "shape", option: "circle", duration: 0.75, advanceOnClick: true } },
    ];
    const back = await readOdp(await writeOdp(src));
    expect(back.slides.map((s) => s.transition)).toEqual([
      { type: "push", option: "l", duration: 0.5, advanceOnClick: true },
      { type: "split", option: "horzIn", duration: 1.25, advanceOnClick: true, advanceAfter: 3 },
      { type: "shape", option: "circle", duration: 0.75, advanceOnClick: true },
    ]);
  });

  it("hazır şablonlar ODP olarak yazılıp okunabilir (Türkçe metin)", async () => {
    for (const t of TEMPLATES) {
      const p = t.build();
      const back = await readOdp(await writeOdp(p));
      expect(back.slides.length, t.id).toBe(p.slides.length);
      expect(texts(back).join(" "), t.id).toBe(texts(p).join(" "));
    }
  });
});

describe("ODP hata durumları", () => {
  it("bozuk paket, içeriksiz paket, başka ODF türü ve şifreli dosya anlaşılır hatayla reddedilir", async () => {
    await expect(readOdp(new Uint8Array([1, 2, 3, 4]))).rejects.toThrow(/ZIP/);
    const empty = new JSZip();
    empty.file("mimetype", "application/vnd.oasis.opendocument.presentation");
    await expect(readOdp(await empty.generateAsync({ type: "uint8array" }))).rejects.toThrow(/content\.xml/);
    const odt = new JSZip();
    odt.file("mimetype", "application/vnd.oasis.opendocument.text");
    odt.file("content.xml", "<x/>");
    const odtBytes = await odt.generateAsync({ type: "uint8array" });
    await expect(readOdp(odtBytes)).rejects.toThrow(/sunusu değil/);
    await expect(sniffFormat(odtBytes, "a.odp")).rejects.toThrow(/metin belgesi/);
    const enc = new JSZip();
    enc.file("mimetype", "application/vnd.oasis.opendocument.presentation");
    enc.file("META-INF/manifest.xml", '<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0"><manifest:file-entry manifest:full-path="content.xml"><manifest:encryption-data/></manifest:file-entry></manifest:manifest>');
    enc.file("content.xml", "gizli");
    await expect(readOdp(await enc.generateAsync({ type: "uint8array" }))).rejects.toThrow(/parola/);
    const bad = new JSZip();
    bad.file("mimetype", "application/vnd.oasis.opendocument.presentation");
    bad.file("content.xml", "<office:document-content");
    await expect(readOdp(await bad.generateAsync({ type: "uint8array" }))).rejects.toThrow(/bozuk/);
  });
});

const hasSoffice = (() => {
  try {
    execFileSync("soffice", ["--version"], { stdio: "ignore", timeout: 20000 });
    return true;
  } catch {
    return false;
  }
})();

describe("LibreOffice doğrulaması", () => {
  it.skipIf(!hasSoffice)(
    "yazılan ODP LibreOffice ile PPTX'e çevrilir ve metin korunur",
    async () => {
      const dir = mkdtempSync(join(tmpdir(), "oh-odp-"));
      try {
        const src = await readOdp(load("lo-bicim.odp"));
        writeFileSync(join(dir, "yazilan.odp"), await writeOdp(src));
        execFileSync("soffice", [`-env:UserInstallation=file://${dir}/profil`, "--headless", "--convert-to", "pptx", join(dir, "yazilan.odp"), "--outdir", dir], { stdio: "ignore", timeout: 120000 });
        expect(existsSync(join(dir, "yazilan.pptx"))).toBe(true);
        const zip = await JSZip.loadAsync(readFileSync(join(dir, "yazilan.pptx")));
        const slides = Object.keys(zip.files).filter((f) => /^ppt\/slides\/slide\d+\.xml$/.test(f));
        expect(slides).toHaveLength(4);
        let all = "";
        for (const f of slides) all += (await zip.file(f)!.async("string")).replace(/<[^>]+>/g, " ");
        for (const w of ["Çeyrek Sonuçları — Iğdır Şubesi", "Geçen çeyreğin özeti", "Akıllı saat çığır açıyor", "Kırmızı Kutu", "Teşekkürler!"]) expect(all).toContain(w);
      } finally {
        rmSync(dir, { recursive: true, force: true });
      }
    },
    180000,
  );
});
