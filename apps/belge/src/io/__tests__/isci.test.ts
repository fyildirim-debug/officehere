// @vitest-environment jsdom
// Belge G/Ç Worker'ı (testte süreç içi taşıyıcı; iletiler structuredClone ile kopyalanır): açma, yazma ve
// otomatik kayıt doğrulaması aynı RPC'den geçer. Kaynak paket yazan iş parçacığında yoksa istenir ve gönderilir.
import { readFileSync } from "node:fs";
import { join } from "node:path";
import JSZip from "jszip";
import { getSource, openSource, releaseSource } from "@officehere/core/opc";
import { defaultSettings, type BelgeDocument } from "../../model/document";
import { encodeInWorker, readInWorker, verifyInWorker, writeDocxInWorker, writeOdtInWorker } from "../client";
import { sourceIds } from "../ioHandlers";
import { writeDocx } from "../index";
import { LoadWarnings, lossMessage } from "../warnings";

const buf = (u: Uint8Array) => u.buffer.slice(u.byteOffset, u.byteOffset + u.byteLength) as ArrayBuffer;
const FIX = join(__dirname, "../__fixtures__");

function simpleDoc(text = "Merhaba dünya"): BelgeDocument {
  return { content: { type: "doc", content: [{ type: "section", content: [{ type: "paragraph", content: [{ type: "text", text }] }] }] }, settings: defaultSettings() };
}

/** Modelin tanımadığı bir parça (tema) içeren .docx. */
async function withTheme(): Promise<Uint8Array> {
  const zip = await JSZip.loadAsync(await writeDocx(simpleDoc("Temalı belge")));
  zip.file("word/theme/theme1.xml", `<?xml version="1.0"?><a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Ofis Teması"><a:themeElements/></a:theme>`);
  const ct = await zip.file("[Content_Types].xml")!.async("string");
  zip.file("[Content_Types].xml", ct.replace("</Types>", `<Override PartName="/word/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/></Types>`));
  const rels = await zip.file("word/_rels/document.xml.rels")!.async("string");
  zip.file("word/_rels/document.xml.rels", rels.replace("</Relationships>", `<Relationship Id="rIdTema" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/></Relationships>`));
  return zip.generateAsync({ type: "uint8array" });
}

describe("Belge G/Ç Worker'ı", () => {
  it("açar: model, biçim ve düz uyarılar; ilerleme bildirir", async () => {
    const steps: string[] = [];
    const r = await readInWorker("rapor.docx", buf(readFileSync(join(FIX, "libreoffice-resaved-report.docx"))), { onProgress: (p) => p.step && steps.push(p.step) });
    if (r.main) throw new Error("Worker'da açılmalıydı");
    expect(r.format).toBe("docx");
    expect(r.doc.settings.kaynak).toMatch(/^[0-9a-f]{64}$/);
    expect(JSON.stringify(r.doc.content)).toContain("text");
    expect(LoadWarnings.fromPlain(r.warnings)).toBeInstanceOf(LoadWarnings);
    expect(steps).toContain("Dosya okunuyor");
  });

  it("HTML ve Markdown ana iş parçacığına bırakılır", async () => {
    const md = new TextEncoder().encode("# Başlık\n\nmetin");
    expect(await readInWorker("not.md", buf(md))).toEqual({ main: true });
    const html = new TextEncoder().encode("<html><body><p>a</p></body></html>");
    expect(await readInWorker("sayfa.html", buf(html))).toEqual({ main: true });
  });

  it("açılamayan dosyada anlaşılır hata (FormatError sınıfıyla)", async () => {
    const bozuk = new Uint8Array([0x50, 0x4b, 0x03, 0x04, 1, 2, 3]);
    await expect(readInWorker("bozuk.docx", buf(bozuk))).rejects.toMatchObject({ name: "FormatError", message: expect.stringContaining("Dosya bozuk") });
  });

  it("dijital imza uyarısı Worker'dan aynen gelir", async () => {
    const r = await readInWorker("imzali.docx", buf(readFileSync(join(__dirname, "../../../../../test-derlem/poi/hello-world-signed.docx"))));
    if (r.main) throw new Error("Worker'da açılmalıydı");
    expect(LoadWarnings.fromPlain(r.warnings).messages()).toContain(lossMessage("signature"));
  });

  it("yazarken eksik kaynak paketi ister; gönderilince tanınmayan parçalar taşınır", async () => {
    const bytes = await withTheme();
    const r = await readInWorker("tema.docx", buf(bytes));
    if (r.main) throw new Error("Worker'da açılmalıydı");
    const id = r.doc.settings.kaynak!;
    // Worker yeniden kurulmuş gibi: kaynak kayıtta yok (süreç içi taşıyıcıda kayıt ortak).
    releaseSource(id);
    const istenen: string[][] = [];
    const out = await writeDocxInWorker(r.doc, false, async (ids) => {
      istenen.push(ids);
      await openSource(bytes, { id });
      return { [id]: getSource(id)!.bytes };
    });
    expect(istenen).toEqual([[id]]);
    const zip = await JSZip.loadAsync(out.bytes);
    expect(zip.file("word/theme/theme1.xml")).not.toBeNull();
    // Kaynak hiçbir yerde yoksa geçiş katmanı olmadan yazılır.
    releaseSource(id);
    const yok = await writeDocxInWorker(r.doc, false, async (ids) => Object.fromEntries(ids.map((x) => [x, null])));
    expect((await JSZip.loadAsync(yok.bytes)).file("word/theme/theme1.xml")).toBeNull();
  });

  it(".odt: kaynak paket kaydedilir; yazarken eksikse istenir, ayarlar ve RDF üst verisi taşınır", async () => {
    const bytes = new Uint8Array(readFileSync(join(FIX, "bicimler/lo-ornek.odt")));
    const r = await readInWorker("lo-ornek.odt", buf(bytes));
    if (r.main) throw new Error("Worker'da açılmalıydı");
    const id = r.doc.settings.kaynak!;
    expect(id).toMatch(/^[0-9a-f]{64}$/);
    releaseSource(id);
    const istenen: string[][] = [];
    const out = await writeOdtInWorker(r.doc, async (ids) => {
      istenen.push(ids);
      await openSource(bytes, { id });
      return { [id]: getSource(id)!.bytes };
    });
    expect(istenen).toEqual([[id]]);
    const zip = await JSZip.loadAsync(out.bytes);
    expect(zip.file("settings.xml")).not.toBeNull();
    expect(zip.file("manifest.rdf")).not.toBeNull();
    expect(zip.file("Thumbnails/thumbnail.png")).toBeNull();
    const manifest = await zip.file("META-INF/manifest.xml")!.async("string");
    expect(manifest).toContain('full-path="settings.xml"');
    expect(manifest).toContain('full-path="manifest.rdf"');
    // Kaynak hiçbir yerde yoksa geçiş katmanı olmadan yazılır.
    releaseSource(id);
    const yok = await writeOdtInWorker(r.doc, async (ids) => Object.fromEntries(ids.map((x) => [x, null])));
    expect((await JSZip.loadAsync(yok.bytes)).file("settings.xml")).toBeNull();
  });

  it("opak öğelerin kaynakları da toplanır", () => {
    const doc = simpleDoc();
    doc.settings.kaynak = "a".repeat(64);
    doc.content.content![0]!.content!.push({ type: "opaqueBlock", attrs: { xml: "<w:p/>", src: "b".repeat(64) } });
    expect([...sourceIds(doc)].sort()).toEqual(["a".repeat(64), "b".repeat(64)]);
  });

  it("Worker yüklenemezse açma ve yazma süreç içinde sürer", async () => {
    const g = globalThis as { Worker?: unknown };
    const prev = g.Worker;
    let made = 0;
    // Yüklenemeyen Worker: ilk iletiden önce "error".
    g.Worker = class {
      private onError: ((e: unknown) => void) | null = null;
      constructor() {
        made++;
        setTimeout(() => this.onError?.({ message: "yüklenemedi" }), 0);
      }
      addEventListener(type: string, fn: (e: unknown) => void) {
        if (type === "error") this.onError = fn;
      }
      postMessage() {}
      terminate() {}
    };
    try {
      vi.resetModules();
      const c = await import("../client");
      const r = await c.readInWorker("rapor.docx", buf(readFileSync(join(FIX, "libreoffice-resaved-report.docx"))));
      expect(r.main).toBeFalsy();
      const out = await c.writeDocxInWorker(simpleDoc(), false, async () => ({}));
      expect(out.bytes.length).toBeGreaterThan(100);
      expect(made).toBe(1);
    } finally {
      g.Worker = prev;
      vi.resetModules();
    }
  });

  it(".odt, .rtf, .md Worker'da yazılır; doğrulama ayrı Worker'da", async () => {
    const odt = (await writeOdtInWorker(simpleDoc(), async () => ({}))).bytes;
    expect(ArrayBuffer.isView(odt) && odt[0] === 0x50 && odt[1] === 0x4b).toBe(true);
    expect(await encodeInWorker("rtf", simpleDoc())).toMatch(/^\{\\rtf1/);
    expect(await encodeInWorker("md", simpleDoc("Merhaba"))).toContain("Merhaba");
    const docx = await writeDocx(simpleDoc());
    expect(await verifyInWorker(buf(docx))).toBe(true);
    await expect(verifyInWorker(buf(new TextEncoder().encode("bozuk")))).rejects.toBeTruthy();
  });
});
