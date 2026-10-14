// @vitest-environment jsdom
// Dosya G/Ç sağlamlığı: bozuk dosyada takılma/yığın taşması olmaz, belge varsayılanları ve açıklamalar
// korunur, okunamayan içerik açılışta uyarılır, ham kitaplık hataları anlaşılır Türkçe iletiye çevrilir.
import { describe, expect, it } from "vitest";
import JSZip from "jszip";
import { execFileSync } from "node:child_process";
import { mkdtempSync, readFileSync, rmSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import type { JSONContent } from "@tiptap/core";
import { BelgeOpenError, FODT_NOTE, classifyOpenError, docToText, openDocument, readDoc, readDocx, readOdt, readRtf, validateContent, writeDocx, writeOdt, writeRtf } from "../index";
import { limitDepth, parseXml, zipText, flattenedDocs } from "../docx/xml";
import { LoadWarnings, lossMessage } from "../warnings";
import { all, fixture, fixtureBuffer, text } from "./bicim-yardimci";

const buf = (u: Uint8Array) => u.buffer.slice(u.byteOffset, u.byteOffset + u.byteLength) as ArrayBuffer;
const enc = (s: string) => buf(new TextEncoder().encode(s));

/** Metin düğümlerinden belirli işaret türünü taşıyanların birleşik metni. */
function marked(doc: JSONContent, type: string): string {
  return all(doc, "text")
    .filter((t) => t.marks?.some((m) => m.type === type))
    .map((t) => t.text)
    .join("|");
}

/** Açıklama kimliği → işaretli metin. */
function commentTexts(doc: JSONContent): Record<string, string> {
  const out: Record<string, string> = {};
  for (const t of all(doc, "text")) for (const m of t.marks ?? []) if (m.type === "comment") out[String(m.attrs?.id)] = (out[String(m.attrs?.id)] ?? "") + t.text;
  return out;
}

/** DOCX paketinde bir parçayı değiştirir. */
async function patchDocx(name: string, part: string, fn: (xml: string) => string): Promise<ArrayBuffer> {
  const zip = await JSZip.loadAsync(fixture(name));
  zip.file(part, fn(await zip.file(part)!.async("string")));
  return buf(await zip.generateAsync({ type: "uint8array" }));
}

/** styles.xml docDefaults → anlamsal değerler (belirtilmeyenler biçim belirtiminin varsayılanları). */
function docDefaults(stylesXml: string) {
  const d = /<w:docDefaults>([\s\S]*?)<\/w:docDefaults>/.exec(stylesXml)?.[1] ?? "";
  const a = (el: string, at: string) => new RegExp(`<w:${el}\\b[^>]*\\bw:${at}="([^"]*)"`).exec(d)?.[1] ?? null;
  return { font: a("rFonts", "ascii"), sz: Number(a("sz", "val") ?? 20), lang: a("lang", "val"), before: Number(a("spacing", "before") ?? 0), after: Number(a("spacing", "after") ?? 0), line: Number(a("spacing", "line") ?? 240) };
}

/** Kaynağın etkin Normal'i: docDefaults üzerine Normal stilinin değerleri. */
function effectiveNormal(stylesXml: string) {
  const dd = docDefaults(stylesXml);
  const normal = /<w:style [^>]*w:styleId="Normal"[^>]*>([\s\S]*?)<\/w:style>/.exec(stylesXml)?.[1] ?? "";
  const a = (el: string, at: string) => new RegExp(`<w:${el}\\b[^>]*\\bw:${at}="([^"]*)"`).exec(normal)?.[1] ?? null;
  return { font: a("rFonts", "ascii") ?? dd.font, sz: Number(a("sz", "val") ?? dd.sz), before: Number(a("spacing", "before") ?? dd.before), after: Number(a("spacing", "after") ?? dd.after), line: Number(a("spacing", "line") ?? dd.line) };
}

async function stylesOfDocx(data: ArrayBuffer | Uint8Array): Promise<string> {
  return (await JSZip.loadAsync(data)).file("word/styles.xml")!.async("string");
}

describe("takılma ve yığın taşması", () => {
  /** WordDocument metninde UTF-16 olarak geçen sözcüğün yerine alan karakteri yazar. */
  function docWithFieldChar(code: number): Uint8Array {
    const b = new Uint8Array(fixture("lo-ornek.doc"));
    const needle = new Uint8Array([...new TextEncoder().encode("paragraf")].flatMap((c) => [c, 0]));
    let at = -1;
    for (let i = 0; i + needle.length <= b.length && at < 0; i++) if (needle.every((v, k) => b[i + k] === v)) at = i;
    expect(at).toBeGreaterThan(0);
    b[at] = code;
    b[at + 1] = 0;
    return b;
  }

  it(".doc: eşi olmayan alan ayırıcısı / alan sonu sonsuz döngüye sokmaz", () => {
    for (const code of [0x14, 0x15]) {
      const t0 = performance.now();
      const w = new LoadWarnings();
      const d = readDoc(docWithFieldChar(code), w);
      expect(performance.now() - t0).toBeLessThan(3000);
      expect(text(d.content)).toContain("Giriş ve Amaç");
      expect(text(d.content)).toContain("aragraf");
      expect(w.has("corrupt")).toBe(true);
    }
  });

  it("DOCX: 20.000 düzey iç içe (bildirilmemiş ön ekli) öğe yığın taşırmaz; içerik açılır, uyarı verilir", async () => {
    const data = await patchDocx("lo-ornek.docx", "word/document.xml", (x) => x.replace("</w:body>", "<x:d>".repeat(20000) + "</x:d>".repeat(20000) + "</w:body>"));
    const t0 = performance.now();
    const r = await openDocument("derin.docx", data);
    expect(performance.now() - t0).toBeLessThan(5000);
    expect(text(r.doc.content)).toContain("Giriş ve Amaç");
    expect(r.warnings).toContain(lossMessage("deep"));
    validateContent(r.doc.content);
  });

  it("DOCX: geçerli ad alanlı derin iç içe yapı (customXml) düzleştirilir", async () => {
    const data = await patchDocx("lo-ornek.docx", "word/document.xml", (x) => x.replace("</w:body>", "<w:customXml>".repeat(5000) + "<w:p><w:r><w:t>derin</w:t></w:r></w:p>" + "</w:customXml>".repeat(5000) + "</w:body>"));
    const r = await openDocument("derin2.docx", data);
    expect(text(r.doc.content)).toContain("Giriş ve Amaç");
    expect(r.warnings).toContain(lossMessage("deep"));
  });

  it("DOCX: dev hücre birleştirmesi (gridSpan) sütun sınırına indirgenir, yazma hızlı kalır", async () => {
    const data = await patchDocx("lo-ornek.docx", "word/document.xml", (x) => x.replace(/<w:tcPr>/g, '<w:tcPr><w:gridSpan w:val="100000"/>'));
    const d = await readDocx(data);
    for (const c of all(d.content, "tableCell")) expect(Number(c.attrs?.colspan)).toBeLessThanOrEqual(63);
    const t0 = performance.now();
    await writeDocx(d);
    expect(performance.now() - t0).toBeLessThan(2000);
  });

  it("limitDepth: sınırın ötesindeki etiketleri atar, iyi biçimli XML üretir", () => {
    const deep = `<?xml version="1.0"?><a>${"<b>".repeat(300)}metin${"</b>".repeat(300)}<c/></a>`;
    const r = limitDepth(deep, 50);
    expect(r.flattened).toBe(true);
    const doc = parseXml(r.text);
    expect(doc.documentElement.textContent).toBe("metin");
    expect(doc.getElementsByTagName("c").length).toBe(1);
    expect(limitDepth("<a><b/></a>", 50)).toEqual({ text: "<a><b/></a>", flattened: false });
    expect(flattenedDocs.has(parseXml(deep))).toBe(true);
  });

  it("bildirilmemiş ad alanı öneki olan XML okunur", () => {
    const doc = parseXml('<w:document xmlns:w="urn:w"><w:body><x:y q:z="1">a</x:y><w:p>b</w:p></w:body></w:document>');
    expect(doc.getElementsByTagName("w:p")[0].textContent).toBe("b");
  });

  it("zip bombası: olağandışı büyük XML parçası açılmadan reddedilir; sondaki boşluklar atılır", async () => {
    const fake = { _data: { uncompressedSize: 600 * 1024 * 1024 }, async: async () => new Uint8Array(0) };
    await expect(zipText(fake)).rejects.toBeInstanceOf(BelgeOpenError);
    // 100 MB'lık, 1000 kat sıkıştırılmış parça (zip bombası imzası) açılmadan reddedilir.
    const bomb = { _data: { uncompressedSize: 100 * 1024 * 1024, compressedSize: 100 * 1024 }, async: async () => new Uint8Array(0) };
    await expect(zipText(bomb)).rejects.toThrow(/zip bombası/);
    const ok = { async: async () => new TextEncoder().encode("<a>x</a>" + " ".repeat(100000) + "\n") };
    expect(await zipText(ok)).toBe("<a>x</a>");
  });

  it("RTF: dosyayı aşan \\bin bozukluk sayılır, içerik okunmaya devam eder; kapanmamış gruplar uyarılır", () => {
    const w = new LoadWarnings();
    const d = readRtf(new TextEncoder().encode("{\\rtf1\\ansi {\\pard Birinci\\par}\\bin2147483647 \\pard Ikinci paragraf\\par}"), w);
    expect(text(d.content)).toContain("Ikinci paragraf");
    expect(w.has("corrupt")).toBe(true);
    const w2 = new LoadWarnings();
    const d2 = readRtf(new TextEncoder().encode("{\\rtf1\\ansi {\\fonttbl{\\f0 Arial;}{\\header Ust\\par {\\pard Metin\\par"), w2);
    expect(w2.has("unclosed")).toBe(true);
    // Kapanmamış tablo / üst bilgi grubundaki gövde metni kurtarılır.
    expect(text(d2.content)).toContain("Metin");
  });

  it("RTF: aşırı derin gruplar hata yerine düzleştirilir", () => {
    const w = new LoadWarnings();
    const d = readRtf(new TextEncoder().encode("{\\rtf1 " + "{".repeat(50000) + "derin metin" + "}".repeat(50000) + "\\par sonra}"), w);
    expect(text(d.content)).toContain("derin metin");
    expect(text(d.content)).toContain("sonra");
    expect(w.has("deep")).toBe(true);
  });
});

describe("belge varsayılanları (sayfa kayması)", () => {
  it("LibreOffice docx: docDefaults ve etkin Normal gidiş-dönüşte korunur", async () => {
    for (const name of ["lo-ornek.docx", "izleme.docx"]) {
      const src = await stylesOfDocx(fixture(name));
      const doc = await readDocx(fixture(name));
      const out = await writeDocx(doc);
      const dst = await stylesOfDocx(out);
      const a = effectiveNormal(src);
      const b = docDefaults(dst);
      expect({ font: b.font, sz: b.sz, before: b.before, after: b.after, line: b.line }).toEqual(a);
      expect(b.lang).toBe(docDefaults(src).lang);
      // İkinci tur değişmez.
      expect(docDefaults(await stylesOfDocx(await writeDocx(await readDocx(out))))).toEqual(b);
    }
  });

  it("yeni belge bizim varsayılanlarımızla yazılır", async () => {
    const out = await writeDocx({ content: { type: "doc", content: [{ type: "paragraph", content: [{ type: "text", text: "x" }] }] }, settings: (await import("../../model/document")).defaultSettings() });
    expect(docDefaults(await stylesOfDocx(out))).toMatchObject({ font: "Aptos", sz: 22, after: 160, line: 259, lang: "tr-TR" });
  });

  it("ODT ve RTF: kaynağın varsayılan paragraf stili Normal olur; paragraflara gereksiz aralık eklenmez", async () => {
    for (const d of [await readOdt(fixture("izleme.odt")), readRtf(fixture("izleme.rtf"))]) {
      expect(d.settings.customStyles?.Normal?.props).toMatchObject({ font: "Liberation Serif", size: 12, spaceAfter: 0, line: 1 });
      for (const p of all(d.content, "paragraph")) expect(p.attrs?.spaceAfter ?? null).toBeNull();
    }
  });

  it("w:contextualSpacing: aynı stildeki ardışık paragrafların arası sıfırlanır", async () => {
    const data = await patchDocx("lo-ornek.docx", "word/styles.xml", (x) => x.replace("</w:styles>", '<w:style w:type="paragraph" w:styleId="ListeOH"><w:name w:val="Liste OH"/><w:pPr><w:spacing w:before="120" w:after="200"/><w:contextualSpacing/></w:pPr></w:style></w:styles>'));
    const zip = await JSZip.loadAsync(data);
    const p = (t: string) => `<w:p><w:pPr><w:pStyle w:val="ListeOH"/></w:pPr><w:r><w:t>${t}</w:t></w:r></w:p>`;
    const xml = await zip.file("word/document.xml")!.async("string");
    zip.file("word/document.xml", xml.replace("<w:body>", `<w:body>${p("A")}${p("B")}${p("C")}`));
    const d = await readDocx(await zip.generateAsync({ type: "uint8array" }));
    const [a, b, c] = all(d.content, "paragraph").slice(0, 3);
    expect([a.attrs?.spaceBefore, a.attrs?.spaceAfter]).toEqual([6, 0]);
    expect([b.attrs?.spaceBefore, b.attrs?.spaceAfter]).toEqual([0, 0]);
    expect([c.attrs?.spaceBefore, c.attrs?.spaceAfter]).toEqual([0, 10]);
  });
});

describe("açıklamalar ve değişiklik izleme", () => {
  it("DOCX: aralıksız (yalnız başvurulu) açıklama konumundaki sözcüğe bağlanır ve kaydedince korunur", async () => {
    const d = await readDocx(fixture("izleme.docx"));
    expect(Object.keys(d.settings.comments)).toHaveLength(2);
    const ct = commentTexts(d.content);
    expect(Object.values(ct).sort()).toEqual(["açıklamalı ifade", "İkinci"]);
    const back = await readDocx(await writeDocx(d));
    expect(Object.keys(back.settings.comments)).toHaveLength(2);
    expect(Object.values(commentTexts(back.content)).sort()).toEqual(["açıklamalı ifade", "İkinci"]);
    expect(Object.values(back.settings.comments).map((c) => c.text).sort()).toEqual(["Aralıklı açıklama metni.", "Nokta açıklaması."]);
  });

  it("DOCX: boş aralıklı açıklama da kaybolmaz", async () => {
    const data = await patchDocx("izleme.docx", "word/document.xml", (x) => x.replace('<w:commentRangeStart w:id="0"/>', "").replace('<w:commentRangeEnd w:id="0"/>', '<w:commentRangeStart w:id="0"/><w:commentRangeEnd w:id="0"/>'));
    const d = await readDocx(data);
    expect(Object.keys(commentTexts(d.content))).toHaveLength(2);
  });

  for (const ext of ["docx", "odt", "rtf", "doc"] as const) {
    it(`${ext}: silinen ve eklenen metin değişiklik işaretiyle okunur; açıklamalar bağlanır`, async () => {
      const r = await openDocument(`izleme.${ext}`, fixtureBuffer(`izleme.${ext}`));
      const c = r.doc.content;
      expect(marked(c, "deletion")).toBe("silinen sözcük");
      expect(marked(c, "insertion")).toBe("eklenen ifade");
      const del = all(c, "text").find((t) => t.text === "silinen sözcük")!.marks!.find((m) => m.type === "deletion")!;
      expect(del.attrs?.author).toBe("Ali Veli");
      expect(String(del.attrs?.date)).toMatch(/^2026-01-02T10:00/);
      expect(Object.keys(r.doc.settings.comments)).toHaveLength(2);
      expect(Object.values(commentTexts(c)).sort()).toContain("İkinci");
      // LibreOffice RTF'e yazarken yazar adındaki Türkçe harfleri "?" yapar; okunan ad kaynaktakidir.
      expect(Object.values(r.doc.settings.comments).map((x) => x.author).sort()).toEqual([ext === "rtf" ? "Ay?e Y?lmaz" : "Ayşe Yılmaz", "Mehmet Kaya"]);
      expect(r.warnings).toEqual([]);
      // Kaydedip açınca (DOCX) izleme ve açıklamalar korunur.
      const back = await readDocx(await writeDocx(r.doc));
      expect(marked(back.content, "deletion")).toBe("silinen sözcük");
      expect(marked(back.content, "insertion")).toBe("eklenen ifade");
      expect(Object.keys(back.settings.comments)).toHaveLength(2);
    });
  }

  it("ODT → ODT: değişiklik izleme yazılır ve geri okunur", async () => {
    const d = await readOdt(fixture("izleme.odt"));
    const back = await readOdt(await writeOdt(d));
    expect(marked(back.content, "deletion")).toBe("silinen sözcük");
    expect(marked(back.content, "insertion")).toBe("eklenen ifade");
    expect(Object.keys(back.settings.comments)).toHaveLength(2);
  });

  it("RTF → RTF: değişiklik izleme, açıklamalar, iki yana yaslı son paragraf ve paragraf sayısı korunur", () => {
    const d = readRtf(fixture("izleme.rtf"));
    expect(all(d.content, "paragraph").at(-1)?.attrs?.textAlign).toBe("justify");
    const back = readRtf(new TextEncoder().encode(writeRtf(d)));
    expect(marked(back.content, "deletion")).toBe("silinen sözcük");
    expect(marked(back.content, "insertion")).toBe("eklenen ifade");
    expect(Object.values(commentTexts(back.content)).sort()).toEqual(Object.values(commentTexts(d.content)).sort());
    expect(Object.values(back.settings.comments).map((x) => x.text).sort()).toEqual(["Aralıklı açıklama metni.", "Nokta açıklaması."]);
    expect(all(back.content, "paragraph").filter((p) => p.attrs?.textAlign === "justify")).toHaveLength(1);
    expect(all(back.content, "paragraph").length).toBe(all(d.content, "paragraph").length);
    // Sonu boş paragrafla biten belge de paragraf kaybetmez.
    const withEmpty = { ...d, content: { ...d.content, content: [{ type: "section", content: [...(d.content.content?.[0]?.content ?? []), { type: "paragraph" }] }] } };
    const back2 = readRtf(new TextEncoder().encode(writeRtf(withEmpty)));
    expect(all(back2.content, "paragraph").length).toBe(all(withEmpty.content, "paragraph").length);
  });

  it("ODT biçim değişikliği kaydı okunamaz: uyarı verilir", async () => {
    const zip = await JSZip.loadAsync(fixture("izleme.odt"));
    const xml = await zip.file("content.xml")!.async("string");
    zip.file("content.xml", xml.replace("</text:tracked-changes>", '<text:changed-region text:id="fc1"><text:format-change><office:change-info><dc:creator>X</dc:creator></office:change-info></text:format-change></text:changed-region></text:tracked-changes>'));
    const r = await openDocument("bicim.odt", buf(await zip.generateAsync({ type: "uint8array" })));
    expect(r.warnings).toContain(lossMessage("trackFormat"));
    expect(r.warning).toContain("kaydederseniz kaybolur");
  });
});

describe("açılış uyarıları", () => {
  it("DOCX: gömülü nesne, makro ve okunamayan şekil uyarılır", async () => {
    const zip = await JSZip.loadAsync(fixture("lo-ornek.docx"));
    const xml = await zip.file("word/document.xml")!.async("string");
    zip.file("word/document.xml", xml.replace("<w:body>", '<w:body><w:p><w:r><w:object><o:OLEObject xmlns:o="urn:schemas-microsoft-com:office:office" ProgID="Paket"/></w:object></w:r></w:p>'));
    zip.file("word/vbaProject.bin", new Uint8Array([1, 2, 3]));
    const r = await openDocument("nesne.docm", buf(await zip.generateAsync({ type: "uint8array" })));
    expect(r.warnings).toEqual(expect.arrayContaining([lossMessage("objects"), lossMessage("macros")]));
    expect(r.warning).toBe(r.warnings.join(" "));
  });

  it("uyarısız belgede warnings boş, warning tanımsız", async () => {
    const r = await openDocument("ornek.docx", fixtureBuffer("lo-ornek.docx"));
    expect(r.warnings).toEqual([]);
    expect(r.warning).toBeUndefined();
  });

  it(".fodt düz XML olarak değil OpenDocument olarak okunur", async () => {
    const r = await openDocument("kaynak.fodt", fixtureBuffer("kaynak.fodt"));
    expect(r.format).toBe("fodt");
    expect(r.warning).toBe(FODT_NOTE);
    expect(r.warnings).toEqual([]);
    expect(all(r.doc.content, "heading").map(text)).toEqual(["Giriş ve Amaç", "Tablo ve Resim"]);
    expect(all(r.doc.content, "table")).toHaveLength(1);
    expect(all(r.doc.content, "image")).toHaveLength(1);
    expect(docToText(r.doc.content)).not.toContain("office:document");
    expect(r.doc.settings.properties.title).toBe("Biçim Örneği");
    const iz = await openDocument("izleme.fodt", fixtureBuffer("izleme.fodt"));
    expect(marked(iz.doc.content, "deletion")).toBe("silinen sözcük");
  });

  it("HTML: göreli yollu resim çözülemezse uyarılır, çözülürse gömülür", async () => {
    const html = enc('<html><body><p>Resim: <img src="resimler/a.png" width="20" height="10"></p></body></html>');
    const r1 = await openDocument("sayfa.html", html);
    expect(r1.warnings).toContain(lossMessage("images"));
    const png = Uint8Array.from(atob("iVBORw0KGgoAAAANSUhEUgAAABAAAAAMCAIAAADkharWAAAAFUlEQVR4nGM4oaFBEmIY1TCqATsCALRr0gH04WUkAAAAAElFTkSuQmCC"), (c) => c.charCodeAt(0));
    const asked: string[] = [];
    const r2 = await openDocument("sayfa.html", html, { resolveImage: async (p) => (asked.push(p), png) });
    expect(asked).toEqual(["resimler/a.png"]);
    expect(r2.warnings).toEqual([]);
    expect(String(all(r2.doc.content, "image")[0].attrs?.src)).toMatch(/^data:image\/png;base64,/);
    // Gömülen resim DOCX'e yazılır.
    expect(all((await readDocx(await writeDocx(r2.doc))).content, "image")).toHaveLength(1);
  });

  it("bozuk RTF (kapanmayan gruplar) boş açılmak yerine uyarı verir", async () => {
    const src = new TextDecoder("latin1").decode(fixture("lo-ornek.rtf")).replace(/\}/g, "");
    const r = await openDocument("dengesiz.rtf", buf(new Uint8Array([...src].map((c) => c.charCodeAt(0)))));
    expect(r.warnings).toContain(lossMessage("unclosed"));
  });
});

describe("hata iletileri", () => {
  const cases: [string, () => Promise<ArrayBuffer>, RegExp][] = [
    ["kesik zip", async () => buf(fixture("lo-ornek.docx").slice(0, 3000)), /ZIP/],
    ["boş document.xml", () => patchDocx("lo-ornek.docx", "word/document.xml", () => ""), /bozuk/i],
    ["çöp document.xml", () => patchDocx("lo-ornek.docx", "word/document.xml", (x) => x.slice(0, x.length >> 1)), /bozuk/i],
    ["document.xml yok", async () => {
      const z = await JSZip.loadAsync(fixture("lo-ornek.docx"));
      z.remove("word/document.xml");
      return buf(await z.generateAsync({ type: "uint8array" }));
    }, /bozuk|geçerli/i],
    ["sahte zip imzası", async () => buf(Uint8Array.from([0x50, 0x4b, 0x03, 0x04, ...Array.from({ length: 500 }, (_, i) => (i * 37) & 255)])), /ZIP|bozuk/],
  ];
  for (const [name, make, re] of cases) {
    it(`${name}: Türkçe ve anlaşılır ileti, teknik ayrıntı ayrı`, async () => {
      const err = await openDocument("bozuk.docx", await make()).then(
        () => null,
        (e: unknown) => e,
      );
      expect(err).toBeInstanceOf(BelgeOpenError);
      const e = err as BelgeOpenError;
      expect(e.message).toMatch(re);
      expect(e.message).toMatch(/[çğıöşüÇĞİÖŞÜ]/);
      expect(e.message).not.toMatch(/docx file|Corrupted|Could not|xmldom|central directory|Microsoft|Word/);
    });
  }

  it("yanlış uzantılı çalışma kitabı / sunu hangi uygulamayla açılacağını söyler", async () => {
    for (const [main, re] of [["xl/workbook.xml", /Tablo/], ["ppt/presentation.xml", /Sunum/]] as const) {
      const z = new JSZip();
      z.file("[Content_Types].xml", '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>');
      z.file("_rels/.rels", `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="${main}"/></Relationships>`);
      z.file(main, "<x/>");
      await expect(openDocument("yanlis.docx", buf(await z.generateAsync({ type: "uint8array" })))).rejects.toThrow(re);
    }
  });

  it("ham kitaplık iletileri sınıflandırılır", () => {
    expect(classifyOpenError(new Error("Could not find the body element: are you sure this is a docx file?"), "docx").message).toMatch(/ana belge bölümü/);
    expect(classifyOpenError(new Error("End of data reached (data length = 10, asked index = 20). Corrupted zip ?"), "docx").message).toMatch(/ZIP/);
    expect(classifyOpenError(new Error("[xmldom error]\tinvalid doc source"), "odt").message).toMatch(/XML/);
    const so = new RangeError("Maximum call stack size exceeded");
    expect(classifyOpenError(so, "docx").message).toMatch(/iç içe/);
    expect(classifyOpenError(new Error("Geçerli bir RTF dosyası değil."), "rtf").message).toBe("Geçerli bir RTF dosyası değil.");
  });
});

/** LibreOffice kuruluysa: kaydedilen DOCX'in sayfa sayısı özgününkiyle aynı olmalı (belge varsayılanları korunur). */
const hasSoffice = (() => {
  try {
    execFileSync("soffice", ["--version"], { stdio: "ignore", timeout: 60_000 });
    return true;
  } catch {
    return false;
  }
})();

describe.skipIf(!hasSoffice)("LibreOffice ile sayfa sayısı", () => {
  it("DOCX gidiş-dönüşünde sayfa sayısı değişmez (LibreOffice ve başka üreticilerin belgeleri)", async () => {
    const dir = mkdtempSync(join(tmpdir(), "oh-sayfa-"));
    const soffice = (fmt: string, ...files: string[]) => execFileSync("soffice", ["--headless", `-env:UserInstallation=file://${join(dir, "profil")}`, "--convert-to", fmt, "--outdir", dir, ...files], { stdio: "ignore", timeout: 200_000 });
    const pages = (pdf: string) => (readFileSync(pdf).toString("latin1").match(/\/Type\s*\/Page(?!s)/g) ?? []).length;
    try {
      // LibreOffice varsayılanlarıyla (Liberation Serif 12, aralıksız) çok sayfalı belge.
      const para = "Ölçüm için üretilmiş uzun paragraf; Türkçe karakterler (ığüşöç İĞÜŞÖÇ) ve sayfa düzenini etkileyen satırlar içerir. ".repeat(4);
      const body = Array.from({ length: 90 }, (_, i) => (i % 15 === 0 ? `<text:h text:outline-level="1">Bölüm ${i / 15 + 1}</text:h>` : `<text:p text:style-name="Standard">${para}</text:p>`)).join("");
      const fodt = new TextDecoder().decode(fixture("izleme.fodt")).replace(/<office:text>[\s\S]*<\/office:text>/, `<office:text>${body}</office:text>`);
      writeFileSync(join(dir, "lo-uzun.fodt"), fodt);
      soffice("docx", join(dir, "lo-uzun.fodt"));
      writeFileSync(join(dir, "cok-bolumlu.docx"), readFileSync(join(__dirname, "../__fixtures__/cok-bolumlu.docx")));
      for (const name of ["lo-uzun", "cok-bolumlu"]) {
        const src = join(dir, `${name}.docx`);
        const ours = join(dir, `${name}-oh.docx`);
        writeFileSync(ours, await writeDocx(await readDocx(readFileSync(src))));
        soffice("pdf", src, ours);
        const a = pages(join(dir, `${name}.pdf`));
        const b = pages(join(dir, `${name}-oh.pdf`));
        expect(a).toBeGreaterThan(1);
        expect(b, name).toBe(a);
      }
    } finally {
      rmSync(dir, { recursive: true, force: true });
    }
  }, 300_000);
});
