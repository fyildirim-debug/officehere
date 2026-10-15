import { deflateSync, crc32 } from "node:zlib";
import { PDFArray, PDFDict, PDFDocument, PDFHexString, PDFName, PDFRawStream, decodePDFRawStream } from "@cantoo/pdf-lib";
import { describe, expect, it } from "vitest";
import {
  applyPdfStandard,
  describeTransparency,
  findTransparency,
  glyphCount,
  hasLargeReal,
  listFonts,
  loadPdf,
  loadSrgbProfile,
  pdfExportRequest,
  prepareForPdfA1,
  rewriteNotdef,
  savePdf,
  splitLargeMatrices,
  stripInvisiblePaint,
  xmpPacket,
} from "../src/main/pdfa";

/** 2×2 RGBA PNG (alfa kanallı görsel). */
function rgbaPng(): Uint8Array {
  const chunk = (type: string, data: Buffer) => {
    const len = Buffer.alloc(4);
    len.writeUInt32BE(data.length);
    const td = Buffer.concat([Buffer.from(type, "latin1"), data]);
    const crc = Buffer.alloc(4);
    crc.writeUInt32BE(crc32(td));
    return Buffer.concat([len, td, crc]);
  };
  const ihdr = Buffer.alloc(13);
  ihdr.writeUInt32BE(2, 0);
  ihdr.writeUInt32BE(2, 4);
  ihdr[8] = 8;
  ihdr[9] = 6;
  const rows = Buffer.from([0, 255, 0, 0, 128, 0, 255, 0, 255, 0, 0, 0, 255, 64, 0, 255, 255, 255, 0]);
  return Buffer.concat([Buffer.from([137, 80, 78, 71, 13, 10, 26, 10]), chunk("IHDR", ihdr), chunk("IDAT", deflateSync(rows)), chunk("IEND", Buffer.alloc(0))]);
}

/** Yalnız tablo dizini ve 'maxp' tablosu olan TrueType (glif sayısı okuması için yeterli). */
function ttf(numGlyphs: number): Uint8Array {
  const b = Buffer.alloc(12 + 16 + 6);
  b.writeUInt32BE(0x00010000, 0);
  b.writeUInt16BE(1, 4);
  b.write("maxp", 12, "latin1");
  b.writeUInt32BE(28, 20);
  b.writeUInt32BE(6, 24);
  b.writeUInt32BE(0x00005000, 28);
  b.writeUInt16BE(numGlyphs, 32);
  return b;
}

/** Çıplak CFF: Name INDEX, Top DICT (CharStrings ofseti, işleç 17) ve n öğeli CharStrings INDEX. */
function cff(n: number): Uint8Array {
  const header = [1, 0, 4, 1];
  const nameIndex = [0, 1, 1, 1, 2, 0x41];
  // Top DICT: <ofset> 17; ofset = başlık + Name INDEX + Top DICT INDEX (5 bayt üst bilgi + 4 bayt DICT).
  const offset = header.length + nameIndex.length + 5 + 4;
  const dict = [28, offset >> 8, offset & 0xff, 17];
  const topIndex = [0, 1, 1, 1, 1 + dict.length, ...dict];
  const charStrings = [n >> 8, n & 0xff, 1, ...Array.from({ length: n + 1 }, (_, i) => i + 1), ...Array(n).fill(14)];
  return Uint8Array.from([...header, ...nameIndex, ...topIndex, ...charStrings]);
}

/** Skia çıktısına benzer CIDFontType2 alt kümesi ile tek sayfalık belge (yazı tipi kaynağı /F1). */
async function cidDocument(glyphs = 300, content = "BT /F1 12 Tf <00010002> Tj ET") {
  const doc = await PDFDocument.create({ updateMetadata: false });
  const ctx = doc.context;
  const page = doc.addPage([200, 200]);
  const file = ctx.register(ctx.flateStream(ttf(glyphs)));
  const fd = ctx.register(ctx.obj({ Type: "FontDescriptor", FontName: "AAAAAA+Deneme", Flags: 4, FontBBox: [0, 0, 1000, 1000], ItalicAngle: 0, Ascent: 800, Descent: -200, CapHeight: 700, StemV: 80, FontFile2: file }));
  const desc = ctx.register(ctx.obj({ Type: "Font", Subtype: "CIDFontType2", BaseFont: "AAAAAA+Deneme", CIDSystemInfo: { Registry: PDFHexString.fromText("Adobe"), Ordering: PDFHexString.fromText("Identity"), Supplement: 0 }, FontDescriptor: fd, DW: 1000, W: [0, [500, 600, 700]], CIDToGIDMap: "Identity" }));
  const font = ctx.register(ctx.obj({ Type: "Font", Subtype: "Type0", BaseFont: "AAAAAA+Deneme", Encoding: "Identity-H", DescendantFonts: [desc] }));
  page.node.set(PDFName.of("Resources"), ctx.obj({ Font: { F1: font } }));
  page.node.set(PDFName.of("Contents"), ctx.register(ctx.flateStream(Buffer.from(content, "latin1"))));
  return doc;
}

const roundTrip = async (doc: PDFDocument) => loadPdf(await savePdf(doc));
const META = { title: "Deneme Yazısı İĞÜŞÖÇ & <ek>", author: "Furkan YILDIRIM", subject: "Konu ş", keywords: "arşiv, PDF/A", creator: "OfficeHere Belge", producer: "OfficeHere 0.1.6; Skia/PDF m152" };
const DATE = new Date("2026-10-15T08:30:00Z");

describe("pdfExportRequest", () => {
  it("bilinmeyen değerleri varsayılana düşürür, metni temizler", () => {
    const r = pdfExportRequest({ standard: "pdfa-3u", pageSize: "B5", after: "sil", title: "  Ba\u0001şlık  ", outline: "evet", pageRanges: "1-3; rm -rf" });
    expect(r).toMatchObject({ standard: "pdf", pageSize: "A4", after: "none", title: "Başlık", outline: false, tagged: false, landscape: false });
    expect(r.pageRanges).toBeUndefined();
    expect(pdfExportRequest(null).standard).toBe("pdf");
  });

  it("geçerli seçenekleri taşır", () => {
    const r = pdfExportRequest({ standard: "pdfa-1b", pageSize: { width: 8.27, height: 11.69 }, pageRanges: "1-3, 5", after: "folder", outline: true, tagged: true, landscape: true });
    expect(r).toMatchObject({ standard: "pdfa-1b", pageSize: { width: 8.27, height: 11.69 }, pageRanges: "1-3, 5", after: "folder", outline: true, tagged: true, landscape: true });
    expect(pdfExportRequest({ pageSize: { width: -1, height: 2 } }).pageSize).toBe("A4");
  });
});

describe("glyphCount", () => {
  it("TrueType 'maxp' tablosundan okur", () => {
    expect(glyphCount(ttf(1234))).toBe(1234);
  });
  it("çıplak CFF'te CharStrings INDEX sayısını okur", () => {
    expect(glyphCount(cff(7))).toBe(7);
  });
  it("tanımadığı veride null döner", () => {
    expect(glyphCount(new Uint8Array([9, 9, 9, 9, 9, 9, 9, 9, 9, 9, 9, 9]))).toBeNull();
  });
});

describe("findTransparency", () => {
  it("saydamlıksız sayfada bulgu yok", async () => {
    const doc = await PDFDocument.create();
    doc.addPage().drawRectangle({ x: 10, y: 10, width: 50, height: 50 });
    expect(findTransparency(await roundTrip(doc))).toEqual([]);
  });

  it("yarı saydam dolguyu ve alfa kanallı görseli sayfasıyla bildirir", async () => {
    const doc = await PDFDocument.create();
    doc.addPage().drawText("ilk");
    doc.addPage().drawRectangle({ x: 10, y: 10, width: 50, height: 50, opacity: 0.5 });
    const img = await doc.embedPng(rgbaPng());
    doc.addPage().drawImage(img, { x: 0, y: 0, width: 20, height: 20 });
    const found = findTransparency(await roundTrip(doc));
    expect(found).toEqual([
      { page: 2, kind: "opaklik" },
      { page: 3, kind: "gorsel" },
    ]);
    expect(describeTransparency(found)).toBe("• 2. sayfa: yarı saydam renk ya da opaklığı düşürülmüş öğe\n• 3. sayfa: saydam arka planlı ya da yarı saydam görsel");
  });
});

describe("applyPdfStandard", () => {
  it("standart PDF: belge bilgileri yazılır, PDF/A işaretleri eklenmez", async () => {
    const doc = await cidDocument();
    applyPdfStandard(doc, { standard: "pdf", meta: META, date: DATE });
    const out = await roundTrip(doc);
    expect(out.getTitle()).toBe(META.title);
    expect(out.getAuthor()).toBe(META.author);
    expect(out.getCreator()).toBe("OfficeHere Belge");
    expect(out.catalog.has(PDFName.of("Metadata"))).toBe(false);
    expect(out.catalog.has(PDFName.of("OutputIntents"))).toBe(false);
  });

  it("PDF/A-2b: XMP Info ile aynı, sRGB OutputIntent ve /ID var; CIDSet eklenmez", async () => {
    const doc = await cidDocument();
    const icc = loadSrgbProfile();
    applyPdfStandard(doc, { standard: "pdfa-2b", meta: META, icc, date: DATE, id: new Uint8Array(16).fill(7) });
    const out = await roundTrip(doc);
    const xmp = Buffer.from((out.catalog.lookup(PDFName.of("Metadata")) as PDFRawStream).getContents()).toString("utf8");
    expect(xmp).toContain("<pdfaid:part>2</pdfaid:part><pdfaid:conformance>B</pdfaid:conformance>");
    expect(xmp).toContain("<rdf:li xml:lang=\"x-default\">Deneme Yazısı İĞÜŞÖÇ &amp; &lt;ek&gt;</rdf:li>");
    expect(xmp).toContain("<rdf:Seq><rdf:li>Furkan YILDIRIM</rdf:li></rdf:Seq>");
    expect(xmp).toContain("<xmp:CreateDate>2026-10-15T08:30:00Z</xmp:CreateDate>");
    expect(out.getTitle()).toBe(META.title);
    expect(out.getCreationDate()?.toISOString()).toBe(DATE.toISOString());
    const oi = (out.catalog.lookup(PDFName.of("OutputIntents")) as PDFArray).lookup(0) as PDFDict;
    expect(oi.get(PDFName.of("S"))?.toString()).toBe("/GTS_PDFA1");
    const profile = oi.lookup(PDFName.of("DestOutputProfile")) as PDFRawStream;
    expect(Buffer.from(decodePDFRawStream(profile).decode())).toEqual(Buffer.from(icc));
    expect(out.context.trailerInfo.ID?.toString()).toContain("07070707");
    expect(listFonts(out)).toEqual([{ name: "AAAAAA+Deneme", embedded: true, cidSet: false }]);
  });

  it("PDF/A-1b: CIDSet tüm glifleri işaretler, boş saydamlık grubu kalkar, bağlantı yazdırılır", async () => {
    const doc = await cidDocument(10);
    const ctx = doc.context;
    const page = doc.getPage(0);
    page.node.set(PDFName.of("Group"), ctx.obj({ S: "Transparency", CS: "DeviceRGB" }));
    page.node.set(PDFName.of("Annots"), ctx.obj([ctx.register(ctx.obj({ Type: "Annot", Subtype: "Link", Rect: [0, 0, 10, 10], F: 2 }))]));
    applyPdfStandard(doc, { standard: "pdfa-1b", meta: META, icc: loadSrgbProfile(), date: DATE });
    const out = await roundTrip(doc);
    const fd = [...out.context.enumerateIndirectObjects()].map(([, o]) => o).find((o): o is PDFDict => o instanceof PDFDict && o.get(PDFName.of("Type"))?.toString() === "/FontDescriptor")!;
    const bits = Buffer.from(decodePDFRawStream(fd.lookup(PDFName.of("CIDSet")) as PDFRawStream).decode());
    expect([...bits]).toEqual([0xff, 0xc0]);
    expect(out.getPage(0).node.has(PDFName.of("Group"))).toBe(false);
    const annot = (out.getPage(0).node.Annots() as PDFArray).lookup(0) as PDFDict;
    expect(annot.get(PDFName.of("F"))?.toString()).toBe("4");
    const xmp = Buffer.from((out.catalog.lookup(PDFName.of("Metadata")) as PDFRawStream).getContents()).toString("utf8");
    expect(xmp).toContain("<pdfaid:part>1</pdfaid:part>");
  });

  it("PDF/A: .notdef glifi aynı genişlikte kaydırmayla değiştirilir", async () => {
    const doc = await cidDocument(10, "BT /F1 12 Tf <000100000002> Tj ET BT /F1 12 Tf [<0000> 120 <0002>] TJ ET");
    applyPdfStandard(doc, { standard: "pdfa-2b", meta: META, icc: loadSrgbProfile(), date: DATE });
    const out = await roundTrip(doc);
    const content = out.context.lookup(out.getPage(0).node.get(PDFName.of("Contents"))) as PDFRawStream;
    expect(Buffer.from(decodePDFRawStream(content).decode()).toString("latin1")).toBe("BT /F1 12 Tf [<0001> -500 <0002>] TJ ET BT /F1 12 Tf [-500 120 <0002>] TJ ET");
  });

  it("PDF/A renk profili olmadan üretilmez", async () => {
    const doc = await cidDocument();
    expect(() => applyPdfStandard(doc, { standard: "pdfa-2b", meta: META })).toThrow();
  });
});

describe("rewriteNotdef", () => {
  it("yalnız Identity-H yazı tiplerinde ve .notdef varken değiştirir", () => {
    const fonts = new Map<string, number | null>([
      ["F1", 250],
      ["F2", null],
    ]);
    expect(rewriteNotdef("/F1 9 Tf <00010002> Tj", fonts)).toBeNull();
    expect(rewriteNotdef("/F2 9 Tf <00000002> Tj", fonts)).toBeNull();
    expect(rewriteNotdef("/F1 9 Tf <0000> Tj /NonStruct <</MCID 3 >>BDC", fonts)).toBe("/F1 9 Tf [-250] TJ /NonStruct <</MCID 3 >>BDC");
  });
});

describe("stripInvisiblePaint", () => {
  const alpha = new Map([
    ["G0", { ca: 0 }],
    ["S0", { CA: 0 }],
    ["G1", { ca: 1, CA: 1 }],
  ]);

  it("opaklığı 0 olan dolgu, metin ve görsel çizimlerini çıkarır; q/Q durumu izlenir", () => {
    const src = ["q /G0 gs 0 0 10 10 re f BT /F1 9 Tf <0001> Tj ET /Im1 Do Q", "0 0 5 5 re f", "/G0 gs 1 1 m 2 2 l B", "/G1 gs 0 0 1 1 re f*"].join("\n");
    expect(stripInvisiblePaint(src, alpha)).toBe(["q /G0 gs 0 0 10 10 re n BT /F1 9 Tf  ET  Q", "0 0 5 5 re f", "/G0 gs 1 1 m 2 2 l S", "/G1 gs 0 0 1 1 re f*"].join("\n"));
  });

  it("çizgisi görünmez dolguda yalnız dolgu kalır; görünmez değilse akış değişmez", () => {
    expect(stripInvisiblePaint("/S0 gs 0 0 m 1 1 l b*", alpha)).toBe("/S0 gs 0 0 m 1 1 l h f*");
    const same = "/G1 gs (a\\)b) Tj [<0001> -20 <0002>] TJ";
    expect(stripInvisiblePaint(same, alpha)).toBe(same);
  });

  it("satır içi görsel ya da kısmen görünen metin varsa dönüştürmez", () => {
    expect(stripInvisiblePaint("BI /W 1 /H 1 ID x EI", alpha)).toBeNull();
    expect(stripInvisiblePaint("/G0 gs 2 Tr (a) Tj", alpha)).toBeNull();
  });
});

describe("PDF/A-1 sayı sınırı", () => {
  it("büyük ötelemeyi sınır içindeki parçalara böler", () => {
    const out = splitLargeMatrices("q\n3.125 0 0 3.125 0 -35093.75 cm\n1 0 0 1 5 5 cm\n")!;
    expect(out).toBe("q\n1 0 0 1 0 -17546.875 cm\n1 0 0 1 0 -17546.875 cm\n3.125 0 0 3.125 0 0 cm\n1 0 0 1 5 5 cm\n");
    expect(hasLargeReal(out)).toBe(false);
  });

  it("yalnız ondalıklı sayıları sayar; glif dizileri ve tamsayılar sınır dışı değildir", () => {
    expect(hasLargeReal("<00890044> Tj /P <</MCID 99999 >>BDC 40000 0 m")).toBe(false);
    expect(hasLargeReal("[<0001> -40000.5 <0002>] TJ")).toBe(true);
    expect(hasLargeReal("0 33000.0 l")).toBe(true);
  });
});

describe("prepareForPdfA1", () => {
  const icc = loadSrgbProfile();
  /** Profil sürümü (ve istenirse kırmızı birincili) değiştirilmiş kopya. */
  const variant = (major: number, shiftRed = false) => {
    const b = Buffer.from(icc);
    b[8] = major;
    if (shiftRed) {
      const n = b.readUInt32BE(128);
      for (let i = 0; i < n; i++) if (b.toString("latin1", 132 + 12 * i, 136 + 12 * i) === "rXYZ") b.writeInt32BE(b.readInt32BE(b.readUInt32BE(136 + 12 * i) + 8) + 4000, b.readUInt32BE(136 + 12 * i) + 8);
    }
    return b;
  };
  const withImage = async (profile: Uint8Array) => {
    const doc = await cidDocument(10);
    const ctx = doc.context;
    const p = ctx.register(ctx.flateStream(profile, { N: 3 }));
    const img = ctx.register(ctx.flateStream(new Uint8Array(3), { Type: "XObject", Subtype: "Image", Width: 1, Height: 1, BitsPerComponent: 8, ColorSpace: ["ICCBased", p] }));
    (doc.getPage(0).node.Resources() as PDFDict).set(PDFName.of("XObject"), ctx.obj({ Im1: img }));
    return { doc, profile: p };
  };

  it("sRGB birincilli ICC v4 profili sRGB v2 ile değişir; engel yok", async () => {
    const { doc, profile } = await withImage(variant(4));
    expect(prepareForPdfA1(doc, icc)).toEqual([]);
    expect(Buffer.from(decodePDFRawStream(doc.context.lookup(profile) as PDFRawStream).decode())).toEqual(Buffer.from(icc));
  });

  it("sRGB dışı ICC v4 profili ve saydamlık engel olarak bildirilir", async () => {
    const { doc } = await withImage(variant(4, true));
    const ctx = doc.context;
    (doc.getPage(0).node.Resources() as PDFDict).set(PDFName.of("ExtGState"), ctx.obj({ G5: ctx.obj({ ca: 0.5 }) }));
    expect(prepareForPdfA1(doc, icc)).toEqual(["• 1. sayfa: yarı saydam renk ya da opaklığı düşürülmüş öğe", "• 1 görselde sRGB dışı (ICC v4) renk profili var."]);
  });

  it("opaklığı 0 olan grafik durumu kullanıldığı akış dönüştürülünce opak yapılır", async () => {
    const doc = await cidDocument(10, "q /G0 gs 0 0 9 9 re f Q BT /F1 12 Tf <0001> Tj ET");
    const ctx = doc.context;
    const gs = ctx.obj({ ca: 0 });
    (doc.getPage(0).node.Resources() as PDFDict).set(PDFName.of("ExtGState"), ctx.obj({ G0: gs }));
    expect(findTransparency(doc)).toEqual([{ page: 1, kind: "opaklik" }]);
    expect(prepareForPdfA1(doc, icc)).toEqual([]);
    expect(gs.get(PDFName.of("ca"))?.toString()).toBe("1");
    const content = ctx.lookup(doc.getPage(0).node.get(PDFName.of("Contents"))) as PDFRawStream;
    expect(Buffer.from(decodePDFRawStream(content).decode()).toString("latin1")).toBe("q /G0 gs 0 0 9 9 re n Q BT /F1 12 Tf <0001> Tj ET");
  });
});

describe("xmpPacket", () => {
  it("boş alanlar yazılmaz; PDF/A işareti yalnız istenince", () => {
    const x = xmpPacket({ producer: "P" }, DATE);
    expect(x).not.toContain("dc:title");
    expect(x).not.toContain("pdfaid");
    expect(x).toContain("<pdf:Producer>P</pdf:Producer>");
    expect(x.startsWith('<?xpacket begin="﻿"')).toBe(true);
  });
});
