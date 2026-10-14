import { describe, expect, it } from "vitest";
import { Carry, carryParts, collectRefs, reattach, rootNamespace } from "../../src/opc/carry";
import { OpcPackage, parseRels, RelsBuilder } from "../../src/opc/package";
import { SourceRegistry } from "../../src/opc/source";
import { BASE_DEFAULTS, contentTypes, CT_MAIN_DOCM, CT_MAIN_DOCX, DOC_XML, minimalDocx, MST, opcHatalari, poi, poiDosyalari, R, rels, T, zipOf, type RelSpec } from "./paket";

const PK = (s: string) => `http://schemas.openxmlformats.org/package/2006/relationships/${s}`;
const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const CT_THEME = "application/vnd.openxmlformats-officedocument.theme+xml";
const CT_FONTTABLE = "application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml";
const CT_CUSTOM = "application/vnd.openxmlformats-officedocument.custom-properties+xml";
const CT_ITEMPROPS = "application/vnd.openxmlformats-officedocument.customXmlProperties+xml";

/** Kaynak docx: ana parça ilişkileri, kök ilişkileri ve ek parçalar. */
async function kaynakDocx(o: {
  body?: string;
  docRels?: RelSpec[];
  rootRels?: RelSpec[];
  parts?: Record<string, string | Uint8Array>;
  overrides?: Record<string, string>;
  defaults?: Record<string, string>;
  mainCt?: string;
}): Promise<Uint8Array> {
  return zipOf({
    "[Content_Types].xml": contentTypes({ ...BASE_DEFAULTS, ...(o.defaults ?? {}) }, { "word/document.xml": o.mainCt ?? CT_MAIN_DOCX, ...(o.overrides ?? {}) }),
    "_rels/.rels": rels([{ id: "rId1", type: T("officeDocument"), target: "word/document.xml" }, ...(o.rootRels ?? [])]),
    "word/document.xml": DOC_XML(o.body ?? ""),
    "word/_rels/document.xml.rels": rels(o.docRels ?? []),
    ...(o.parts ?? {}),
  });
}

const text = async (b: Uint8Array, name: string) => (await OpcPackage.open(b)).text(name);
const relsOf = async (b: Uint8Array, part: string) => (await OpcPackage.open(b)).rels(part);

describe("paket düzeyi taşıma: sınıflama", () => {
  it("örtük ilişki taşınır, açık ilişki modele kalır, sahip türler yazıcıda kalır", async () => {
    const src = await kaynakDocx({
      body: `<w:p><w:r><w:object><o:OLEObject xmlns:o="urn:schemas-microsoft-com:office:office" r:id="rId4"/></w:object><w:drawing><a:blip r:embed="rId5"/></w:drawing></w:r></w:p>`,
      docRels: [
        { id: "rId1", type: T("styles"), target: "styles.xml" },
        { id: "rId2", type: T("theme"), target: "theme/theme1.xml" },
        { id: "rId3", type: T("fontTable"), target: "fontTable.xml" },
        { id: "rId4", type: T("oleObject"), target: "embeddings/o.bin" },
        { id: "rId5", type: T("image"), target: "media/a.png" },
        { id: "rId6", type: "http://ornek.com/iliski/bilinmeyen", target: "ozel/veri.xml" },
      ],
      rootRels: [{ id: "rId2", type: PK("metadata/core-properties"), target: "docProps/core.xml" }],
      parts: {
        "word/styles.xml": "<w:styles/>",
        "word/theme/theme1.xml": "<a:theme/>",
        "word/fontTable.xml": "<w:fonts/>",
        "word/embeddings/o.bin": new Uint8Array([1, 2]),
        "word/media/a.png": new Uint8Array([3]),
        "word/ozel/veri.xml": "<veri/>",
        "docProps/core.xml": "<cp:coreProperties/>",
      },
      overrides: { "word/theme/theme1.xml": CT_THEME, "word/fontTable.xml": CT_FONTTABLE },
      defaults: { png: "image/png", bin: "application/vnd.openxmlformats-officedocument.oleObject" },
    });
    const out = await minimalDocx();
    const { bytes, report } = await carryParts({ app: "belge", source: await OpcPackage.open(src), output: out });
    expect(report.carried.map((c) => [c.from, c.to, c.kind])).toEqual([
      ["word/theme/theme1.xml", "word/theme/theme1.xml", "implicit"],
      ["word/fontTable.xml", "word/fontTable.xml", "implicit"],
      ["word/ozel/veri.xml", "word/ozel/veri.xml", "implicit"],
    ]);
    expect(report.explicit.map((e) => e.id)).toEqual(["rId4"]);
    expect(report.owned.map((e) => e.id).sort()).toEqual(["rId1", "rId2", "rId5"]);
    expect(report.dropped).toEqual([]);
    expect(opcHatalari(bytes)).toEqual([]);
    const pkg = await OpcPackage.open(bytes);
    expect(pkg.contentType("word/theme/theme1.xml")).toBe(CT_THEME);
    expect(pkg.contentType("word/fontTable.xml")).toBe(CT_FONTTABLE);
    expect(pkg.has("word/embeddings/o.bin")).toBe(false);
    // yazıcının ilişkileri korunur, yeniler yeni kimlik alır
    const docRels = await pkg.rels("word/document.xml");
    expect(docRels[0]).toMatchObject({ id: "rId1", target: "styles.xml" });
    expect(new Set(docRels.map((r) => r.id)).size).toBe(docRels.length);
    expect(docRels.find((r) => r.part === "word/theme/theme1.xml")?.target).toBe("theme/theme1.xml");
    // yazıcının parçaları değişmeden (ham) kopyalanır
    expect(await pkg.text("word/document.xml")).toBe(await text(out, "word/document.xml"));
  });

  it("[Content_Types].xml ilk parça; yazıcı aynı türü ürettiyse taşınmaz", async () => {
    const src = await kaynakDocx({
      docRels: [{ id: "rId2", type: T("theme"), target: "theme/theme1.xml" }],
      parts: { "word/theme/theme1.xml": "<a:theme name='özgün'/>" },
      overrides: { "word/theme/theme1.xml": CT_THEME },
    });
    const out = await minimalDocx({ "word/theme/tema.xml": "<a:theme name='yazıcı'/>" }, { docRels: [{ id: "rId9", type: T("theme"), target: "theme/tema.xml" }], overrides: { "word/theme/tema.xml": CT_THEME } });
    const { bytes, report } = await carryParts({ app: "belge", source: await OpcPackage.open(src), output: out });
    expect(report.carried).toEqual([]);
    expect(report.owned.map((o) => o.part)).toEqual(["word/theme/theme1.xml"]);
    expect(opcHatalari(out).map((h) => h.tur)).toContain("ct-ilk-değil");
    expect(opcHatalari(bytes)).toEqual([]);
    expect((await OpcPackage.open(bytes)).has("word/theme/theme1.xml")).toBe(false);
  });

  it("düşürme politikası: imza, küçük resim, stylesWithEffects, yorum genişletmeleri", async () => {
    const src = await kaynakDocx({
      rootRels: [
        { id: "rId5", type: PK("digital-signature/origin"), target: "_xmlsignatures/origin.sigs" },
        { id: "rId6", type: PK("metadata/thumbnail"), target: "docProps/thumbnail.jpeg" },
      ],
      docRels: [
        { id: "rId1", type: MST("2007/relationships/stylesWithEffects"), target: "stylesWithEffects.xml" },
        { id: "rId2", type: MST("2011/relationships/commentsExtended"), target: "commentsExtended.xml" },
        { id: "rId3", type: MST("2011/relationships/people"), target: "people.xml" },
      ],
      parts: {
        "_xmlsignatures/origin.sigs": new Uint8Array(),
        "_xmlsignatures/_rels/origin.sigs.rels": rels([{ id: "rId1", type: PK("digital-signature/signature"), target: "sig1.xml" }]),
        "_xmlsignatures/sig1.xml": "<Signature/>",
        "docProps/thumbnail.jpeg": new Uint8Array([0xff, 0xd8]),
        "word/stylesWithEffects.xml": "<w:styles/>",
        "word/commentsExtended.xml": "<w15:commentsEx/>",
        "word/people.xml": "<w15:people/>",
      },
      defaults: { sigs: "application/vnd.openxmlformats-package.digital-signature-origin", jpeg: "image/jpeg" },
      overrides: { "_xmlsignatures/sig1.xml": "application/vnd.openxmlformats-package.digital-signature-xmlsignature+xml" },
    });
    const pkgSrc = await OpcPackage.open(src);
    expect(pkgSrc.signed).toBe(true);
    const { bytes, report } = await carryParts({ app: "belge", source: pkgSrc, output: await minimalDocx() });
    expect(report.dropped.map((d) => [d.reason, d.part])).toEqual([
      ["signature", "_xmlsignatures/origin.sigs"],
      ["thumbnail", "docProps/thumbnail.jpeg"],
      ["stale", "word/stylesWithEffects.xml"],
      ["comments", "word/commentsExtended.xml"],
      ["comments", "word/people.xml"],
    ]);
    expect(report.carried).toEqual([]);
    const pkg = await OpcPackage.open(bytes);
    expect(pkg.names.some((n) => n.startsWith("_xmlsignatures"))).toBe(false);
    expect(pkg.signed).toBe(false);
  });

  it("hedefsiz ilişki atılır; dış hedefli örtük ilişki taşınır", async () => {
    const src = await kaynakDocx({
      docRels: [
        { id: "rId1", type: "http://ornek.com/iliski/a", target: "yok.xml" },
        { id: "rId2", type: "http://ornek.com/iliski/b", target: "https://ornek.com/kaynak", external: true },
        { id: "rId3", type: "http://ornek.com/iliski/c", target: "c.xml" },
        { id: "rId4", type: "http://ornek.com/iliski/d", target: "../../../dışarı.xml" },
      ],
      parts: { "word/c.xml": "<c/>", "word/_rels/c.xml.rels": rels([{ id: "rId1", type: T("x"), target: "kayip.xml" }, { id: "rId2", type: T("y"), target: "https://a.b", external: true }]) },
    });
    const { bytes, report } = await carryParts({ app: "belge", source: await OpcPackage.open(src), output: await minimalDocx() });
    expect(report.dropped.map((d) => [d.reason, d.part])).toEqual([
      ["missing", "word/yok.xml"],
      ["missing", "../../../dışarı.xml"],
      ["missing", "word/kayip.xml"],
    ]);
    expect(report.carried.map((c) => [c.kind, c.to])).toEqual([
      ["external", "https://ornek.com/kaynak"],
      ["implicit", "word/c.xml"],
    ]);
    expect(opcHatalari(bytes)).toEqual([]);
    const docRels = await relsOf(bytes, "word/document.xml");
    expect(docRels.find((r) => r.external)).toMatchObject({ target: "https://ornek.com/kaynak", type: "http://ornek.com/iliski/b" });
    expect((await relsOf(bytes, "word/c.xml")).map((r) => r.id)).toEqual(["rId2"]);
  });
});

describe("paket düzeyi taşıma: kapanış ve adlar", () => {
  it("döngülü .rels: kapanış sonlanır, her parça bir kez taşınır, ana parçaya geri başvuru eşlenir", async () => {
    const src = await kaynakDocx({
      docRels: [{ id: "rId7", type: T("glossaryDocument"), target: "glossary/document.xml" }],
      parts: {
        "word/glossary/document.xml": `<w:glossaryDocument xmlns:w="${W}"/>`,
        "word/glossary/_rels/document.xml.rels": rels([
          { id: "rId1", type: T("styles"), target: "styles.xml" },
          { id: "rId2", type: "http://ornek.com/geri", target: "../document.xml" },
          { id: "rId3", type: T("glossaryDocument"), target: "/word/glossary/document.xml" },
        ]),
        "word/glossary/styles.xml": "<w:styles/>",
        "word/glossary/_rels/styles.xml.rels": rels([
          { id: "rId1", type: "http://ornek.com/dongu", target: "document.xml" },
          { id: "rId2", type: "http://ornek.com/kendi", target: "styles.xml" },
          { id: "rId3", type: "http://ornek.com/yok", target: "yok.xml" },
        ]),
      },
      overrides: { "word/glossary/document.xml": "application/vnd.openxmlformats-officedocument.wordprocessingml.document.glossary+xml" },
    });
    const { bytes, report } = await carryParts({ app: "belge", source: await OpcPackage.open(src), output: await minimalDocx() });
    expect(report.carried.map((c) => [c.from, c.kind])).toEqual([
      ["word/glossary/document.xml", "implicit"],
      ["word/glossary/styles.xml", "closure"],
    ]);
    expect(report.dropped.map((d) => d.reason)).toEqual(["missing"]);
    expect(opcHatalari(bytes)).toEqual([]);
    const g = await relsOf(bytes, "word/glossary/document.xml");
    expect(g.map((r) => [r.id, r.part])).toEqual([
      ["rId1", "word/glossary/styles.xml"],
      ["rId2", "word/document.xml"],
      ["rId3", "word/glossary/document.xml"],
    ]);
    const s = await relsOf(bytes, "word/glossary/styles.xml");
    expect(s.map((r) => r.part)).toEqual(["word/glossary/document.xml", "word/glossary/styles.xml"]);
  });

  it("ad çakışması: oh_tN_ önekiyle yeniden adlandırılır, ilişki hedefleri ve içerik türleri yeni ada göre yazılır", async () => {
    const item = (ns: string) => `<?xml version="1.0"?><k:Kok xmlns:k="${ns}"/>`;
    const src = await kaynakDocx({
      docRels: [{ id: "rId3", type: T("customXml"), target: "../customXml/item1.xml" }],
      parts: {
        "customXml/item1.xml": item("urn:kaynak"),
        "customXml/_rels/item1.xml.rels": rels([{ id: "rId1", type: T("customXmlProps"), target: "itemProps1.xml" }]),
        "customXml/itemProps1.xml": "<ds:datastoreItem/>",
      },
      overrides: { "customXml/itemProps1.xml": CT_ITEMPROPS },
    });
    const out = await minimalDocx(
      {
        "customXml/item1.xml": item("urn:yazici"),
        "customXml/_rels/item1.xml.rels": rels([{ id: "rId1", type: T("customXmlProps"), target: "itemProps1.xml" }]),
        "customXml/itemProps1.xml": "<ds:datastoreItem yazici='1'/>",
      },
      { docRels: [{ id: "rId2", type: T("customXml"), target: "../customXml/item1.xml" }], overrides: { "customXml/itemProps1.xml": CT_ITEMPROPS } },
    );
    const { bytes, report } = await carryParts({ app: "belge", source: await OpcPackage.open(src), output: out });
    const [a, b] = report.carried;
    expect(a.from).toBe("customXml/item1.xml");
    expect(a.to).toMatch(/^customXml\/oh_t\d+_item1\.xml$/);
    expect(b.to).toMatch(/^customXml\/oh_t\d+_itemProps1\.xml$/);
    expect(opcHatalari(bytes)).toEqual([]);
    const pkg = await OpcPackage.open(bytes);
    expect(await pkg.text("customXml/itemProps1.xml")).toContain("yazici");
    expect(rootNamespace((await pkg.text(a.to))!)).toBe("urn:kaynak");
    expect((await pkg.rels(a.to)).map((r) => r.part)).toEqual([b.to]);
    expect(pkg.contentType(b.to)).toBe(CT_ITEMPROPS);
  });

  it("customXml: aynı kök ad alanlı öğe yazıcıda varsa yazıcınınki geçerli", async () => {
    const src = await kaynakDocx({
      docRels: [{ id: "rId3", type: T("customXml"), target: "../customXml/item1.xml" }],
      parts: { "customXml/item1.xml": `<b:Sources xmlns:b="urn:kaynakca" eski="1"/>` },
    });
    const out = await minimalDocx({ "customXml/item9.xml": `<b:Sources xmlns:b="urn:kaynakca"/>` }, { docRels: [{ id: "rId2", type: T("customXml"), target: "../customXml/item9.xml" }] });
    const { report } = await carryParts({ app: "belge", source: await OpcPackage.open(src), output: out });
    expect(report.carried).toEqual([]);
    expect(report.owned.map((o) => o.part)).toEqual(["customXml/item1.xml"]);
  });

  it("içerik türü: Override taşınır, kaynak varsayılanı çıktıda yoksa eklenir, çakışan varsayılanda Override yazılır", async () => {
    const src = await kaynakDocx({
      docRels: [
        { id: "rId1", type: T("fontTable"), target: "fontTable.xml" },
        { id: "rId2", type: "http://ornek.com/ikili", target: "veri/k.bin" },
      ],
      parts: {
        "word/fontTable.xml": "<w:fonts/>",
        "word/_rels/fontTable.xml.rels": rels([{ id: "rId1", type: T("font"), target: "fonts/font1.odttf" }]),
        "word/fonts/font1.odttf": new Uint8Array([9, 9]),
        "word/veri/k.bin": new Uint8Array([1]),
        "word/adsiz": "uzantısız",
      },
      overrides: { "word/fontTable.xml": CT_FONTTABLE },
      defaults: { odttf: "application/vnd.openxmlformats-officedocument.obfuscatedFont", bin: "application/x-kaynak" },
    });
    const out = await zipOf({
      "[Content_Types].xml": contentTypes({ ...BASE_DEFAULTS, bin: "application/x-yazici" }, { "word/document.xml": CT_MAIN_DOCX }),
      "_rels/.rels": rels([{ id: "rId1", type: T("officeDocument"), target: "word/document.xml" }]),
      "word/document.xml": DOC_XML(),
      "word/media/b.bin": new Uint8Array([2]),
    });
    const { bytes } = await carryParts({ app: "belge", source: await OpcPackage.open(src), output: out });
    expect(opcHatalari(bytes)).toEqual([]);
    const pkg = await OpcPackage.open(bytes);
    expect(pkg.contentTypes.getOverride("word/fontTable.xml")).toBe(CT_FONTTABLE);
    expect(pkg.contentTypes.getDefault("odttf")).toBe("application/vnd.openxmlformats-officedocument.obfuscatedFont");
    expect(pkg.contentTypes.getDefault("bin")).toBe("application/x-yazici");
    expect(pkg.contentTypes.getOverride("word/veri/k.bin")).toBe("application/x-kaynak");
    expect(pkg.contentType("word/media/b.bin")).toBe("application/x-yazici");
  });

  it("içerik türü olmayan kaynak parça application/octet-stream ile taşınır", async () => {
    const src = await kaynakDocx({ docRels: [{ id: "rId1", type: "http://ornek.com/x", target: "ham.qqq" }], parts: { "word/ham.qqq": "?" } });
    const { bytes } = await carryParts({ app: "belge", source: await OpcPackage.open(src), output: await minimalDocx() });
    expect(opcHatalari(bytes)).toEqual([]);
    expect((await OpcPackage.open(bytes)).contentType("word/ham.qqq")).toBe("application/octet-stream");
  });
});

describe("makro, özel özellikler, eşleme", () => {
  const docm = () =>
    kaynakDocx({
      mainCt: CT_MAIN_DOCM,
      docRels: [{ id: "rId9", type: MST("2006/relationships/vbaProject"), target: "vbaProject.bin" }],
      parts: {
        "word/vbaProject.bin": new Uint8Array([0xd0, 0xcf, 0x11, 0xe0]),
        "word/_rels/vbaProject.bin.rels": rels([{ id: "rId1", type: MST("2006/relationships/wordVbaData"), target: "vbaData.xml" }]),
        "word/vbaData.xml": "<wne:vbaSuppData/>",
      },
      defaults: { bin: "application/vnd.ms-office.vbaProject" },
      overrides: { "word/vbaData.xml": "application/vnd.ms-word.vbaData+xml" },
    });

  it("makrolu hedef: vbaProject kapanışıyla taşınır, ana parça türü makrolu olur", async () => {
    const src = await OpcPackage.open(await docm());
    expect(src.macroEnabled).toBe(true);
    const { bytes, report } = await carryParts({ app: "belge", source: src, output: await minimalDocx(), macro: true });
    expect(report.macro).toBe(true);
    expect(report.carried.map((c) => c.from)).toEqual(["word/vbaProject.bin", "word/vbaData.xml"]);
    expect(opcHatalari(bytes, ".docm")).toEqual([]);
    const pkg = await OpcPackage.open(bytes);
    expect(pkg.mainContentType).toBe(CT_MAIN_DOCM);
    expect(pkg.contentType("word/vbaProject.bin")).toBe("application/vnd.ms-office.vbaProject");
    expect(pkg.macroEnabled).toBe(true);
  });

  it("makrosuz hedef: makro projesi 'macro' gerekçesiyle düşürülür, tür değişmez", async () => {
    const { bytes, report } = await carryParts({ app: "belge", source: await OpcPackage.open(await docm()), output: await minimalDocx(), macro: false });
    expect(report.dropped.map((d) => [d.reason, d.part])).toEqual([["macro", "word/vbaProject.bin"]]);
    expect(report.macro).toBe(false);
    expect(opcHatalari(bytes, ".docx")).toEqual([]);
    expect((await OpcPackage.open(bytes)).mainContentType).toBe(CT_MAIN_DOCX);
  });

  const custom = (props: [string, string][]) =>
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">${props
      .map(([n, v], i) => `<property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="${i + 2}" name="${n}"><vt:lpwstr>${v}</vt:lpwstr></property>`)
      .join("")}</Properties>`;
  const kaynakOzel = () =>
    kaynakDocx({
      rootRels: [{ id: "rId3", type: T("custom-properties"), target: "docProps/custom.xml" }],
      parts: { "docProps/custom.xml": custom([["OfficeHereBelge", "eski"], ["Proje", "Kırmızı"], ["Sürüm", "3"]]) },
      overrides: { "docProps/custom.xml": CT_CUSTOM },
    });

  it("özel özellikler: kullanıcınınkiler yazıcınınkilerle birleşir, yazıcının adları kaynakla ezilmez", async () => {
    const out = await minimalDocx({ "docProps/custom.xml": custom([["OfficeHereBelge", "yeni"], ["Proje", "Mavi"]]) }, { rootRels: [{ id: "rId4", type: T("custom-properties"), target: "docProps/custom.xml" }], overrides: { "docProps/custom.xml": CT_CUSTOM } });
    const { bytes, report } = await carryParts({ app: "belge", source: await OpcPackage.open(await kaynakOzel()), output: out });
    expect(report.customProperties).toBe(1);
    expect(opcHatalari(bytes)).toEqual([]);
    const x = (await text(bytes, "docProps/custom.xml"))!;
    const props = [...x.matchAll(/pid="(\d+)" name="([^"]+)"><vt:lpwstr>([^<]*)</g)].map((m) => [m[1], m[2], m[3]]);
    expect(props).toEqual([
      ["2", "OfficeHereBelge", "yeni"],
      ["3", "Proje", "Mavi"],
      ["4", "Sürüm", "3"],
    ]);
    expect(x).toContain('xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"');
  });

  it("özel özellikler: yazıcıda custom.xml yoksa kullanıcı özellikleriyle yeni parça kurulur", async () => {
    const { bytes, report } = await carryParts({ app: "belge", source: await OpcPackage.open(await kaynakOzel()), output: await minimalDocx() });
    expect(report.customProperties).toBe(2);
    expect(opcHatalari(bytes)).toEqual([]);
    const pkg = await OpcPackage.open(bytes);
    expect(pkg.contentType("docProps/custom.xml")).toBe(CT_CUSTOM);
    expect((await pkg.rels("")).some((r) => r.part === "docProps/custom.xml")).toBe(true);
    const x = (await pkg.text("docProps/custom.xml"))!;
    expect(x).not.toContain("OfficeHereBelge");
    expect(x).toContain('pid="2" name="Proje"');
    expect(x).toContain('pid="3" name="Sürüm"');
  });

  it("partMap: eşlenen parçanın örtük ilişkileri taşınır; kapanıştaki eşlemesiz yapısal başvuru atılır", async () => {
    const CT_WB = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml";
    const CT_WS = "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml";
    const src = await zipOf({
      "[Content_Types].xml": contentTypes(BASE_DEFAULTS, { "xl/workbook.xml": CT_WB, "xl/worksheets/sheet1.xml": CT_WS, "xl/worksheets/sheet2.xml": CT_WS }),
      "_rels/.rels": rels([{ id: "rId1", type: T("officeDocument"), target: "xl/workbook.xml" }]),
      "xl/workbook.xml": `<workbook xmlns:r="${R}"><sheets><sheet r:id="rId1"/><sheet r:id="rId2"/></sheets></workbook>`,
      "xl/_rels/workbook.xml.rels": rels([
        { id: "rId1", type: T("worksheet"), target: "worksheets/sheet1.xml" },
        { id: "rId2", type: T("worksheet"), target: "worksheets/sheet2.xml" },
      ]),
      "xl/worksheets/sheet1.xml": "<worksheet/>",
      "xl/worksheets/sheet2.xml": "<worksheet/>",
      "xl/worksheets/_rels/sheet1.xml.rels": rels([{ id: "rId1", type: "http://ornek.com/sayfaEki", target: "../ek/ek1.xml" }]),
      "xl/ek/ek1.xml": "<ek/>",
      "xl/ek/_rels/ek1.xml.rels": rels([
        { id: "rId1", type: T("worksheet"), target: "../worksheets/sheet1.xml" },
        { id: "rId2", type: T("worksheet"), target: "../worksheets/sheet2.xml" },
      ]),
    });
    const out = await zipOf({
      "[Content_Types].xml": contentTypes(BASE_DEFAULTS, { "xl/workbook.xml": CT_WB, "xl/worksheets/sayfa_a.xml": CT_WS }),
      "_rels/.rels": rels([{ id: "rId1", type: T("officeDocument"), target: "xl/workbook.xml" }]),
      "xl/workbook.xml": `<workbook xmlns:r="${R}"><sheets><sheet r:id="rId1"/></sheets></workbook>`,
      "xl/_rels/workbook.xml.rels": rels([{ id: "rId1", type: T("worksheet"), target: "worksheets/sayfa_a.xml" }]),
      "xl/worksheets/sayfa_a.xml": "<worksheet/>",
    });
    const { bytes, report } = await carryParts({ app: "tablo", source: await OpcPackage.open(src), output: out, partMap: { "xl/worksheets/sheet1.xml": "xl/worksheets/sayfa_a.xml" } });
    expect(report.carried.map((c) => c.from)).toEqual(["xl/ek/ek1.xml"]);
    expect(report.dropped.map((d) => [d.reason, d.part])).toEqual([["structural", "xl/worksheets/sheet2.xml"]]);
    expect(opcHatalari(bytes, ".xlsx")).toEqual([]);
    expect((await relsOf(bytes, "xl/worksheets/sayfa_a.xml")).map((r) => r.part)).toEqual(["xl/ek/ek1.xml"]);
    expect((await relsOf(bytes, "xl/ek/ek1.xml")).map((r) => r.part)).toEqual(["xl/worksheets/sayfa_a.xml"]);
  });

  it("kaynak kimliği kayıtta yoksa çıktı bozulmaz, rapora yazılır", async () => {
    const out = await minimalDocx();
    const { bytes, report } = await carryParts({ app: "belge", source: "kaynak-yok", output: out, registry: new SourceRegistry() });
    expect(report.sourceMissing).toBe(true);
    expect(opcHatalari(bytes)).toEqual([]);
    expect((await OpcPackage.open(bytes)).names.sort()).toEqual((await OpcPackage.open(out)).names.sort());
  });
});

describe("öğe düzeyi taşıyıcı (Carry)", () => {
  const OPAK = `<w:r><w:object><o:OLEObject xmlns:o="urn:schemas-microsoft-com:office:office" r:id="rId7"/></w:object><w:drawing><a:blip r:embed="rId8"/><dgm:relIds r:dm="rId9"/><a:hlinkClick r:id="rId10"/></w:drawing></w:r>`;
  const kaynak = () =>
    kaynakDocx({
      body: `<w:p>${OPAK}</w:p>`,
      docRels: [
        { id: "rId7", type: T("oleObject"), target: "embeddings/ole1.bin" },
        { id: "rId8", type: T("image"), target: "media/oh_t3_onizleme.emf" },
        { id: "rId9", type: T("diagramData"), target: "diagrams/data1.xml" },
        { id: "rId10", type: T("hyperlink"), target: "https://ornek.com", external: true },
      ],
      parts: {
        "word/embeddings/ole1.bin": new Uint8Array([1, 2, 3]),
        "word/media/oh_t3_onizleme.emf": new Uint8Array([4]),
        "word/media/kullanilmayan.png": new Uint8Array([5]),
        "word/diagrams/data1.xml": `<dgm:dataModel><dsp:dataModelExt relId="ESKI"/></dgm:dataModel>`,
        "word/diagrams/_rels/data1.xml.rels": rels([
          { id: "rId1", type: T("image"), target: "../media/oh_t3_onizleme.emf" },
          { id: "rId2", type: T("image"), target: "../media/ic.png" },
        ]),
        "word/media/ic.png": new Uint8Array([6]),
      },
      defaults: { bin: "application/vnd.openxmlformats-officedocument.oleObject", emf: "image/x-emf", png: "image/png" },
      overrides: { "word/diagrams/data1.xml": "application/vnd.openxmlformats-officedocument.drawingml.diagramData+xml" },
    });

  it("ilişkiler yeniden kurulur, kopyalar oh_tN_ alır, kapanış ve yama uygulanır, ulaşılamayan kopya atılır", async () => {
    const registry = new SourceRegistry();
    const src = await OpcPackage.open(await kaynak());
    const id = registry.register(src, { name: "kaynak.docx" });
    // okurken: opak XML'in ilişkileri saklanır
    const refs = await collectRefs(src, "word/document.xml", OPAK, { prefixes: ["r"] });
    expect(refs.map((r) => r.id).sort()).toEqual(["rId10", "rId7", "rId8", "rId9"]);
    // yazarken: yazıcı kendi .rels'ini kurar
    const carry = new Carry(registry);
    const inst = carry.instance(id);
    const docRels = new RelsBuilder("word/document.xml", parseRels(rels([{ id: "rId1", type: T("styles"), target: "styles.xml" }]), "word/document.xml"));
    const { xml, ids, missing } = reattach(OPAK, refs, inst, (r) => (r.external ? docRels.addExternal(r.type, r.target) : docRels.addPart(r.type, r.target)));
    expect(missing).toEqual([]);
    expect(ids.get("rId7")).not.toBe("rId7");
    const dataTo = inst.mapped("word/diagrams/data1.xml")!;
    expect(dataTo).toMatch(/^word\/diagrams\/oh_t\d+_data1\.xml$/);
    // önek birikmez: kaynakta zaten oh_t3_ olan ad yeniden tek önek alır
    expect(inst.mapped("word/media/oh_t3_onizleme.emf")).toMatch(/^word\/media\/oh_t\d+_onizleme\.emf$/);
    carry.patch(dataTo, (s) => s.replace('relId="ESKI"', 'relId="YENI"'));
    // yazıcının hiç bağlamadığı kopya: çöp toplanmalı
    expect(inst.copy("word/media/kullanilmayan.png")).toBeTruthy();
    expect(inst.copy("word/media/olmayan.png")).toBeNull();
    const out = await minimalDocx({ "word/document.xml": DOC_XML(`<w:p>${xml}</w:p>`), "word/_rels/document.xml.rels": docRels.xml() });
    const { bytes, report } = await carryParts({ app: "belge", source: id, output: out, carry, registry });
    expect(report.collected).toBe(1);
    expect(report.carried.filter((c) => c.kind === "element").map((c) => c.from).sort()).toEqual(["word/diagrams/data1.xml", "word/embeddings/ole1.bin", "word/media/oh_t3_onizleme.emf"]);
    // kapanış: data1 → önizleme (aynı örneğin kopyasına bağlanır) ve ic.png (öğe düzeyinde önekli)
    expect(report.carried.filter((c) => c.kind === "closure").map((c) => c.to)).toEqual([expect.stringMatching(/^word\/media\/oh_t\d+_ic\.png$/)]);
    expect(opcHatalari(bytes)).toEqual([]);
    const pkg = await OpcPackage.open(bytes);
    expect(await pkg.text(dataTo)).toContain('relId="YENI"');
    expect(pkg.names.some((n) => n.includes("kullanilmayan"))).toBe(false);
    const dr = await pkg.rels(dataTo);
    expect(dr[0].part).toBe(inst.mapped("word/media/oh_t3_onizleme.emf"));
    // belge XML'indeki kimlikler yeni .rels'te çözülüyor
    const d = await pkg.rels("word/document.xml");
    for (const nid of ids.values()) expect(d.some((r) => r.id === nid)).toBe(true);
  });

  it("örnekler ayrı kopya alır; kayıtta olmayan kaynakta kopya null", async () => {
    const registry = new SourceRegistry();
    const id = registry.register(await OpcPackage.open(await kaynak()));
    const carry = new Carry(registry);
    const a = carry.instance(id).copy("word/embeddings/ole1.bin");
    const b = carry.instance(id).copy("word/embeddings/ole1.bin");
    expect(a).not.toBe(b);
    expect(carry.size).toBe(2);
    expect(carry.instance("kaynak-yok").copy("word/embeddings/ole1.bin")).toBeNull();
    expect(carry.instance(id).copy("[Content_Types].xml")).toBeNull();
    expect(carry.instance(id).copy("word/_rels/document.xml.rels")).toBeNull();
  });

  it("öğe düzeyi kopya adı yazıcının parçasıyla çakışırsa yazıcınınki geçerli ('conflict')", async () => {
    const registry = new SourceRegistry();
    const id = registry.register(await OpcPackage.open(await kaynak()));
    const carry = new Carry(registry);
    const to = carry.instance(id).copy("word/embeddings/ole1.bin")!;
    const out = await minimalDocx({ [to]: new Uint8Array([7]) }, { docRels: [{ id: "rId5", type: T("oleObject"), target: to.slice(5) }], overrides: { [to]: "application/x-y" } });
    const { bytes, report } = await carryParts({ app: "belge", source: id, output: out, carry, registry });
    expect(report.dropped.map((d) => d.reason)).toContain("conflict");
    expect(Array.from((await (await OpcPackage.open(bytes)).data(to))!)).toEqual([7]);
  });
});

describe("POI dosyaları", () => {
  const uz = (f: string) => f.slice(f.lastIndexOf("."));
  const appOf = (f: string) => (/\.do[ct][xm]$/.test(f) ? "belge" : /\.xl[st][xm]$/.test(f) ? "tablo" : "sunum");

  it("paket kendine taşınınca hiçbir şey eklenmez, yeni paket hatası çıkmaz", async () => {
    for (const f of poiDosyalari(/\.(docx|docm|xlsx|xlsm|pptx|pptm)$/)) {
      const b = poi(f);
      const src = await OpcPackage.open(b);
      const { bytes, report } = await carryParts({ app: appOf(f), source: src, output: b, macro: src.macroEnabled });
      expect(report.carried, f).toEqual([]);
      const once = new Set(opcHatalari(b, uz(f)).map((h) => `${h.tur} ${h.ayrinti}`));
      const sonra = opcHatalari(bytes, uz(f)).filter((h) => h.tur !== "ct-ilk-değil" && !once.has(`${h.tur} ${h.ayrinti}`));
      expect(sonra, f).toEqual([]);
    }
  });

  it("docx: en küçük yazıcı çıktısına taşıma paket hatasız; imza düşer, makro korunur", async () => {
    for (const f of poiDosyalari(/\.(docx|docm)$/)) {
      const src = await OpcPackage.open(poi(f));
      const macro = /\.docm$/.test(f);
      const { bytes, report } = await carryParts({ app: "belge", source: src, output: await minimalDocx(), macro });
      // Kaynaktaki kırık kimlikler (taşınan parçanın kendi XML'inde) yeni hata sayılmaz
      const srcErr = new Set(opcHatalari(poi(f), uz(f)).map((h) => h.tur));
      expect(opcHatalari(bytes, macro ? ".docm" : ".docx").filter((h) => !srcErr.has(h.tur)), f).toEqual([]);
      const pkg = await OpcPackage.open(bytes);
      expect(pkg.signed, f).toBe(false);
      if (src.signed) expect(report.dropped.some((d) => d.reason === "signature"), f).toBe(true);
      if (macro) expect(pkg.macroEnabled, f).toBe(true);
      expect(report.ms, f).toBeLessThan(1000);
    }
  });
});
