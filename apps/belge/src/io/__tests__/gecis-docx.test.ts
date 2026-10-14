// @vitest-environment jsdom
// Geçiş katmanı, paket düzeyi (P3.1): açılan .docx'in modelin tanımadığı parçaları kaydederken taşınır
// (tema, fontTable + gömülü yazı tipi, webSettings, glossary, customXml, klavye özelleştirmesi, .docm'de
// makro projesi), settings.xml baştan yazılmaz (yama), custom.xml birleşir; dijital imza ve açıklama yanıt
// zinciri düşer ve açılışta uyarılır. Kaynak paket yoksa (kurtarma) yalnız yazıcının çıktısı yazılır.
import { describe, expect, it } from "vitest";
import JSZip from "jszip";
import { OpcPackage, getSource, type CarryReport } from "@officehere/core/opc";
import { openDocument, readDocx, writeDocx, writeOdt } from "../index";
import { defaultSettings, type BelgeDocument } from "../../model/document";
import { MACROS_REMOVED, lossMessage, saveNotes } from "../warnings";

const buf = (u: Uint8Array) => u.buffer.slice(u.byteOffset, u.byteOffset + u.byteLength) as ArrayBuffer;
const R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const PKG = "http://schemas.openxmlformats.org/package/2006/relationships";
const MS = "http://schemas.microsoft.com/office";
const CT_WML = "application/vnd.openxmlformats-officedocument.wordprocessingml";

function simpleDoc(text = "Merhaba dünya"): BelgeDocument {
  return { content: { type: "doc", content: [{ type: "section", content: [{ type: "paragraph", content: [{ type: "text", text }] }] }] }, settings: defaultSettings() };
}

const SETTINGS =
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:settings xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:r="${R}" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" xmlns:w="${W}" xmlns:w14="${MS}/word/2010/wordml" mc:Ignorable="w14">` +
  `<w:zoom w:percent="90"/><w:proofState w:spelling="clean"/><w:attachedTemplate r:id="rId1"/><w:mailMerge><w:mainDocumentType w:val="formLetters"/><w:dataSource r:id="rId7"/></w:mailMerge>` +
  `<w:documentProtection w:edit="forms" w:enforcement="1"/><w:defaultTabStop w:val="720"/><w:characterSpacingControl w:val="doNotCompress"/>` +
  `<w:footnotePr><w:pos w:val="beneathText"/><w:footnote w:id="0"/><w:footnote w:id="1"/></w:footnotePr>` +
  `<w:compat><w:cachedColBalance/><w:splitPgBreakAndParaMark/><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="14"/></w:compat>` +
  `<w:docVars><w:docVar w:name="Musteri" w:val="ACME"/></w:docVars><w:rsids><w:rsidRoot w:val="00AB12CD"/><w:rsid w:val="00AB12CD"/></w:rsids>` +
  `<m:mathPr><m:mathFont m:val="Cambria Math"/></m:mathPr><w:uiCompat97To2003/><w:themeFontLang w:val="en-US"/><w14:docId w14:val="1A2B3C4D"/></w:settings>`;

const rel = (id: string, type: string, target: string, external = false) => `<Relationship Id="${id}" Type="${type}" Target="${target}"${external ? ' TargetMode="External"' : ""}/>`;
const addRels = (xml: string, rels: string) => xml.replace("</Relationships>", `${rels}</Relationships>`);
const addTypes = (xml: string, types: string) => xml.replace("</Types>", `${types}</Types>`);
const ov = (part: string, type: string) => `<Override PartName="/${part}" ContentType="${type}"/>`;

/**
 * Office'in ürettiği bir belgeyi taklit eder: yazıcımızın çıktısına modelin tanımadığı parçalar eklenir.
 * `macro`: makrolu belge (.docm: vbaProject + vbaData, makrolu ana tür).
 */
async function zengin(opts: { macro?: boolean } = {}): Promise<Uint8Array> {
  const zip = await JSZip.loadAsync(await writeDocx(simpleDoc()));
  const text = async (p: string) => (await zip.file(p)!.async("string"))!;
  zip.file("word/theme/theme1.xml", `<?xml version="1.0"?><a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Ofis Teması"><a:themeElements/></a:theme>`);
  zip.file("word/fontTable.xml", `<?xml version="1.0"?><w:fonts xmlns:w="${W}" xmlns:r="${R}"><w:font w:name="Gömülü"><w:embedRegular r:id="rId1" w:fontKey="{00000000-0000-0000-0000-000000000000}"/></w:font></w:fonts>`);
  zip.file("word/_rels/fontTable.xml.rels", `<?xml version="1.0"?><Relationships xmlns="${PKG}">${rel("rId1", `${R}/font`, "fonts/font1.odttf")}</Relationships>`);
  zip.file("word/fonts/font1.odttf", new Uint8Array([1, 2, 3, 4, 5]));
  zip.file("word/webSettings.xml", `<?xml version="1.0"?><w:webSettings xmlns:w="${W}"><w:optimizeForBrowser/></w:webSettings>`);
  zip.file("word/glossary/document.xml", `<?xml version="1.0"?><w:glossaryDocument xmlns:w="${W}"><w:docParts/></w:glossaryDocument>`);
  zip.file("word/customizations.xml", `<?xml version="1.0"?><wne:tcg xmlns:wne="${MS}/word/2006/wordml"/>`);
  zip.file("customXml/item1.xml", `<?xml version="1.0"?><x:veri xmlns:x="urn:ornek:veri"><x:a>1</x:a></x:veri>`);
  zip.file("customXml/itemProps1.xml", `<?xml version="1.0"?><ds:datastoreItem ds:itemID="{11111111-1111-1111-1111-111111111111}" xmlns:ds="http://schemas.openxmlformats.org/officeDocument/2006/customXml"/>`);
  zip.file("customXml/_rels/item1.xml.rels", `<?xml version="1.0"?><Relationships xmlns="${PKG}">${rel("rId1", `${R}/customXmlProps`, "itemProps1.xml")}</Relationships>`);
  zip.file("word/commentsExtended.xml", `<?xml version="1.0"?><w15:commentsEx xmlns:w15="${MS}/word/2012/wordml"/>`);
  zip.file("word/people.xml", `<?xml version="1.0"?><w15:people xmlns:w15="${MS}/word/2012/wordml"/>`);
  zip.file("word/header9.xml", `<?xml version="1.0"?><w:hdr xmlns:w="${W}"><w:p/></w:hdr>`);
  zip.file("word/footnotes.xml", `<?xml version="1.0"?><w:footnotes xmlns:w="${W}"><w:footnote w:type="separator" w:id="0"><w:p><w:r><w:separator/></w:r></w:p></w:footnote><w:footnote w:type="continuationSeparator" w:id="1"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote></w:footnotes>`);
  zip.file("word/settings.xml", SETTINGS);
  zip.file("word/_rels/settings.xml.rels", `<?xml version="1.0"?><Relationships xmlns="${PKG}">${rel("rId1", `${R}/attachedTemplate`, "file:///C:/Sablonlar/Normal.dotm", true)}</Relationships>`);
  zip.file("_xmlsignatures/origin.sigs", "");
  zip.file("_xmlsignatures/sig1.xml", `<?xml version="1.0"?><Signature xmlns="http://www.w3.org/2000/09/xmldsig#"/>`);
  zip.file("_xmlsignatures/_rels/origin.sigs.rels", `<?xml version="1.0"?><Relationships xmlns="${PKG}">${rel("rId1", `${PKG}/digital-signature/signature`, "sig1.xml")}</Relationships>`);
  let docRels = addRels(
    await text("word/_rels/document.xml.rels"),
    rel("rId901", `${R}/theme`, "theme/theme1.xml") +
      rel("rId902", `${R}/fontTable`, "fontTable.xml") +
      rel("rId903", `${R}/webSettings`, "webSettings.xml") +
      rel("rId904", `${R}/glossaryDocument`, "glossary/document.xml") +
      rel("rId905", `${R}/customXml`, "../customXml/item1.xml") +
      rel("rId906", `${MS}/2011/relationships/commentsExtended`, "commentsExtended.xml") +
      rel("rId907", `${MS}/2011/relationships/people`, "people.xml") +
      rel("rId908", `${R}/header`, "header9.xml") +
      rel("rId909", `${R}/footnotes`, "footnotes.xml") +
      rel("rId910", `${MS}/2006/relationships/keyMapCustomizations`, "customizations.xml"),
  );
  let types =
    ov("word/theme/theme1.xml", "application/vnd.openxmlformats-officedocument.theme+xml") +
    ov("word/fontTable.xml", `${CT_WML}.fontTable+xml`) +
    '<Default Extension="odttf" ContentType="application/vnd.openxmlformats-officedocument.obfuscatedFont"/>' +
    ov("word/webSettings.xml", `${CT_WML}.webSettings+xml`) +
    ov("word/glossary/document.xml", `${CT_WML}.document.glossary+xml`) +
    ov("word/customizations.xml", "application/vnd.ms-word.keyMapCustomizations+xml") +
    ov("customXml/itemProps1.xml", "application/vnd.openxmlformats-officedocument.customXmlProperties+xml") +
    ov("word/commentsExtended.xml", `${CT_WML}.commentsExtended+xml`) +
    ov("word/people.xml", `${CT_WML}.people+xml`) +
    ov("word/header9.xml", `${CT_WML}.header+xml`) +
    ov("word/footnotes.xml", `${CT_WML}.footnotes+xml`) +
    '<Default Extension="sigs" ContentType="application/vnd.openxmlformats-package.digital-signature-origin"/>' +
    ov("_xmlsignatures/sig1.xml", "application/vnd.openxmlformats-package.digital-signature-xmlsignature+xml");
  let ct = await text("[Content_Types].xml");
  if (opts.macro) {
    zip.file("word/vbaProject.bin", new Uint8Array([0xd0, 0xcf, 0x11, 0xe0, 9, 9, 9]));
    zip.file("word/vbaData.xml", `<?xml version="1.0"?><wne:vbaSuppData xmlns:wne="${MS}/word/2006/wordml"/>`);
    zip.file("word/_rels/vbaProject.bin.rels", `<?xml version="1.0"?><Relationships xmlns="${PKG}">${rel("rId1", `${MS}/2006/relationships/wordVbaData`, "vbaData.xml")}</Relationships>`);
    docRels = addRels(docRels, rel("rId911", `${MS}/2006/relationships/vbaProject`, "vbaProject.bin"));
    types += '<Default Extension="bin" ContentType="application/vnd.ms-office.vbaProject"/>' + ov("word/vbaData.xml", "application/vnd.ms-word.vbaData+xml");
    ct = ct.replace(`${CT_WML}.document.main+xml`, "application/vnd.ms-word.document.macroEnabled.main+xml");
  }
  zip.file("word/_rels/document.xml.rels", docRels);
  zip.file("[Content_Types].xml", addTypes(ct, types));
  zip.file("_rels/.rels", addRels(await text("_rels/.rels"), rel("rId90", `${PKG}/digital-signature/origin`, "_xmlsignatures/origin.sigs")));
  zip.file("word/document.xml", (await text("word/document.xml")).replace("<w:sectPr>", '<w:sectPr><w:headerReference w:type="first" r:id="rId908"/>'));
  zip.file(
    "docProps/custom.xml",
    (await text("docProps/custom.xml")).replace(
      "</Properties>",
      '<property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="3" name="Proje"><vt:lpwstr>Köprü</vt:lpwstr></property></Properties>',
    ),
  );
  return zip.generateAsync({ type: "uint8array", compression: "DEFLATE" });
}

/** Paket tutarlılığı: her parçanın içerik türü var, iç ilişki hedefleri var. Hatalar döner. */
async function opcErrors(bytes: Uint8Array): Promise<string[]> {
  const pkg = await OpcPackage.open(bytes);
  const out: string[] = [];
  for (const n of pkg.names) {
    if (n === "[Content_Types].xml" || n.endsWith("/")) continue;
    if (!pkg.contentType(n)) out.push(`içerik türü yok: ${n}`);
    if (!n.endsWith(".rels")) for (const r of await pkg.rels(n)) if (!r.external && (!r.part || !pkg.has(r.part))) out.push(`hedef yok: ${n} → ${r.target}`);
  }
  for (const r of await pkg.rels("")) if (!r.external && (!r.part || !pkg.has(r.part))) out.push(`hedef yok: kök → ${r.target}`);
  return out;
}

async function open(bytes: Uint8Array, name = "zengin.docx") {
  return openDocument(name, buf(bytes));
}

describe("geçiş katmanı: kaynak paket kaydı", () => {
  it("açılan .docx kaynak kaydına girer; modelde yalnız kimlik (SHA-256) tutulur", async () => {
    const src = await zengin();
    const r = await open(src);
    expect(r.doc.settings.kaynak).toMatch(/^[0-9a-f]{64}$/);
    expect(getSource(r.doc.settings.kaynak)?.bytes.byteLength).toBe(src.byteLength);
    // Doğrulama okuması kaynak kaydetmez.
    const v = await readDocx(buf(await writeDocx(simpleDoc("başka"))), undefined, { kaynak: false });
    expect(v.settings.kaynak).toBeUndefined();
  });

  it("dijital imza ve açıklama yanıt zinciri açılışta uyarılır", async () => {
    const r = await open(await zengin());
    expect(r.warnings).toContain(lossMessage("signature"));
    expect(r.warnings).toContain(lossMessage("commentThreads"));
    expect(r.warnings[0]).toBe(lossMessage("signature"));
  });
});

describe("geçiş katmanı: kaydederken taşıma", () => {
  it("tema, yazı tipleri, webSettings, glossary, customXml ve klavye özelleştirmesi taşınır; imza ve yanıt zinciri düşer", async () => {
    const r = await open(await zengin());
    let report: CarryReport | undefined;
    const out = await writeDocx(r.doc, { onReport: (x) => (report = x) });
    const z = await JSZip.loadAsync(out);
    expect(Object.keys(z.files)[0]).toBe("[Content_Types].xml");
    for (const p of ["word/theme/theme1.xml", "word/fontTable.xml", "word/fonts/font1.odttf", "word/webSettings.xml", "word/glossary/document.xml", "customXml/item1.xml", "customXml/itemProps1.xml", "word/customizations.xml"])
      expect(z.file(p), p).toBeTruthy();
    for (const p of ["_xmlsignatures/origin.sigs", "_xmlsignatures/sig1.xml", "word/commentsExtended.xml", "word/people.xml"]) expect(z.file(p), p).toBeNull();
    expect(new Set(report!.dropped.map((d) => d.reason))).toEqual(new Set(["signature", "comments"]));
    expect(report!.sourceMissing).toBe(false);
    expect(await z.file("word/_rels/document.xml.rels")!.async("string")).toMatch(/relationships\/theme" Target="theme\/theme1.xml"/);
    expect(await opcErrors(out)).toEqual([]);
    expect(saveNotes(report)).toEqual([]);
  });

  it("custom.xml birleşir: kullanıcı özelliği korunur, OfficeHere ayarı yazıcınınki", async () => {
    const r = await open(await zengin());
    const z = await JSZip.loadAsync(await writeDocx(r.doc));
    const custom = await z.file("docProps/custom.xml")!.async("string");
    expect(custom).toContain('name="Proje"');
    expect(custom).toContain("Köprü");
    expect(custom.match(/name="OfficeHereBelge"/g)).toHaveLength(1);
  });

  it("settings.xml yamalanır: compat, rsid, docVars, attachedTemplate, mathPr, defaultTabStop özgünden; sahip öğeler modelden", async () => {
    const r = await open(await zengin());
    r.doc.settings.oddEven = true;
    const z = await JSZip.loadAsync(await writeDocx(r.doc));
    const s = await z.file("word/settings.xml")!.async("string");
    for (const part of ['<w:zoom w:percent="90"/>', '<w:defaultTabStop w:val="720"/>', '<w:docVar w:name="Musteri" w:val="ACME"/>', '<w:rsidRoot w:val="00AB12CD"/>', "<m:mathPr>", '<w14:docId w14:val="1A2B3C4D"/>', '<w:themeFontLang w:val="en-US"/>'])
      expect(s, part).toContain(part);
    expect(s).toContain('w:val="14"/>'); // uyumluluk kipi özgünden
    // Model okumadığı kısıtlamayı (form) bozmaz; tek/çift üst bilgi modelden.
    expect(s).toContain('<w:documentProtection w:edit="forms" w:enforcement="1"/>');
    expect(s).toContain("<w:evenAndOddHeaders/>");
    // Dipnot numaralama ayarı korunur, ayraç başvuruları yazıcınınki.
    expect(s).toContain('<w:footnotePr><w:pos w:val="beneathText"/><w:footnote w:id="-1"/><w:footnote w:id="0"/></w:footnotePr>');
    // Şema sırası: uyumluluk ayarları dizilir, şemada olmayan ilk sürüm ayarı yazılmaz.
    expect(s).toMatch(/<w:splitPgBreakAndParaMark\/><w:cachedColBalance\/><w:compatSetting/);
    expect(s).not.toContain("uiCompat97To2003");
    expect(s.indexOf("<w:defaultTabStop")).toBeLessThan(s.indexOf("<w:evenAndOddHeaders"));
    // Şablon ilişkisi yeniden kurulur; karşılığı olmayan ilişkiye başvuran öğe (adres mektup kaynağı rId7) atılır.
    const sr = await z.file("word/_rels/settings.xml.rels")!.async("string");
    expect(sr).toContain('Target="file:///C:/Sablonlar/Normal.dotm" TargetMode="External"');
    const id = /attachedTemplate r:id="([^"]+)"/.exec(s)?.[1];
    expect(sr).toContain(`Id="${id}"`);
    expect(s).not.toContain("<w:mailMerge>");
  });

  it("şemaya aykırı yakınlaştırma (yüzdesiz) yüzdeyle yazılır", async () => {
    const zip = await JSZip.loadAsync(await zengin());
    zip.file("word/settings.xml", SETTINGS.replace('<w:zoom w:percent="90"/>', '<w:zoom w:val="bestFit"/>'));
    const r = await open(await zip.generateAsync({ type: "uint8array" }));
    const s = await (await JSZip.loadAsync(await writeDocx(r.doc))).file("word/settings.xml")!.async("string");
    expect(s).toContain('<w:zoom w:val="bestFit" w:percent="100"/>');
  });

  it("model kısıtlamayı değiştirince yazılan kısıtlama modelinki olur", async () => {
    const r = await open(await zengin());
    r.doc.settings.protection = { mode: "readOnly", enforced: true };
    const s = await (await JSZip.loadAsync(await writeDocx(r.doc))).file("word/settings.xml")!.async("string");
    expect(s).toContain('<w:documentProtection w:edit="readOnly" w:enforcement="1"/>');
    expect(s).not.toContain('w:edit="forms"');
  });

  it("özgünün ilk bölümündeki boş üst bilgi türü ve dipnot ayraçları korunur; ikinci kayıt kararlı", async () => {
    const r = await open(await zengin());
    const out = await writeDocx(r.doc);
    const z = await JSZip.loadAsync(out);
    const doc = await z.file("word/document.xml")!.async("string");
    expect(doc).toMatch(/<w:headerReference w:type="first" r:id="rId\d+"\/>/);
    const notes = await z.file("word/footnotes.xml")!.async("string");
    expect(notes).toContain('w:type="separator" w:id="-1"');
    expect(notes).not.toMatch(/<w:footnote w:id="1">/);
    const names = (zz: JSZip) => Object.keys(zz.files).filter((n) => !zz.files[n].dir).sort();
    const r2 = await open(out, "ikinci.docx");
    const z2 = await JSZip.loadAsync(await writeDocx(r2.doc));
    expect(names(z2)).toEqual(names(z));
    for (const p of ["word/settings.xml", "[Content_Types].xml", "word/_rels/document.xml.rels", "docProps/custom.xml"]) expect(await z2.file(p)!.async("string"), p).toBe(await z.file(p)!.async("string"));
  });

  it(".docm: makrolu hedefte makro projesi korunur; .docx hedefinde düşer ve kaydederken bildirilir", async () => {
    const r = await open(await zengin({ macro: true }), "makro.docm");
    expect(r.warnings).toContain(lossMessage("macros"));
    let report: CarryReport | undefined;
    const docm = await JSZip.loadAsync(await writeDocx(r.doc, { macro: true, onReport: (x) => (report = x) }));
    expect(docm.file("word/vbaProject.bin")).toBeTruthy();
    expect(docm.file("word/vbaData.xml")).toBeTruthy();
    expect(await docm.file("[Content_Types].xml")!.async("string")).toContain('PartName="/word/document.xml" ContentType="application/vnd.ms-word.document.macroEnabled.main+xml"');
    expect(report!.macro).toBe(true);
    expect(saveNotes(report)).toEqual([]);
    const docxBytes = await writeDocx(r.doc, { onReport: (x) => (report = x) });
    const docx = await JSZip.loadAsync(docxBytes);
    expect(docx.file("word/vbaProject.bin")).toBeNull();
    // Makrolu kaynaktaki klavye özelleştirmesi makrolarla birlikte gider.
    expect(docx.file("word/customizations.xml")).toBeNull();
    expect(await docx.file("[Content_Types].xml")!.async("string")).not.toContain("macroEnabled");
    expect(saveNotes(report)).toEqual([MACROS_REMOVED]);
    expect(await opcErrors(docxBytes)).toEqual([]);
  });

  it("kaynak paket yoksa (kurtarma deposunda bulunamadı) yazıcının çıktısı yazılır ve raporlanır", async () => {
    const doc = simpleDoc();
    doc.settings.kaynak = "f".repeat(64);
    let report: CarryReport | undefined;
    const out = await writeDocx(doc, { onReport: (x) => (report = x) });
    expect(report!.sourceMissing).toBe(true);
    expect(report!.carried).toEqual([]);
    expect((await readDocx(buf(out))).content).toBeTruthy();
    expect(await opcErrors(out)).toEqual([]);
  });

  it("yeni belgede taşıma yapılmaz; içerik türleri yine ilk parça", async () => {
    let called = false;
    const z = await JSZip.loadAsync(await writeDocx(simpleDoc(), { onReport: () => (called = true) }));
    expect(called).toBe(false);
    expect(Object.keys(z.files)[0]).toBe("[Content_Types].xml");
  });

  it("biçim değişince (.odt) taşıma yapılmaz", async () => {
    const r = await open(await zengin());
    const z = await JSZip.loadAsync(await writeOdt(r.doc));
    expect(Object.keys(z.files).some((n) => /theme|fontTable|customXml|glossary/.test(n))).toBe(false);
  });
});
