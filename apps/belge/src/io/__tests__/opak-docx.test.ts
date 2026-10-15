// @vitest-environment jsdom
// Geçiş katmanı, öğe düzeyi (P3.2): modelin karşılamadığı .docx öğeleri (gömülü OLE nesnesi, anlaşılmayan
// çizim / mc:AlternateContent, içe aktarılan içerik, bilinmeyen blok/paragraf/koşu öğeleri) opak düğüm olarak
// özgün XML'leri ve ilişkileriyle taşınır; kaydederken parçaları kopyalanır, belge genelinde benzersiz kimlikler
// yeniden numaralanır. Paragraf/koşu özelliklerinin okunmayan çocukları (ooxmlExtra) şema sırasıyla korunur.
import { readFileSync } from "node:fs";
import { join } from "node:path";
import { describe, expect, it } from "vitest";
import JSZip from "jszip";
import { generateHTML, generateJSON, type JSONContent } from "@tiptap/core";
import { openDocument, readDocx, validateContent, writeDocx, writeMarkdown, writeOdt, writeRtf } from "../index";
import { schemaExtensions } from "../../editor/extensions";
import { hasOpaque, opaquePlaceholder } from "../../editor/opaque";
import { topElements } from "../docx/partWriter";
import { defaultSettings, type BelgeDocument } from "../../model/document";
import { lossMessage } from "../warnings";
import { all, text } from "./bicim-yardimci";

const POI = join(__dirname, "../../../../../test-derlem/poi");
const poi = (name: string) => new Uint8Array(readFileSync(join(POI, name)));
const buf = (u: Uint8Array) => u.buffer.slice(u.byteOffset, u.byteOffset + u.byteLength) as ArrayBuffer;

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const MS = "http://schemas.microsoft.com/office";
const PNG = "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAQAAAADCAIAAAA7ljmRAAAAEElEQVR4nGM4IScHRww4OQD1xwwxB+utmQAAAABJRU5ErkJggg==";
const ROOT = `<w:document xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:r="${R}" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" xmlns:w="${W}" xmlns:w14="${MS}/word/2010/wordml" xmlns:w15="${MS}/word/2012/wordml" xmlns:wpg="${MS}/word/2010/wordprocessingGroup" xmlns:wps="${MS}/word/2010/wordprocessingShape" mc:Ignorable="w14 w15">`;

async function zipOf(bytes: Uint8Array) {
  return JSZip.loadAsync(bytes);
}
async function entry(bytes: Uint8Array, name: string): Promise<string | null> {
  return (await (await zipOf(bytes)).file(name)?.async("string")) ?? null;
}

/** Yazıcımızın çıktısından, gövdesi ve ek parçaları verilen bir .docx (Office'in yazdığı belgeyi taklit eder). */
async function docxWith(body: string, opts: { rels?: string; files?: Record<string, string | Uint8Array>; types?: string } = {}): Promise<Uint8Array> {
  const base = await writeDocx({ content: { type: "doc", content: [{ type: "section", content: [{ type: "paragraph" }] }] }, settings: defaultSettings() });
  const zip = await zipOf(base);
  zip.file("word/document.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n${ROOT}<w:body>${body}<w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr></w:body></w:document>`);
  if (opts.rels) zip.file("word/_rels/document.xml.rels", (await zip.file("word/_rels/document.xml.rels")!.async("string")).replace("</Relationships>", `${opts.rels}</Relationships>`));
  if (opts.types) zip.file("[Content_Types].xml", (await zip.file("[Content_Types].xml")!.async("string")).replace("</Types>", `${opts.types}</Types>`));
  for (const [k, v] of Object.entries(opts.files ?? {})) zip.file(k, v);
  return zip.generateAsync({ type: "uint8array" });
}

/** Belge ilişkileri: kimlik → hedef. */
async function relsOf(bytes: Uint8Array): Promise<Map<string, { type: string; target: string }>> {
  const xml = (await entry(bytes, "word/_rels/document.xml.rels")) ?? "";
  return new Map([...xml.matchAll(/<Relationship Id="([^"]+)" Type="([^"]+)" Target="([^"]+)"/g)].map((m) => [m[1]!, { type: m[2]!, target: m[3]! }]));
}

const opaques = (doc: JSONContent) => [...all(doc, "opaqueInline"), ...all(doc, "opaqueBlock")];

describe("gömülü OLE nesnesi", () => {
  it("opak satır içi öğe olur, kaydederken nesnesi ve önizlemesiyle korunur; açılışta kayıp uyarısı verilmez", async () => {
    const src = poi("EmbeddedDocument.docx");
    const r = await openDocument("EmbeddedDocument.docx", buf(src));
    expect(r.warnings).not.toContain(lossMessage("objects"));
    const [op] = opaques(r.doc.content);
    expect(op?.type).toBe("opaqueInline");
    expect(op!.attrs).toMatchObject({ kind: "ole", label: "Gömülü nesne (hesap tablosu)", level: "run", preview: null, width: 277, height: 21 });
    expect((op!.attrs!.refs as { type: string }[]).map((x) => x.type.replace(/^.*\//, "")).sort()).toEqual(["image", "oleObject"]);
    expect(op!.attrs!.xml).toMatch(/^<w:object\b[^>]*xmlns:w="/);
    expect(() => validateContent(r.doc.content)).not.toThrow();

    const out = await writeDocx(r.doc);
    const doc = (await entry(out, "word/document.xml"))!;
    expect(doc).toContain('ProgID="Excel.Sheet.8"');
    const rels = await relsOf(out);
    const zip = await zipOf(out);
    // XML'deki her ilişki kimliği yeni pakette var ve hedefi kopyalanmış parça; baytlar özgünle aynı.
    const ids = [...doc.matchAll(/r:id="([^"]+)"/g)].map((m) => m[1]!);
    expect(ids).toHaveLength(2);
    const orig = await zipOf(src);
    for (const id of ids) {
      const rel = rels.get(id)!;
      expect(rel.target).toMatch(/^(embeddings|media)\/oh_t\d+_/);
      const copied = await zip.file(`word/${rel.target}`)!.async("uint8array");
      const from = await orig.file(`word/${rel.target.replace(/oh_t\d+_/, "")}`)!.async("uint8array");
      expect(copied).toEqual(from);
    }
    const ct = (await entry(out, "[Content_Types].xml"))!;
    expect(ct).toMatch(/Extension="xls"|\.xls"/);
    expect(ct).toMatch(/Extension="emf"|\.emf"/);

    // İkinci tur kararlı: aynı XML ve aynı parça kümesi.
    const r2 = await readDocx(out);
    expect(opaques(r2.content).map((n) => n.attrs!.label)).toEqual(["Gömülü nesne (hesap tablosu)"]);
    const out2 = await writeDocx(r2);
    expect(Object.keys((await zipOf(out2)).files).sort()).toEqual(Object.keys(zip.files).sort());
    expect(await entry(out2, "word/document.xml")).toBe(doc);
  });

  it("birden çok nesne ve paket (ekli dosya) korunur; gövdedeki bütün nesneler alındıysa uyarı yok", async () => {
    const src = poi("WordWithAttachments.docx");
    const r = await openDocument("WordWithAttachments.docx", buf(src));
    expect(r.warnings).not.toContain(lossMessage("objects"));
    const ole = opaques(r.doc.content).filter((n) => n.attrs!.kind === "ole");
    expect(ole.map((n) => n.attrs!.label)).toEqual(["Gömülü nesne (ekli dosya)", "Gömülü nesne (belge)", "Gömülü nesne (belge)", "Gömülü nesne (hesap tablosu)", "Gömülü nesne (hesap tablosu)"]);
    const out = await writeDocx(r.doc);
    const names = Object.keys((await zipOf(out)).files);
    expect(names.filter((n) => n.startsWith("word/embeddings/"))).toHaveLength(5);
    // Başka nesnelerin tanımladığı VML şekil türü (_x0000_t75) her nesneye eklenir: tek başına silinse de çizilir.
    const doc = (await entry(out, "word/document.xml"))!;
    const objects = [...doc.matchAll(/<w:object\b[\s\S]*?<\/w:object>/g)].map((m) => m[0]);
    expect(objects).toHaveLength(5);
    for (const o of objects) expect(o).toContain('<v:shapetype id="_x0000_t75"');
  });

  it("çoğaltılan nesne kendi parça kopyalarını alır; VML şekil ve OLE nesne kimlikleri belgede benzersizdir", async () => {
    const r = await openDocument("EmbeddedDocument.docx", buf(poi("EmbeddedDocument.docx")));
    const [op] = opaques(r.doc.content);
    const section = r.doc.content.content![0]!;
    section.content!.push({ type: "paragraph", content: [op!, { type: "text", text: " ve " }, op!] });
    const out = await writeDocx(r.doc);
    const doc = (await entry(out, "word/document.xml"))!;
    const objects = [...doc.matchAll(/<w:object\b[\s\S]*?<\/w:object>/g)].map((m) => m[0]);
    expect(objects).toHaveLength(3);
    const rels = await relsOf(out);
    const targets = objects.map((o) => rels.get(/<o:OLEObject\b[^>]*r:id="([^"]+)"/.exec(o)![1]!)!.target);
    expect(new Set(targets).size).toBe(3);
    const shapeIds = objects.map((o) => /<v:shape id="([^"]+)"/.exec(o)![1]);
    expect(new Set(shapeIds).size).toBe(3);
    for (const o of objects) expect(/ShapeID="([^"]+)"/.exec(o)![1]).toBe(/<v:shape id="([^"]+)"/.exec(o)![1]);
    expect(new Set(objects.map((o) => /ObjectID="([^"]+)"/.exec(o)![1])).size).toBe(3);
    const zip = await zipOf(out);
    for (const t of targets) expect(zip.file(`word/${t}`)).not.toBeNull();
  });

  it("özgün paket yoksa yedek görünüm yazılır (önizleme resmi ya da yer tutucu metin); kırık ilişki yazılmaz", async () => {
    const r = await openDocument("EmbeddedDocument.docx", buf(poi("EmbeddedDocument.docx")));
    const [op] = opaques(r.doc.content);
    const doc: BelgeDocument = {
      settings: { ...r.doc.settings, kaynak: "kayitta-yok" },
      content: { type: "doc", content: [{ type: "section", content: [{ type: "paragraph", content: [{ ...op!, attrs: { ...op!.attrs, src: "kayitta-yok" } }, { ...op!, attrs: { ...op!.attrs, src: "kayitta-yok", preview: PNG } }] }] }] },
    };
    const out = await writeDocx(doc);
    const xml = (await entry(out, "word/document.xml"))!;
    expect(xml).not.toContain("<w:object");
    expect(xml).toContain(opaquePlaceholder(op!.attrs));
    expect(xml).toContain("<pic:pic>");
    const rels = [...(await relsOf(out)).values()];
    expect(rels.some((x) => /oleObject/.test(x.type))).toBe(false);
  });
});

describe("anlaşılmayan çizim, içe aktarılan içerik ve bilinmeyen öğeler", () => {
  const group =
    `<w:p><w:r><mc:AlternateContent><mc:Choice Requires="wpg"><w:drawing><wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="1" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1"><wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="column"><wp:posOffset>0</wp:posOffset></wp:positionH><wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV><wp:extent cx="952500" cy="476250"/><wp:wrapNone/><wp:docPr id="1" name="Grup 1"/><wp:cNvGraphicFramePr/>` +
    `<a:graphic><a:graphicData uri="${MS}/word/2010/wordprocessingGroup"><wpg:wgp><wpg:cNvGrpSpPr/><wpg:grpSpPr/><wps:wsp><wps:cNvSpPr/><wps:spPr/><wps:txbx><w:txbxContent><w:p w14:paraId="1A2B3C4D" w14:textId="77777777"><w:bookmarkStart w:id="7" w:name="_Grup"/><w:r><w:t>Grup metni</w:t></w:r><w:bookmarkEnd w:id="7"/></w:p></w:txbxContent></wps:txbx><wps:bodyPr/></wps:wsp></wpg:wgp></a:graphicData></a:graphic></wp:anchor></w:drawing></mc:Choice>` +
    `<mc:Fallback><w:pict><v:group style="width:75pt;height:37.5pt"><v:rect style="width:75pt;height:37.5pt"/></v:group></w:pict></mc:Fallback></mc:AlternateContent></w:r></w:p>`;

  it("çizim grubu bütünüyle opak; çoğaltınca docPr, w14:paraId ve yer işareti kimlikleri yeniden numaralanır", async () => {
    const bytes = await docxWith(group);
    const r = await openDocument("grup.docx", buf(bytes));
    const [op] = opaques(r.doc.content);
    expect(op!.attrs).toMatchObject({ kind: "drawing", label: "Çizim grubu", width: 100, height: 50 });
    // Kapsamda yok sayılabilir ilan edilen ad alanı (w14) öğenin kökünde de yok sayılabilir.
    expect(op!.attrs!.xml).toMatch(/^<mc:AlternateContent\b[^>]*mc:Ignorable="w14"/);
    const section = r.doc.content.content![0]!;
    section.content!.push({ type: "paragraph", content: [op!, { type: "image", attrs: { src: PNG, width: 4, height: 3 } }, op!] });
    const out = await writeDocx(r.doc);
    const doc = (await entry(out, "word/document.xml"))!;
    const docPr = [...doc.matchAll(/<wp:docPr id="(\d+)"/g)].map((m) => m[1]);
    expect(docPr).toHaveLength(4);
    expect(new Set(docPr).size).toBe(4);
    const paraIds = [...doc.matchAll(/w14:paraId="([0-9A-F]{8})"/g)].map((m) => m[1]);
    expect(paraIds).toHaveLength(3);
    expect(new Set(paraIds).size).toBe(3);
    const starts = [...doc.matchAll(/<w:bookmarkStart w:id="(\d+)"/g)].map((m) => m[1]);
    const ends = [...doc.matchAll(/<w:bookmarkEnd w:id="(\d+)"/g)].map((m) => m[1]);
    expect(new Set(starts).size).toBe(3);
    expect(ends.sort()).toEqual(starts.sort());
    // İkinci tur kararlı.
    const out2 = await writeDocx(await readDocx(out));
    expect(await entry(out2, "word/document.xml")).toBe(doc);
  });

  it("içe aktarılan içerik (altChunk) opak blok olur; parçası taşınır", async () => {
    const bytes = await docxWith(`<w:p><w:r><w:t>Önce</w:t></w:r></w:p><w:altChunk r:id="rIdAlt"/><w:p><w:r><w:t>Sonra</w:t></w:r></w:p>`, {
      rels: `<Relationship Id="rIdAlt" Type="${R}/aFChunk" Target="parca.htm"/>`,
      types: `<Default Extension="htm" ContentType="text/html"/>`,
      files: { "word/parca.htm": "<html><body><p>İçe aktarılan</p></body></html>" },
    });
    const r = await openDocument("altchunk.docx", buf(bytes));
    expect(r.warnings).toEqual([]);
    const [op] = opaques(r.doc.content);
    expect(op).toMatchObject({ type: "opaqueBlock", attrs: { kind: "altChunk", label: "İçe aktarılan içerik (.htm)" } });
    const out = await writeDocx(r.doc);
    const doc = (await entry(out, "word/document.xml"))!;
    expect(doc).toMatch(/<w:t xml:space="preserve">Önce<\/w:t><\/w:r><\/w:p><w:altChunk\b/);
    const rid = /<w:altChunk\b[^>]*r:id="([^"]+)"/.exec(doc)![1]!;
    const rel = (await relsOf(out)).get(rid)!;
    expect(rel.type).toBe(`${R}/aFChunk`);
    expect(await entry(out, `word/${rel.target}`)).toContain("İçe aktarılan");
    expect(await entry(out, "[Content_Types].xml")).toContain("text/html");
  });

  it("bilinmeyen blok, paragraf ve koşu öğeleri opak; çıktının bilmediği ad alanı mc:AlternateContent ile sarılır", async () => {
    const bytes = await docxWith(
      `<w:p><w:r><w:t xml:space="preserve">Bugün </w:t></w:r><w:r><w:dayLong/></w:r><w15:ozel w15:val="1"/></w:p>` +
        `<m:oMathPara><m:oMath><m:r><m:t>x</m:t></m:r></m:oMath></m:oMathPara><w:p><w:r><w:pict><v:rect style="width:0;height:2.25pt" o:hr="t" o:hrstd="t" fillcolor="gray" stroked="f"/></w:pict></w:r></w:p>`,
    );
    const r = await openDocument("bilinmeyen.docx", buf(bytes));
    const ops = opaques(r.doc.content);
    expect(ops.map((n) => [n.type, n.attrs!.kind, n.attrs!.label, n.attrs!.level ?? null])).toEqual([
      ["opaqueInline", "other", "w:dayLong", "run"],
      ["opaqueInline", "other", "w15:ozel", "para"],
      ["opaqueInline", "hr", "Yatay çizgi", "run"],
      ["opaqueBlock", "math", "Denklem", null],
    ]);
    const ozel = ops.find((n) => n.attrs!.label === "w15:ozel")!.attrs!.xml as string;
    expect(ozel).toMatch(/^<mc:AlternateContent xmlns:mc="[^"]+" xmlns:w15="[^"]+"><mc:Choice Requires="w15"><w15:ozel\b/);
    const out = await writeDocx(r.doc);
    const doc = (await entry(out, "word/document.xml"))!;
    expect(doc).toContain("<w:r><w:dayLong");
    expect(doc).toContain("<m:oMathPara");
    expect(doc).toContain('o:hr="t"');
    const r2 = await readDocx(out);
    expect(opaques(r2.content).map((n) => n.attrs!.xml)).toEqual(ops.map((n) => n.attrs!.xml));
  });

  it("yapıyı bozacak opak XML yazılmaz, yerine yer tutucu yazılır", async () => {
    const bad = (xml: string): JSONContent => ({ type: "opaqueInline", attrs: { xml, refs: [], src: null, kind: "other", label: "deneme", level: "para", preview: null, width: null, height: null } });
    const doc: BelgeDocument = {
      settings: defaultSettings(),
      content: { type: "doc", content: [{ type: "section", content: [{ type: "paragraph", content: [bad("<w:r/></w:p><w:p>"), bad("<!DOCTYPE x><w:r/>"), bad("<w:r/><w:r/>"), bad("metin")] }] }] },
    };
    const xml = (await entry(await writeDocx(doc), "word/document.xml"))!;
    expect(xml.match(/\[Desteklenmeyen öğe: deneme\]/g)).toHaveLength(4);
    expect(xml).not.toContain("DOCTYPE");
    expect(topElements("<w:r/></w:p><w:p>")).toBeNull();
    expect(topElements('<w:a x="1"><w:b>t</w:b></w:a><w:c/>')?.map((e) => e.local)).toEqual(["a", "c"]);
  });
});

describe("paragraf ve koşu ek özellikleri (ooxmlExtra)", () => {
  it("modelin okumadığı pPr/rPr çocukları şema sırasıyla korunur; paragraf işaretindeki düzeltme atılır", async () => {
    const bytes = await docxWith(
      `<w:p><w:pPr><w:keepNext/><w:framePr w:w="2000" w:h="500" w:hAnchor="page" w:vAnchor="page" w:x="1000" w:y="1000"/><w:widowControl w:val="0"/><w:snapToGrid w:val="0"/><w:spacing w:after="0"/><w:jc w:val="center"/><w:textAlignment w:val="center"/><w:rPr><w:lang w:val="en-US"/><w:ins w:id="5" w:author="x" w:date="2020-01-01T00:00:00Z"/></w:rPr></w:pPr>` +
        `<w:r><w:rPr><w:b/><w:noProof/><w:w w:val="90"/><w:kern w:val="2"/><w:lang w:val="de-DE"/><w14:shadow w14:blurRad="38100" w14:dist="19050" w14:dir="2700000" w14:sx="100000" w14:sy="100000" w14:kx="0" w14:ky="0" w14:algn="tl"/></w:rPr><w:t>Metin</w:t></w:r></w:p>`,
    );
    const r = await readDocx(bytes);
    const p = all(r.content, "paragraph")[0]!;
    expect(p.attrs!.ooxmlExtra).toBe('<w:framePr w:w="2000" w:h="500" w:hAnchor="page" w:vAnchor="page" w:x="1000" w:y="1000"/><w:widowControl w:val="0"/><w:snapToGrid w:val="0"/><w:textAlignment w:val="center"/><w:rPr><w:lang w:val="en-US"/></w:rPr>');
    const ts = p.content![0]!.marks!.find((m) => m.type === "textStyle")!;
    expect(ts.attrs!.ooxmlExtra).toBe('<w:noProof/><w:w w:val="90"/><w:kern w:val="2"/><w:lang w:val="de-DE"/><w14:shadow w14:blurRad="38100" w14:dist="19050" w14:dir="2700000" w14:sx="100000" w14:sy="100000" w14:kx="0" w14:ky="0" w14:algn="tl"/>');
    expect(() => validateContent(r.content)).not.toThrow();
    const out = await writeDocx(r);
    const doc = (await entry(out, "word/document.xml"))!;
    const pPr = /<w:pPr>([\s\S]*?)<\/w:pPr>/.exec(doc)![1]!;
    expect([...pPr.matchAll(/<w:(\w+)[\s/>]/g)].map((m) => m[1]).filter((n) => n !== "lang")).toEqual(["keepNext", "framePr", "widowControl", "snapToGrid", "spacing", "jc", "textAlignment", "rPr"]);
    // Ana ad alanı dışındaki ek özellik kapsayıcıda yok sayılabilir ilan edilir (üst/alt bilgi kökünde bildirim yok).
    expect(doc).toContain('<w:r><w:rPr mc:Ignorable="w14">');
    const rPr = /<w:r><w:rPr[^>]*>([\s\S]*?)<\/w:rPr><w:t/.exec(doc)![1]!;
    expect([...rPr.matchAll(/<(w(?:14)?):(\w+)/g)].map((m) => `${m[1]}:${m[2]}`)).toEqual(["w:b", "w:bCs", "w:noProof", "w:w", "w:kern", "w:lang", "w14:shadow"]);
    expect(doc).not.toContain('w:author="x"');
    // İkinci tur kararlı.
    expect(await entry(await writeDocx(await readDocx(out)), "word/document.xml")).toBe(doc);
  });
});

describe("başka biçimler, pano ve açıklama çapası", () => {
  it(".odt/.rtf/.md'de opak öğenin yerine yer tutucu metin yazılır", async () => {
    const r = await openDocument("EmbeddedDocument.docx", buf(poi("EmbeddedDocument.docx")));
    expect(hasOpaque(r.doc)).toBe(true);
    const block: JSONContent = { type: "opaqueBlock", attrs: { xml: "<w:altChunk/>", refs: [], src: null, kind: "altChunk", label: "İçe aktarılan içerik", preview: null, width: null, height: null } };
    r.doc.content.content![0]!.content!.push(block);
    const phInline = opaquePlaceholder({ label: "Gömülü nesne (hesap tablosu)" });
    const phBlock = opaquePlaceholder(block.attrs);
    const odt = await JSZip.loadAsync(await writeOdt(r.doc));
    const content = await odt.file("content.xml")!.async("string");
    expect(content).toContain(phInline);
    expect(content).toContain(phBlock);
    const rtf = writeRtf(r.doc);
    expect(rtf).toContain("Desteklenmeyen \\u");
    const md = writeMarkdown(r.doc);
    expect(md).toContain("\\[Desteklenmeyen öğe: Gömülü nesne (hesap tablosu)\\]");
  });

  it("pano: opak öğe ve ek özellikler HTML'de kayıt defteri anahtarıyla taşınır; dış HTML XML sokamaz", async () => {
    const r = await readDocx(await docxWith(`<w:p><w:pPr><w:widowControl w:val="0"/></w:pPr><w:r><w:rPr><w:lang w:val="en-US"/></w:rPr><w:t>Dil</w:t></w:r><w:r><w:dayLong/></w:r></w:p><m:oMathPara><m:oMath><m:r><m:t>y</m:t></m:r></m:oMath></m:oMathPara>`));
    const section = r.content.content![0]!;
    const html = generateHTML({ type: "doc", content: [section] }, schemaExtensions());
    // XML HTML'e yazılmaz (yalnız anahtar); kartın görünen adı yazılır.
    expect(html).not.toContain("<w:dayLong");
    expect(html).not.toContain("widowControl");
    expect(html).not.toContain("xmlns:");
    expect(html).toContain("Desteklenmeyen öğe: w:dayLong");
    const back = generateJSON(html, schemaExtensions());
    expect(opaques(back).map((n) => n.attrs!.xml)).toEqual(opaques(r.content).map((n) => n.attrs!.xml));
    const p = all(back, "paragraph")[0]!;
    expect(p.attrs!.ooxmlExtra).toBe('<w:widowControl w:val="0"/>');
    expect(p.content![0]!.marks!.find((m) => m.type === "textStyle")?.attrs?.ooxmlExtra).toBe('<w:lang w:val="en-US"/>');

    const foreign = generateJSON(
      `<p>a<span data-opaque="sahte" data-xml="&lt;w:r/&gt;">Desteklenmeyen öğe: x</span><span data-ooxml="sahte">b</span></p><div data-opaque="sahte"><p>gizli</p></div>`,
      schemaExtensions(),
    );
    expect(opaques(foreign)).toEqual([]);
    expect(JSON.stringify(foreign)).not.toContain("ooxmlExtra\":\"");
    expect(all(foreign, "paragraph").map(text)).toEqual(["ab"]);
  });

  it("ardından metin gelmeyen nokta açıklaması kaybolmaz: çapaya bağlanır, ikinci turda çapa birikmez", async () => {
    const r = await openDocument("comment.docx", buf(poi("comment.docx")));
    expect(r.warnings).not.toContain(lossMessage("comments"));
    const anchors = all(r.doc.content, "bookmark");
    expect(anchors).toHaveLength(1);
    expect(anchors[0]!.marks).toEqual([{ type: "comment", attrs: { id: "c0" } }]);
    const out = await writeDocx(r.doc);
    expect(await entry(out, "word/comments.xml")).toContain("<w:comment ");
    const doc = (await entry(out, "word/document.xml"))!;
    expect(doc).toMatch(/<w:commentRangeStart w:id="0"\/><w:bookmarkStart w:id="0" w:name="_Aciklama_c0"\/><w:bookmarkEnd w:id="0"\/><w:commentRangeEnd w:id="0"\/>/);
    const r2 = await openDocument("comment2.docx", buf(out));
    expect(all(r2.doc.content, "bookmark")).toHaveLength(1);
    expect(await entry(await writeDocx(r2.doc), "word/document.xml")).toBe(doc);
  });
});
