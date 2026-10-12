// @vitest-environment jsdom
// 3. tur DOCX gidiş-dönüşleri: metin kaydırma, sekme durakları, anahat düzeyi, metin efektleri, özel stiller,
// alanlar (SEQ/REF/CITATION/MERGEFIELD), kaynakça, şekiller tablosu, belge özellikleri, tablo seçenekleri;
// tablo dönüştürme/sıralama ve XLSX alıcı listesi.
import { describe, expect, it } from "vitest";
import JSZip from "jszip";
import type { JSONContent } from "@tiptap/core";
import { Node as PMNode } from "@tiptap/pm/model";
import { belgeSchema, readDocx, writeDocx } from "../index";
import { defaultSettings, type BelgeDocument } from "../../model/document";
import { buildTable, sortTableRows, tableToParagraphs, textRows } from "../../editor/tableTools";
import { parseXlsx, rowsToList } from "../../mailmerge/data";

const PNG = "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAQAAAADCAIAAAA7ljmRAAAAEElEQVR4nGM4IScHRww4OQD1xwwxB+utmQAAAABJRU5ErkJggg==";
const t = (text: string, marks?: JSONContent["marks"]): JSONContent => ({ type: "text", text, ...(marks ? { marks } : {}) });
const p = (content: JSONContent[] | string, attrs: Record<string, unknown> = {}): JSONContent => ({ type: "paragraph", attrs, content: typeof content === "string" ? [t(content)] : content });

function docOf(content: JSONContent[], patch: Partial<BelgeDocument["settings"]> = {}): BelgeDocument {
  return { content: { type: "doc", content }, settings: { ...defaultSettings(), ...patch } };
}

async function roundtrip(doc: BelgeDocument) {
  const bytes = await writeDocx(doc);
  return { back: await readDocx(bytes), zip: await JSZip.loadAsync(bytes) };
}

/** Bölüm içindeki üst düzey bloklar. */
const blocks = (d: BelgeDocument) => d.content.content![0].content!;
/** Belgedeki belirli türdeki ilk satır içi düğüm. */
function findNode(json: JSONContent, type: string): JSONContent | undefined {
  if (json.type === type) return json;
  for (const c of json.content ?? []) {
    const f = findNode(c, type);
    if (f) return f;
  }
  return undefined;
}

describe("DOCX gidiş-dönüş (3. tur)", () => {
  it("metin kaydırma: kare, metnin arkasında (wp:anchor)", async () => {
    const content = [
      p([{ type: "image", attrs: { src: PNG, width: 80, height: 60, wrap: "square", floatSide: "right" } }, t("Kare kaydırılan metin")]),
      p([{ type: "image", attrs: { src: PNG, width: 80, height: 60, wrap: "behind", posX: 40, posY: 20, z: 3 } }, t("Arkada")]),
      p([{ type: "image", attrs: { src: PNG, width: 40, height: 30 } }, t("Satır içi")]),
    ];
    const { back, zip } = await roundtrip(docOf(content));
    const xml = await zip.file("word/document.xml")!.async("string");
    expect(xml).toContain("<wp:anchor");
    expect(xml).toContain("<wp:wrapSquare");
    expect(xml).toMatch(/behindDoc="1"/);
    const imgs = blocks(back).map((b) => b.content![0].attrs!);
    expect(imgs[0]).toMatchObject({ wrap: "square", floatSide: "right" });
    expect(imgs[1]).toMatchObject({ wrap: "behind", posX: 40, posY: 20 });
    expect(imgs[1].z).toBeGreaterThan(0);
    expect(imgs[2].wrap ?? "inline").toBe("inline");
  });

  it("sekme durakları ve anahat düzeyi", async () => {
    const tabs = [
      { pos: 72, align: "left", leader: "none" },
      { pos: 216, align: "decimal", leader: "dot" },
      { pos: 432, align: "right", leader: "hyphen" },
    ];
    const content = [p("Ad\tFiyat\tSayfa", { tabs }), p("Gövde başlığı", { outlineLevel: 2 })];
    const { back, zip } = await roundtrip(docOf(content));
    const xml = await zip.file("word/document.xml")!.async("string");
    expect(xml).toContain('<w:tab w:val="decimal" w:leader="dot" w:pos="4320"/>');
    expect(xml).toContain('<w:outlineLvl w:val="1"/>');
    expect(blocks(back)[0].attrs?.tabs).toEqual(tabs);
    expect(blocks(back)[1].attrs?.outlineLevel).toBe(2);
  });

  it("metin efektleri", async () => {
    const content = [p([t("Gölge", [{ type: "textStyle", attrs: { shadow: true, outline: true } }]), t(" kabartma", [{ type: "textStyle", attrs: { emboss: true } }])])];
    const { back, zip } = await roundtrip(docOf(content));
    const xml = await zip.file("word/document.xml")!.async("string");
    expect(xml).toContain("<w:shadow/>");
    expect(xml).toContain("<w:emboss/>");
    const marks = blocks(back)[0].content!.map((x) => x.marks?.find((m) => m.type === "textStyle")?.attrs);
    expect(marks[0]).toMatchObject({ shadow: true, outline: true });
    expect(marks[1]).toMatchObject({ emboss: true });
  });

  it("özel stiller, kaynaklar ve belge özellikleri", async () => {
    const customStyles = { Vurgulu: { id: "Vurgulu", name: "Vurgulu Metin", type: "paragraph" as const, basedOn: "Normal", next: "Normal", props: { bold: true, color: "#C00000", size: 13 } } };
    const sources = [{ tag: "Kaya2019", type: "book" as const, author: "Kaya, Ali", title: "Kitap", year: "2019", publisher: "Yayınevi" }];
    const doc = docOf([p("Özel", { styleId: "Vurgulu" })], { customStyles, sources, bibStyle: "ieee" });
    doc.settings.properties = { ...doc.settings.properties, title: "Rapor", subject: "Deneme", author: "Ayşe Yılmaz", keywords: "rapor; test" };
    const { back, zip } = await roundtrip(doc);
    const styles = await zip.file("word/styles.xml")!.async("string");
    expect(styles).toContain('w:styleId="Vurgulu"');
    expect(styles).toContain('<w:name w:val="Vurgulu Metin"/>');
    expect(back.settings.customStyles?.Vurgulu).toMatchObject({ name: "Vurgulu Metin", props: { bold: true, color: "#C00000", size: 13 } });
    expect(back.settings.sources).toEqual(sources);
    expect(back.settings.bibStyle).toBe("ieee");
    expect(blocks(back)[0].attrs?.styleId).toBe("Vurgulu");
    const core = await zip.file("docProps/core.xml")!.async("string");
    expect(core).toContain("<dc:title>Rapor</dc:title>");
    expect(back.settings.properties).toMatchObject({ title: "Rapor", subject: "Deneme", author: "Ayşe Yılmaz", keywords: "rapor; test" });
  });

  it("alanlar: SEQ, REF, CITATION, MERGEFIELD", async () => {
    const content = [
      p([{ type: "bookmark", attrs: { name: "_Ref1" } }, t("Şekil "), { type: "seqField", attrs: { label: "Şekil", num: 1 } }, t(": Grafik")], { styleId: "Caption" }),
      p([
        t("Bkz. "),
        { type: "crossRef", attrs: { ref: "_Ref1", kind: "label", text: "Şekil 1", link: true } },
        t(" ve "),
        { type: "citation", attrs: { tag: "Kaya2019", pages: "", text: "(Kaya, 2019)" } },
      ]),
      p([t("Sayın "), { type: "mergeField", attrs: { name: "Ad" } }]),
    ];
    const { back, zip } = await roundtrip(docOf(content));
    const xml = await zip.file("word/document.xml")!.async("string");
    expect(xml).toMatch(/SEQ Şekil/);
    expect(xml).toMatch(/REF _Ref1/);
    expect(xml).toMatch(/CITATION Kaya2019/);
    expect(xml).toMatch(/MERGEFIELD/);
    expect(findNode(back.content, "seqField")?.attrs).toMatchObject({ label: "Şekil", num: 1 });
    expect(findNode(back.content, "crossRef")?.attrs).toMatchObject({ ref: "_Ref1", text: "Şekil 1" });
    expect(findNode(back.content, "citation")?.attrs).toMatchObject({ tag: "Kaya2019" });
    expect(findNode(back.content, "mergeField")?.attrs).toMatchObject({ name: "Ad" });
    expect(blocks(back)[0].attrs?.styleId).toBe("Caption");
  });

  it("kaynakça ve şekiller tablosu blokları", async () => {
    const entries = [{ before: "Kaya, A. (2019). ", italic: "Kitap", after: ". Yayınevi." }];
    const content = [
      { type: "toc", attrs: { kind: "tof", label: "Tablo", title: "", entries: [{ level: 1, text: "Tablo 1: Veriler", page: 2 }], showPages: true } },
      { type: "bibliography", attrs: { title: "Kaynakça", entries } },
      p("Son"),
    ];
    const { back, zip } = await roundtrip(docOf(content));
    const xml = await zip.file("word/document.xml")!.async("string");
    expect(xml).toContain('TOC \\h \\z \\c "Tablo"');
    expect(xml).toContain("BIBLIOGRAPHY");
    const b = blocks(back);
    expect(b[0]).toMatchObject({ type: "toc", attrs: { kind: "tof", label: "Tablo" } });
    expect(b[1]).toMatchObject({ type: "bibliography", attrs: { title: "Kaynakça" } });
    expect((b[1].attrs?.entries as unknown[]).length).toBe(1);
  });

  it("tablo: başlık satırını yinele ve hücre kenar boşlukları", async () => {
    const cell = (text: string): JSONContent => ({ type: "tableCell", content: [p(text)] });
    const content = [
      {
        type: "table",
        attrs: { tableStyle: "grid", repeatHeader: true, cellMargins: { top: 2, bottom: 1, left: 7.2, right: 3.6 } },
        content: [
          { type: "tableRow", content: [cell("Ad"), cell("Puan")] },
          { type: "tableRow", content: [cell("Ali"), cell("9")] },
        ],
      },
      p(""),
    ];
    const { back, zip } = await roundtrip(docOf(content));
    const xml = await zip.file("word/document.xml")!.async("string");
    expect(xml).toContain("<w:trPr><w:tblHeader/></w:trPr>");
    expect(xml).toContain('<w:left w:w="144" w:type="dxa"/>');
    expect(blocks(back)[0].attrs).toMatchObject({ repeatHeader: true, cellMargins: { top: 2, bottom: 1, left: 7.2, right: 3.6 } });
  });
});

describe("Tablo dönüştürme ve sıralama", () => {
  const schema = belgeSchema();
  const para = (text: string) => schema.nodes.paragraph.create(null, text ? schema.text(text) : null);

  it("metni tabloya ve geri dönüştürür", () => {
    const nodes = [para("Ad\tPuan"), para("Ali\t9"), para("Ayşe")];
    const rows = textRows(nodes, "\t", 2);
    expect(rows.map((r) => r.map((c) => c.textContent))).toEqual([["Ad", "Puan"], ["Ali", "9"], ["Ayşe"]]);
    const table = buildTable(schema, rows, 2);
    expect(table.childCount).toBe(3);
    expect(table.child(2).childCount).toBe(2);
    expect(table.child(2).child(1).textContent).toBe("");
    expect(tableToParagraphs(schema, table, ";").map((n) => n.textContent)).toEqual(["Ad;Puan", "Ali;9", "Ayşe;"]);
    expect(tableToParagraphs(schema, table, "para").map((n) => n.textContent)).toEqual(["Ad", "Puan", "Ali", "9", "Ayşe", ""]);
    // Paragraf ayracı: sütun sayısı kadar paragraf bir satır olur.
    expect(textRows([para("a"), para("b"), para("c")], "para", 2).map((r) => r.length)).toEqual([2, 1]);
  });

  it("tablo satırlarını üst bilgi satırı korunarak sıralar", () => {
    const table = buildTable(
      schema,
      [
        [para("Ad"), para("Puan")],
        [para("Zeynep"), para("85")],
        [para("Ali"), para("100")],
        [para("Çiğdem"), para("7")],
      ],
      2,
    );
    const names = (rows: PMNode[] | null) => rows!.map((r) => r.child(0).textContent);
    expect(names(sortTableRows(table, { col: 0, type: "text", desc: false, header: true }))).toEqual(["Ad", "Ali", "Çiğdem", "Zeynep"]);
    expect(names(sortTableRows(table, { col: 1, type: "number", desc: true, header: true }))).toEqual(["Ad", "Ali", "Zeynep", "Çiğdem"]);
    expect(names(sortTableRows(table, { col: 0, type: "text", desc: false, header: false }))).toEqual(["Ad", "Ali", "Çiğdem", "Zeynep"]);
  });
});

describe("XLSX alıcı listesi", () => {
  it("paylaşılan dizeler, sayılar ve boş hücreler", async () => {
    const zip = new JSZip();
    zip.file(
      "xl/workbook.xml",
      '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Liste" sheetId="1" r:id="rId1"/></sheets></workbook>',
    );
    zip.file(
      "xl/_rels/workbook.xml.rels",
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="worksheet" Target="worksheets/liste.xml"/></Relationships>',
    );
    zip.file("xl/sharedStrings.xml", '<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><si><t>Ad</t></si><si><t>Tutar</t></si><si><r><t>Ay</t></r><r><t>şe</t></r></si></sst>');
    zip.file(
      "xl/worksheets/liste.xml",
      '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1" t="s"><v>1</v></c></row><row r="2"><c r="A2" t="s"><v>2</v></c><c r="C2"><v>12.5</v></c></row><row r="3"/></sheetData></worksheet>',
    );
    const rows = await parseXlsx(await zip.generateAsync({ type: "uint8array" }));
    expect(rows).toEqual([
      ["Ad", "Tutar"],
      ["Ayşe", "", "12,5"],
    ]);
    expect(rowsToList("liste.xlsx", rows).records[0]).toEqual({ Ad: "Ayşe", Tutar: "" });
  });
});
