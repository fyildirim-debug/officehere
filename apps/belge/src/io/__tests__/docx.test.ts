// @vitest-environment jsdom
// DOCX yazıcı/okuyucu gidiş-dönüş testleri.
import { describe, expect, it } from "vitest";
import JSZip from "jszip";
import type { JSONContent } from "@tiptap/core";
import { Node as PMNode } from "@tiptap/pm/model";
import { belgeSchema, readDocx, writeDocx } from "../index";
import { defaultSettings, normalizeDoc, type BelgeDocument } from "../../model/document";

const PNG = "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAQAAAADCAIAAAA7ljmRAAAAEElEQVR4nGM4IScHRww4OQD1xwwxB+utmQAAAABJRU5ErkJggg==";

const t = (text: string, marks?: JSONContent["marks"]): JSONContent => ({ type: "text", text, ...(marks ? { marks } : {}) });
const p = (content: JSONContent[] | string, attrs: Record<string, unknown> = {}): JSONContent => ({ type: "paragraph", attrs, content: typeof content === "string" ? [t(content)] : content });

/** Şema üzerinden normalleştirme: varsayılan öznitelikler doldurulur, bitişik metinler birleşir. */
function norm(json: JSONContent) {
  return PMNode.fromJSON(belgeSchema(), normalizeDoc(json)).toJSON();
}

async function roundtrip(doc: BelgeDocument) {
  const bytes = await writeDocx(doc);
  return readDocx(bytes);
}

function docOf(...content: JSONContent[]): BelgeDocument {
  return { content: { type: "doc", content }, settings: defaultSettings() };
}

describe("DOCX gidiş-dönüş", () => {
  it("biçimli metin", async () => {
    const content = [
      p([
        t("Düz "),
        t("kalın", [{ type: "bold" }]),
        t(" italik", [{ type: "italic" }]),
        t(" altı çizili", [{ type: "underline", attrs: { type: "single", color: null } }]),
        t(" çift", [{ type: "underline", attrs: { type: "dashed", color: "#FF0000" } }]),
        t(" üstü çizili", [{ type: "strike" }]),
        t("2", [{ type: "superscript" }]),
        t("x", [{ type: "subscript" }]),
        t(" renkli", [{ type: "textStyle", attrs: { color: "#C00000", fontFamily: "Georgia", fontSize: "14pt" } }]),
        t(" vurgulu", [{ type: "highlight", attrs: { color: "#00FF00" } }]),
        t(" büyük", [{ type: "textStyle", attrs: { caps: true, letterSpacing: 1, backgroundColor: "#FFC000" } }]),
        t(" güçlü", [{ type: "charStyle", attrs: { styleId: "Strong" } }]),
        t("\tsekme"),
        { type: "hardBreak" },
        t("İĞÜŞÖÇ ığüşöç"),
      ]),
    ];
    const back = await roundtrip(docOf(...content));
    expect(norm(back.content)).toEqual(norm({ type: "doc", content }));
  });

  it("başlıklar, paragraf stilleri ve paragraf biçimi", async () => {
    const content = [
      { type: "heading", attrs: { level: 1 }, content: [t("Başlık 1")] },
      { type: "heading", attrs: { level: 2 }, content: [t("Başlık 2")] },
      { type: "heading", attrs: { level: 3 }, content: [t("Başlık 3")] },
      p("Konu", { styleId: "Title" }),
      p("Alt konu", { styleId: "Subtitle" }),
      p("Alıntı", { styleId: "Quote" }),
      p("Belirgin", { styleId: "IntenseQuote" }),
      p("Aralıksız", { styleId: "NoSpacing" }),
      p("Ortalı ve girintili", { textAlign: "center", indentLeft: 36, indentRight: 18, indentFirst: 35.4, spaceBefore: 12, spaceAfter: 6, lineHeight: 1.5 }),
      p("Asılı", { textAlign: "justify", indentLeft: 35.4, indentFirst: -35.4, lineRule: "exact", lineHeight: 18 }),
      p("Gölgeli", {
        shading: "#DAE9F7",
        borders: { top: { style: "single", size: 0.5, color: "#000000", space: 1 }, bottom: { style: "double", size: 0.75, color: "#0F4761", space: 2 } },
        keepNext: true,
        pageBreakBefore: true,
      }),
      { type: "horizontalRule" },
      p("Son"),
    ];
    const back = await roundtrip(docOf(...content));
    expect(norm(back.content)).toEqual(norm({ type: "doc", content }));
  });

  it("listeler", async () => {
    const li = (text: string, ...more: JSONContent[]): JSONContent => ({ type: "listItem", content: [p(text), ...more] });
    const content = [
      { type: "bulletList", attrs: { listStyle: "arrow" }, content: [li("Bir"), li("İki")] },
      p("Ara"),
      {
        type: "orderedList",
        attrs: { listStyle: "upper-roman", start: 3 },
        content: [li("Üç"), li("Dört", { type: "bulletList", content: [li("İç madde"), li("İç madde 2")] })],
      },
      { type: "orderedList", attrs: { multilevel: "outline", start: 1 }, content: [li("Anahat", { type: "orderedList", attrs: { multilevel: "outline", start: 1 }, content: [li("Alt")] })] },
    ];
    const back = await roundtrip(docOf(...content));
    expect(norm(back.content)).toEqual(norm({ type: "doc", content }));
  });

  it("tablolar (birleştirilmiş hücreler, gölgelendirme, stil)", async () => {
    const cell = (text: string, attrs: Record<string, unknown> = {}): JSONContent => ({ type: "tableCell", attrs: { colwidth: [120], ...attrs }, content: [p(text)] });
    const content = [
      {
        type: "table",
        attrs: { tableStyle: "grid-blue", headerRow: true, bandedRows: false, firstColumn: true, totalRow: false, align: "center" },
        content: [
          { type: "tableRow", content: [cell("A1", { colspan: 2, colwidth: [120, 120] }), cell("C1")] },
          { type: "tableRow", content: [cell("A2", { rowspan: 2, background: "#FFC000" }), cell("B2", { verticalAlign: "middle" }), cell("C2")] },
          { type: "tableRow", content: [cell("B3"), cell("C3")] },
        ],
      },
      p("Sonra"),
    ];
    const back = await roundtrip(docOf(...content));
    expect(norm(back.content)).toEqual(norm({ type: "doc", content }));
  });

  it("pencereye sığan tablo genişlikleri boş kalır", async () => {
    const content = [
      { type: "table", attrs: { tableStyle: "grid" }, content: [{ type: "tableRow", content: [{ type: "tableCell", content: [p("x")] }, { type: "tableCell", content: [p("y")] }] }] },
      p(""),
    ];
    const back = await roundtrip(docOf(...content));
    // belge > bölüm > tablo > satır > hücre
    const cell = back.content.content![0].content![0].content![0].content![0];
    expect(cell.attrs?.colwidth).toBeNull();
  });

  it("resim, köprü, yer işareti, sayfa sonu", async () => {
    const content = [
      p([t("Resim: "), { type: "image", attrs: { src: PNG, width: 40, height: 30, alt: "Kırmızı kare", rotate: 90, flipH: true } }]),
      p([t("Site", [{ type: "link", attrs: { href: "https://ornek.com/a?b=1&c=2" } }]), t(" ve "), t("iç bağlantı", [{ type: "link", attrs: { href: "#yerimi" } }])]),
      { type: "pageBreak" },
      p([{ type: "bookmark", attrs: { name: "yerimi" } }, t("Hedef")]),
    ];
    const back = await roundtrip(docOf(...content));
    expect(norm(back.content)).toEqual(norm({ type: "doc", content }));
  });

  it("açıklamalar ve izlenen değişiklikler", async () => {
    const date = "2026-10-11T10:00:00Z";
    const doc = docOf(
      p([t("Önce "), t("açıklamalı", [{ type: "comment", attrs: { id: "c0" } }]), t(" metin")]),
      p([t("eklenen", [{ type: "insertion", attrs: { author: "Ayşe", date } }]), t("silinen", [{ type: "deletion", attrs: { author: "Ayşe", date } }])]),
    );
    doc.settings.comments = { c0: { id: "c0", author: "Furkan Yıldırım", initials: "FY", date, text: "Bunu kontrol et.\nİkinci satır." } };
    const back = await roundtrip(doc);
    expect(norm(back.content)).toEqual(norm(doc.content));
    expect(back.settings.comments.c0.text).toBe("Bunu kontrol et.\nİkinci satır.");
    expect(back.settings.comments.c0.author).toBe("Furkan Yıldırım");
  });

  it("izlenen biçim değişikliği (w:rPrChange)", async () => {
    const date = "2026-10-11T10:00:00Z";
    const old = JSON.stringify([{ type: "italic" }]);
    const content = [p([t("biçimi değişen", [{ type: "bold" }, { type: "formatChange", attrs: { author: "Ayşe", date, old } }]), t(" metin")])];
    const bytes = await writeDocx(docOf(...content));
    const xml = await (await JSZip.loadAsync(bytes)).file("word/document.xml")!.async("string");
    expect(xml).toMatch(/<w:rPr><w:b\/><w:bCs\/><w:rPrChange w:id="\d+" w:author="Ayşe" w:date="2026-10-11T10:00:00Z"><w:rPr><w:i\/><w:iCs\/><\/w:rPr><\/w:rPrChange><\/w:rPr>/);
    const back = await readDocx(bytes);
    const run = back.content.content![0].content![0].content![0];
    expect(run.marks!.map((m) => m.type)).toEqual(["bold", "formatChange"]);
    expect(JSON.parse(String(run.marks![1].attrs!.old))).toEqual([{ type: "italic" }]);
  });

  it("dipnotlar ve son notlar", async () => {
    const content = [
      p([t("Birinci iddia"), { type: "note", attrs: { kind: "footnote", text: "Kaynak: Örnek Yayınları, 2026.", id: "nf1" } }, t(" ve ikinci"), { type: "note", attrs: { kind: "footnote", text: "İkinci dipnot.", id: "nf2" } }]),
      p([t("Son not burada"), { type: "note", attrs: { kind: "endnote", text: "Belgenin sonundaki açıklama.", id: "ne1" } }]),
    ];
    const bytes = await writeDocx(docOf(...content));
    const zip = await JSZip.loadAsync(bytes);
    expect(zip.file("word/footnotes.xml")).toBeTruthy();
    expect(zip.file("word/endnotes.xml")).toBeTruthy();
    const back = await readDocx(bytes);
    expect(norm(back.content)).toEqual(norm({ type: "doc", content }));
  });

  it("içindekiler", async () => {
    const entries = [
      { level: 1, text: "Giriş", page: 1, pos: 0 },
      { level: 2, text: "Alt bölüm", page: 2, pos: 0 },
    ];
    const content = [{ type: "toc", attrs: { entries, title: "İçindekiler" } }, { type: "heading", attrs: { level: 1 }, content: [t("Giriş")] }];
    const back = await roundtrip(docOf(...content));
    expect(norm(back.content)).toEqual(norm({ type: "doc", content }));
  });

  it("sayfa yapısı, üst/alt bilgi, filigran, sayfa rengi ve kenarlığı", async () => {
    const doc = docOf(p("Gövde"));
    const s = doc.settings;
    s.page = { ...s.page, width: 16838, height: 11906, orientation: "landscape", margin: { top: 720, bottom: 1080, left: 1440, right: 900, header: 500, footer: 600, gutter: 0 }, columns: 2, columnSpace: 567 };
    s.header = { type: "doc", content: [p([t("Üst bilgi metni")], { textAlign: "center" })] };
    s.footer = { type: "doc", content: [p([t("Sayfa "), { type: "pageField", attrs: { field: "PAGE" } }, t(" / "), { type: "pageField", attrs: { field: "NUMPAGES" } }], { textAlign: "right" })] };
    s.watermark = { text: "GİZLİ", color: "#C0C0C0", layout: "diagonal", font: "Aptos", semitransparent: true };
    s.pageColor = "#FFF2CC";
    s.pageBorder = { style: "double", color: "#156082", width: 1.5 };
    s.styleSet = "lines";
    s.theme = "Okyanus";
    s.properties.title = "Deneme Başlığı";
    s.properties.author = "Furkan";
    const back = await roundtrip(doc);
    const b = back.settings;
    expect(b.page).toEqual(s.page);
    expect(norm(b.header!)).toEqual(norm(s.header));
    expect(norm(b.footer!)).toEqual(norm(s.footer));
    expect(b.watermark).toEqual(s.watermark);
    expect(b.pageColor).toBe("#FFF2CC");
    expect(b.pageBorder).toEqual(s.pageBorder);
    expect(b.styleSet).toBe("lines");
    expect(b.theme).toBe("Okyanus");
    expect(b.properties.title).toBe("Deneme Başlığı");
    expect(b.properties.author).toBe("Furkan");
  });

  it("paketteki tüm XML parçaları iyi biçimli", async () => {
    const doc = docOf(p([t("a & b < c > \"d\" 'e'")]), { type: "table", content: [{ type: "tableRow", content: [{ type: "tableCell", content: [p("x")] }] }] });
    doc.settings.header = { type: "doc", content: [p("Üst")] };
    const bytes = await writeDocx(doc);
    const zip = await JSZip.loadAsync(bytes);
    const names = Object.keys(zip.files).filter((n) => n.endsWith(".xml") || n.endsWith(".rels"));
    expect(names).toContain("word/document.xml");
    expect(names).toContain("[Content_Types].xml");
    for (const n of names) {
      const xml = await zip.file(n)!.async("string");
      const parsed = new DOMParser().parseFromString(xml, "application/xml");
      expect(parsed.getElementsByTagName("parsererror").length, n).toBe(0);
    }
    // Tablo gövdenin sonundaysa ardından paragraf gelmeli.
    const docXml = await zip.file("word/document.xml")!.async("string");
    expect(docXml).toMatch(/<\/w:tbl><w:p(\/>|><\/w:p>)<w:sectPr>/);
  });
});

describe("DOCX bölümler", () => {
  const land = { ...defaultSettings().page, width: 16838, height: 11906, orientation: "landscape" as const };
  const cols = { ...defaultSettings().page, columns: 2, columnSpace: 567, columnLine: true };

  it("bölüm sonları, yönlendirme, sütunlar ve ilk sayfa farklı gidiş-dönüş", async () => {
    const doc: BelgeDocument = {
      settings: defaultSettings(),
      content: {
        type: "doc",
        content: [
          { type: "section", attrs: { kind: "nextPage", page: { ...defaultSettings().page, titlePage: true } }, content: [p("Birinci bölüm")] },
          { type: "section", attrs: { kind: "nextPage", page: land }, content: [p("Yatay bölüm")] },
          { type: "section", attrs: { kind: "continuous", page: cols }, content: [p("Sütun 1"), { type: "columnBreak" }, p("Sütun 2")] },
          { type: "section", attrs: { kind: "oddPage", page: null }, content: [{ type: "bulletList", content: [{ type: "listItem", content: [p("madde")] }] }] },
        ],
      },
    };
    const bytes = await writeDocx(doc);
    const xml = await (await JSZip.loadAsync(bytes)).file("word/document.xml")!.async("string");
    expect(xml.match(/<w:sectPr>/g)?.length).toBe(4);
    expect(xml).toContain('<w:type w:val="continuous"/>');
    expect(xml).toContain('<w:type w:val="oddPage"/>');
    expect(xml).toContain('w:orient="landscape"');
    expect(xml).toContain('w:num="2"');
    expect(xml).toContain('w:sep="1"');
    expect(xml).toContain("<w:titlePg/>");
    expect(xml).toContain('<w:br w:type="column"/>');
    // sectPr son paragrafın pPr'ında (ayrı boş paragraf yok)
    expect(xml).toMatch(/Birinci bölüm<\/w:t><\/w:r><\/w:p>/);
    const back = await readDocx(bytes);
    expect(norm(back.content)).toEqual(norm(doc.content));
  });

  it("ilk sayfa ve çift sayfa üst/alt bilgileri", async () => {
    const doc = docOf(p("Gövde"));
    doc.settings.header = { type: "doc", content: [p("Tek üst")] };
    doc.settings.headerFirst = { type: "doc", content: [p("İlk üst")] };
    doc.settings.headerEven = { type: "doc", content: [p("Çift üst")] };
    doc.settings.footerEven = { type: "doc", content: [p("Çift alt")] };
    doc.settings.oddEven = true;
    doc.settings.page = { ...doc.settings.page, titlePage: true };
    const bytes = await writeDocx(doc);
    const zip = await JSZip.loadAsync(bytes);
    expect(await zip.file("word/settings.xml")!.async("string")).toContain("<w:evenAndOddHeaders/>");
    const xml = await zip.file("word/document.xml")!.async("string");
    expect(xml).toContain('w:type="first"');
    expect(xml).toContain('w:type="even"');
    const back = await readDocx(bytes);
    expect(back.settings.oddEven).toBe(true);
    expect(back.settings.page.titlePage).toBe(true);
    expect(JSON.stringify(back.settings.headerFirst)).toContain("İlk üst");
    expect(JSON.stringify(back.settings.headerEven)).toContain("Çift üst");
    expect(JSON.stringify(back.settings.footerEven)).toContain("Çift alt");
    expect(JSON.stringify(back.settings.header)).toContain("Tek üst");
  });

  it("zengin dipnot metni", async () => {
    const rich = [p([t("Kalın", [{ type: "bold" }]), t(" ve "), t("italik", [{ type: "italic" }])])];
    const content = [p([t("Metin"), { type: "note", attrs: { id: "n1", kind: "footnote", text: "Kalın ve italik", content: rich } }])];
    const back = await roundtrip(docOf(...content));
    const s = JSON.stringify(back.content);
    expect(s).toContain('"bold"');
    expect(s).toContain('"italic"');
    expect(s).toContain("Kalın ve italik");
  });
});

describe("DOCX metin kutusu", () => {
  it("konum, boyut, kenarlık, dolgu ve zengin içerik gidiş-dönüş", async () => {
    const content = [
      { type: "textBox", attrs: { x: 120, y: 30, w: 260, h: 140, border: "#C00000", borderWidth: 1.5, fill: "#DAE9F7" }, content: [p([t("Kutu ", []), t("kalın", [{ type: "bold" }])])] },
      p("Kutunun bağlı olduğu paragraf"),
      { type: "textBox", attrs: { x: 0, y: 0, w: 200, h: 100, border: null, borderWidth: 0.75, fill: null }, content: [p("Kenarsız")] },
      p("İkinci çapa"),
      { type: "table", content: [{ type: "tableRow", content: [{ type: "tableCell", content: [p("hücre")] }] }] },
      { type: "paragraph" },
    ];
    const bytes = await writeDocx(docOf(...content));
    const xml = await (await JSZip.loadAsync(bytes)).file("word/document.xml")!.async("string");
    expect(xml).toContain("<wps:txbx><w:txbxContent>");
    expect(xml).toContain("<v:textbox><w:txbxContent>");
    // Çapa sonraki paragrafın içinde (ayrı boş paragraf yok)
    expect(xml).toMatch(/<\/mc:AlternateContent><\/w:r><w:r><w:t xml:space="preserve">Kutunun/);
    const back = await readDocx(bytes);
    expect(norm(back.content)).toEqual(norm({ type: "doc", content }));
  });
});
