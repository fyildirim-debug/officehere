// DOCX parça yazıcısı (partWriter) anlık görüntüleri: paragraf, koşu ve tablo XML'i. Geçiş katmanının öğe
// düzeyi değişikliklerinden (opak öğeler, ek özellikler, kimlik yeniden numaralama) önce bugünkü çıktıyı sabitler;
// bilinçli bir yazıcı değişikliğinde anlık görüntü gözden geçirilerek güncellenir (vitest -u).
import { afterAll, beforeAll, describe, expect, it, vi } from "vitest";
import type { JSONContent } from "@tiptap/core";
import { defaultSettings } from "../../model/document";
import { Ctx, PartWriter, Rels } from "../docx/partWriter";

const PNG = "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAQAAAADCAIAAAA7ljmRAAAAEElEQVR4nGM4IScHRww4OQD1xwwxB+utmQAAAABJRU5ErkJggg==";

const t = (text: string, marks?: JSONContent["marks"]): JSONContent => ({ type: "text", text, ...(marks ? { marks } : {}) });
const p = (content: JSONContent[] | string, attrs: Record<string, unknown> = {}): JSONContent => ({ type: "paragraph", attrs, content: typeof content === "string" ? [t(content)] : content });

function writer() {
  const ctx = new Ctx(defaultSettings());
  const rels = new Rels();
  return { ctx, rels, pw: new PartWriter(ctx, rels, 9026) };
}

/** Blokları yazar; XML ve parçanın ilişkileri. */
function blocks(nodes: JSONContent[]) {
  const { pw, rels, ctx } = writer();
  return { xml: pw.blocks(nodes), rels: rels.xml(), ctx };
}

beforeAll(() => {
  vi.useFakeTimers();
  vi.setSystemTime(new Date("2026-03-04T05:06:07Z"));
});
afterAll(() => {
  vi.useRealTimers();
});

describe("partWriter: paragraf", () => {
  it("düz paragraf, başlık ve boş paragraf", () => {
    expect(blocks([p("Merhaba dünya"), { type: "heading", attrs: { level: 2 }, content: [t("Başlık")] }, { type: "paragraph" }]).xml).toMatchInlineSnapshot(`"<w:p><w:r><w:t xml:space="preserve">Merhaba dünya</w:t></w:r></w:p><w:p><w:pPr><w:pStyle w:val="Heading2"/></w:pPr><w:r><w:t xml:space="preserve">Başlık</w:t></w:r></w:p><w:p></w:p>"`);
  });

  it("stil, hizalama, aralık, girinti ve satır aralığı", () => {
    expect(
      blocks([
        p("Alıntı", { styleId: "Quote", textAlign: "justify", spaceBefore: 6, spaceAfter: 12, lineHeight: 1.5, indentLeft: 36, indentRight: 18, indentFirst: -18 }),
        p("Tam", { textAlign: "center", lineHeight: 14, lineRule: "exact", indentFirst: 24 }),
        p("En az", { textAlign: "left", lineHeight: 12, lineRule: "atLeast" }),
      ]).xml,
    ).toMatchInlineSnapshot(`"<w:p><w:pPr><w:pStyle w:val="Quote"/><w:spacing w:before="120" w:after="240" w:line="360" w:lineRule="auto"/><w:ind w:left="720" w:right="360" w:hanging="360"/><w:jc w:val="both"/></w:pPr><w:r><w:t xml:space="preserve">Alıntı</w:t></w:r></w:p><w:p><w:pPr><w:spacing w:line="280" w:lineRule="exact"/><w:ind w:firstLine="480"/><w:jc w:val="center"/></w:pPr><w:r><w:t xml:space="preserve">Tam</w:t></w:r></w:p><w:p><w:pPr><w:spacing w:line="240" w:lineRule="atLeast"/><w:jc w:val="left"/></w:pPr><w:r><w:t xml:space="preserve">En az</w:t></w:r></w:p>"`);
  });

  it("kenarlık, gölgelendirme, sekme durakları ve sayfalama seçenekleri", () => {
    const borders = { top: { style: "single", size: 1, color: "#FF0000", space: 1 }, bottom: { style: "double", size: 1.5, color: "#000000", space: 2 }, left: null, right: null };
    expect(
      blocks([
        p("Kutulu", {
          borders,
          shading: "#EEEEEE",
          tabs: [
            { pos: 72, align: "left", leader: "none" },
            { pos: 216, align: "right", leader: "dot" },
          ],
          keepNext: true,
          keepLines: true,
          pageBreakBefore: true,
          suppressLineNumbers: true,
          outlineLevel: 2,
        }),
      ]).xml,
    ).toMatchInlineSnapshot(`"<w:p><w:pPr><w:keepNext/><w:keepLines/><w:pageBreakBefore/><w:suppressLineNumbers/><w:pBdr><w:top w:val="single" w:sz="8" w:space="1" w:color="FF0000"/><w:bottom w:val="double" w:sz="12" w:space="2" w:color="000000"/></w:pBdr><w:shd w:val="clear" w:color="auto" w:fill="EEEEEE"/><w:tabs><w:tab w:val="left" w:pos="1440"/><w:tab w:val="right" w:leader="dot" w:pos="4320"/></w:tabs><w:outlineLvl w:val="1"/></w:pPr><w:r><w:t xml:space="preserve">Kutulu</w:t></w:r></w:p>"`);
  });

  it("listeler: madde, numaralı, iç içe ve ikinci paragraf", () => {
    const li = (...content: JSONContent[]): JSONContent => ({ type: "listItem", content });
    const { xml, ctx } = blocks([
      { type: "bulletList", attrs: { listStyle: "square" }, content: [li(p("Bir")), li(p("İki"), { type: "orderedList", attrs: { start: 3 }, content: [li(p("İç"))] })] },
      { type: "orderedList", attrs: { start: 1, multilevel: "outline" }, content: [li(p("Birinci"), p("Devam"))] },
    ]);
    expect(xml).toMatchInlineSnapshot(`"<w:p><w:pPr><w:pStyle w:val="ListParagraph"/><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr><w:r><w:t xml:space="preserve">Bir</w:t></w:r></w:p><w:p><w:pPr><w:pStyle w:val="ListParagraph"/><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr><w:r><w:t xml:space="preserve">İki</w:t></w:r></w:p><w:p><w:pPr><w:pStyle w:val="ListParagraph"/><w:numPr><w:ilvl w:val="1"/><w:numId w:val="2"/></w:numPr></w:pPr><w:r><w:t xml:space="preserve">İç</w:t></w:r></w:p><w:p><w:pPr><w:pStyle w:val="ListParagraph"/><w:numPr><w:ilvl w:val="0"/><w:numId w:val="3"/></w:numPr></w:pPr><w:r><w:t xml:space="preserve">Birinci</w:t></w:r></w:p><w:p><w:pPr><w:pStyle w:val="ListParagraph"/><w:ind w:left="720"/></w:pPr><w:r><w:t xml:space="preserve">Devam</w:t></w:r></w:p>"`);
    expect(ctx.nums).toMatchInlineSnapshot(`
      [
        {
          "abstractId": 0,
          "numId": 1,
          "start": 1,
        },
        {
          "abstractId": 1,
          "numId": 2,
          "start": 3,
        },
        {
          "abstractId": 2,
          "numId": 3,
          "start": 1,
        },
      ]
    `);
  });

  it("bölüm sonu paragrafı ve büyük ilk harf", () => {
    const { pw } = writer();
    const last = p("Bölüm sonu");
    pw.sectTarget = last;
    pw.sectXml = '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr>';
    const dc = p([t("Yazı", [{ type: "textStyle", attrs: { fontFamily: "Georgia" } }])], { dropCap: { mode: "drop", lines: 3, distance: 2 } });
    expect(pw.blocks([dc, last])).toMatchInlineSnapshot(`"<w:p><w:pPr><w:keepNext/><w:framePr w:dropCap="drop" w:lines="3" w:hSpace="40" w:wrap="around" w:vAnchor="text" w:hAnchor="text"/><w:spacing w:before="0" w:after="0" w:line="855" w:lineRule="exact"/><w:textAlignment w:val="baseline"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Georgia" w:hAnsi="Georgia" w:eastAsia="Georgia" w:cs="Georgia"/><w:sz w:val="103"/><w:szCs w:val="103"/></w:rPr><w:t xml:space="preserve">Y</w:t></w:r></w:p><w:p><w:r><w:rPr><w:rFonts w:ascii="Georgia" w:hAnsi="Georgia" w:eastAsia="Georgia" w:cs="Georgia"/></w:rPr><w:t xml:space="preserve">azı</w:t></w:r></w:p><w:p><w:pPr><w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr></w:pPr><w:r><w:t xml:space="preserve">Bölüm sonu</w:t></w:r></w:p>"`);
  });

  it("yatay çizgi, sayfa ve sütun sonu", () => {
    expect(blocks([{ type: "horizontalRule" }, { type: "pageBreak" }, { type: "columnBreak" }]).xml).toMatchInlineSnapshot(`"<w:p><w:pPr><w:pBdr><w:bottom w:val="single" w:sz="6" w:space="1" w:color="auto"/></w:pBdr></w:pPr></w:p><w:p><w:r><w:br w:type="page"/></w:r></w:p><w:p><w:r><w:br w:type="column"/></w:r></w:p>"`);
  });
});

describe("partWriter: koşu", () => {
  it("karakter biçimleri", () => {
    const ts = { fontFamily: "Arial", fontSize: "13pt", color: "#336699", backgroundColor: "#FFEEDD", caps: true, smallCaps: true, hidden: true, doubleStrike: true, shadow: true, outline: true, emboss: true, imprint: true, letterSpacing: 1.5 };
    expect(
      blocks([
        p([
          t("kalın", [{ type: "bold" }]),
          t("eğik", [{ type: "italic" }]),
          t("altı", [{ type: "underline", attrs: { type: "dashed", color: "#00FF00" } }]),
          t("üstü", [{ type: "strike" }]),
          t("üst", [{ type: "superscript" }]),
          t("alt", [{ type: "subscript" }]),
          t("sarı", [{ type: "highlight", attrs: { color: "#FFFF00" } }]),
          t("özel", [{ type: "highlight", attrs: { color: "#123456" } }]),
          t("stil", [{ type: "textStyle", attrs: ts }]),
          t("vurgu", [{ type: "charStyle", attrs: { styleId: "Strong" } }]),
          t("a\tb"),
          { type: "hardBreak" },
          t("son"),
        ]),
      ]).xml,
    ).toMatchInlineSnapshot(`"<w:p><w:r><w:rPr><w:b/><w:bCs/></w:rPr><w:t xml:space="preserve">kalın</w:t></w:r><w:r><w:rPr><w:i/><w:iCs/></w:rPr><w:t xml:space="preserve">eğik</w:t></w:r><w:r><w:rPr><w:u w:val="dash" w:color="00FF00"/></w:rPr><w:t xml:space="preserve">altı</w:t></w:r><w:r><w:rPr><w:strike/></w:rPr><w:t xml:space="preserve">üstü</w:t></w:r><w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:t xml:space="preserve">üst</w:t></w:r><w:r><w:rPr><w:vertAlign w:val="subscript"/></w:rPr><w:t xml:space="preserve">alt</w:t></w:r><w:r><w:rPr><w:highlight w:val="yellow"/></w:rPr><w:t xml:space="preserve">sarı</w:t></w:r><w:r><w:rPr><w:shd w:val="clear" w:color="auto" w:fill="123456"/></w:rPr><w:t xml:space="preserve">özel</w:t></w:r><w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:eastAsia="Arial" w:cs="Arial"/><w:caps/><w:smallCaps/><w:dstrike/><w:outline/><w:shadow/><w:emboss/><w:imprint/><w:vanish/><w:color w:val="336699"/><w:spacing w:val="30"/><w:sz w:val="26"/><w:szCs w:val="26"/><w:shd w:val="clear" w:color="auto" w:fill="FFEEDD"/></w:rPr><w:t xml:space="preserve">stil</w:t></w:r><w:r><w:rPr><w:rStyle w:val="Strong"/></w:rPr><w:t xml:space="preserve">vurgu</w:t></w:r><w:r><w:t xml:space="preserve">a</w:t><w:tab/><w:t xml:space="preserve">b</w:t></w:r><w:r><w:br/></w:r><w:r><w:t xml:space="preserve">son</w:t></w:r></w:p>"`);
  });

  it("köprüler, izlenen değişiklikler ve biçim değişikliği", () => {
    const { xml, rels } = blocks([
      p([
        t("dış", [{ type: "link", attrs: { href: "https://ornek.org/a?b=1&c=2" } }]),
        t("iç", [{ type: "link", attrs: { href: "#_Toc1" } }]),
        t("eklenen", [{ type: "insertion", attrs: { author: "Ayşe", date: "2025-01-02T03:04:05Z" } }]),
        t("silinen", [{ type: "deletion", attrs: { author: "", date: "" } }]),
        t("biçim", [
          { type: "bold" },
          { type: "formatChange", attrs: { author: "Ali", date: "2025-01-02T03:04:05Z", old: JSON.stringify([{ type: "italic" }]) } },
        ]),
      ]),
    ]);
    expect(xml).toMatchInlineSnapshot(`"<w:p><w:hyperlink r:id="rId1" w:history="1"><w:r><w:rPr><w:rStyle w:val="Hyperlink"/></w:rPr><w:t xml:space="preserve">dış</w:t></w:r></w:hyperlink><w:hyperlink w:anchor="_Toc1" w:history="1"><w:r><w:rPr><w:rStyle w:val="Hyperlink"/></w:rPr><w:t xml:space="preserve">iç</w:t></w:r></w:hyperlink><w:ins w:id="1" w:author="Ayşe" w:date="2025-01-02T03:04:05Z"><w:r><w:t xml:space="preserve">eklenen</w:t></w:r></w:ins><w:del w:id="2" w:author="Kullanıcı" w:date="2026-03-04T05:06:07Z"><w:r><w:delText xml:space="preserve">silinen</w:delText></w:r></w:del><w:r><w:rPr><w:b/><w:bCs/><w:rPrChange w:id="3" w:author="Ali" w:date="2025-01-02T03:04:05Z"><w:rPr><w:i/><w:iCs/></w:rPr></w:rPrChange></w:rPr><w:t xml:space="preserve">biçim</w:t></w:r></w:p>"`);
    expect(rels).toMatchInlineSnapshot(`
      "<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
      <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://ornek.org/a?b=1&amp;c=2" TargetMode="External"/></Relationships>"
    `);
  });

  it("açıklama aralığı, yer işareti ve resim yazısı başvurusu", () => {
    const c = [{ type: "comment", attrs: { id: "c1" } }];
    expect(
      blocks([
        p([t("önce "), t("açıklamalı", c), t(" sonra"), { type: "bookmark", attrs: { name: "isaret" } }]),
        p([{ type: "bookmark", attrs: { name: "_Ref123" } }, t("Şekil "), { type: "seqField", attrs: { label: "Şekil", num: 2 } }, t(": açıklama")]),
      ]).xml,
    ).toMatchInlineSnapshot(`"<w:p><w:r><w:t xml:space="preserve">önce </w:t></w:r><w:commentRangeStart w:id="0"/><w:r><w:t xml:space="preserve">açıklamalı</w:t></w:r><w:commentRangeEnd w:id="0"/><w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr><w:commentReference w:id="0"/></w:r><w:r><w:t xml:space="preserve"> sonra</w:t></w:r><w:bookmarkStart w:id="0" w:name="isaret"/><w:bookmarkEnd w:id="0"/></w:p><w:p><w:bookmarkStart w:id="1" w:name="_Ref123"/><w:r><w:t xml:space="preserve">Şekil </w:t></w:r><w:fldSimple w:instr=" SEQ Şekil \\* ARABIC "><w:r><w:t xml:space="preserve">2</w:t></w:r></w:fldSimple><w:bookmarkEnd w:id="1"/><w:r><w:t xml:space="preserve">: açıklama</w:t></w:r></w:p>"`);
  });

  it("alanlar, dipnot, denklem ve resim", () => {
    const { xml, rels, ctx } = blocks([
      p([
        { type: "pageField", attrs: { field: "PAGE" } },
        { type: "pageField", attrs: { field: "NUMPAGES" } },
        { type: "pageField", attrs: { field: "DATE", format: "dd.MM.yyyy" } },
        { type: "crossRef", attrs: { ref: "_Ref123", kind: "page", link: true, text: "3" } },
        { type: "citation", attrs: { tag: "Yaz01", pages: "12", text: "(Yazar, 2001)" } },
        { type: "mergeField", attrs: { name: "Ad" } },
        { type: "note", attrs: { kind: "footnote", text: "Dipnot metni", content: null } },
        { type: "equation", attrs: { latex: "x^2", display: false } },
        { type: "image", attrs: { src: PNG, width: 40, height: 30, alt: "Küçük resim" } },
        { type: "image", attrs: { src: PNG, width: 40, height: 30, wrap: "square", floatSide: "right", posX: 10, posY: 5 } },
      ]),
    ]);
    expect(xml).toMatchInlineSnapshot(`"<w:p><w:fldSimple w:instr=" PAGE \\* MERGEFORMAT "><w:r><w:t xml:space="preserve">1</w:t></w:r></w:fldSimple><w:fldSimple w:instr=" NUMPAGES \\* MERGEFORMAT "><w:r><w:t xml:space="preserve">1</w:t></w:r></w:fldSimple><w:fldSimple w:instr=" DATE \\@ &quot;dd.MM.yyyy&quot; \\* MERGEFORMAT "><w:r><w:t xml:space="preserve">04.03.2026</w:t></w:r></w:fldSimple><w:fldSimple w:instr=" PAGEREF _Ref123 \\h "><w:r><w:t xml:space="preserve">3</w:t></w:r></w:fldSimple><w:fldSimple w:instr=" CITATION Yaz01 \\p 12 \\l 1055 "><w:r><w:t xml:space="preserve">(Yazar, 2001)</w:t></w:r></w:fldSimple><w:fldSimple w:instr=" MERGEFIELD Ad "><w:r><w:t xml:space="preserve">«Ad»</w:t></w:r></w:fldSimple><w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteReference w:id="1"/></w:r><m:oMath><m:sSup><m:e><m:r><m:t xml:space="preserve">x</m:t></m:r></m:e><m:sup><m:r><m:t xml:space="preserve">2</m:t></m:r></m:sup></m:sSup></m:oMath><w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="381000" cy="285750"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:docPr id="1" name="Resim 1" descr="Küçük resim"/><wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/></wp:cNvGraphicFramePr><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:nvPicPr><pic:cNvPr id="1" name="image1.png" descr="Küçük resim"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="rId1"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="381000" cy="285750"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r><w:r><w:drawing><wp:anchor distT="0" distB="0" distL="114300" distR="114300" simplePos="0" relativeHeight="251658242" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1"><wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="column"><wp:posOffset>5255260</wp:posOffset></wp:positionH><wp:positionV relativeFrom="paragraph"><wp:posOffset>47625</wp:posOffset></wp:positionV><wp:extent cx="381000" cy="285750"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:wrapSquare wrapText="bothSides"/><wp:docPr id="2" name="Resim 2"/><wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/></wp:cNvGraphicFramePr><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:nvPicPr><pic:cNvPr id="2" name="image2.png"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="rId2"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="381000" cy="285750"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:anchor></w:drawing></w:r></w:p>"`);
    expect(rels).toMatchInlineSnapshot(`
      "<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
      <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.png"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image2.png"/></Relationships>"
    `);
    expect({ media: ctx.media.map((m) => m.name), footnotes: ctx.footnotes, drawingId: ctx.drawingId }).toMatchInlineSnapshot(`
      {
        "drawingId": 3,
        "footnotes": [
          {
            "content": null,
            "text": "Dipnot metni",
          },
        ],
        "media": [
          "image1.png",
          "image2.png",
        ],
      }
    `);
  });
});

describe("partWriter: tablo", () => {
  const cell = (text: string, attrs: Record<string, unknown> = {}): JSONContent => ({ type: "tableCell", attrs: { colspan: 1, rowspan: 1, colwidth: null, ...attrs }, content: [p(text)] });
  const row = (...cells: JSONContent[]): JSONContent => ({ type: "tableRow", content: cells });

  it("ızgara, birleştirme, hücre zemini ve dikey hizalama", () => {
    const table: JSONContent = {
      type: "table",
      attrs: { tableStyle: "grid", headerRow: true, bandedRows: true, firstColumn: false, totalRow: false, align: null, repeatHeader: true, cellMargins: null },
      content: [
        row(cell("A1", { colwidth: [100] }), cell("B1", { colwidth: [150] }), cell("C1", { colwidth: [120] })),
        row(cell("A2", { colspan: 2, colwidth: [100, 150], background: "#DDEEFF" }), cell("C2", { rowspan: 2, colwidth: [120], verticalAlign: "middle" })),
        row(cell("A3", { colwidth: [100] }), cell("B3", { colwidth: [150], verticalAlign: "bottom" })),
      ],
    };
    expect(blocks([table]).xml).toMatchInlineSnapshot(`"<w:tbl><w:tblPr><w:tblStyle w:val="TableGrid"/><w:tblW w:w="5550" w:type="dxa"/><w:tblLayout w:type="fixed"/><w:tblLook w:val="0420" w:firstRow="1" w:lastRow="0" w:firstColumn="0" w:lastColumn="0" w:noHBand="0" w:noVBand="1"/></w:tblPr><w:tblGrid><w:gridCol w:w="1500"/><w:gridCol w:w="2250"/><w:gridCol w:w="1800"/></w:tblGrid><w:tr><w:trPr><w:tblHeader/></w:trPr><w:tc><w:tcPr><w:tcW w:w="1500" w:type="dxa"/></w:tcPr><w:p><w:r><w:t xml:space="preserve">A1</w:t></w:r></w:p></w:tc><w:tc><w:tcPr><w:tcW w:w="2250" w:type="dxa"/></w:tcPr><w:p><w:r><w:t xml:space="preserve">B1</w:t></w:r></w:p></w:tc><w:tc><w:tcPr><w:tcW w:w="1800" w:type="dxa"/></w:tcPr><w:p><w:r><w:t xml:space="preserve">C1</w:t></w:r></w:p></w:tc></w:tr><w:tr><w:tc><w:tcPr><w:tcW w:w="3750" w:type="dxa"/><w:gridSpan w:val="2"/><w:shd w:val="clear" w:color="auto" w:fill="DDEEFF"/></w:tcPr><w:p><w:r><w:t xml:space="preserve">A2</w:t></w:r></w:p></w:tc><w:tc><w:tcPr><w:tcW w:w="1800" w:type="dxa"/><w:vMerge w:val="restart"/><w:vAlign w:val="center"/></w:tcPr><w:p><w:r><w:t xml:space="preserve">C2</w:t></w:r></w:p></w:tc></w:tr><w:tr><w:tc><w:tcPr><w:tcW w:w="1500" w:type="dxa"/></w:tcPr><w:p><w:r><w:t xml:space="preserve">A3</w:t></w:r></w:p></w:tc><w:tc><w:tcPr><w:tcW w:w="2250" w:type="dxa"/><w:vAlign w:val="bottom"/></w:tcPr><w:p><w:r><w:t xml:space="preserve">B3</w:t></w:r></w:p></w:tc><w:tc><w:tcPr><w:tcW w:w="1800" w:type="dxa"/><w:vMerge/></w:tcPr><w:p/></w:tc></w:tr></w:tbl>"`);
  });

  it("genişliksiz tablo, ortalama, girinti ve hücre kenar boşlukları", () => {
    const table = (attrs: Record<string, unknown>): JSONContent => ({
      type: "table",
      attrs: { tableStyle: "none", headerRow: false, bandedRows: false, firstColumn: true, totalRow: true, repeatHeader: false, cellMargins: null, ...attrs },
      content: [row(cell("x"), cell("y"))],
    });
    expect(blocks([table({ align: "center", cellMargins: { top: 1, bottom: 2, left: 3, right: 4 } }), table({ align: null, indent: 20 })]).xml).toMatchInlineSnapshot(`"<w:tbl><w:tblPr><w:tblStyle w:val="TableNormal"/><w:tblW w:w="5000" w:type="pct"/><w:jc w:val="center"/><w:tblLayout w:type="fixed"/><w:tblCellMar><w:top w:w="20" w:type="dxa"/><w:left w:w="60" w:type="dxa"/><w:bottom w:w="40" w:type="dxa"/><w:right w:w="80" w:type="dxa"/></w:tblCellMar><w:tblLook w:val="06C0" w:firstRow="0" w:lastRow="1" w:firstColumn="1" w:lastColumn="0" w:noHBand="1" w:noVBand="1"/></w:tblPr><w:tblGrid><w:gridCol w:w="4513"/><w:gridCol w:w="4513"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="4513" w:type="dxa"/></w:tcPr><w:p><w:r><w:t xml:space="preserve">x</w:t></w:r></w:p></w:tc><w:tc><w:tcPr><w:tcW w:w="4513" w:type="dxa"/></w:tcPr><w:p><w:r><w:t xml:space="preserve">y</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:tbl><w:tblPr><w:tblStyle w:val="TableNormal"/><w:tblW w:w="5000" w:type="pct"/><w:tblInd w:w="300" w:type="dxa"/><w:tblLayout w:type="fixed"/><w:tblLook w:val="06C0" w:firstRow="0" w:lastRow="1" w:firstColumn="1" w:lastColumn="0" w:noHBand="1" w:noVBand="1"/></w:tblPr><w:tblGrid><w:gridCol w:w="4513"/><w:gridCol w:w="4513"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="4513" w:type="dxa"/></w:tcPr><w:p><w:r><w:t xml:space="preserve">x</w:t></w:r></w:p></w:tc><w:tc><w:tcPr><w:tcW w:w="4513" w:type="dxa"/></w:tcPr><w:p><w:r><w:t xml:space="preserve">y</w:t></w:r></w:p></w:tc></w:tr></w:tbl>"`);
  });
});

describe("partWriter: diğer bloklar", () => {
  it("metin kutusu sonraki paragrafa bağlanır, İçindekiler", () => {
    const box: JSONContent = { type: "textBox", attrs: { x: 10, y: 20, w: 200, h: 80, border: "#000000", fill: "#FFFFCC", borderWidth: 1 }, content: [p("Kutu içi")] };
    const toc: JSONContent = { type: "toc", attrs: { title: "İçindekiler", entries: [{ level: 1, text: "Giriş", page: 1, pos: 0 }, { level: 2, text: "Alt & bölüm", page: 3, pos: 0 }] } };
    expect(blocks([box, p("Çapa paragrafı"), toc]).xml).toMatchInlineSnapshot(`"<w:p><w:r><mc:AlternateContent><mc:Choice Requires="wps"><w:drawing><wp:anchor distT="0" distB="0" distL="114300" distR="114300" simplePos="0" relativeHeight="251659265" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1"><wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="column"><wp:posOffset>95250</wp:posOffset></wp:positionH><wp:positionV relativeFrom="paragraph"><wp:posOffset>190500</wp:posOffset></wp:positionV><wp:extent cx="1905000" cy="762000"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:wrapNone/><wp:docPr id="1" name="Metin Kutusu 1"/><wp:cNvGraphicFramePr/><a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp><wps:cNvSpPr txBox="1"/><wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1905000" cy="762000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="FFFFCC"/></a:solidFill><a:ln w="12700"><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:ln></wps:spPr><wps:txbx><w:txbxContent><w:p><w:r><w:t xml:space="preserve">Kutu içi</w:t></w:r></w:p></w:txbxContent></wps:txbx><wps:bodyPr rot="0" vert="horz" wrap="square" lIns="91440" tIns="45720" rIns="91440" bIns="45720" anchor="t" anchorCtr="0"><a:noAutofit/></wps:bodyPr></wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing></mc:Choice><mc:Fallback><w:pict><v:shapetype id="_x0000_t202" coordsize="21600,21600" o:spt="202" path="m,l,21600r21600,l21600,xe"><v:stroke joinstyle="miter"/><v:path gradientshapeok="t" o:connecttype="rect"/></v:shapetype><v:shape id="Metin Kutusu 1" o:spid="_x0000_s1025" type="#_x0000_t202" style="position:absolute;margin-left:7.5pt;margin-top:15pt;width:150pt;height:60pt;z-index:1;mso-position-horizontal-relative:text;mso-position-vertical-relative:text" fillcolor="#FFFFCC" strokecolor="#000000" strokeweight="1pt"><v:textbox><w:txbxContent><w:p><w:r><w:t xml:space="preserve">Kutu içi</w:t></w:r></w:p></w:txbxContent></v:textbox></v:shape></w:pict></mc:Fallback></mc:AlternateContent></w:r><w:r><w:t xml:space="preserve">Çapa paragrafı</w:t></w:r></w:p><w:sdt><w:sdtPr><w:docPartObj><w:docPartGallery w:val="Table of Contents"/><w:docPartUnique/></w:docPartObj></w:sdtPr><w:sdtContent><w:p><w:pPr><w:pStyle w:val="TOCHeading"/></w:pPr><w:r><w:t xml:space="preserve">İçindekiler</w:t></w:r></w:p><w:p><w:pPr><w:pStyle w:val="TOC1"/><w:tabs><w:tab w:val="right" w:leader="dot" w:pos="9026"/></w:tabs></w:pPr><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> TOC \\o "1-3" \\h \\z \\u </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t xml:space="preserve">Giriş</w:t></w:r><w:r><w:tab/></w:r><w:r><w:t>1</w:t></w:r></w:p><w:p><w:pPr><w:pStyle w:val="TOC2"/><w:tabs><w:tab w:val="right" w:leader="dot" w:pos="9026"/></w:tabs></w:pPr><w:r><w:t xml:space="preserve">Alt &amp; bölüm</w:t></w:r><w:r><w:tab/></w:r><w:r><w:t>3</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p></w:sdtContent></w:sdt>"`);
  });
});
