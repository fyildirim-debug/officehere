// @vitest-environment jsdom
// Tablo satır yüksekliği (height, px; heightRule "exact"): HTML, DOCX (w:trHeight), ODT (style:min-row-height /
// style:row-height) ve RTF (\trrh) gidiş-dönüşü.
import { describe, expect, it } from "vitest";
import JSZip from "jszip";
import type { JSONContent } from "@tiptap/core";
import { DOMSerializer, Node as PMNode } from "@tiptap/pm/model";
import { belgeSchema, htmlToDoc, readDocx, readOdt, readRtf, writeDocx, writeOdt, writeRtf } from "../index";
import { normalizeDoc } from "../../model/document";
import { all, docOf, p } from "./bicim-yardimci";

const cell = (s: string): JSONContent => ({ type: "tableCell", attrs: { colspan: 1, rowspan: 1, colwidth: [120] }, content: [p(s)] });
const row = (s: string, attrs?: Record<string, unknown>): JSONContent => ({ type: "tableRow", ...(attrs ? { attrs } : {}), content: [cell(`${s}1`), cell(`${s}2`)] });
const table = (): JSONContent => ({
  type: "table",
  attrs: { tableStyle: "grid" },
  content: [row("A", { height: 60 }), row("B"), row("C", { height: 37.8, heightRule: "exact" })],
});
const heights = (c: JSONContent) => all(c, "tableRow").map((r) => [r.attrs?.height ?? null, r.attrs?.heightRule ?? null]);

describe("Tablo satır yüksekliği", () => {
  it("HTML: <tr> CSS yüksekliği (px ya da pt) okunur, yazılırken data-height ve height eklenir", () => {
    const doc = htmlToDoc('<table><tbody><tr style="height:22.5pt"><td><p>x</p></td></tr><tr data-height="48"><td><p>y</p></td></tr><tr><td><p>z</p></td></tr></tbody></table>');
    expect(heights(doc)).toEqual([
      [30, null],
      [48, null],
      [null, null],
    ]);
    const node = PMNode.fromJSON(belgeSchema(), normalizeDoc({ type: "doc", content: [table()] }));
    const host = document.createElement("div");
    host.appendChild(DOMSerializer.fromSchema(belgeSchema()).serializeFragment(node.content));
    const trs = Array.from(host.querySelectorAll("tr"));
    expect(trs.map((tr) => tr.style.height)).toEqual(["60px", "", "37.8px"]);
    expect(trs.map((tr) => tr.getAttribute("data-height-rule"))).toEqual([null, null, "exact"]);
  });

  it("DOCX: w:trHeight (twip, hRule) yazılır ve okunur; hRule=auto yok sayılır", async () => {
    const bytes = await writeDocx(docOf(table()));
    const xml = await (await JSZip.loadAsync(bytes)).file("word/document.xml")!.async("string");
    expect(xml).toContain('<w:tr><w:trPr><w:trHeight w:val="900"/></w:trPr>');
    expect(xml).toContain('<w:trHeight w:val="567" w:hRule="exact"/>');
    // Yüksekliksiz satır trPr taşımaz.
    expect(xml.match(/<w:trPr>/g)).toHaveLength(2);
    const back = await readDocx(bytes);
    expect(heights(back.content)).toEqual([
      [60, null],
      [null, null],
      [37.8, "exact"],
    ]);
    // Başlık satırı yinelenirken yükseklik de korunur (aynı trPr içinde).
    const hdr = docOf({ ...table(), attrs: { tableStyle: "grid", repeatHeader: true } });
    const xml2 = await (await JSZip.loadAsync(await writeDocx(hdr))).file("word/document.xml")!.async("string");
    expect(xml2).toContain('<w:trPr><w:trHeight w:val="900"/><w:tblHeader/></w:trPr>');
    // Word'ün "auto" kuralı yükseklik vermez.
    const auto = await writeDocx(docOf(table()));
    const zip = await JSZip.loadAsync(auto);
    const doc = (await zip.file("word/document.xml")!.async("string")).replace('<w:trHeight w:val="900"/>', '<w:trHeight w:val="900" w:hRule="auto"/>');
    zip.file("word/document.xml", doc);
    const autoBack = await readDocx(await zip.generateAsync({ type: "uint8array" }));
    expect(heights(autoBack.content)[0]).toEqual([null, null]);
  });

  it("ODT: style:min-row-height / style:row-height ile gidiş-dönüş", async () => {
    const bytes = await writeOdt(docOf(table()));
    const xml = await (await JSZip.loadAsync(bytes)).file("content.xml")!.async("string");
    expect(xml).toContain('style:min-row-height="1.5875cm"');
    expect(xml).toContain('style:row-height="1.0001cm"');
    const back = await readOdt(bytes);
    expect(heights(back.content)).toEqual([
      [60, null],
      [null, null],
      [37.8, "exact"],
    ]);
  });

  it("RTF: \\trrh ile gidiş-dönüş (eksi değer tam yükseklik)", () => {
    const out = writeRtf(docOf(table()));
    expect(out).toContain("\\trrh900");
    expect(out).toContain("\\trrh-567");
    const back = readRtf(new TextEncoder().encode(out));
    expect(heights(back.content)).toEqual([
      [60, null],
      [null, null],
      [37.8, "exact"],
    ]);
  });
});
