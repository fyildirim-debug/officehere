// @vitest-environment jsdom
// Tablo soldan girintisi (indent, px): HTML, DOCX (w:tblInd), ODT (fo:margin-left) ve RTF (\trleft) gidiş-dönüşü.
import { describe, expect, it } from "vitest";
import JSZip from "jszip";
import type { JSONContent } from "@tiptap/core";
import { DOMSerializer, Node as PMNode } from "@tiptap/pm/model";
import { belgeSchema, htmlToDoc, readDocx, readOdt, readRtf, writeDocx, writeOdt, writeRtf } from "../index";
import { normalizeDoc } from "../../model/document";
import { all, docOf, p } from "./bicim-yardimci";

const cell = (s: string, w: number): JSONContent => ({ type: "tableCell", attrs: { colspan: 1, rowspan: 1, colwidth: [w] }, content: [p(s)] });
const table = (attrs: Record<string, unknown>): JSONContent => ({
  type: "table",
  attrs: { tableStyle: "grid", ...attrs },
  content: [
    { type: "tableRow", content: [cell("A1", 120), cell("B1", 160)] },
    { type: "tableRow", content: [cell("A2", 120), cell("B2", 160)] },
  ],
});
const tableAttrs = (c: JSONContent) => all(c, "table")[0]?.attrs ?? {};

describe("Tablo girintisi", () => {
  it("HTML: data-indent okunur, yazılırken sola yaslı tabloya margin-left eklenir", () => {
    const doc = htmlToDoc('<table data-indent="48"><tbody><tr><td><p>x</p></td></tr></tbody></table>');
    expect(tableAttrs(doc).indent).toBe(48);
    const node = PMNode.fromJSON(belgeSchema(), normalizeDoc({ type: "doc", content: [table({ indent: 48 })] }));
    const host = document.createElement("div");
    host.appendChild(DOMSerializer.fromSchema(belgeSchema()).serializeFragment(node.content));
    const el = host.querySelector("table")!;
    expect(el.getAttribute("data-indent")).toBe("48");
    expect(el.style.marginLeft).toBe("48px");
    // Ortalı tabloda girinti yok sayılır (otomatik kenar boşluğu ortalar).
    const centered = PMNode.fromJSON(belgeSchema(), normalizeDoc({ type: "doc", content: [table({ indent: 48, align: "center" })] }));
    const host2 = document.createElement("div");
    host2.appendChild(DOMSerializer.fromSchema(belgeSchema()).serializeFragment(centered.content));
    expect(host2.querySelector("table")!.style.marginLeft).toBe("");
  });

  it("DOCX: w:tblInd (twip) yazılır ve okunur", async () => {
    const bytes = await writeDocx(docOf(table({ indent: 64 })));
    const xml = await (await JSZip.loadAsync(bytes)).file("word/document.xml")!.async("string");
    expect(xml).toContain('<w:tblInd w:w="960" w:type="dxa"/>');
    // Öğe sırası: jc'den sonra, tblLayout'tan önce (şema sırası).
    expect(xml.indexOf("<w:tblInd")).toBeLessThan(xml.indexOf("<w:tblLayout"));
    const back = await readDocx(bytes);
    expect(tableAttrs(back.content)).toMatchObject({ indent: 64, align: null });
  });

  it("DOCX: ortalı tabloda girinti yazılmaz; girintisiz tablo öznitelik taşımaz", async () => {
    const bytes = await writeDocx(docOf(table({ indent: 64, align: "center" }), table({})));
    const xml = await (await JSZip.loadAsync(bytes)).file("word/document.xml")!.async("string");
    expect(xml).not.toContain("<w:tblInd");
    const back = await readDocx(bytes);
    const [a, b] = all(back.content, "table");
    expect(a.attrs?.align).toBe("center");
    expect(a.attrs?.indent ?? 0).toBe(0);
    expect(b.attrs?.indent ?? 0).toBe(0);
  });

  it("ODT: fo:margin-left ile gidiş-dönüş", async () => {
    const bytes = await writeOdt(docOf(table({ indent: 96 })));
    const xml = await (await JSZip.loadAsync(bytes)).file("content.xml")!.async("string");
    expect(xml).toMatch(/table:align="left" fo:margin-left="2\.54cm"/);
    const back = await readOdt(bytes);
    expect(tableAttrs(back.content)).toMatchObject({ indent: 96 });
  });

  it("RTF: \\trleft ile gidiş-dönüş; \\trqc ortalı tablo olarak okunur", () => {
    const out = writeRtf(docOf(table({ indent: 40 })));
    expect(out).toContain("\\trleft600");
    const back = readRtf(new TextEncoder().encode(out));
    const a = tableAttrs(back.content);
    expect(a.indent).toBe(40);
    expect(all(back.content, "tableCell").map((c) => c.attrs?.colwidth)).toEqual([[120], [160], [120], [160]]);
    const centered = readRtf(new TextEncoder().encode(writeRtf(docOf(table({ align: "center" })))));
    expect(tableAttrs(centered.content)).toMatchObject({ align: "center" });
  });
});
