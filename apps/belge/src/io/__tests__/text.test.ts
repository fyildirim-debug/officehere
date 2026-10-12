// @vitest-environment jsdom
import { describe, expect, it } from "vitest";
import { decodeText, docToText, extractHtmlBody, htmlDocument, textToDoc } from "../text";
import { htmlToDoc } from "../index";
import { defaultSettings } from "../../model/document";

describe("düz metin", () => {
  it("UTF-8, BOM ve Windows-1254 çözümleme", () => {
    expect(decodeText(new TextEncoder().encode("﻿ığüşöç İĞÜŞÖÇ"))).toBe("ığüşöç İĞÜŞÖÇ");
    // Windows-1254: ğ=0xF0, ı=0xFD, ş=0xFE, İ=0xDD, Ş=0xDE, Ğ=0xD0
    expect(decodeText(new Uint8Array([0xf0, 0xfd, 0xfe, 0xdd, 0xde, 0xd0]))).toBe("ğışİŞĞ");
  });

  it("metin → belge → metin", () => {
    const doc = textToDoc("Birinci satır\r\nİkinci satır\n\nDördüncü");
    expect(doc.content!.length).toBe(4);
    expect(docToText(doc, "\n")).toBe("Birinci satır\nİkinci satır\n\nDördüncü\n");
  });

  it("listeler ve tablolar düz metne", () => {
    const doc = {
      type: "doc",
      content: [
        { type: "bulletList", content: [{ type: "listItem", content: [{ type: "paragraph", content: [{ type: "text", text: "Madde" }] }] }] },
        { type: "orderedList", attrs: { start: 1 }, content: [{ type: "listItem", content: [{ type: "paragraph", content: [{ type: "text", text: "Bir" }] }] }, { type: "listItem", content: [{ type: "paragraph", content: [{ type: "text", text: "İki" }] }] }] },
        { type: "table", content: [{ type: "tableRow", content: [{ type: "tableCell", content: [{ type: "paragraph", content: [{ type: "text", text: "A" }] }] }, { type: "tableCell", content: [{ type: "paragraph", content: [{ type: "text", text: "B" }] }] }] }] },
      ],
    };
    expect(docToText(doc, "\n")).toBe("• Madde\n1. Bir\n2. İki\nA\tB\n");
  });
});

describe("HTML", () => {
  it("dışa aktarılan HTML yeniden içe aktarılabilir", () => {
    const html = htmlDocument('<h1>Başlık</h1><p style="text-align: center">Ortalı <strong>kalın</strong></p><ul><li><p>Madde</p></li></ul>', defaultSettings(), "Deneme");
    expect(html).toContain("<!DOCTYPE html>");
    expect(html).toContain("@page");
    const body = extractHtmlBody(html);
    expect(body).toContain("<h1>Başlık</h1>");
    const doc = htmlToDoc(html);
    expect(doc.content![0].type).toBe("heading");
    expect(doc.content![1].attrs?.textAlign).toBe("center");
    expect(doc.content![2].type).toBe("bulletList");
  });

  it("dış HTML: satır içi biçimler ve tablolar", () => {
    const doc = htmlToDoc('<p><span style="color: rgb(192, 0, 0); font-family: Georgia">renkli</span> <u>altı</u></p><table><tr><td>x</td><td>y</td></tr></table>');
    const s = JSON.stringify(doc);
    expect(s).toContain("Georgia");
    expect(s).toContain('"underline"');
    expect(s).toContain('"table"');
  });
});
