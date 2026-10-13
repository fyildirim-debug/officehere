// @vitest-environment jsdom
// Zengin Metin (RTF): LibreOffice örneği, kod sayfaları, \uN atlama, hedefler, tablolar, gidiş-dönüş.
import { describe, expect, it } from "vitest";
import { readRtf, validateContent, writeRtf } from "../index";
import { all, docOf, fixture, markAttr, marksOf, norm, p, richDoc, text } from "./bicim-yardimci";

const rtf = (body: string) => readRtf(new TextEncoder().encode(body));

describe("RTF okuma", () => {
  it("LibreOffice örneği: başlıklar, biçimler, liste, tablo, resim, dipnot, üst/alt bilgi", () => {
    const d = readRtf(fixture("lo-ornek.rtf"));
    validateContent(d.content);
    const c = d.content;
    expect(all(c, "heading").map((h) => [h.attrs?.level, text(h)])).toEqual([
      [1, "Giriş ve Amaç"],
      [2, "Tablo ve Resim"],
    ]);
    expect(text(c)).toContain("Türkçe: İĞÜŞÖÇ ığüşöç.");
    expect(marksOf(c, "kalın")).toContain("bold");
    expect(marksOf(c, "italik")).toContain("italic");
    expect(marksOf(c, "altı çizili")).toContain("underline");
    expect(marksOf(c, "üstü çizili")).toContain("strike");
    expect(markAttr(c, "kırmızı", "textStyle")?.color).toBe("#C00000");
    expect(markAttr(c, "ornek.com.tr", "link")?.href).toBe("https://ornek.com.tr/");
    expect(all(all(c, "orderedList")[0], "listItem").map(text)).toEqual(["Elma", "Armut", "Çilek"]);
    expect(all(c, "bulletList").length).toBe(2);
    const table = all(c, "table")[0];
    expect(table.content![1].content![2].attrs?.rowspan).toBe(2);
    expect(table.content![2].content![0].attrs?.colspan).toBe(2);
    expect(table.content![0].content![0].attrs?.background).toBe("#DAE9F7");
    expect(all(c, "image")).toHaveLength(1);
    expect(all(c, "image")[0].attrs).toMatchObject({ width: 60, height: 45 });
    expect(all(c, "note")[0].attrs).toMatchObject({ kind: "footnote", text: "Dipnot açıklaması burada." });
    expect(text(d.settings.header!)).toBe("Üst bilgi — Sayfa ");
    expect(text(d.settings.footer!)).toBe("Alt bilgi: Gizlilik İçerir");
  });

  it("ANSI kod sayfaları: varsayılan 1252, Türkçe 1254 ve yazı tipine göre \\fcharset", () => {
    const a = rtf("{\\rtf1\\ansi\\deff0{\\fonttbl{\\f0 Arial;}}caf\\'e9 na\\'efve\\par}");
    expect(text(a.content)).toBe("café naïve");
    const tr = rtf("{\\rtf1\\ansi\\ansicpg1254\\deff0{\\fonttbl{\\f0 Arial;}}\\'dd\\'fd\\'d0\\'f0\\'de\\'fe\\'dc\\'fc\\'d6\\'f6\\'c7\\'e7\\par}");
    expect(text(tr.content)).toBe("İıĞğŞşÜüÖöÇç");
    const perFont = rtf("{\\rtf1\\ansi\\ansicpg1252\\deff0{\\fonttbl{\\f0 Arial;}{\\f1\\fcharset162 Arial TUR;}}{\\f1 \\'fd\\'f0}{\\f0 \\'e9}\\par}");
    expect(text(perFont.content)).toBe("ığé");
  });

  it("\\uN ve \\uc atlama sayısı, \\* hedefleri ve özel karakterler", () => {
    const d = rtf("{\\rtf1\\ansi{\\*\\generator Deneme;}{\\info{\\title Ba\\u351?l\\u305?k}}\\uc1\\u305?\\u287?{\\uc2\\u350\\'53\\'53}\\u-10179?\\u-8704? a\\tab b\\~c\\emdash d\\par}");
    expect(text(d.content)).toBe("ığŞ😀 a\tb c—d");
    expect(text(d.content)).not.toContain("Deneme");
    expect(d.settings.properties.title).toBe("Başlık");
  });

  it("paragraf biçimi: hizalama, girintiler, aralık, satır aralığı, sayfa sonu", () => {
    const d = rtf("{\\rtf1\\ansi\\pard\\qc\\li720\\ri360\\fi-360\\sb120\\sa240\\sl360\\slmult1 Orta\\par\\pard\\qj\\sl-280\\slmult0 Yaslı\\par\\page\\pard Son\\par}");
    const ps = all(d.content, "paragraph");
    expect(ps[0].attrs).toMatchObject({ textAlign: "center", indentLeft: 36, indentRight: 18, indentFirst: -18, spaceBefore: 6, spaceAfter: 12, lineHeight: 1.5 });
    expect(ps[1].attrs).toMatchObject({ textAlign: "justify", lineHeight: 14, lineRule: "exact" });
    expect(all(d.content, "pageBreak")).toHaveLength(1);
  });

  it("tablo: \\trowd \\cellx \\cell \\row, yatay ve dikey birleştirme", () => {
    const d = rtf(
      "{\\rtf1\\ansi\\trowd\\clmgf\\cellx2000\\clmrg\\cellx4000\\clvmgf\\cellx6000\\pard\\intbl A\\cell\\cell B\\cell\\row" +
        "\\trowd\\cellx2000\\cellx4000\\clvmrg\\cellx6000\\pard\\intbl C\\cell D\\cell\\cell\\row\\pard Sonra\\par}",
    );
    validateContent(d.content);
    const table = all(d.content, "table")[0];
    expect(table.content!.map((r) => r.content!.map((c) => [text(c), c.attrs?.colspan, c.attrs?.rowspan]))).toEqual([
      [
        ["A", 2, 1],
        ["B", 1, 2],
      ],
      [
        ["C", 1, 1],
        ["D", 1, 1],
      ],
    ]);
    expect(text(all(d.content, "paragraph").at(-1)!)).toBe("Sonra");
  });

  it("bozuk ya da RTF olmayan veri reddedilir", () => {
    expect(() => rtf("Merhaba dünya")).toThrow(/RTF/);
  });
});

describe("RTF yazma", () => {
  it("Türkçe karakterler \\uN ile, başlıklar stil tablosuyla yazılır", () => {
    const out = writeRtf(docOf({ type: "heading", attrs: { level: 1 }, content: [{ type: "text", text: "Başlık" }] }, p("ığüşöç İĞ")));
    expect(out.startsWith("{\\rtf1\\ansi\\ansicpg1254")).toBe(true);
    expect(out).toContain("\\u305\\'fd");
    expect(out).toContain("heading 1;");
    expect(out).toMatch(/[\x20-\x7e\n]+/);
    expect([...out].every((ch) => ch.charCodeAt(0) < 128)).toBe(true);
  });

  it("gidiş-dönüş: Belge → RTF → Belge aynı içerik ve biçim", () => {
    const src = richDoc();
    const back = readRtf(new TextEncoder().encode(writeRtf(src)));
    validateContent(back.content);
    expect(norm(back.content)).toEqual(norm(src.content));
    expect(back.settings.properties.title).toBe("Gidiş Dönüş");
  });

  it("üst/alt bilgi ve sayfa yapısı yazılır", () => {
    const d = docOf(p("Gövde"));
    d.settings.header = { type: "doc", content: [p([{ type: "text", text: "Üst " }, { type: "pageField", attrs: { field: "PAGE" } }])] };
    d.settings.page = { ...d.settings.page, orientation: "landscape", width: 16838, height: 11906 };
    const back = readRtf(new TextEncoder().encode(writeRtf(d)));
    expect(text(back.settings.header!)).toBe("Üst ");
    expect(JSON.stringify(back.settings.header)).toContain("pageField");
    expect(back.settings.page).toMatchObject({ width: 16838, height: 11906, orientation: "landscape" });
  });
});
