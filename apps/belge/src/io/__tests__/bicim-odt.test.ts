// @vitest-environment jsdom
// OpenDocument Metin: LibreOffice örneğini okuma, yazma, paket geçerliliği ve gidiş-dönüş.
import { describe, expect, it } from "vitest";
import JSZip from "jszip";
import { readDocx, readOdt, validateContent, writeOdt } from "../index";
import { all, docOf, fixture, markAttr, marksOf, norm, p, richDoc, text } from "./bicim-yardimci";

describe("ODT okuma (LibreOffice örneği)", () => {
  it("metin, başlıklar, karakter biçimleri ve Türkçe karakterler", async () => {
    const d = await readOdt(fixture("lo-ornek.odt"));
    validateContent(d.content);
    const c = d.content;
    expect(all(c, "heading").map((h) => [h.attrs?.level, text(h)])).toEqual([
      [1, "Giriş ve Amaç"],
      [2, "Tablo ve Resim"],
    ]);
    expect(all(c, "paragraph").some((x) => x.attrs?.styleId === "Title" && text(x) === "Biçim Örneği")).toBe(true);
    expect(text(c)).toContain("Türkçe: İĞÜŞÖÇ ığüşöç.");
    expect(marksOf(c, "kalın")).toContain("bold");
    expect(marksOf(c, "italik")).toContain("italic");
    expect(marksOf(c, "altı çizili")).toContain("underline");
    expect(marksOf(c, "üstü çizili")).toContain("strike");
    expect(marksOf(c, "2")).toContain("superscript");
    expect(markAttr(c, "kırmızı", "textStyle")?.color).toBe("#C00000");
    expect(markAttr(c, "büyük yazı tipi", "textStyle")).toMatchObject({ fontFamily: "Liberation Sans", fontSize: "14pt" });
    expect(markAttr(c, "vurgulu", "highlight")?.color).toBe("#FFFF00");
    expect(markAttr(c, "ornek.com.tr", "link")?.href).toBe("https://ornek.com.tr/");
    // Sekme, satır sonu ve art arda boşluklar (text:s)
    expect(text(c)).toContain("sekme\tsonrası");
    expect(all(c, "hardBreak").length).toBeGreaterThanOrEqual(1);
    expect(text(c)).toContain("üç  boşluk");
  });

  it("listeler, tablo (birleştirilmiş hücreler), resim, dipnot, sayfa sonu ve hizalama", async () => {
    const d = await readOdt(fixture("lo-ornek.odt"));
    const c = d.content;
    const bullets = all(c, "bulletList");
    expect(all(bullets[0], "listItem").map((li) => text(li.content![0]))).toEqual(["Birinci madde", "İkinci madde", "İç madde"]);
    expect(all(bullets[0], "bulletList").length).toBe(2);
    expect(all(all(c, "orderedList")[0], "listItem").map(text)).toEqual(["Elma", "Armut", "Çilek"]);
    const table = all(c, "table")[0];
    expect(table.content!.length).toBe(3);
    expect(table.content![0].content![0].attrs?.background).toBe("#DAE9F7");
    expect(table.content![1].content![2].attrs?.rowspan).toBe(2);
    expect(table.content![2].content![0].attrs?.colspan).toBe(2);
    expect(text(table.content![2])).toBe("Yatay birleşik hücre");
    const img = all(c, "image")[0];
    expect(img.attrs?.src).toMatch(/^data:image\/png;base64,/);
    expect(img.attrs?.width).toBe(61);
    const note = all(c, "note")[0];
    expect(note.attrs).toMatchObject({ kind: "footnote", text: "Dipnot açıklaması burada." });
    expect(all(c, "pageBreak").length).toBe(1);
    const just = all(c, "paragraph").find((x) => text(x).startsWith("İki yana"));
    expect(just?.attrs).toMatchObject({ textAlign: "justify", indentLeft: 28.35, indentFirst: 14.15 });
    expect(all(c, "paragraph").find((x) => text(x) === "Sağa hizalı paragraf")?.attrs?.textAlign).toBe("right");
  });

  it("sayfa yapısı, üst/alt bilgi ve belge özellikleri", async () => {
    const d = await readOdt(fixture("lo-ornek.odt"));
    const s = d.settings;
    expect(s.page.width).toBe(11906);
    expect(s.page.margin.left).toBe(1417);
    expect(s.page.margin.right).toBe(1134);
    expect(text(s.header!)).toBe("Üst bilgi — Sayfa ");
    expect(JSON.stringify(s.header)).toContain('"field":"PAGE"');
    expect(text(s.footer!)).toBe("Alt bilgi: Gizlilik İçerir");
    expect(s.properties.title).toBe("Biçim Örneği");
  });
});

describe("ODT yazma", () => {
  it("geçerli paket: mimetype ilk ve sıkıştırmasız, manifest ve parçalar", async () => {
    const bytes = await writeOdt(richDoc());
    const head = new TextDecoder().decode(bytes.subarray(30, 38));
    expect(head).toBe("mimetype");
    // Yerel dosya başlığında sıkıştırma yöntemi 0 (STORE) ve içerik düz yazılır.
    expect(bytes[8] | (bytes[9] << 8)).toBe(0);
    expect(new TextDecoder().decode(bytes.subarray(38, 38 + 39))).toBe("application/vnd.oasis.opendocument.text");
    const zip = await JSZip.loadAsync(bytes);
    const manifest = await zip.file("META-INF/manifest.xml")!.async("string");
    for (const f of ["content.xml", "styles.xml", "meta.xml"]) expect(manifest).toContain(`manifest:full-path="${f}"`);
    expect(manifest).toMatch(/Pictures\/resim1\.png/);
    expect(zip.file("Pictures/resim1.png")).toBeTruthy();
    const content = await zip.file("content.xml")!.async("string");
    expect(() => new DOMParser().parseFromString(content, "application/xml")).not.toThrow();
    expect(new DOMParser().parseFromString(content, "application/xml").getElementsByTagName("parsererror").length).toBe(0);
    expect(content).toContain('text:outline-level="1"');
    expect(content).toContain("table:number-columns-spanned=\"2\"");
    expect(content).toContain("table:number-rows-spanned=\"2\"");
    expect(content).toContain('text:note-class="footnote"');
    const meta = await zip.file("meta.xml")!.async("string");
    expect(meta).toContain("<dc:title>Gidiş Dönüş</dc:title>");
  });

  it("gidiş-dönüş: Belge → ODT → Belge aynı içerik ve biçim", async () => {
    const src = richDoc();
    const back = await readOdt(await writeOdt(src));
    expect(norm(back.content)).toEqual(norm(src.content));
    expect(back.settings.properties.title).toBe("Gidiş Dönüş");
  });

  it("gidiş-dönüş: DOCX → Belge → ODT → Belge aynı metin ve biçim", async () => {
    const fromDocx = await readDocx(fixture("lo-ornek.docx"));
    const back = await readOdt(await writeOdt(fromDocx));
    validateContent(back.content);
    expect(text(back.content)).toBe(text(fromDocx.content));
    for (const [needle, mark] of [["kalın", "bold"], ["italik", "italic"], ["üstü çizili", "strike"], ["ornek.com.tr", "link"]]) {
      expect(marksOf(back.content, needle)).toContain(mark);
      expect(marksOf(fromDocx.content, needle)).toContain(mark);
    }
    expect(all(back.content, "heading").map(text)).toEqual(all(fromDocx.content, "heading").map(text));
    expect(all(back.content, "table").length).toBe(all(fromDocx.content, "table").length);
    expect(all(back.content, "listItem").length).toBe(all(fromDocx.content, "listItem").length);
  });

  it("üst/alt bilgi, kenar boşlukları ve bölüm sayfa yapısı korunur", async () => {
    const d = docOf(p("Gövde"));
    d.settings.header = { type: "doc", content: [p([{ type: "text", text: "Başlık " }, { type: "pageField", attrs: { field: "PAGE" } }])] };
    d.settings.footer = { type: "doc", content: [p("Alt")] };
    d.settings.page.margin.top = 1700;
    const back = await readOdt(await writeOdt(d));
    expect(text(back.settings.header!)).toBe("Başlık ");
    expect(JSON.stringify(back.settings.header)).toContain("pageField");
    expect(text(back.settings.footer!)).toBe("Alt");
    expect(back.settings.page.margin.top).toBe(1700);
    expect(back.settings.page.margin.header).toBe(d.settings.page.margin.header);
  });

  it("bozuk dosya anlaşılır hatayla reddedilir", async () => {
    await expect(readOdt(new TextEncoder().encode("PK\u0003\u0004bozuk"))).rejects.toThrow(/ODT/);
    const zip = new JSZip();
    zip.file("mimetype", "application/vnd.oasis.opendocument.text");
    await expect(readOdt(await zip.generateAsync({ type: "uint8array" }))).rejects.toThrow(/content\.xml/);
  });

  it("parolayla şifrelenmiş .odt anlaşılır hatayla reddedilir", async () => {
    const zip = new JSZip();
    zip.file("mimetype", "application/vnd.oasis.opendocument.text");
    zip.file("content.xml", new Uint8Array([0x8f, 0x12, 0x00, 0xfe]));
    zip.file("META-INF/manifest.xml", '<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0"><manifest:file-entry manifest:full-path="content.xml"><manifest:encryption-data/></manifest:file-entry></manifest:manifest>');
    await expect(readOdt(await zip.generateAsync({ type: "uint8array" }))).rejects.toThrow(/parolayla şifrelenmiş/);
  });
});
