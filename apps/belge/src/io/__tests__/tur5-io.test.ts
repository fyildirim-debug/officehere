// @vitest-environment jsdom
// 5. tur (QA): kaydet-aç turunda bulunan kayıpların düzeltmeleri.
import { describe, expect, it } from "vitest";
import JSZip from "jszip";
import type { JSONContent } from "@tiptap/core";
import { readDocx, readOdt, readRtf, writeDocx, writeOdt, writeRtf } from "../index";
import { fieldNowText } from "../../editor/nodes";
import { all, docOf, p, t } from "./bicim-yardimci";

const buf = async (b: Blob | ArrayBuffer | Uint8Array): Promise<ArrayBuffer> => (b instanceof Blob ? b.arrayBuffer() : b instanceof Uint8Array ? (b.buffer.slice(b.byteOffset, b.byteOffset + b.byteLength) as ArrayBuffer) : b);
const blocks = (c: JSONContent) => c.content!.flatMap((s) => s.content ?? []);

describe("DOCX: belge sonundaki açıklama", () => {
  it("son paragrafta biten açıklama kaydet-aç turunda boş paragraf eklemez", async () => {
    const d = docOf({ type: "section", content: [p("bir"), p([t("açıklamalı", [{ type: "comment", attrs: { id: "c1" } }])])] });
    d.settings.comments = { c1: { id: "c1", author: "Ali", initials: "A", date: "2026-01-01T00:00:00Z", text: "Not" } };
    let cur = d;
    for (let i = 0; i < 3; i++) cur = await readDocx(await buf(await writeDocx(cur)));
    expect(blocks(cur.content)).toHaveLength(2);
    expect(Object.values(cur.settings.comments).map((c) => c.text)).toEqual(["Not"]);
    const xml = await (await JSZip.loadAsync(await buf(await writeDocx(d)))).file("word/document.xml")!.async("string");
    expect(xml).toMatch(/açıklamalı<\/w:t><\/w:r><w:commentRangeEnd w:id="0"\/><w:r>.*<w:commentReference w:id="0"\/><\/w:r><\/w:p><w:sectPr/);
  });
});

describe("RTF: alanlar ve başlık satırı", () => {
  it("tarih ve saat alanı sonucu güncel değerdir ve alan olarak geri okunur", () => {
    const d = docOf({ type: "section", content: [p([t("Tarih: "), { type: "pageField", attrs: { field: "DATE" } }, t(" Saat: "), { type: "pageField", attrs: { field: "TIME" } }])] });
    const rtf = writeRtf(d);
    expect(rtf).toContain(`{\\fldrslt{\\f0\\fs22 ${fieldNowText("DATE")}}}`);
    const back = readRtf(new TextEncoder().encode(rtf));
    const fields = all(back.content, "pageField").map((f) => f.attrs?.field);
    expect(fields).toEqual(["DATE", "TIME"]);
    expect(all(back.content, "text").map((x) => x.text).join("")).toBe("Tarih:  Saat: ");
  });

  it("başlık satırlarını yinele (\\trhdr) korunur", () => {
    const table: JSONContent = {
      type: "table",
      attrs: { repeatHeader: true },
      content: [
        { type: "tableRow", content: [{ type: "tableCell", content: [p("Başlık")] }] },
        { type: "tableRow", content: [{ type: "tableCell", content: [p("Değer")] }] },
      ],
    };
    const back = readRtf(new TextEncoder().encode(writeRtf(docOf({ type: "section", content: [table] }))));
    expect(all(back.content, "table")[0].attrs?.repeatHeader).toBe(true);
  });
});

describe("ODT: sayfa rengi, karakter aralığı, anahtar sözcükler", () => {
  it("sayfa rengi sayfa düzenine yazılır ve geri okunur", async () => {
    const d = docOf({ type: "section", content: [p("renkli sayfa")] });
    d.settings.pageColor = "#FFF2CC";
    const out = await buf(await writeOdt(d));
    const styles = await (await JSZip.loadAsync(out)).file("styles.xml")!.async("string");
    expect(styles).toContain('fo:background-color="#FFF2CC"');
    expect((await readOdt(out)).settings.pageColor).toBe("#FFF2CC");
  });

  it("karakter aralığı yuvarlama artığı taşımaz; anahtar sözcük ayırıcısı korunur", async () => {
    const d = docOf({ type: "section", content: [p([t("aralıklı", [{ type: "textStyle", attrs: { letterSpacing: 2 } }])])] });
    d.settings.properties = { ...d.settings.properties, keywords: "rapor; test" };
    const back = await readOdt(await buf(await writeOdt(d)));
    const ts = all(back.content, "text")[0].marks?.find((m) => m.type === "textStyle");
    expect(ts?.attrs?.letterSpacing).toBe(2);
    expect(back.settings.properties.keywords).toBe("rapor; test");
  });
});

describe("Tarih ve Saat alanı biçimi", () => {
  const fieldDoc = (format: string | null, field = "DATE") => docOf({ type: "section", content: [p([t("Bugün: "), { type: "pageField", attrs: { field, format } }])] });
  const fieldsOf = (c: JSONContent) => all(c, "pageField").map((f) => [f.attrs?.field, f.attrs?.format ?? null]);

  it("DOCX: \\@ biçim resmi yazılır ve geri okunur; sonuç biçimli değerdir", async () => {
    const d = fieldDoc("d MMMM yyyy dddd");
    const out = await buf(await writeDocx(d));
    const xml = await (await JSZip.loadAsync(out)).file("word/document.xml")!.async("string");
    expect(xml).toContain('w:instr=" DATE \\@ &quot;d MMMM yyyy dddd&quot; \\* MERGEFORMAT "');
    expect(xml).toContain(`>${fieldNowText("DATE", new Date(), "d MMMM yyyy dddd")}</w:t>`);
    expect(fieldsOf((await readDocx(out)).content)).toEqual([["DATE", "d MMMM yyyy dddd"]]);
    // Varsayılan biçim düğümde saklanmaz (gidiş-dönüş kararlı).
    expect(fieldsOf((await readDocx(await buf(await writeDocx(fieldDoc(null))))).content)).toEqual([["DATE", null]]);
  });

  it("RTF ve ODT gidiş-dönüşünde biçim korunur", async () => {
    for (const [field, format] of [
      ["DATE", "dd.MM.yy"],
      ["DATE", "d MMMM yyyy"],
      ["TIME", "HH:mm:ss"],
    ] as const) {
      const d = fieldDoc(format, field);
      expect(fieldsOf(readRtf(new TextEncoder().encode(writeRtf(d))).content)).toEqual([[field, format]]);
      expect(fieldsOf((await readOdt(await buf(await writeOdt(d)))).content)).toEqual([[field, format]]);
    }
  });

  it("ODT: sabit tarih (text:fixed) yazıldığı günün değerini metin olarak korur", async () => {
    const d = fieldDoc(null);
    const zip = await JSZip.loadAsync(await buf(await writeOdt(d)));
    const content = (await zip.file("content.xml")!.async("string")).replace(/<text:date [^>]*>[^<]*<\/text:date>/, '<text:date text:fixed="true" text:date-value="2020-02-03">03.02.2020</text:date>');
    zip.file("content.xml", content);
    const back = await readOdt(await zip.generateAsync({ type: "arraybuffer" }));
    expect(all(back.content, "pageField")).toHaveLength(0);
    expect(all(back.content, "text").map((x) => x.text).join("")).toBe("Bugün: 03.02.2020");
  });
});

describe("Büyük İlk Harf", () => {
  const dcDoc = (dropCap: Record<string, unknown>, text = "Şimdi uzun bir paragraf metni.") =>
    docOf({ type: "section", content: [p("Önceki"), { type: "paragraph", attrs: { dropCap, spaceAfter: 6 }, content: [t(text, [{ type: "bold" }])] }, p("Sonraki")] });

  it("DOCX: harf çerçeveli ayrı paragrafa yazılır, açılınca paragrafa geri katılır", async () => {
    const d = dcDoc({ mode: "drop", lines: 3, distance: 2, font: "Georgia" });
    const out = await buf(await writeDocx(d));
    const xml = await (await JSZip.loadAsync(out)).file("word/document.xml")!.async("string");
    expect(xml).toMatch(/<w:framePr w:dropCap="drop" w:lines="3" w:hSpace="40" w:wrap="around" w:vAnchor="text" w:hAnchor="text"\/>/);
    // Harf: kalın, Georgia; geri kalan metin sonraki paragrafta.
    expect(xml).toMatch(/<w:framePr[^>]*\/>.*?<w:rFonts w:ascii="Georgia"[^>]*\/><w:b\/>.*?<w:t xml:space="preserve">Ş<\/w:t>.*?<\/w:p><w:p>.*?<w:t xml:space="preserve">imdi uzun bir paragraf metni.<\/w:t>/);
    const back = await readDocx(out);
    const paras = blocks(back.content);
    expect(paras.map((x) => all(x, "text").map((y) => y.text).join(""))).toEqual(["Önceki", "Şimdi uzun bir paragraf metni.", "Sonraki"]);
    expect(paras[1].attrs?.dropCap).toEqual({ mode: "drop", lines: 3, distance: 2, font: "Georgia" });
    expect(paras[1].attrs?.spaceAfter).toBe(6);
    expect(all(paras[1], "text")[0].marks?.map((m) => m.type)).toEqual(["bold"]);
  });

  it("DOCX: kenar boşluğu kipi ve bölüm sonundaki paragraf korunur", async () => {
    const d = docOf(
      { type: "section", content: [{ type: "paragraph", attrs: { dropCap: { mode: "margin", lines: 2, distance: 0, font: null } }, content: [t("Kenar")] }] },
      { type: "section", content: [p("İkinci bölüm")] },
    );
    const back = await readDocx(await buf(await writeDocx(d)));
    expect(back.content.content).toHaveLength(2);
    const first = back.content.content![0].content![0];
    expect(first.attrs?.dropCap).toEqual({ mode: "margin", lines: 2, distance: 0, font: null });
    expect(all(first, "text").map((x) => x.text).join("")).toBe("Kenar");
  });

  it("ODT: style:drop-cap yazılır ve geri okunur (yazı tipi harf stilinden)", async () => {
    const d = dcDoc({ mode: "drop", lines: 4, distance: 5.67, font: "Georgia" });
    const out = await buf(await writeOdt(d));
    const content = await (await JSZip.loadAsync(out)).file("content.xml")!.async("string");
    expect(content).toMatch(/<style:drop-cap style:lines="4" style:length="1" style:distance="0\.200cm" style:style-name="[^"]+"\/>/);
    const back = await readOdt(out);
    const para = blocks(back.content)[1];
    expect(para.attrs?.dropCap).toEqual({ mode: "drop", lines: 4, distance: 5.67, font: "Georgia" });
    expect(all(para, "text").map((x) => x.text).join("")).toBe("Şimdi uzun bir paragraf metni.");
  });
});

describe("ODT açıklamaları", () => {
  const cm = (id: string) => [{ type: "comment", attrs: { id } }];
  const comments = {
    a: { id: "a", author: "Ayşe Yılmaz", initials: "AY", date: "2026-03-04T10:20:30Z", text: "Birinci açıklama\nikinci satır" },
    b: { id: "b", author: "Ali", initials: "A", date: "2026-03-05T08:00:00Z", text: "Çözüldü", resolved: true },
  };
  const commentedDoc = () => {
    const d = docOf({
      type: "section",
      content: [
        p([t("Başta "), t("açıklamalı metin", cm("a")), t(" sonra.")]),
        p([t("İki paragrafa ", cm("b"))]),
        p([t("yayılan", cm("b")), t(" son")]),
      ],
    });
    d.settings.comments = structuredClone(comments);
    return d;
  };

  it("office:annotation aralıklarıyla yazılır ve geri okunur", async () => {
    const out = await buf(await writeOdt(commentedDoc()));
    const content = await (await JSZip.loadAsync(out)).file("content.xml")!.async("string");
    expect(content).toContain('<office:annotation office:name="__Annotation__1"><dc:creator>Ayşe Yılmaz</dc:creator><dc:date>2026-03-04T10:20:30</dc:date><meta:creator-initials>AY</meta:creator-initials><text:p>Birinci açıklama</text:p><text:p>ikinci satır</text:p></office:annotation>');
    expect(content).toContain('<office:annotation-end office:name="__Annotation__1"/>');
    expect(content).toContain('loext:resolved="true"');
    const back = await readOdt(out);
    const texts = all(back.content, "text");
    const ids = (s: string) => texts.find((x) => x.text === s)?.marks?.filter((m) => m.type === "comment").map((m) => m.attrs?.id) ?? [];
    const idA = ids("açıklamalı metin")[0];
    const idB = ids("İki paragrafa ")[0];
    expect(idA).toBeTruthy();
    expect(ids("Başta ")).toEqual([]);
    expect(ids(" sonra.")).toEqual([]);
    expect(ids("yayılan")).toEqual([idB]);
    expect(ids(" son")).toEqual([]);
    expect(back.settings.comments[idA as string]).toMatchObject({ author: "Ayşe Yılmaz", initials: "AY", date: "2026-03-04T10:20:30Z", text: "Birinci açıklama\nikinci satır" });
    expect(back.settings.comments[idB as string]).toMatchObject({ author: "Ali", text: "Çözüldü", resolved: true });
  });

  it("belge sonunda biten açıklama son paragrafta kapanır", async () => {
    const d = docOf({ type: "section", content: [p([t("Son "), t("kelime", cm("a"))])] });
    d.settings.comments = structuredClone(comments);
    const out = await buf(await writeOdt(d));
    const content = await (await JSZip.loadAsync(out)).file("content.xml")!.async("string");
    expect(content).toMatch(/kelime(<\/text:span>)?<office:annotation-end office:name="__Annotation__1"\/><\/text:p>/);
    const back = await readOdt(out);
    expect(blocks(back.content)).toHaveLength(1);
    expect(Object.values(back.settings.comments)).toHaveLength(1);
  });

  it("LibreOffice'in aralıksız (nokta) açıklaması sonraki sözcüğe, paragraf sonundaysa son metne bağlanır", async () => {
    const zip = await JSZip.loadAsync(await buf(await writeOdt(docOf({ type: "section", content: [p("x")] }))));
    const body =
      '<text:p text:style-name="Standard">Önce <office:annotation><dc:creator>Ali Veli</dc:creator><dc:date>2026-01-02T03:04:05.123456789</dc:date><text:p>Nokta</text:p></office:annotation>sonraki sözcük</text:p>' +
      '<text:p text:style-name="Standard">Bitiş<office:annotation><dc:creator>Can</dc:creator><text:p>Sonda</text:p></office:annotation></text:p>';
    const content = (await zip.file("content.xml")!.async("string")).replace(/<office:text>[\s\S]*<\/office:text>/, `<office:text>${body}</office:text>`);
    zip.file("content.xml", content);
    const back = await readOdt(await zip.generateAsync({ type: "arraybuffer" }));
    const texts = all(back.content, "text");
    const commented = texts.filter((x) => x.marks?.some((m) => m.type === "comment")).map((x) => x.text);
    expect(commented).toEqual(["sonraki", "Bitiş"]);
    const list = Object.values(back.settings.comments);
    expect(list.map((c) => [c.author, c.initials, c.text, c.date])).toEqual([
      ["Ali Veli", "AV", "Nokta", "2026-01-02T03:04:05Z"],
      ["Can", "C", "Sonda", ""],
    ]);
  });
});
