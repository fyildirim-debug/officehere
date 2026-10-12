// @vitest-environment jsdom
// 4. tur DOCX gidiş-dönüşleri: kaynakça (customXml b:Sources), birleştirme veri kaynağı, bölüm başına
// üst/alt bilgi (Öncekine Bağla), DATE/TIME alanları, sürüklenmiş kaydırmalı resim konumu.
import { describe, expect, it } from "vitest";
import JSZip from "jszip";
import type { JSONContent } from "@tiptap/core";
import { readDocx, writeDocx } from "../index";
import { defaultSettings, type BelgeDocument } from "../../model/document";
import type { Source } from "../../model/sources";
import { parseXml } from "../docx/xml";
import { sourcesFromXml, sourcesToXml } from "../docx/customXml";

const PNG = "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAQAAAADCAIAAAA7ljmRAAAAEElEQVR4nGM4IScHRww4OQD1xwwxB+utmQAAAABJRU5ErkJggg==";
const t = (text: string): JSONContent => ({ type: "text", text });
const p = (content: JSONContent[] | string): JSONContent => ({ type: "paragraph", content: typeof content === "string" ? [t(content)] : content });
const H = (text: string): JSONContent => ({ type: "doc", content: [p(text)] });
const textOf = (j: JSONContent | null | undefined): string => (j ? (j.text ?? "") + (j.content ?? []).map(textOf).join("") : "");

async function roundtrip(doc: BelgeDocument) {
  const bytes = await writeDocx(doc);
  return { back: await readDocx(bytes), zip: await JSZip.loadAsync(bytes) };
}

const SOURCES: Source[] = [
  { tag: "Yil2024", type: "book", author: "Yılmaz, Ayşe; Demir, Can", title: "Türkçe Yazım", year: "2024", publisher: "Örnek Yayınları", city: "İstanbul" },
  { tag: "Kaya2023", type: "article", author: "Kaya, Ali", title: "Bir Makale", year: "2023", publisher: "Dergi", volume: "5", issue: "2", pages: "10-20" },
  { tag: "Web2025", type: "web", author: "Şahin, Elif", title: "Bir Sayfa", year: "2025", publisher: "Site", url: "https://example.com", accessed: "03.04.2025" },
];

describe("DOCX gidiş-dönüş (4. tur)", () => {
  it("kaynaklar customXml b:Sources parçasına yazılır ve geri okunur", async () => {
    const doc: BelgeDocument = { content: { type: "doc", content: [{ type: "section", content: [p("Metin")] }] }, settings: { ...defaultSettings(), sources: SOURCES, bibStyle: "ieee" } };
    const { back, zip } = await roundtrip(doc);
    const item = Object.keys(zip.files).find((n) => /^customXml\/item\d+\.xml$/.test(n))!;
    const xml = await zip.file(item)!.async("string");
    expect(xml).toContain("<b:Sources");
    expect(xml).toContain("<b:Tag>Yil2024</b:Tag>");
    expect(xml).toContain("<b:SourceType>JournalArticle</b:SourceType>");
    expect(await zip.file("[Content_Types].xml")!.async("string")).toContain("customXmlProperties+xml");
    expect(await zip.file("word/_rels/document.xml.rels")!.async("string")).toContain("../customXml/item1.xml");
    expect(zip.file("customXml/_rels/item1.xml.rels")).toBeTruthy();
    expect(back.settings.sources?.map((s) => s.tag)).toEqual(["Yil2024", "Kaya2023", "Web2025"]);
    // Yalnız b:Sources'u olan (başka düzenleyicide oluşturulmuş) belge: kaynaklar oradan alınır.
    const parsed = sourcesFromXml(parseXml(sourcesToXml(SOURCES, "apa")));
    expect(parsed.style).toBe("apa");
    expect(parsed.sources[0]).toMatchObject({ tag: "Yil2024", type: "book", author: "Yılmaz, Ayşe; Demir, Can", city: "İstanbul" });
    expect(parsed.sources[1]).toMatchObject({ type: "article", publisher: "Dergi", volume: "5", issue: "2", pages: "10-20" });
    expect(parsed.sources[2]).toMatchObject({ type: "web", url: "https://example.com", accessed: "03.04.2025" });
  });

  it("birleştirme veri kaynağı belgeyle saklanır (alıcılar ve hariç tutulanlar)", async () => {
    const mailMerge = { list: { name: "alicilar.csv", fields: ["Ad", "Şehir"], records: [{ Ad: "Ayşe", Şehir: "İzmir" }, { Ad: "Gül, Şule", Şehir: "Iğdır" }] }, excluded: [1] };
    const doc: BelgeDocument = {
      content: { type: "doc", content: [{ type: "section", content: [p([t("Sayın "), { type: "mergeField", attrs: { name: "Ad" } }])] }] },
      settings: { ...defaultSettings(), mailMerge },
    };
    const { back } = await roundtrip(doc);
    expect(back.settings.mailMerge).toEqual(mailMerge);
  });

  it("bölüm başına üst bilgi: bağlı bölüm başvuruyu paylaşır, bağı kaldırılan kendi parçasını alır", async () => {
    const sec = (text: string, sid: string | null = null): JSONContent => ({ type: "section", attrs: { kind: "nextPage", page: null, sid }, content: [p(text)] });
    const doc: BelgeDocument = {
      content: { type: "doc", content: [sec("Bir"), sec("İki"), sec("Üç", "x3")] },
      settings: { ...defaultSettings(), header: H("Genel üst bilgi"), footer: H("Genel alt bilgi"), sectionHf: { x3: { header: H("Üçüncü bölüm") } } },
    };
    const { back, zip } = await roundtrip(doc);
    const xml = await zip.file("word/document.xml")!.async("string");
    const refs = [...xml.matchAll(/<w:sectPr>(.*?)<\/w:sectPr>/g)].map((m) => [...m[1].matchAll(/headerReference w:type="default" r:id="(rId\d+)"/g)].map((x) => x[1])[0]);
    expect(refs[0]).toBe(refs[1]);
    expect(refs[2]).not.toBe(refs[0]);
    expect(textOf(back.settings.header)).toBe("Genel üst bilgi");
    const sids = back.content.content!.map((s) => s.attrs?.sid ?? null);
    expect(sids[1]).toBeNull();
    expect(sids[2]).toBeTruthy();
    expect(textOf(back.settings.sectionHf?.[sids[2] as string]?.header)).toBe("Üçüncü bölüm");
    // Alt bilgi üçüncü bölümde de bağlı kalır.
    expect(back.settings.sectionHf?.[sids[2] as string]).not.toHaveProperty("footer");
  });

  it("DATE/TIME alanları ve sürüklenmiş kare/üst-alt resim konumu", async () => {
    const doc: BelgeDocument = {
      content: {
        type: "doc",
        content: [
          {
            type: "section",
            content: [
              p([t("Tarih: "), { type: "pageField", attrs: { field: "DATE" } }, t(" Saat: "), { type: "pageField", attrs: { field: "TIME" } }]),
              p([{ type: "image", attrs: { src: PNG, width: 80, height: 60, wrap: "square", floatSide: "left", posX: 50, posY: 24 } }, t("Sola, kenardan 50 px")]),
              p([{ type: "image", attrs: { src: PNG, width: 80, height: 60, wrap: "square", floatSide: "right", posX: 30, posY: 0 } }, t("Sağa, kenardan 30 px")]),
              p([{ type: "image", attrs: { src: PNG, width: 80, height: 60, wrap: "topBottom", floatSide: "center" } }, t("Ortada")]),
            ],
          },
        ],
      },
      settings: defaultSettings(),
    };
    const { back, zip } = await roundtrip(doc);
    const xml = await zip.file("word/document.xml")!.async("string");
    expect(xml).toContain('w:instr=" DATE \\@ &quot;dd.MM.yyyy&quot; \\* MERGEFORMAT "');
    expect(xml).toContain("TIME \\@ &quot;HH:mm&quot;");
    const blocks = back.content.content![0].content!;
    const fields = (blocks[0].content ?? []).filter((n) => n.type === "pageField").map((n) => n.attrs?.field);
    expect(fields).toEqual(["DATE", "TIME"]);
    const img = (i: number) => blocks[i].content!.find((n) => n.type === "image")!.attrs!;
    expect(img(1)).toMatchObject({ wrap: "square", floatSide: "left", posX: 50, posY: 24 });
    expect(img(2)).toMatchObject({ wrap: "square", floatSide: "right", posX: 30 });
    expect(img(3)).toMatchObject({ wrap: "topBottom", floatSide: "center" });
  });

  it("satır numaraları (w:lnNumType), satır numarası gizleme, denklem (OMML), grafik, kırpma, dizin, koruma", async () => {
    const page = { ...defaultSettings().page, lineNumbers: { restart: "newSection" as const, start: 5, countBy: 2, distance: 360 } };
    const doc: BelgeDocument = {
      content: {
        type: "doc",
        content: [
          {
            type: "section",
            content: [
              { type: "paragraph", attrs: { suppressLineNumbers: true }, content: [t("Numarasız paragraf")] },
              p([t("Denklem: "), { type: "equation", attrs: { latex: "x=\\frac{1}{2}", display: false } }]),
              p([{ type: "equation", attrs: { latex: "\\sum_{i=1}^n i", display: true } }]),
              p([{ type: "chart", attrs: { data: { type: "bar", title: "Satış", categories: ["A", "B"], series: [{ name: "2025", values: [3, 4.5], color: "#156082" }], width: 400, height: 240 } } }]),
              p([{ type: "image", attrs: { src: PNG, width: 50, height: 30, crop: { l: 0.25, t: 0, r: 0.25, b: 0 } } }]),
              p([t("Ankara"), { type: "indexEntry", attrs: { main: "Ankara", sub: null } }]),
              { type: "indexBlock", attrs: { columns: 2, entries: [{ main: "Ankara", pages: [1] }] } },
            ],
          },
        ],
      },
      settings: { ...defaultSettings(), page, protection: { mode: "readOnly", enforced: true } },
    };
    const { back, zip } = await roundtrip(doc);
    const xml = await zip.file("word/document.xml")!.async("string");
    expect(xml).toContain('<w:lnNumType w:countBy="2" w:start="4" w:distance="360" w:restart="newSection"/>');
    expect(xml).toMatch(/<w:lnNumType[^>]*\/><w:cols/);
    expect(xml).toContain("<w:suppressLineNumbers/>");
    expect(xml).toContain("<m:oMath><m:r>");
    expect(xml).toContain("<m:oMathPara>");
    expect(xml).toContain('<a:srcRect l="25000" t="0" r="25000" b="0"/>');
    expect(xml).toContain(" XE &quot;Ankara&quot; ");
    expect(xml).toContain("INDEX \\c &quot;2&quot;");
    expect(zip.file("word/charts/chart1.xml")).toBeTruthy();
    expect(await zip.file("word/settings.xml")!.async("string")).toContain('<w:documentProtection w:edit="readOnly" w:enforcement="1"/>');
    expect(back.settings.page.lineNumbers).toEqual({ restart: "newSection", start: 5, countBy: 2, distance: 360 });
    expect(back.settings.protection).toEqual({ mode: "readOnly", enforced: true });
    const blocks = back.content.content![0].content!;
    expect(blocks[0].attrs?.suppressLineNumbers).toBe(true);
    const eq = blocks[1].content!.find((n) => n.type === "equation")!;
    expect(eq.attrs).toMatchObject({ latex: "x=\\frac{1}{2}", display: false });
    expect(blocks[2].content![0].attrs).toMatchObject({ display: true });
    expect(blocks[3].content![0].attrs!.data).toMatchObject({ type: "bar", title: "Satış", categories: ["A", "B"], series: [{ name: "2025", values: [3, 4.5] }] });
    expect(blocks[4].content![0].attrs!.crop).toEqual({ l: 0.25, t: 0, r: 0.25, b: 0 });
    expect(blocks[5].content!.some((n) => n.type === "indexEntry" && n.attrs?.main === "Ankara")).toBe(true);
    expect(blocks[6]).toMatchObject({ type: "indexBlock", attrs: { columns: 2, entries: [{ main: "Ankara", pages: [1] }] } });
  });
});
