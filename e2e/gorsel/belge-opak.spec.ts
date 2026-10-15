import { crc32 } from "node:zlib";
import { expect, test, type Page } from "@playwright/test";
import { UZUN_PENCERE, karsilastir, sayfaHatalari, uygulamaAc, type Tema } from "./ortak";

// Belge görsel regresyonu: desteklenmeyen (opak) öğe kartları. Gömülü nesne (önizlemesiz ve PNG önizlemeli),
// içe aktarılan içerik (blok kart), anlaşılmayan çizim grubu, bilinmeyen koşu öğesi ve VML yatay çizgi tek
// sayfada; açık temada ve koyu tema + Koyu Sayfa'da. Belge, derleme dosya eklenmesin diye burada kurulur.

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const MS = "http://schemas.microsoft.com/office/word/2010";
// 4×3 renkli PNG (önizleme yedeği).
const PNG = Buffer.from("iVBORw0KGgoAAAANSUhEUgAAAAQAAAADCAIAAAA7ljmRAAAAEElEQVR4nGM4IScHRww4OQD1xwwxB+utmQAAAABJRU5ErkJggg==", "base64");

const nesne = (rid: string, w: number, h: number, progId: string, img = "") =>
  `<w:object w:dxaOrig="${w * 20}" w:dyaOrig="${h * 20}"><v:shape id="_x0000_i10${rid.slice(-1)}" type="#_x0000_t75" style="width:${w}pt;height:${h}pt" o:ole="">${img}</v:shape><o:OLEObject Type="Embed" ProgID="${progId}" ShapeID="_x0000_i10${rid.slice(-1)}" DrawAspect="Content" ObjectID="_10000000${rid.slice(-1)}" r:id="${rid}"/></w:object>`;

const BELGE =
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:document xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:r="${R}" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:w="${W}" xmlns:wpg="${MS}/wordprocessingGroup" xmlns:wps="${MS}/wordprocessingShape" mc:Ignorable="">` +
  `<w:body>` +
  `<w:p><w:r><w:t xml:space="preserve">Önizlemesiz gömülü nesne: </w:t></w:r><w:r>${nesne("rId2", 240, 30, "Excel.Sheet.12")}</w:r></w:p>` +
  `<w:p><w:r><w:t xml:space="preserve">Önizlemeli nesne: </w:t></w:r><w:r>${nesne("rId4", 120, 60, "Word.Document.12", '<v:imagedata r:id="rId1" o:title=""/>')}</w:r><w:r><w:t xml:space="preserve"> ve bilinmeyen öğe </w:t></w:r><w:r><w:dayLong/></w:r></w:p>` +
  `<w:altChunk r:id="rId3"/>` +
  `<w:p><w:r><mc:AlternateContent><mc:Choice Requires="wpg"><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="1905000" cy="571500"/><wp:docPr id="1" name="Grup 1"/><a:graphic><a:graphicData uri="${MS}/wordprocessingGroup"><wpg:wgp><wpg:cNvGrpSpPr/><wpg:grpSpPr/></wpg:wgp></a:graphicData></a:graphic></wp:inline></w:drawing></mc:Choice><mc:Fallback><w:pict><v:rect style="width:150pt;height:45pt"/></w:pict></mc:Fallback></mc:AlternateContent></w:r></w:p>` +
  `<w:p><w:r><w:pict><v:rect style="width:0;height:2.25pt" o:hr="t" o:hrstd="t" fillcolor="gray" stroked="f"/></w:pict></w:r></w:p>` +
  `<w:p><w:r><w:t>Son paragraf.</w:t></w:r></w:p>` +
  `<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1417" w:right="1417" w:bottom="1417" w:left="1417" w:header="708" w:footer="708" w:gutter="0"/></w:sectPr></w:body></w:document>`;

const rel = (id: string, type: string, target: string) => `<Relationship Id="${id}" Type="${R}/${type}" Target="${target}"/>`;
const PARCALAR: [string, Buffer][] = [
  [
    "[Content_Types].xml",
    Buffer.from(
      `<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/><Default Extension="bin" ContentType="application/vnd.openxmlformats-officedocument.oleObject"/><Default Extension="htm" ContentType="text/html"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>`,
    ),
  ],
  ["_rels/.rels", Buffer.from(`<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${rel("rId1", "officeDocument", "word/document.xml")}</Relationships>`)],
  [
    "word/_rels/document.xml.rels",
    Buffer.from(
      `<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${rel("rId1", "image", "media/image1.png")}${rel("rId2", "oleObject", "embeddings/oleObject1.bin")}${rel("rId3", "aFChunk", "parca.htm")}${rel("rId4", "oleObject", "embeddings/oleObject2.bin")}</Relationships>`,
    ),
  ],
  ["word/document.xml", Buffer.from(BELGE)],
  ["word/media/image1.png", PNG],
  ["word/embeddings/oleObject1.bin", Buffer.from([1, 2, 3, 4])],
  ["word/embeddings/oleObject2.bin", Buffer.from([5, 6, 7, 8])],
  ["word/parca.htm", Buffer.from("<html><body><p>İçe aktarılan</p></body></html>")],
];

/** Sıkıştırmasız ZIP (yalnız bu sınama belgesi için). */
function zip(files: [string, Buffer][]): Buffer {
  const yerel: Buffer[] = [];
  const merkez: Buffer[] = [];
  let konum = 0;
  for (const [ad, veri] of files) {
    const n = Buffer.from(ad);
    const crc = crc32(veri);
    const l = Buffer.alloc(30);
    l.writeUInt32LE(0x04034b50, 0);
    l.writeUInt16LE(20, 4);
    l.writeUInt32LE(crc, 14);
    l.writeUInt32LE(veri.length, 18);
    l.writeUInt32LE(veri.length, 22);
    l.writeUInt16LE(n.length, 26);
    const c = Buffer.alloc(46);
    c.writeUInt32LE(0x02014b50, 0);
    c.writeUInt16LE(20, 4);
    c.writeUInt16LE(20, 6);
    c.writeUInt32LE(crc, 16);
    c.writeUInt32LE(veri.length, 20);
    c.writeUInt32LE(veri.length, 24);
    c.writeUInt16LE(n.length, 28);
    c.writeUInt32LE(konum, 42);
    yerel.push(l, n, veri);
    merkez.push(c, n);
    konum += 30 + n.length + veri.length;
  }
  const cd = Buffer.concat(merkez);
  const son = Buffer.alloc(22);
  son.writeUInt32LE(0x06054b50, 0);
  son.writeUInt16LE(files.length, 8);
  son.writeUInt16LE(files.length, 10);
  son.writeUInt32LE(cd.length, 12);
  son.writeUInt32LE(konum, 16);
  return Buffer.concat([...yerel, cd, son]);
}

async function belgeAc(page: Page) {
  await page.setInputFiles("#bl-open-input", { name: "opak-kartlar.docx", mimeType: "application/octet-stream", buffer: zip(PARCALAR) });
  await expect(page.locator(".oh-titlebar")).toContainText("opak-kartlar", { timeout: 30_000 });
  await expect(page.locator(".bl-body .bl-opaque")).toHaveCount(6);
}

test.use({ viewport: UZUN_PENCERE });

const DURUMLAR: [Tema, string, boolean][] = [
  ["colorful", "acik", false],
  ["dark", "koyu", true],
];

for (const [tema, ad, koyuSayfa] of DURUMLAR) {
  test(`Belge: opak öğe kartları, ${ad} tema${koyuSayfa ? " ve Koyu Sayfa" : ""}`, async ({ page }) => {
    const hatalar = sayfaHatalari(page);
    await uygulamaAc(page, "belge", tema);
    await belgeAc(page);
    if (koyuSayfa) {
      await page.getByRole("tab", { name: "Görünüm", exact: true }).click();
      const dugme = page.locator(".oh-ribbon-body").getByRole("button", { name: "Koyu Sayfa", exact: true });
      await dugme.click();
      await expect(dugme).toHaveAttribute("aria-pressed", "true");
    }
    // Düzenleyici odağı ve imleç görüntüye girmesin: seçim belgenin sonuna alınır, odak bırakılır.
    await page.locator(".bl-body .ProseMirror").evaluate((e) => (e as HTMLElement).blur());
    await karsilastir(page, page.locator('.bl-page[data-page="0"]'), `belge-opak-kartlar-${ad}`);
    expect(hatalar).toEqual([]);
  });
}
