import { readFileSync } from "node:fs";
import { resolve } from "node:path";
import { crc32 } from "node:zlib";
import { expect, test, type Page } from "@playwright/test";
import { URL, collectErrors, interceptFileChooser } from "./ortak";

// Artımlı sayfalama: yazma, Enter, silme ve sayfa sonundan sonra sayfa sınırları (ara parçaların türü,
// yüksekliği ve konumu), aynı belge baştan sayfalandığında bulunanla aynı olmalı. Baştan sayfalama Web
// Düzeni'nden Yazdırma Düzeni'ne dönülerek istenir.

const DOCX_MIME = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
const DERLEM = resolve(__dirname, "..", "test-derlem");

/** Sıkıştırmasız ZIP (yalnız bu sınama belgesi için). */
function zip(files: [string, string][]): Buffer {
  const yerel: Buffer[] = [];
  const merkez: Buffer[] = [];
  let konum = 0;
  for (const [ad, metin] of files) {
    const n = Buffer.from(ad);
    const veri = Buffer.from(metin);
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

/** Desteklenmeyen öğeli çok sayfalı belge: paragraflar arasında satır içi gömülü nesne ve blok (içe aktarılan içerik) kartları. */
function opaqueDocx(): Buffer {
  const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
  const R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
  const rel = (id: string, type: string, target: string) => `<Relationship Id="${id}" Type="${R}/${type}" Target="${target}"/>`;
  const p = (t: string) => `<w:p><w:r><w:t xml:space="preserve">${t}</w:t></w:r></w:p>`;
  const cumle = "Sayfa sınırlarını sınamak için yazılmış uzunca bir paragraf cümlesi; satır sonlarına kadar akar. ";
  let body = "";
  const rels: string[] = [rel("rId1", "styles", "styles.xml")];
  const parts: [string, string][] = [];
  for (let k = 0; k < 14; k++) {
    for (let j = 0; j < 4; j++) body += p(`${k + 1}.${j + 1} ${cumle.repeat(2 + ((k + j) % 4))}`);
    const ole = `rId${10 + k}`;
    const chunk = `rId${100 + k}`;
    body += `<w:p><w:r><w:t xml:space="preserve">Nesne ${k + 1}: </w:t></w:r><w:r><w:object w:dxaOrig="${2400 + k * 300}" w:dyaOrig="${600 + (k % 3) * 900}"><v:shape id="_x0000_i10${k}" type="#_x0000_t75" style="width:${120 + k * 15}pt;height:${30 + (k % 3) * 45}pt" o:ole=""/><o:OLEObject Type="Embed" ProgID="Excel.Sheet.12" ShapeID="_x0000_i10${k}" DrawAspect="Content" ObjectID="_10000000${k}" r:id="${ole}"/></w:object></w:r></w:p>`;
    if (k % 2 === 0) body += `<w:altChunk r:id="${chunk}"/>`;
    rels.push(rel(ole, "oleObject", `embeddings/oleObject${k}.bin`));
    parts.push([`word/embeddings/oleObject${k}.bin`, `nesne ${k}`]);
    if (k % 2 === 0) {
      rels.push(rel(chunk, "aFChunk", `parca${k}.htm`));
      parts.push([`word/parca${k}.htm`, `<html><body><p>İçe aktarılan ${k}</p></body></html>`]);
    }
  }
  const doc = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:r="${R}" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:w="${W}"><w:body>${body}${p("Son paragraf.")}<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1417" w:right="1417" w:bottom="1417" w:left="1417" w:header="708" w:footer="708" w:gutter="0"/></w:sectPr></w:body></w:document>`;
  return zip([
    ["[Content_Types].xml", `<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="bin" ContentType="application/vnd.openxmlformats-officedocument.oleObject"/><Default Extension="htm" ContentType="text/html"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/></Types>`],
    ["_rels/.rels", `<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${rel("rId1", "officeDocument", "word/document.xml")}</Relationships>`],
    ["word/_rels/document.xml.rels", `<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${rels.join("")}</Relationships>`],
    ["word/document.xml", doc],
    ["word/styles.xml", `<?xml version="1.0" encoding="UTF-8"?><w:styles xmlns:w="${W}"><w:docDefaults><w:pPrDefault><w:pPr><w:spacing w:after="160"/></w:pPr></w:pPrDefault></w:docDefaults></w:styles>`],
    ...parts,
  ]);
}

test.describe("OfficeHere Belge — artımlı sayfalama", () => {
  let errors: string[] = [];

  test.beforeEach(async ({ page }) => {
    errors = collectErrors(page);
    interceptFileChooser(page);
    await page.goto(URL.belge);
    await expect(page.locator(".bl-body .ProseMirror")).toBeVisible();
  });

  test.afterEach(() => {
    expect(errors).toEqual([]);
  });

  const ribbon = (page: Page) => page.locator(".oh-ribbon-body");
  const tab = (page: Page, name: string) => page.getByRole("tab", { name, exact: true }).click();
  const button = (page: Page, name: string) => ribbon(page).getByRole("button", { name, exact: true }).first();

  /** Sayfa ara parçaları (tür, yükseklik, ardından gelen ilk metin) ve sayfa sayısı. */
  const layout = (page: Page) =>
    page.evaluate(() => {
      const root = document.querySelector(".bl-body .ProseMirror")!;
      const tw = document.createTreeWalker(root, NodeFilter.SHOW_TEXT);
      const parts = Array.from(root.querySelectorAll<HTMLElement>(".bl-pg")).map((e) => {
        tw.currentNode = e;
        let t = tw.nextNode();
        while (t && t.parentElement?.closest(".bl-pg")) t = tw.nextNode();
        return `${e.className}|${Math.round(parseFloat(e.style.height || "0"))}|${(t?.textContent ?? "").slice(0, 24)}`;
      });
      return { pages: document.querySelectorAll(".bl-pages > .bl-page").length, parts };
    });

  /** Yerleşim iki ardışık okumada aynı kalana dek bekler. */
  const settled = async (page: Page) => {
    let prev = "";
    let cur = "";
    await expect
      .poll(
        async () => {
          prev = cur;
          cur = JSON.stringify(await layout(page));
          return cur === prev;
        },
        { intervals: [300], timeout: 20_000 },
      )
      .toBe(true);
    return JSON.parse(cur) as Awaited<ReturnType<typeof layout>>;
  };

  /** Baştan (tam) sayfalama: Web Düzeni'ne geçip Yazdırma Düzeni'ne dönülür. */
  const repaginate = async (page: Page) => {
    await tab(page, "Görünüm");
    await button(page, "Web Düzeni").click();
    await expect(page.locator(".bl-pages > .bl-page")).toHaveCount(0);
    await button(page, "Yazdırma Düzeni").click();
    await tab(page, "Giriş");
    return settled(page);
  };

  for (const [file, min] of [
    ["lo/lo-uzun.docx", 20],
    ["lo/lo-rapor.docx", 6],
    ["opak-kartlar.docx", 5],
  ] as const) {
    test(`${file}: düzenlemelerden sonra sayfa sınırları baştan sayfalamayla aynı`, async ({ page }) => {
      const buffer = file.includes("/") ? readFileSync(resolve(DERLEM, file)) : opaqueDocx();
      await page.setInputFiles("#bl-open-input", { name: file.split("/").pop()!, mimeType: DOCX_MIME, buffer });
      await expect.poll(() => page.locator(".bl-pages > .bl-page").count(), { timeout: 30_000 }).toBeGreaterThanOrEqual(min);
      // Desteklenmeyen öğe kartları (satır içi ve blok) da ölçülen birimler arasında.
      if (!file.includes("/")) {
        await expect(page.locator(".bl-body .bl-opaque.is-inline")).toHaveCount(14);
        await expect(page.locator(".bl-body .bl-opaque.is-block")).toHaveCount(7);
      }
      const first = await settled(page);
      expect(first.parts.length).toBeGreaterThan(0);
      const paras = page.locator(".bl-body .ProseMirror p");
      const count = await paras.count();
      // Ortadaki bir paragrafın sonuna satır ekleyecek kadar yazı, sonra yeni paragraf ve silme.
      await paras.nth(Math.floor(count / 2)).click();
      await page.keyboard.press("End");
      await page.keyboard.type(" sayfa sınırlarını kaydıracak kadar uzun bir ek metin, bir iki üç dört beş altı yedi sekiz", { delay: 5 });
      await page.keyboard.press("Enter");
      await page.keyboard.type("Araya eklenen yeni paragraf", { delay: 5 });
      // Baştaki bir paragrafta silme (sonraki bütün sayfalar kayar).
      await paras.nth(3).click();
      await page.keyboard.press("End");
      for (let i = 0; i < 30; i++) await page.keyboard.press("Backspace");
      const edited = await settled(page);
      expect(await repaginate(page)).toEqual(edited);
      // Sayfa sonu ve geri alma
      await paras.nth(Math.floor(count / 3)).click();
      await page.keyboard.press("Control+Enter");
      const broken = await settled(page);
      expect(broken.pages).toBeGreaterThan(edited.pages - 1);
      expect(await repaginate(page)).toEqual(broken);
      await page.locator(".bl-body .ProseMirror").focus();
      await page.keyboard.press("Control+z");
      const undone = await settled(page);
      expect(await repaginate(page)).toEqual(undone);
    });
  }
});
