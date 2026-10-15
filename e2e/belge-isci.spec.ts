import { readFileSync } from "node:fs";
import { resolve } from "node:path";
import { crc32, deflateRawSync } from "node:zlib";
import { expect, test, type Page } from "@playwright/test";
import { URL, collectErrors, zipEntry } from "./ortak";

// OfficeHere Belge dosya G/Ç Worker'ı (P4.1): büyük belge ana iş parçacığını dondurmadan açılır (Worker'da
// okunur, ilk ekran önce yerleşir, kalanı parça parça eklenir), durum çubuğunda ilerleme ve İptal görünür,
// iptal açık belgeyi korur; kaydetme de Worker'da yapılır.

const FIX = resolve(__dirname, "../apps/belge/src/io/__fixtures__/bicimler");

/** En küçük ZIP yazıcısı (sıkıştırılmış girdiler, merkez dizin). */
function zip(files: Record<string, string>): Buffer {
  const locals: Buffer[] = [];
  const centrals: Buffer[] = [];
  let offset = 0;
  for (const [name, text] of Object.entries(files)) {
    const raw = Buffer.from(text, "utf8");
    const data = deflateRawSync(raw);
    const n = Buffer.from(name, "utf8");
    const crc = crc32(raw);
    const local = Buffer.alloc(30);
    local.writeUInt32LE(0x04034b50, 0);
    local.writeUInt16LE(20, 4);
    local.writeUInt16LE(0x0800, 6);
    local.writeUInt16LE(8, 8);
    local.writeUInt32LE(crc, 14);
    local.writeUInt32LE(data.length, 18);
    local.writeUInt32LE(raw.length, 22);
    local.writeUInt16LE(n.length, 26);
    const central = Buffer.alloc(46);
    central.writeUInt32LE(0x02014b50, 0);
    central.writeUInt16LE(20, 4);
    central.writeUInt16LE(20, 6);
    central.writeUInt16LE(0x0800, 8);
    central.writeUInt16LE(8, 10);
    central.writeUInt32LE(crc, 16);
    central.writeUInt32LE(data.length, 20);
    central.writeUInt32LE(raw.length, 24);
    central.writeUInt16LE(n.length, 28);
    central.writeUInt32LE(offset, 42);
    locals.push(local, n, data);
    centrals.push(central, n);
    offset += local.length + n.length + data.length;
  }
  const dir = Buffer.concat(centrals);
  const end = Buffer.alloc(22);
  end.writeUInt32LE(0x06054b50, 0);
  end.writeUInt16LE(Object.keys(files).length, 8);
  end.writeUInt16LE(Object.keys(files).length, 10);
  end.writeUInt32LE(dir.length, 12);
  end.writeUInt32LE(offset, 16);
  return Buffer.concat([...locals, dir, end]);
}

/** ~300 sayfalık .docx: ilk ve son paragrafta işaret. */
function bigDocx(paragraphs = 3000): Buffer {
  const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
  const p = (t: string) => `<w:p><w:r><w:t xml:space="preserve">${t}</w:t></w:r></w:p>`;
  const body = [p("İLK-İŞARET büyük belge")];
  for (let i = 1; i < paragraphs; i++) {
    if (i % 60 === 0) body.push(`<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>Bölüm ${i / 60}</w:t></w:r></w:p>`);
    body.push(p(`${i}. paragraf: ${"Çalışma raporu ığüşöç İĞÜŞÖÇ ile uzun bir satır metni. ".repeat(5)}`));
  }
  body.push(p("SON-İŞARET"));
  return zip({
    "[Content_Types].xml": `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>`,
    "_rels/.rels": `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`,
    "word/document.xml": `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="${W}"><w:body>${body.join("")}<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1417" w:right="1417" w:bottom="1417" w:left="1417" w:header="708" w:footer="708" w:gutter="0"/></w:sectPr></w:body></w:document>`,
  });
}

test.describe("Belge dosya G/Ç Worker'ı", () => {
  let errors: string[] = [];
  const big = bigDocx();

  test.beforeEach(async ({ page, context }) => {
    errors = collectErrors(page);
    await context.addInitScript(() => {
      const w = window as unknown as Record<string, unknown>;
      delete w.showSaveFilePicker;
      delete w.showOpenFilePicker;
      const lt: number[] = [];
      w.__lt = lt;
      try {
        new PerformanceObserver((l) => {
          for (const e of l.getEntries()) lt.push(Math.round(e.duration));
        }).observe({ type: "longtask", buffered: true });
      } catch {
        /* gözlemci yoksa ölçülmez */
      }
    });
    await page.goto(URL.belge);
    await expect(body(page)).toBeVisible();
  });

  test.afterEach(() => {
    expect(errors).toEqual([]);
  });

  const body = (page: Page) => page.locator(".bl-body .ProseMirror");
  const open = (page: Page, name: string, buffer: Buffer) => page.setInputFiles("#bl-open-input", { name, mimeType: "application/octet-stream", buffer });
  const longTasks = (page: Page) => page.evaluate(() => (window as unknown as { __lt: number[] }).__lt.slice());
  /**
   * İptal'e basılır; sayfa içinde ölçülen süre (tıklama olayından ilerleme göstergesinin kalkmasına) döner.
   * Playwright'ın tıklama hazırlığı meşgul sayfada uzun sürebilir; ölçüme girmez.
   */
  const cancel = async (page: Page) => {
    await page.evaluate(() => {
      const w = window as unknown as { __iptal: { t0: number; t1: number } };
      w.__iptal = { t0: 0, t1: 0 };
      document.addEventListener("click", (e) => {
        if ((e.target as Element).closest?.(".oh-progress-cancel")) w.__iptal.t0 = performance.now();
      }, { capture: true, once: true });
      new MutationObserver((_r, mo) => {
        if (w.__iptal.t0 && !document.querySelector(".oh-progress")) {
          w.__iptal.t1 = performance.now();
          mo.disconnect();
        }
      }).observe(document.body, { childList: true, subtree: true });
    });
    await page.locator(".oh-progress").getByRole("button", { name: "İptal" }).click();
    await expect(page.locator(".oh-progress")).toHaveCount(0, { timeout: 5000 });
    return page.evaluate(() => {
      const t = (window as unknown as { __iptal: { t0: number; t1: number } }).__iptal;
      return Math.round(t.t1 - t.t0);
    });
  };

  test("büyük belge donmadan açılır: ilk ekran önce, sonuna kadar yüklenir; kaydetme Worker'da", async ({ page }) => {
    test.setTimeout(180_000);
    const before = (await longTasks(page)).length;
    await open(page, "buyuk-belge.docx", big);
    await expect(page.locator(".oh-titlebar")).toContainText("buyuk-belge", { timeout: 60_000 });
    await expect(body(page)).toContainText("İLK-İŞARET");
    await expect(body(page)).toContainText("SON-İŞARET", { timeout: 120_000 });
    await expect(page.locator(".oh-progress")).toHaveCount(0);
    const lt = (await longTasks(page)).slice(before);
    console.log(`açılış uzun görevleri: ${JSON.stringify(lt)}`);
    // Worker'sız açılışta tek görev ~2 sn sürüyordu.
    expect(Math.max(0, ...lt)).toBeLessThan(500);

    await body(page).locator("p").first().click();
    await page.keyboard.press("Control+Home");
    await page.keyboard.type("YENİ ");
    const download = page.waitForEvent("download", { timeout: 120_000 });
    await page.keyboard.press("Control+s");
    const xml = zipEntry(readFileSync((await (await download).path())!), "word/document.xml") ?? "";
    expect(xml).toContain("YENİ");
    expect(xml).toContain("SON-İŞARET");
  });

  test("İptal: okunurken ya da sayfalar yerleştirilirken önceki belge korunur (1 sn içinde)", async ({ page }) => {
    test.setTimeout(180_000);
    await open(page, "lo-ornek.docx", readFileSync(resolve(FIX, "lo-ornek.docx")));
    await expect(page.locator(".oh-titlebar")).toContainText("lo-ornek");
    await expect(body(page)).toContainText("Giriş ve Amaç");
    const kept = async () => {
      await expect(page.locator(".oh-titlebar")).toContainText("lo-ornek");
      await expect(body(page)).toContainText("Giriş ve Amaç");
      await expect(body(page)).not.toContainText("İLK-İŞARET");
      await expect(page.locator(".oh-titlebar-dirty")).toHaveCount(0);
    };

    // 1) İlk görünüşte (okuma ya da yerleştirmenin başı)
    await open(page, "buyuk-belge.docx", big);
    await expect(page.getByRole("progressbar", { name: "Açılıyor…" })).toBeVisible({ timeout: 30_000 });
    let ms = await cancel(page);
    console.log(`iptal (ilk): ${ms} ms`);
    expect(ms).toBeLessThan(1000);
    await kept();

    // 2) Sayfalar yerleştirilirken
    await open(page, "buyuk-belge.docx", big);
    await expect(page.getByRole("progressbar", { name: "Açılıyor…" })).toHaveAttribute("aria-valuetext", /yerleştiriliyor/, { timeout: 60_000 });
    ms = await cancel(page);
    console.log(`iptal (yerleştirme): ${ms} ms`);
    expect(ms).toBeLessThan(1000);
    await kept();

    // İptalden sonra açma yine çalışır.
    await open(page, "buyuk-belge.docx", big);
    await expect(body(page)).toContainText("SON-İŞARET", { timeout: 120_000 });
  });
});
