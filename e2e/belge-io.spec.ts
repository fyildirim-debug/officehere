import { readFileSync } from "node:fs";
import { resolve } from "node:path";
import { expect, test, type Page } from "@playwright/test";
import { URL, collectErrors, zipEntry } from "./ortak";

// OfficeHere Belge: dosya G/Ç sağlamlığı — kötü niyetli HTML yapıştırma, açılışta kayıp uyarıları ve
// bozuk dosyalarda anlaşılır Türkçe hata iletileri.

const FIX = resolve(__dirname, "../apps/belge/src/io/__fixtures__/bicimler");
const fixture = (name: string) => readFileSync(resolve(FIX, name));

const EVIL =
  '<p onclick="window.__xss=1">Güvenli metin <b>kalın</b><img src="olmayan-resim.png" onerror="window.__xss=2"></p>' +
  '<script>window.__xss=3</script><iframe srcdoc="<script>parent.__xss=4</script>"></iframe>' +
  '<a href="javascript:window.__xss=5">bağlantı</a><svg><image href="x" onerror="window.__xss=6"/></svg>';

test.describe("OfficeHere Belge dosya G/Ç sağlamlığı", () => {
  let errors: string[] = [];

  test.beforeEach(async ({ page }) => {
    errors = collectErrors(page);
    await page.addInitScript(() => {
      const w = window as unknown as Record<string, unknown>;
      delete w.showSaveFilePicker;
      delete w.showOpenFilePicker;
    });
    await page.goto(URL.belge);
    await expect(body(page)).toBeVisible();
  });

  test.afterEach(() => {
    // Var olmayan resmin yüklenememesi beklenen bir ağ hatasıdır; başka hata olmamalı.
    expect(errors.filter((e) => !/olmayan-resim|Failed to load resource/.test(e))).toEqual([]);
  });

  const body = (page: Page) => page.locator(".bl-body .ProseMirror");
  const bar = (page: Page) => page.locator(".oh-messagebar");
  const open = (page: Page, name: string, buffer: Buffer) => page.setInputFiles("#bl-open-input", { name, mimeType: "application/octet-stream", buffer });
  const xss = (page: Page) => page.evaluate(() => (window as unknown as { __xss?: number }).__xss ?? null);

  test("Ctrl+V: kötü niyetli HTML yapıştırılınca betik çalışmaz, içerik korunur", async ({ page }) => {
    await body(page).click();
    await page.evaluate((html) => {
      const dt = new DataTransfer();
      dt.setData("text/html", html);
      dt.setData("text/plain", "Güvenli metin kalın bağlantı");
      document.querySelector(".bl-body .ProseMirror")!.dispatchEvent(new ClipboardEvent("paste", { clipboardData: dt, bubbles: true, cancelable: true }));
    }, EVIL);
    await expect(body(page)).toContainText("Güvenli metin");
    await expect(body(page).locator("strong", { hasText: "kalın" })).toBeVisible();
    await page.waitForTimeout(500);
    expect(await xss(page)).toBeNull();
    expect(await body(page).innerHTML()).not.toMatch(/onerror|onclick|javascript:|<script|<iframe/i);
  });

  for (const mode of ["Kaynak Biçimlendirmesini Koru (K)", "Biçimlendirmeyi Birleştir (B)"]) {
    test(`şerit Yapıştır › ${mode}: panodaki kötü niyetli HTML betik çalıştırmaz`, async ({ page }) => {
      // Chromium pano okumasında HTML'i kendisi temizler; temizlemeyen ortamları (başka tarayıcılar, ham pano)
      // sınamak için pano okuma işlevi ham HTML döndürecek biçimde değiştirilir.
      await page.evaluate((html) => {
        const item = { types: ["text/html", "text/plain"], getType: async (t: string) => new Blob([t === "text/html" ? html : "Güvenli metin"], { type: t }) };
        Object.defineProperty(navigator, "clipboard", { configurable: true, value: { read: async () => [item], readText: async () => "Güvenli metin" } });
      }, EVIL);
      await body(page).click();
      // Yapıştır bölünmüş düğmesinin menüsü
      await page.locator(".oh-ribbon").getByRole("button", { name: /Yapıştır/ }).last().click();
      await page.getByRole("menuitem", { name: mode }).click();
      await expect(body(page)).toContainText("Güvenli metin");
      await page.waitForTimeout(500);
      expect(await xss(page)).toBeNull();
      expect(await body(page).innerHTML()).not.toMatch(/onerror|onclick|javascript:/i);
    });
  }

  test("açıklama tarihi: LibreOffice'in geçersiz tarihi ekranda görünmez, kaydedilen .docx'e yazılmaz", async ({ page }, testInfo) => {
    // LibreOffice tarihsiz açıklamaya w:date="0-00-00T00:00:00Z" yazar (şemaya aykırı).
    const src = readFileSync(resolve(__dirname, "../test-derlem/lo/lo-karma.docx"));
    expect(zipEntry(src, "word/comments.xml")).toContain('w:date="0-00-00T00:00:00Z"');
    await open(page, "lo-karma.docx", src);
    await expect(page.locator(".oh-titlebar")).toContainText("lo-karma");
    const balloons = page.locator(".bl-balloon:not(.is-revision)");
    await expect(balloons).toHaveCount(2);
    await expect(balloons.locator(".bl-balloon-date")).toHaveText(["", ""]);
    await expect(page.locator(".bl-balloon-date", { hasText: "Invalid" })).toHaveCount(0);
    const download = page.waitForEvent("download");
    await page.keyboard.press("Control+s");
    const out = testInfo.outputPath("lo-karma.docx");
    await (await download).saveAs(out);
    const comments = zipEntry(readFileSync(out), "word/comments.xml") ?? "";
    expect(comments.match(/<w:comment /g)).toHaveLength(2);
    expect(comments).not.toContain("w:date=");
  });

  test("okunamayan içerik açılışta ileti çubuğunda uyarılır", async ({ page }) => {
    // Kapanmamış gruplu (bozuk) RTF: boş açılmak yerine uyarı.
    const rtf = new TextDecoder("latin1").decode(fixture("lo-ornek.rtf")).replace(/\}/g, "");
    await open(page, "dengesiz.rtf", Buffer.from(rtf, "latin1"));
    await expect(bar(page)).toContainText(/kaydederseniz/i);
    await expect(bar(page)).toContainText("bozuk");
    // İçerik kurtarılır (boş açılmaz).
    await expect(body(page)).toContainText("Giriş ve Amaç");
  });

  test("değişiklik izleme ve açıklamalar okunur; uyarı çıkmaz", async ({ page }) => {
    await open(page, "izleme.odt", fixture("izleme.odt"));
    await expect(body(page)).toContainText("eklenen ifade");
    await expect(body(page)).toContainText("silinen sözcük");
    await expect(bar(page)).toHaveCount(0);
  });

  test(".fodt OpenDocument olarak açılır (düz XML metni olarak değil)", async ({ page }) => {
    await open(page, "kaynak.fodt", fixture("kaynak.fodt"));
    await expect(body(page).locator("h1")).toHaveText("Giriş ve Amaç");
    await expect(body(page)).not.toContainText("office:document");
    await expect(bar(page)).toContainText(".fodt");
  });

  test("bozuk dosya anlaşılır Türkçe iletiyle reddedilir", async ({ page }) => {
    await open(page, "kesik.docx", fixture("lo-ornek.docx").subarray(0, 3000));
    await expect(bar(page)).toContainText("Dosya bozuk");
    await expect(bar(page)).not.toContainText(/Corrupted|central directory|docx file|Error/);
  });
});
