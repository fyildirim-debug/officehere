import { readFileSync } from "node:fs";
import { resolve } from "node:path";
import { expect, test, type Page } from "@playwright/test";
import { URL, collectErrors } from "./ortak";

// OfficeHere Belge: yeni dosya biçimleri (.odt, .rtf, .md, .doc) web'de açma, kaydetme ve dışarı aktarma.

const FIX = resolve(__dirname, "../apps/belge/src/io/__fixtures__/bicimler");
const fixture = (name: string) => readFileSync(resolve(FIX, name));

test.describe("OfficeHere Belge dosya biçimleri", () => {
  let errors: string[] = [];

  test.beforeEach(async ({ page }) => {
    errors = collectErrors(page);
    // Tarayıcı dosya seçicileri yerine indirme / <input type=file> yolu kullanılsın.
    await page.addInitScript(() => {
      const w = window as unknown as Record<string, unknown>;
      delete w.showSaveFilePicker;
      delete w.showOpenFilePicker;
    });
    await page.goto(URL.belge);
    await expect(body(page)).toBeVisible();
  });

  test.afterEach(() => {
    expect(errors).toEqual([]);
  });

  const body = (page: Page) => page.locator(".bl-body .ProseMirror");
  const titlebar = (page: Page) => page.locator(".oh-titlebar");
  const open = (page: Page, name: string, buffer: Buffer) => page.setInputFiles("#bl-open-input", { name, mimeType: "application/octet-stream", buffer });

  test("giriş kutusu yeni biçimleri kabul eder", async ({ page }) => {
    const accept = await page.locator("#bl-open-input").getAttribute("accept");
    for (const ext of [".docx", ".docm", ".dotx", ".dotm", ".doc", ".odt", ".ott", ".rtf", ".md", ".markdown", ".html", ".htm", ".txt"]) expect(accept).toContain(ext);
  });

  for (const [file, ext] of [
    ["lo-ornek.odt", "odt"],
    ["lo-ornek.rtf", "rtf"],
    ["lo-ornek.doc", "doc"],
  ] as const) {
    test(`.${ext} açılır: başlık, biçim, tablo ve resim görünür`, async ({ page }) => {
      await open(page, file, fixture(file));
      await expect(body(page).locator("h1")).toHaveText("Giriş ve Amaç");
      await expect(body(page).locator("h2")).toHaveText("Tablo ve Resim");
      await expect(body(page).locator("strong", { hasText: "kalın" })).toBeVisible();
      await expect(body(page)).toContainText("Türkçe: İĞÜŞÖÇ ığüşöç.");
      await expect(body(page).locator("table td")).toHaveCount(7);
      await expect(body(page).locator("table td[colspan='2']")).toHaveText("Yatay birleşik hücre");
      await expect(body(page).locator("img[src^='data:image']").first()).toBeVisible();
      await expect(titlebar(page)).toContainText(`lo-ornek.${ext}`);
      if (ext === "doc") {
        await expect(titlebar(page)).toContainText("Uyumluluk Modu");
        await expect(page.locator(".oh-messagebar")).toContainText("Uyumluluk modu");
        await expect(page.locator(".oh-messagebar")).not.toContainText("Word");
      }
    });
  }

  test(".md açılır: başlık, liste, tablo ve kod bloğu Markdown'dan çözülür", async ({ page }) => {
    await open(page, "ornek.md", fixture("ornek.md"));
    await expect(body(page).locator("h1")).toHaveText("Ana Başlık İĞÜŞÖÇ");
    await expect(body(page).locator("h2")).toHaveText("Alt Çizgili Başlık");
    await expect(body(page).locator("strong").first()).toHaveText("kalın");
    await expect(body(page).locator("ul > li").first()).toContainText("Birinci madde");
    await expect(body(page).locator("table tr")).toHaveCount(3);
    await expect(body(page)).toContainText('const selam = "Merhaba";');
    await expect(body(page)).not.toContainText("**");
    await expect(titlebar(page)).toContainText("ornek.md");
  });

  test("Dışarı Aktar › OpenDocument Metin: indirilen .odt yeniden açılır", async ({ page, browser }, testInfo) => {
    await page.keyboard.type("Dışa aktarılan ");
    await page.keyboard.press("Control+b");
    await page.keyboard.type("kalın");
    await page.keyboard.press("Control+b");
    await page.keyboard.type(" metin İıĞğ.");
    await page.locator(".oh-ribbon-file").click();
    await page.locator(".oh-bs-item", { hasText: "Dışarı Aktar" }).first().click();
    await expect(page.locator(".oh-bs-action", { hasText: "Markdown (.md)" })).toBeVisible();
    await expect(page.locator(".oh-bs-action", { hasText: "Zengin Metin (.rtf)" })).toBeVisible();
    const download = page.waitForEvent("download");
    await page.locator(".oh-bs-action", { hasText: "OpenDocument Metin (.odt)" }).click();
    const file = await download;
    expect(file.suggestedFilename()).toMatch(/\.odt$/);
    const path = testInfo.outputPath(file.suggestedFilename());
    await file.saveAs(path);
    const bytes = readFileSync(path);
    expect(bytes.subarray(30, 38).toString()).toBe("mimetype");

    const other = await browser.newPage();
    const otherErrors = collectErrors(other);
    await other.goto(URL.belge);
    await other.setInputFiles("#bl-open-input", { name: file.suggestedFilename(), mimeType: "application/vnd.oasis.opendocument.text", buffer: bytes });
    const ob = other.locator(".bl-body .ProseMirror");
    await expect(ob).toContainText("Dışa aktarılan kalın metin İıĞğ.");
    await expect(ob.locator("strong")).toHaveText("kalın");
    expect(otherErrors).toEqual([]);
    await other.close();
  });

  test("Farklı Kaydet: açılan .odt aynı biçimde indirilir ve yeniden açılır", async ({ page, browser }, testInfo) => {
    await open(page, "lo-ornek.odt", fixture("lo-ornek.odt"));
    await expect(body(page).locator("h1")).toHaveText("Giriş ve Amaç");
    const download = page.waitForEvent("download");
    await page.keyboard.press("F12");
    const file = await download;
    expect(file.suggestedFilename()).toBe("lo-ornek.odt");
    const path = testInfo.outputPath("kayit.odt");
    await file.saveAs(path);
    const other = await browser.newPage();
    await other.goto(URL.belge);
    await other.setInputFiles("#bl-open-input", { name: "kayit.odt", mimeType: "application/vnd.oasis.opendocument.text", buffer: readFileSync(path) });
    const ob = other.locator(".bl-body .ProseMirror");
    await expect(ob.locator("h1")).toHaveText("Giriş ve Amaç");
    await expect(ob.locator("table td[colspan='2']")).toHaveText("Yatay birleşik hücre");
    await expect(ob.locator("img[src^='data:image']").first()).toBeVisible();
    await other.close();
  });

  test("Kaydet: .md belgede bir kez biçim sorulur, sonra aynı biçimde kaydedilir", async ({ page }) => {
    await open(page, "not.md", Buffer.from("# Not\n\nİlk satır.\n"));
    await expect(body(page).locator("h1")).toHaveText("Not");
    await body(page).locator("p").first().click();
    await page.keyboard.press("End");
    await page.keyboard.type(" Eklenen.");
    await page.keyboard.press("Control+s");
    const dialog = page.getByRole("dialog", { name: "Biçimi koru" });
    await expect(dialog).toContainText("Markdown");
    const download = page.waitForEvent("download");
    await dialog.getByRole("button", { name: ".md olarak kaydet" }).click();
    const file = await download;
    expect(file.suggestedFilename()).toBe("not.md");
    const saved = readFileSync(await file.path());
    expect(saved.toString("utf8")).toContain("# Not\n\nİlk satır. Eklenen.");
    // İkinci kayıtta soru yinelenmez.
    await page.keyboard.type(" İki");
    const second = page.waitForEvent("download");
    await page.keyboard.press("Control+s");
    expect((await second).suggestedFilename()).toBe("not.md");
    await expect(dialog).toHaveCount(0);
  });

  test("Kaydet: .doc belge özgün biçimin üzerine yazılmaz, .docx olarak kaydedilir", async ({ page }) => {
    await open(page, "eski.doc", fixture("lo-ornek.doc"));
    await expect(body(page).locator("h1")).toHaveText("Giriş ve Amaç");
    const download = page.waitForEvent("download");
    await page.keyboard.press("Control+s");
    const file = await download;
    expect(file.suggestedFilename()).toBe("eski.docx");
    await expect(titlebar(page)).not.toContainText("Uyumluluk Modu");
    await expect(titlebar(page)).toContainText("eski");
  });
});
