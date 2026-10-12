import { expect, test, type Page } from "@playwright/test";
import { URL, collectErrors } from "./ortak";

// Başlık çubuğundaki kayıt durumu rozeti (tarayıcı sürümü: dosyaya doğrudan
// yazılamadığından otomatik kaydetme yoktur ve seçeneği gizlidir).

const badge = (page: Page) => page.locator(".oh-titlebar .oh-savebadge");

async function openApp(page: Page, url: string) {
  // Tarayıcı dosya seçicileri yerine indirme yolu kullanılsın.
  await page.addInitScript(() => {
    const w = window as unknown as Record<string, unknown>;
    delete w.showSaveFilePicker;
    delete w.showOpenFilePicker;
  });
  await page.goto(url);
  await expect(page.locator(".oh-ribbon")).toBeVisible();
}

test.describe("Kayıt durumu rozeti", () => {
  let errors: string[] = [];

  test.beforeEach(({ page }) => {
    errors = collectErrors(page);
  });

  test.afterEach(() => {
    expect(errors).toEqual([]);
  });

  test("Belge: düzenleyince 'Kaydedilmedi', Ctrl+S sonrası 'Kaydedildi SS:DD'; otomatik kaydet seçeneği yok", async ({ page }) => {
    await openApp(page, URL.belge);
    await expect(page.locator(".bl-body .ProseMirror")).toBeFocused();
    await expect(badge(page)).toHaveCount(0);
    await page.keyboard.type("Rozet denemesi");
    await expect(badge(page)).toHaveText("Kaydedilmedi");
    await expect(badge(page)).toHaveClass(/is-dirty/);
    // Tarayıcıda kendiliğinden kaydedilmez: boşta kalma süresinden sonra da kirli kalır.
    await page.waitForTimeout(2600);
    await expect(badge(page)).toHaveText("Kaydedilmedi");

    const [download] = await Promise.all([page.waitForEvent("download"), page.keyboard.press("Control+s")]);
    expect(download.suggestedFilename()).toMatch(/\.docx$/);
    await expect(badge(page)).toHaveText(/^Kaydedildi \d{2}:\d{2}$/);
    await expect(badge(page)).toHaveClass(/is-saved/);
    await expect(page.locator(".oh-titlebar")).not.toContainText("Kaydedilmedi");

    // Yeniden düzenleyince rozet yine "Kaydedilmedi" olur.
    await page.locator(".bl-body .ProseMirror").click();
    await page.keyboard.type(" ek");
    await expect(badge(page)).toHaveText("Kaydedilmedi");

    await page.locator(".oh-ribbon-file").click();
    await page.locator(".oh-bs-item", { hasText: "Seçenekler" }).first().click();
    await expect(page.getByText("Arayüz teması")).toBeVisible();
    await expect(page.getByRole("checkbox", { name: "Otomatik kaydet" })).toHaveCount(0);
  });

  test("Tablo: hücre değişince 'Kaydedilmedi'", async ({ page }) => {
    await openApp(page, URL.tablo);
    await expect(badge(page)).toHaveCount(0);
    const box = page.getByRole("textbox", { name: "Ad Kutusu" });
    await box.fill("B2");
    await box.press("Enter");
    await page.keyboard.type("42");
    await page.keyboard.press("Enter");
    await expect(badge(page)).toHaveText("Kaydedilmedi");
  });

  test("Sunum: başlık yazılınca 'Kaydedilmedi'", async ({ page }) => {
    await openApp(page, URL.sunum);
    await expect(badge(page)).toHaveCount(0);
    await page.locator(".stage .sv [data-shape-id]").first().click();
    await page.keyboard.type("Rozet");
    await page.keyboard.press("Escape");
    await page.keyboard.press("Escape");
    await expect(badge(page)).toHaveText("Kaydedilmedi");
  });
});
