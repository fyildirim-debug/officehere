import { expect, test } from "@playwright/test";
import { URL, collectErrors } from "./ortak";

// Üç uygulamada ortak kabuk davranışı: başlık çubuğu, şerit, Dosya görünümü, durum çubuğu.
const cases = [
  { app: "belge", url: URL.belge, tabs: ["Giriş", "Ekle", "Tasarım", "Düzen", "Başvurular", "Gözden Geçir", "Görünüm", "Yardım"], name: "OfficeHere Belge" },
  { app: "tablo", url: URL.tablo, tabs: ["Giriş", "Ekle", "Sayfa Düzeni", "Formüller", "Veri", "Gözden Geçir", "Görünüm", "Yardım"], name: "OfficeHere Tablo" },
  { app: "sunum", url: URL.sunum, tabs: ["Giriş", "Ekle", "Tasarım", "Geçişler", "Animasyonlar", "Slayt Gösterisi", "Gözden Geçir", "Görünüm", "Yardım"], name: "OfficeHere Sunum" },
] as const;

for (const c of cases) {
  test.describe(c.name, () => {
    test("kabuk yüklenir, konsol hatası yok", async ({ page }) => {
      const errors = collectErrors(page);
      await page.goto(c.url);
      await expect(page.locator(".oh-titlebar")).toContainText(c.name);
      await expect(page.locator("html")).toHaveAttribute("data-app", c.app);
      for (const t of c.tabs) {
        await expect(page.getByRole("tab", { name: t, exact: true })).toBeVisible();
      }
      await expect(page.locator(".oh-statusbar")).toBeVisible();
      expect(errors).toEqual([]);
    });

    test("tüm şerit sekmeleri hatasız açılır", async ({ page }) => {
      const errors = collectErrors(page);
      await page.goto(c.url);
      for (const t of c.tabs) {
        await page.getByRole("tab", { name: t, exact: true }).click();
        await expect(page.locator(".oh-ribbon-body .oh-ribbon-group").first()).toBeVisible();
      }
      expect(errors).toEqual([]);
    });

    test("Dosya görünümü açılır ve geri dönülür", async ({ page }) => {
      await page.goto(c.url);
      await page.locator(".oh-ribbon-file").click();
      await expect(page.locator(".oh-backstage")).toBeVisible();
      for (const item of ["Yeni", "Aç", "Bilgi", "Yazdır", "Dışarı Aktar", "Seçenekler", "Hakkında"]) {
        await page.locator(".oh-bs-item", { hasText: item }).first().click();
        await expect(page.locator(".oh-bs-h1")).toHaveText(item);
      }
      await page.locator(".oh-bs-back").click();
      await expect(page.locator(".oh-backstage")).toHaveCount(0);
    });

    test("arayüzde Microsoft marka adları geçmez", async ({ page }) => {
      await page.goto(c.url);
      const text = await page.locator("body").innerText();
      expect(text).not.toMatch(/\b(Microsoft|Word|Excel|PowerPoint)\b/);
    });

    test("koyu tema uygulanır", async ({ page }) => {
      await page.goto(c.url);
      await page.evaluate(() => document.documentElement.setAttribute("data-oh-theme", "dark"));
      const bg = await page.locator(".oh-ribbon-body").evaluate((el) => getComputedStyle(el).backgroundColor);
      expect(bg).toBe("rgb(51, 51, 51)");
    });
  });
}
