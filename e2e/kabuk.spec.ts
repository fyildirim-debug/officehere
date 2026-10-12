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

    test("Hakkında üçüncü taraf lisans bildirimini açar", async ({ page }) => {
      await page.goto(c.url);
      await page.locator(".oh-ribbon-file").click();
      await page.locator(".oh-bs-item", { hasText: "Hakkında" }).first().click();
      const link = page.getByRole("link", { name: "Üçüncü taraf yazılım bildirimleri" });
      await expect(link).toBeVisible();
      // Geliştirici satırı
      await expect(page.locator(".oh-credit")).toContainText("Furkan YILDIRIM");
      await expect(page.getByRole("link", { name: "furkanyildirim.com" })).toHaveAttribute("href", "https://furkanyildirim.com");
      const res = await page.request.get(new globalThis.URL((await link.getAttribute("href")) ?? "", page.url()).toString());
      expect(res.ok()).toBe(true);
      const body = await res.text();
      expect(body).toContain("Third-Party Software Notices");
      for (const pkg of ["react", "@fluentui/react-icons"]) expect(body).toMatch(new RegExp(`^${pkg.replace(/[/.]/g, "\\$&")} \\S+ — MIT$`, "m"));
    });

    test("koyu tema uygulanır", async ({ page }) => {
      await page.goto(c.url);
      await page.evaluate(() => document.documentElement.setAttribute("data-oh-theme", "dark"));
      const bg = await page.locator(".oh-ribbon-body").evaluate((el) => getComputedStyle(el).backgroundColor);
      expect(bg).toBe("rgb(51, 51, 51)");
    });
  });
}
