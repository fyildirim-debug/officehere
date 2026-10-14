import { expect, test, type Page } from "@playwright/test";
import { URL, collectErrors } from "./ortak";

// "Hata bildir" (web): sürüm/ortam ve son günlük satırlarını GÖSTEREN, düzenlenebilir pencere.
// Hiçbir şey kendiliğinden gönderilmez; "GitHub'da bildir" önceden doldurulmuş issue adresini açar.

async function open(page: Page, url: string) {
  await page.addInitScript(() => {
    const w = window as unknown as { __acilan: string[]; open: (u?: string) => null };
    w.__acilan = [];
    w.open = (u?: string) => {
      w.__acilan.push(String(u));
      return null;
    };
  });
  await page.goto(url);
  await expect(page.locator(".oh-ribbon")).toBeVisible();
}

for (const app of ["belge", "tablo", "sunum"] as const) {
  test(`${app}: Yardım › Hata Bildir penceresi ortamı ve günlüğü gösterir, GitHub adresini doldurur`, async ({ page }) => {
    const errors = collectErrors(page);
    await open(page, URL[app]);
    // Sayfada yakalanmamış bir hata: günlüğe (bellekte) yazılır.
    await Promise.all([
      page.waitForEvent("pageerror"),
      page.evaluate(() => {
        setTimeout(() => {
          throw new Error("web deneme hatası");
        });
      }),
    ]);
    errors.length = 0;
    await page.getByRole("tab", { name: "Yardım" }).click();
    await page.getByRole("button", { name: "Hata Bildir" }).click();
    const dlg = page.getByRole("dialog", { name: "Hata bildir" });
    await expect(dlg).toBeVisible();
    await expect(dlg).toContainText("yalnız siz göndermeyi seçerseniz paylaşılır");
    const body = dlg.getByRole("textbox", { name: "Rapor" });
    await expect(body).toHaveValue(/\(web\)/);
    await expect(body).toHaveValue(/Tarayıcı: /);
    await expect(body).toHaveValue(/web deneme hatası/);
    await dlg.getByRole("textbox", { name: "Başlık" }).fill("Deneme başlığı");
    await dlg.getByRole("button", { name: "GitHub'da bildir" }).click();
    const opened = await page.evaluate(() => (window as unknown as { __acilan: string[] }).__acilan);
    expect(opened).toHaveLength(1);
    expect(opened[0].startsWith("https://github.com/fyildirim-debug/officehere/issues/new?title=Deneme%20ba")).toBe(true);
    expect(opened[0].length).toBeLessThanOrEqual(7000);
    await dlg.locator(".oh-dialog-footer").getByRole("button", { name: "Kapat" }).click();
    await expect(dlg).toHaveCount(0);
    expect(errors).toEqual([]);
  });
}

test("Dosya › Hakkında › Hata bildir", async ({ page }) => {
  await open(page, URL.belge);
  await page.locator(".oh-ribbon-file").click();
  await page.locator(".oh-bs-item", { hasText: "Hakkında" }).first().click();
  await page.getByRole("button", { name: /Hata bildir/ }).click();
  await expect(page.getByRole("dialog", { name: "Hata bildir" })).toBeVisible();
});
