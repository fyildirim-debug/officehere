import { expect, test, type Page } from "@playwright/test";
import { UZUN_PENCERE, fareyiCek, karsilastir, sayfaHatalari, uygulamaAc } from "./ortak";

// Belge görsel regresyonu: Resmî Yazışma (2020) şablon kümesi. Her şablon Dosya › Yeni › "Resmî yazışma"
// galerisinden, tarih sabitlenerek (bugünün tarihi görüntüyü değiştirmesin) oluşturulur; tüm sayfaları karşılaştırılır.
// Ek olarak okul başlığı + logo yeri + Arial 11 türevi. Temel görüntüler: resmi-<şablon>-s<n>.png.

const SABLONLAR: [ad: string, galeri: string, sayfa: number][] = [
  ["ust-yazi", "Üst yazı", 1],
  ["dagitimli", "Dağıtımlı yazı", 1],
  ["olur", "Olur yazısı", 1],
  ["tekit", "Tekit yazısı", 1],
  ["ek-dagitim", "Ek ve dağıtım listesi", 3],
  ["dilekce", "Dilekçe", 1],
  ["tutanak", "Tutanak ve rapor", 1],
  ["zarf", "Zarf", 1],
];

test.use({ viewport: UZUN_PENCERE });

async function sablonAc(page: Page, galeri: string, doldur?: (pencere: ReturnType<Page["getByRole"]>) => Promise<void>) {
  await page.locator(".oh-ribbon-file").click();
  await page.locator(".oh-bs-item", { hasText: "Yeni" }).first().click();
  await page.locator(".oh-template", { hasText: "Resmî yazışma" }).click();
  const pencere = page.getByRole("dialog", { name: "Resmî Yazışma Şablonları" });
  await pencere.getByRole("button", { name: galeri, exact: true }).click();
  await pencere.getByPlaceholder("GG.AA.YYYY").fill("15.10.2026");
  await doldur?.(pencere);
  await pencere.getByRole("button", { name: "Oluştur", exact: true }).click();
  await expect(pencere).toBeHidden();
  await expect(page.locator(".oh-backstage")).toBeHidden();
}

for (const [ad, galeri, sayfa] of SABLONLAR) {
  test(`Belge: resmî yazı şablonu — ${galeri}`, async ({ page }) => {
    const hatalar = sayfaHatalari(page);
    await uygulamaAc(page, "belge");
    await sablonAc(page, galeri);
    await expect(page.locator(".oh-titlebar")).toContainText(galeri.split(" ")[0]);
    await expect.poll(() => page.locator(".bl-page").count()).toBe(sayfa);
    await fareyiCek(page);
    for (let i = 0; i < sayfa; i++) await karsilastir(page, page.locator(`.bl-page[data-page="${i}"]`), `resmi-${ad}-s${i + 1}`);
    expect(hatalar).toEqual([]);
  });
}

test("Belge: resmî yazı şablonu — okul başlığı, logo yeri, Arial 11", async ({ page }) => {
  const hatalar = sayfaHatalari(page);
  await uygulamaAc(page, "belge");
  await sablonAc(page, "Üst yazı", async (pencere) => {
    await pencere.getByLabel("Başlık türü").selectOption({ label: "Okul müdürlüğü" });
    await pencere.getByLabel("Yazı tipi").selectOption({ label: "Arial 11" });
    await pencere.getByLabel("Sol üstte logo yeri bırak").check();
  });
  await expect.poll(() => page.locator(".bl-page").count()).toBe(1);
  await fareyiCek(page);
  await karsilastir(page, page.locator('.bl-page[data-page="0"]'), "resmi-ust-yazi-okul-s1");
  expect(hatalar).toEqual([]);
});
