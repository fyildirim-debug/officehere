import { expect, test, type Page } from "@playwright/test";
import { URL, collectErrors } from "./ortak";

// Dosya görünümündeki son dosyalar listesi (tarayıcı: yerel depodaki kayıtlar).
const now = Date.now();
const files = [
  { path: "İSTANBUL ofis raporu.docx", name: "İSTANBUL ofis raporu.docx", app: "belge", openedAt: now - 2 * 60_000 },
  { path: "Işık ölçümü.docx", name: "Işık ölçümü.docx", app: "belge", openedAt: now - 3 * 3_600_000 },
  { path: "Toplantı notları.docx", name: "Toplantı notları.docx", app: "belge", openedAt: now - 2 * 86_400_000 },
  { path: "Invoice 2026.docx", name: "Invoice 2026.docx", app: "belge", openedAt: now - 30 * 86_400_000 },
  { path: "Bütçe.xlsx", name: "Bütçe.xlsx", app: "tablo", openedAt: now - 60_000 },
];

async function openRecentPage(page: Page) {
  // Kayıtlar yalnız ilk yüklemede yazılır; sonraki değişiklikler yeniden yüklemede korunur.
  await page.addInitScript((list) => {
    if (sessionStorage.getItem("oh-e2e-son")) return;
    sessionStorage.setItem("oh-e2e-son", "1");
    localStorage.setItem("officehere.recent", JSON.stringify(list));
  }, files);
  await page.goto(URL.belge);
  await page.locator(".oh-ribbon-file").click();
  await page.locator(".oh-bs-item", { hasText: "Aç" }).first().click();
  await expect(page.locator(".oh-bs-h1")).toHaveText("Aç");
}

const names = (page: Page) => page.locator(".oh-recent-row .oh-recent-filename").allTextContents();

test("son dosyalar: arama ad ve yolda Türkçe büyük/küçük harf duyarsız süzer", async ({ page }) => {
  const errors = collectErrors(page);
  await openRecentPage(page);
  // Yalnız bu uygulamanın dosyaları, en son açılan önce; göreli zaman gösterilir.
  await expect.poll(() => names(page)).toEqual(["İSTANBUL ofis raporu.docx", "Işık ölçümü.docx", "Toplantı notları.docx", "Invoice 2026.docx"]);
  const times = page.locator(".oh-recent-time");
  await expect(times).toHaveText(["2 dakika önce", /saat önce|dün/, "2 gün önce", /^\d{2}\.\d{2}\.\d{4}$/]);

  const search = page.getByRole("searchbox", { name: "Son dosyalarda ara" });
  await search.fill("istanbul");
  await expect.poll(() => names(page)).toEqual(["İSTANBUL ofis raporu.docx"]);
  await search.fill("IŞIK");
  await expect.poll(() => names(page)).toEqual(["Işık ölçümü.docx"]);
  await search.fill("invoice");
  await expect.poll(() => names(page)).toEqual(["Invoice 2026.docx"]);
  await search.fill("TOPLANTI NOT");
  await expect.poll(() => names(page)).toEqual(["Toplantı notları.docx"]);
  await search.fill("böyle bir dosya yok");
  await expect(page.locator(".oh-recent .oh-bs-empty")).toContainText("ile eşleşen dosya yok");

  // Escape önce aramayı temizler, Dosya görünümü açık kalır.
  await search.press("Escape");
  await expect(search).toHaveValue("");
  await expect(page.locator(".oh-backstage")).toBeVisible();
  await expect.poll(() => names(page)).toHaveLength(4);

  // Klavye: aşağı ok ilk satıra gider, oklarla satırlar arasında gezilir.
  await search.press("ArrowDown");
  await expect(page.locator(".oh-recent-open").first()).toBeFocused();
  await page.keyboard.press("ArrowDown");
  await expect(page.locator(".oh-recent-open").nth(1)).toBeFocused();
  expect(errors).toEqual([]);
});

test("son dosyalar: tarayıcıda sabitleme ve listeden kaldırma; klasörde göster yok", async ({ page }) => {
  await openRecentPage(page);
  await page.locator(".oh-recent-row", { hasText: "Toplantı notları.docx" }).click({ button: "right" });
  const menu = page.getByRole("menu");
  await expect(menu.getByRole("menuitem")).toHaveText(["Aç", "Sabitle", "Listeden kaldır"]);
  await menu.getByRole("menuitem", { name: "Sabitle" }).click();
  await expect.poll(() => names(page)).toEqual(["Toplantı notları.docx", "İSTANBUL ofis raporu.docx", "Işık ölçümü.docx", "Invoice 2026.docx"]);
  await expect(page.locator(".oh-recent-group")).toHaveText(["Sabitlenmiş", "Son"]);

  await page.locator(".oh-recent-row", { hasText: "Işık ölçümü.docx" }).click({ button: "right" });
  await page.getByRole("menuitem", { name: "Listeden kaldır" }).click();
  await expect.poll(() => names(page)).toEqual(["Toplantı notları.docx", "İSTANBUL ofis raporu.docx", "Invoice 2026.docx"]);

  // Değişiklikler yerel depoya yazılır: Giriş sayfasında ve yeniden yüklemede korunur.
  await page.locator(".oh-bs-item", { hasText: "Giriş" }).first().click();
  await expect.poll(() => names(page)).toEqual(["Toplantı notları.docx", "İSTANBUL ofis raporu.docx", "Invoice 2026.docx"]);
  const stored = await page.evaluate(() => JSON.parse(localStorage.getItem("officehere.recent") ?? "[]") as { name: string; pinned?: boolean }[]);
  expect(stored.map((f) => [f.name, f.pinned === true])).toEqual([
    ["Toplantı notları.docx", true],
    ["Bütçe.xlsx", false],
    ["İSTANBUL ofis raporu.docx", false],
    ["Invoice 2026.docx", false],
  ]);
  await page.reload();
  await page.locator(".oh-ribbon-file").click();
  await expect.poll(() => names(page)).toEqual(["Toplantı notları.docx", "İSTANBUL ofis raporu.docx", "Invoice 2026.docx"]);

  // Raptiye düğmesiyle sabitleme kaldırılır.
  await page.getByRole("button", { name: "Toplantı notları.docx: sabitlemeyi kaldır" }).click();
  await expect.poll(() => names(page)).toEqual(["İSTANBUL ofis raporu.docx", "Toplantı notları.docx", "Invoice 2026.docx"]);
  await expect(page.locator(".oh-recent-group")).toHaveCount(0);
});
