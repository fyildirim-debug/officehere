import { expect, test, type Page } from "@playwright/test";
import { URL, collectErrors } from "./ortak";

// Başlık çubuğundaki ortak komut araması (Alt+Q): şerit komutları, Dosya eylemleri,
// son kullanılanlar ve listenin sonundaki "bul" satırı; Belge, Tablo ve Sunum'da.

const search = (page: Page) => page.getByRole("combobox", { name: "Komut ara" });
const list = (page: Page) => page.getByRole("listbox", { name: "Arama sonuçları" });
const options = (page: Page) => list(page).getByRole("option");
const recentGroup = (page: Page) => list(page).getByRole("group", { name: "Son kullanılanlar" });
/** Komut çalıştıktan sonra uygulamanın ertelenmiş odak dönüşü (bir sonraki kare) tamamlansın. */
const settle = (page: Page) => page.evaluate(() => new Promise<void>((r) => requestAnimationFrame(() => requestAnimationFrame(() => r()))));

let errors: string[] = [];

test.beforeEach(async ({ page }) => {
  errors = collectErrors(page);
  // Tarayıcı dosya seçicileri yerine indirme yolu kullanılsın.
  await page.addInitScript(() => {
    const w = window as unknown as Record<string, unknown>;
    delete w.showSaveFilePicker;
    delete w.showOpenFilePicker;
  });
});

test.afterEach(() => {
  expect(errors).toEqual([]);
});

test.describe("Komut araması — Belge", () => {
  const body = (page: Page) => page.locator(".bl-body .ProseMirror");

  test.beforeEach(async ({ page }) => {
    await page.goto(URL.belge);
    await expect(body(page)).toBeFocused();
  });

  test("“KALIN” aranıp seçilince seçili metin kalınlaşır; ↑/↓, aria-activedescendant ve Esc", async ({ page }) => {
    await page.keyboard.type("Merhaba dünya");
    await page.keyboard.press("Shift+Home");
    await page.keyboard.press("Alt+q");
    await expect(search(page)).toBeFocused();
    // Türkçe büyük harf: "KALIN" → "Kalın".
    await page.keyboard.type("KALIN");
    const first = options(page).first();
    await expect(first).toContainText("Kalın");
    await expect(first).toContainText("Giriş › Yazı Tipi");
    await expect(first.locator("kbd")).toHaveText("Ctrl+B");
    await expect(options(page).last()).toContainText("Belgede bul: KALIN");
    // Belge'de satır önceden seçili gelmez (Enter "Belgede bul" yapar).
    await expect(search(page)).not.toHaveAttribute("aria-activedescendant", /./);
    await page.keyboard.press("ArrowDown");
    await expect(first).toHaveAttribute("aria-selected", "true");
    await expect(search(page)).toHaveAttribute("aria-activedescendant", (await first.getAttribute("id"))!);
    // Yukarı ok listenin sonuna sarar, aşağı ok başa döner.
    await page.keyboard.press("ArrowUp");
    await expect(options(page).last()).toHaveAttribute("aria-selected", "true");
    await page.keyboard.press("ArrowDown");
    await expect(first).toHaveAttribute("aria-selected", "true");
    await page.keyboard.press("Enter");
    await expect(list(page)).toHaveCount(0);
    await expect(body(page).locator("strong")).toHaveText("Merhaba dünya");
    await expect(body(page)).toBeFocused();
    await settle(page);
    // Esc listeyi kapatır, odak belgeye döner.
    await page.keyboard.press("Alt+q");
    await page.keyboard.type("italik");
    await expect(options(page).first()).toContainText("İtalik");
    await page.keyboard.press("Escape");
    await expect(list(page)).toHaveCount(0);
    await expect(search(page)).toHaveValue("");
    await expect(body(page)).toBeFocused();
  });

  test("seçim yokken Enter belgede bulur (gezinti bölmesi); Ctrl+K bağlantı eklemeye devam eder", async ({ page }) => {
    await page.keyboard.type("Elma armut elma");
    await search(page).click();
    await page.keyboard.type("armut");
    await expect(options(page).last()).toContainText("Belgede bul: armut");
    await page.keyboard.press("Enter");
    await expect(page.locator(".bl-nav-search input")).toHaveValue("armut");
    await body(page).click();
    await page.keyboard.press("Control+k");
    await expect(page.getByRole("dialog")).toContainText("Köprü Ekle");
    await expect(search(page)).not.toBeFocused();
  });

  test("son kullanılanlar kutu boşken listelenir ve yeniden açılışta korunur", async ({ page }) => {
    await page.keyboard.press("Alt+q");
    // Henüz komut kullanılmadı: boş kutuda liste açılmaz.
    await expect(list(page)).toHaveCount(0);
    for (const q of ["italik", "altı çizili"]) {
      await page.keyboard.type(q);
      await page.keyboard.press("ArrowDown");
      await page.keyboard.press("Enter");
      await expect(body(page)).toBeFocused();
      await settle(page);
      await page.keyboard.press("Alt+q");
    }
    await expect(recentGroup(page)).toContainText("Son kullanılanlar");
    await expect(recentGroup(page).getByRole("option")).toHaveCount(2);
    await expect(recentGroup(page).getByRole("option").first()).toContainText("Altı Çizili");
    await page.reload();
    await expect(body(page)).toBeFocused();
    await page.keyboard.press("Alt+q");
    await expect(recentGroup(page).getByRole("option").nth(1)).toContainText("İtalik");
    // Son kullanılandan çalıştırma.
    await page.keyboard.press("ArrowDown");
    await page.keyboard.press("ArrowDown");
    await page.keyboard.press("Enter");
    await expect(body(page)).toBeFocused();
    await page.keyboard.type("eğik");
    await expect(body(page).locator("em")).toHaveText("eğik");
  });

  test("Dosya eylemleri: Seçenekler Dosya görünümünü açar, Farklı Kaydet dosyayı kaydeder", async ({ page }) => {
    await page.keyboard.press("Alt+q");
    await page.keyboard.type("seçenekler");
    await expect(options(page).first()).toContainText("Dosya");
    await page.keyboard.press("ArrowDown");
    await page.keyboard.press("Enter");
    await expect(page.locator(".oh-backstage")).toBeVisible();
    await expect(page.locator(".oh-backstage")).toContainText("Arayüz teması");
    // Dosya görünümü açıkken de şerit komutu çalışır (görünüm kapanır, sekmeye geçilir).
    await settle(page);
    await page.keyboard.press("Alt+q");
    await page.keyboard.type("farklı kaydet");
    const saveAs = options(page).filter({ hasText: "Farklı Kaydet" }).first();
    await expect(saveAs.locator("kbd")).toHaveText("F12");
    const download = page.waitForEvent("download");
    await saveAs.click();
    expect((await download).suggestedFilename()).toMatch(/\.docx$/);
    await settle(page);
    await page.keyboard.press("Alt+q");
    await page.keyboard.type("sözcük sayısı");
    await page.keyboard.press("ArrowDown");
    await page.keyboard.press("Enter");
    await expect(page.locator(".oh-backstage")).toHaveCount(0);
    await expect(page.getByRole("tab", { name: "Gözden Geçir", exact: true })).toHaveAttribute("aria-selected", "true");
    await expect(page.getByRole("dialog")).toContainText("Sözcük");
  });
});

test.describe("Komut araması — Tablo", () => {
  test.beforeEach(async ({ page }) => {
    await page.goto(URL.tablo);
    await page.waitForFunction(() => !!(window as unknown as { __ohTablo?: unknown }).__ohTablo);
  });

  const bold = (page: Page, ref: string) => page.evaluate((r) => (window as unknown as { __ohTablo: { cell(r: string): { bold: boolean } | null } }).__ohTablo.cell(r)?.bold ?? false, ref);

  test("şerit komutu seçili hücrede çalışır; seçim yokken Enter sayfada bulur; Ctrl+K bağlantıda kalır", async ({ page }) => {
    const box = page.getByRole("textbox", { name: "Ad Kutusu" });
    await box.fill("A1");
    await box.press("Enter");
    await page.keyboard.type("elma");
    await page.keyboard.press("Enter");
    await box.fill("A1");
    await box.press("Enter");
    await page.keyboard.press("Alt+q");
    await page.keyboard.type("kalın");
    await expect(options(page).first()).toContainText("Kalın");
    await options(page).first().click();
    await expect.poll(() => bold(page, "A1")).toBe(true);
    await settle(page);
    // Seçim yokken Enter: "Sayfada bul" (eski davranış).
    await page.keyboard.press("Alt+q");
    await page.keyboard.type("elma");
    await expect(options(page).last()).toContainText("Sayfada bul: elma");
    await page.keyboard.press("Enter");
    const find = page.getByRole("dialog", { name: "Bul ve Değiştir" });
    await expect(find).toBeVisible();
    await expect(find.locator("input").first()).toHaveValue("elma");
    await find.getByRole("button", { name: "Kapat" }).first().click();
    await expect(find).toHaveCount(0);
    await box.fill("B2");
    await box.press("Enter");
    await page.keyboard.press("Control+k");
    await expect(page.getByRole("dialog", { name: "Köprü Ekle" })).toBeVisible();
    await expect(search(page)).not.toBeFocused();
  });

  test("son kullanılanlar ve Dosya eylemi (Yazdır)", async ({ page }) => {
    await page.keyboard.press("Alt+q");
    await page.keyboard.type("italik");
    await page.keyboard.press("ArrowDown");
    await page.keyboard.press("Enter");
    await settle(page);
    await page.keyboard.press("Alt+q");
    await expect(recentGroup(page).getByRole("option").first()).toContainText("İtalik");
    await page.keyboard.type("yazdır");
    await expect(options(page).first()).toContainText("Yazdır");
    await expect(options(page).first()).toContainText("Dosya");
    await page.keyboard.press("ArrowDown");
    await page.keyboard.press("Enter");
    await expect(page.locator(".oh-backstage")).toBeVisible();
    await settle(page);
    await page.keyboard.press("Alt+q");
    await expect(recentGroup(page).getByRole("option").first()).toContainText("Yazdır");
    await expect(recentGroup(page).getByRole("option").nth(1)).toContainText("İtalik");
  });
});

test.describe("Komut araması — Sunum", () => {
  const thumbs = (page: Page) => page.locator(".thumb-pane .thumb-item");

  test.beforeEach(async ({ page }) => {
    await page.goto(URL.sunum);
    await expect(page.locator(".stage .sv")).toBeVisible();
  });

  test("mevcut davranış: ilk sonuç seçili gelir, Enter çalıştırır; son kullanılanlar ve Dosya eylemi", async ({ page }) => {
    await page.keyboard.press("Alt+q");
    await expect(search(page)).toBeFocused();
    await page.keyboard.type("SLAYTI GİZLE");
    await expect(options(page).first()).toContainText("Slaytı Gizle");
    await expect(options(page).first()).toHaveAttribute("aria-selected", "true");
    await expect(options(page).last()).toContainText("için Bul");
    await page.keyboard.press("Enter");
    await expect(thumbs(page).first()).toHaveClass(/is-hidden|hidden/);
    await settle(page);
    await page.keyboard.press("Alt+q");
    await expect(recentGroup(page).getByRole("option").first()).toContainText("Slaytı Gizle");
    await page.keyboard.type("seçenekler");
    await page.keyboard.press("Enter");
    await expect(page.locator(".oh-backstage")).toBeVisible();
    await expect(page.locator(".oh-backstage")).toContainText("Arayüz teması");
  });
});
