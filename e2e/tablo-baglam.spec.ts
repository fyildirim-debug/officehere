import { readFileSync } from "node:fs";
import { resolve } from "node:path";
import { expect, test, type Page } from "@playwright/test";
import { URL, collectErrors, interceptFileChooser } from "./ortak";

// OfficeHere Tablo: yüzen nesnelerde (resim, şekil, grafik) sağ tık menüsü nesnenin türüne göre değişir.

const PNG = readFileSync(resolve(__dirname, "../apps/desktop/build/icons/128x128.png"));

interface TabloWindow {
  __ohTablo?: { cellRect(ref: string): { x: number; y: number; w: number; h: number } | null; chartCount(): number };
}

const ribbon = (page: Page) => page.locator(".oh-ribbon-body");
/** Açık menüdeki öğe (onay işaretli öğeler dahil), etiketine göre. */
const item = (page: Page, name: string | RegExp) =>
  page.locator(".oh-popover [role^=menuitem]").filter({ has: page.locator(".oh-menu-label", { hasText: typeof name === "string" ? new RegExp(`^${name.replace(/[.*+?^${}()|[\]\\]/g, "\\$&")}$`) : name }) });
const menuLabels = (page: Page) => page.locator(".oh-popover [role^=menuitem] .oh-menu-label").allTextContents();
/** Yüzen nesnelerin çizim sırası (arkadan öne). */
const stack = (page: Page) => page.locator(".tb-float").evaluateAll((els) => els.map((e) => (e.classList.contains("is-image") ? "resim" : e.classList.contains("is-shape") ? "şekil" : "grafik")));

async function clickCell(page: Page, ref: string) {
  const r = (await page.evaluate((x) => (window as unknown as TabloWindow).__ohTablo!.cellRect(x), ref))!;
  await page.mouse.click(r.x + r.w / 2, r.y + r.h / 2);
}

async function rightClick(page: Page, sel: string) {
  const b = (await page.locator(sel).first().boundingBox())!;
  await page.mouse.click(b.x + Math.min(30, b.width / 2), b.y + Math.min(30, b.height / 2), { button: "right" });
  await expect(page.locator(".oh-popover [role=menu]").first()).toBeVisible();
}

async function insertImage(page: Page) {
  await page.getByRole("tab", { name: "Ekle", exact: true }).click();
  const [chooser] = await Promise.all([page.waitForEvent("filechooser"), ribbon(page).getByRole("button", { name: "Resimler" }).click()]);
  await chooser.setFiles({ name: "simge.png", mimeType: "image/png", buffer: PNG });
  await expect(page.locator(".tb-float.is-image")).toBeVisible();
}

async function insertShape(page: Page, at: string, label: string) {
  await clickCell(page, at);
  await page.getByRole("tab", { name: "Ekle", exact: true }).click();
  await ribbon(page).getByRole("button", { name: "Şekiller" }).click();
  await page.getByRole("menuitem", { name: label, exact: true }).click();
  await expect(page.locator(".tb-float.is-shape").last()).toBeVisible();
}

test.describe("Tablo › nesneye özgü sağ tık menüleri", () => {
  let errors: string[] = [];

  test.beforeEach(async ({ page }) => {
    errors = collectErrors(page);
    interceptFileChooser(page);
    await page.addInitScript(() => {
      delete (window as unknown as Record<string, unknown>).showSaveFilePicker;
    });
    await page.goto(URL.tablo);
    await page.waitForFunction(() => !!(window as unknown as TabloWindow).__ohTablo?.cellRect("A1"));
  });

  test.afterEach(() => {
    expect(errors).toEqual([]);
  });

  test("resim: Resmi Değiştir, Resim Olarak Kaydet ve katman alt menüleri; şekil/grafik komutları yok", async ({ page }) => {
    await insertImage(page);
    await rightClick(page, ".tb-float.is-image");
    const labels = await menuLabels(page);
    for (const l of ["Kes", "Kopyala", "Yapıştır", "Resmi Değiştir...", "Öne Getir", "Arkaya Gönder", "Resim Olarak Kaydet...", "Sil"]) expect(labels).toContain(l);
    for (const l of ["Metni Düzenle", "Grafik Türünü Değiştir...", "Veri Seç..."]) expect(labels).not.toContain(l);
    await item(page, /^Arkaya Gönder/).hover();
    await expect(item(page, "En Arkaya Gönder").locator("svg.oh-arrange-icon")).toHaveCount(1);
    // Alt menüdeki bir adımlık komut (üst menüdeki aynı adlı öğeden sonra gelir).
    await expect(item(page, "Arkaya Gönder").last()).toBeVisible();
    await page.keyboard.press("Escape");
    await page.keyboard.press("Escape");
    await rightClick(page, ".tb-float.is-image");
    const [download] = await Promise.all([page.waitForEvent("download"), item(page, "Resim Olarak Kaydet...").click()]);
    expect(download.suggestedFilename()).toBe("simge.png");
    expect(readFileSync(await download.path()).equals(PNG)).toBe(true);
  });

  test("resim: Resmi Değiştir yeni dosyayı aynı yere koyar", async ({ page }) => {
    await insertImage(page);
    const img = page.locator(".tb-float.is-image img");
    const before = await img.getAttribute("src");
    await rightClick(page, ".tb-float.is-image");
    const svg = Buffer.from('<svg xmlns="http://www.w3.org/2000/svg" width="60" height="30"><rect width="60" height="30" fill="#3a3"/></svg>');
    const [chooser] = await Promise.all([page.waitForEvent("filechooser"), item(page, "Resmi Değiştir...").click()]);
    await chooser.setFiles({ name: "yesil.svg", mimeType: "image/svg+xml", buffer: svg });
    await expect.poll(() => img.getAttribute("src")).not.toBe(before);
    await expect(page.locator(".tb-float.is-image")).toHaveCount(1);
    await page.keyboard.press("Control+z");
    await expect.poll(() => img.getAttribute("src")).toBe(before);
  });

  test("şekil: Metni Düzenle ve Şekli Biçimlendir; resim komutları yok", async ({ page }) => {
    await insertShape(page, "F3", "Oval");
    await rightClick(page, ".tb-float.is-shape");
    const labels = await menuLabels(page);
    for (const l of ["Kes", "Kopyala", "Yapıştır", "Metni Düzenle", "Öne Getir", "Arkaya Gönder", "Şekli Biçimlendir...", "Sil"]) expect(labels).toContain(l);
    for (const l of ["Resmi Değiştir...", "Resim Olarak Kaydet...", "Grafik Türünü Değiştir..."]) expect(labels).not.toContain(l);
    await item(page, "Şekli Biçimlendir...").click();
    await expect(page.getByRole("dialog", { name: "Şekli Biçimlendir" })).toBeVisible();
  });

  test("katman sırası: En Arkaya Gönder ve bir adım Öne Getir", async ({ page }) => {
    await insertImage(page);
    await insertShape(page, "B3", "Dikdörtgen");
    await expect.poll(() => stack(page)).toEqual(["resim", "şekil"]);
    await rightClick(page, ".tb-float.is-shape");
    await item(page, /^Arkaya Gönder/).hover();
    await item(page, "En Arkaya Gönder").click();
    await expect.poll(() => stack(page)).toEqual(["şekil", "resim"]);
    // Şekil artık resmin arkasında: resmin örtmediği kenarından tutulur.
    const b = (await page.locator(".tb-float.is-shape").boundingBox())!;
    const ib = (await page.locator(".tb-float.is-image").boundingBox())!;
    const x = b.x + b.width > ib.x + ib.width ? b.x + b.width - 6 : b.x + 6;
    await page.mouse.click(x, b.y + b.height - 6, { button: "right" });
    await item(page, /^Öne Getir/).hover();
    await expect(item(page, "En Öne Getir")).toBeVisible();
    await item(page, "Öne Getir").last().click();
    await expect.poll(() => stack(page)).toEqual(["resim", "şekil"]);
  });

  test("grafik: grafik komutları korunur, katman alt menüleri eklenir; Shift+F10 aynı menüyü açar", async ({ page }) => {
    const box = page.getByRole("textbox", { name: "Ad Kutusu" });
    await box.fill("A1");
    await box.press("Enter");
    for (const row of [["Ay", "Satış"], ["Ocak", "100"], ["Şubat", "150"]]) {
      await page.keyboard.type(row[0]);
      await page.keyboard.press("Tab");
      await page.keyboard.type(row[1]);
      await page.keyboard.press("Enter");
    }
    await clickCell(page, "A1");
    await page.getByRole("tab", { name: "Ekle", exact: true }).click();
    await ribbon(page).getByRole("button", { name: "Sütun veya Çubuk" }).click();
    await page.locator(".oh-menu-item", { hasText: "Kümelenmiş Sütun" }).first().click();
    await expect(page.locator(".tb-float.is-chart")).toBeVisible();
    await rightClick(page, ".tb-float.is-chart");
    const labels = await menuLabels(page);
    for (const l of ["Grafik Türünü Değiştir...", "Veri Seç...", "Grafiği Taşı...", "Öne Getir", "Arkaya Gönder", "Grafik Alanını Biçimlendir...", "Sil"]) expect(labels).toContain(l);
    for (const l of ["Resmi Değiştir...", "Metni Düzenle"]) expect(labels).not.toContain(l);
    await item(page, /^Öne Getir/).hover();
    await expect(item(page, "En Öne Getir")).toBeVisible();
    await page.keyboard.press("Escape");
    await page.keyboard.press("Escape");
    await expect(page.locator(".oh-popover")).toHaveCount(0);
    // Klavye: seçili grafikte Shift+F10.
    await page.locator(".tb-float.is-chart").focus();
    await page.keyboard.press("Shift+F10");
    await expect(item(page, "Veri Seç...")).toBeVisible();
    await item(page, "Sil").click();
    await expect(page.locator(".tb-float.is-chart")).toHaveCount(0);
  });
});
