import { readFileSync } from "node:fs";
import { resolve } from "node:path";
import { expect, test, type Page } from "@playwright/test";
import { URL, collectErrors } from "./ortak";

// OfficeHere Sunum: .odp, .ppt, .ppsx, .otp dosyalarını açma ve ODP olarak kaydetme / dışarı aktarma.

const fixture = (n: string) => readFileSync(resolve(__dirname, "..", "apps/sunum/src/io/__fixtures__", n));
const thumbs = (page: Page) => page.locator(".thumb-pane .thumb-item");
const stage = (page: Page) => page.locator(".stage .sv");
const bar = (page: Page) => page.locator(".oh-messagebar");

async function openApp(page: Page) {
  await page.addInitScript(() => {
    const w = window as unknown as Record<string, unknown>;
    delete w.showSaveFilePicker;
    delete w.showOpenFilePicker;
  });
  await page.goto(URL.sunum);
  await expect(stage(page)).toBeVisible();
  await page.locator(".stage").click({ position: { x: 4, y: 4 } });
}

/** Ctrl+O ile dosya seçiciden açar. */
async function openFile(page: Page, name: string, buffer: Buffer) {
  const [chooser] = await Promise.all([page.waitForEvent("filechooser"), page.keyboard.press("Control+o")]);
  expect(chooser.isMultiple()).toBe(false);
  await chooser.setFiles({ name, mimeType: "application/octet-stream", buffer });
}

test.describe("OfficeHere Sunum dosya biçimleri", () => {
  let errors: string[] = [];

  test.beforeEach(async ({ page }) => {
    errors = collectErrors(page);
    await openApp(page);
  });

  test.afterEach(() => {
    expect(errors).toEqual([]);
  });

  test("Aç süzgeci tüm sunu biçimlerini kabul eder", async ({ page }) => {
    const [chooser] = await Promise.all([page.waitForEvent("filechooser"), page.keyboard.press("Control+o")]);
    const accept = await chooser.element().getAttribute("accept");
    for (const ext of ["pptx", "pptm", "ppt", "odp", "ppsx", "ppsm", "potx", "potm", "otp"]) expect(accept).toContain(`.${ext}`);
    await chooser.setFiles([]);
  });

  test(".odp açılır: slaytlar, metin, resim ve notlar görünür", async ({ page }) => {
    await openFile(page, "lo-bicim.odp", fixture("lo-bicim.odp"));
    await expect(page.locator(".oh-titlebar")).toContainText("lo-bicim");
    await expect(thumbs(page)).toHaveCount(4);
    await expect(stage(page)).toContainText("Çeyrek Sonuçları — Iğdır Şubesi");
    await expect(page.locator(".notes-pane textarea")).toHaveValue("Açılış: herkese hoş geldiniz deyin.");
    await thumbs(page).nth(1).click();
    await expect(stage(page)).toContainText("Akıllı saat çığır açıyor");
    await thumbs(page).nth(2).click();
    await expect(stage(page).locator("img").first()).toBeVisible();
    await expect(stage(page)).toContainText("Kırmızı Kutu");
  });

  test(".ppt açılır (uyumluluk modu); Ctrl+S .pptx olarak farklı kaydeder", async ({ page }) => {
    await openFile(page, "lo-bicim.ppt", fixture("lo-bicim.ppt"));
    await expect(thumbs(page)).toHaveCount(4);
    await expect(bar(page)).toContainText("uyumluluk modu");
    await expect(stage(page)).toContainText("Çeyrek Sonuçları — Iğdır Şubesi");
    await thumbs(page).nth(1).click();
    await expect(stage(page)).toContainText("Soru ve cevap");
    const [download] = await Promise.all([page.waitForEvent("download"), page.keyboard.press("Control+s")]);
    expect(download.suggestedFilename()).toBe("lo-bicim.pptx");
    const bytes = readFileSync((await download.path())!);
    expect(bytes.subarray(0, 2).toString()).toBe("PK");
  });

  test(".ppsx açılır; ileti çubuğundan slayt gösterisi başlatılır", async ({ page }) => {
    await openFile(page, "lo-bicim.ppsx", fixture("lo-bicim.ppsx"));
    await expect(thumbs(page)).toHaveCount(4);
    const start = bar(page).getByRole("button", { name: "Slayt Gösterisini Başlat" });
    await expect(start).toBeVisible();
    await start.click();
    await expect(page.locator(".show-root")).toBeVisible();
    await expect(page.locator(".show-stage")).toContainText("Çeyrek Sonuçları");
    await page.keyboard.press("Escape");
    await expect(page.locator(".show-root")).toHaveCount(0);
  });

  test(".otp şablonu adsız yeni sunu olarak açılır", async ({ page }) => {
    await openFile(page, "lo-bicim.otp", fixture("lo-bicim.otp"));
    await expect(thumbs(page)).toHaveCount(4);
    await expect(page.locator(".oh-titlebar")).not.toContainText("lo-bicim");
    await expect(bar(page)).toContainText("şablonundan yeni sunu");
    const [download] = await Promise.all([page.waitForEvent("download"), page.keyboard.press("Control+s")]);
    expect(download.suggestedFilename()).toMatch(/^Sunu\d+\.pptx$/);
  });

  test(".odp kaydedilirken bir kez biçim sorulur; indirilen dosya yeniden açılır", async ({ page }) => {
    await openFile(page, "lo-bicim.odp", fixture("lo-bicim.odp"));
    await expect(thumbs(page)).toHaveCount(4);
    await page.keyboard.press("Control+s");
    const dlg = page.getByRole("dialog", { name: "Biçimi Koru" });
    await expect(dlg).toContainText("kaybolabilir");
    const [download] = await Promise.all([page.waitForEvent("download"), dlg.getByRole("button", { name: ".odp olarak kaydet" }).click()]);
    expect(download.suggestedFilename()).toBe("lo-bicim.odp");
    const bytes = readFileSync((await download.path())!);
    expect(bytes.subarray(30, 38).toString()).toBe("mimetype");
    // İkinci kayıtta soru yok.
    const [again] = await Promise.all([page.waitForEvent("download"), page.keyboard.press("Control+s")]);
    expect(again.suggestedFilename()).toBe("lo-bicim.odp");
    await expect(dlg).toHaveCount(0);
    // Yeni sunu açıp indirileni yeniden aç.
    await page.keyboard.press("Control+n");
    await expect(thumbs(page)).toHaveCount(1);
    await openFile(page, "lo-bicim.odp", bytes);
    await expect(thumbs(page)).toHaveCount(4);
    await expect(stage(page)).toContainText("Çeyrek Sonuçları — Iğdır Şubesi");
    await thumbs(page).nth(2).click();
    await expect(stage(page)).toContainText("Kırmızı Kutu");
  });

  test("Dışarı Aktar › ODP: yeni sunu .odp olarak indirilir ve yeniden açılır", async ({ page }) => {
    await stage(page).locator("[data-shape-id]").first().click();
    await page.keyboard.type("Açık Belge Başlığı Ğüşİöç");
    await page.keyboard.press("Escape");
    await page.keyboard.press("Escape");
    await page.locator(".oh-ribbon-file").click();
    await page.getByText("Dışarı Aktar", { exact: true }).click();
    const card = page.locator(".oh-bs-action", { hasText: "OpenDocument Sunu (.odp)" });
    await expect(card).toBeVisible();
    await expect(page.locator(".oh-bs-action", { hasText: "Slayt Gösterisi (.ppsx)" })).toBeVisible();
    await expect(page.locator(".oh-bs-action", { hasText: "Şablon (.potx)" })).toBeVisible();
    const [download] = await Promise.all([page.waitForEvent("download"), card.click()]);
    expect(download.suggestedFilename()).toBe("Sunu1.odp");
    const bytes = readFileSync((await download.path())!);
    await page.keyboard.press("Control+n");
    await openFile(page, "Sunu1.odp", bytes);
    await expect(thumbs(page)).toHaveCount(1);
    await expect(stage(page)).toContainText("Açık Belge Başlığı Ğüşİöç");
  });

  test("bozuk dosya anlaşılır bir hatayla reddedilir", async ({ page }) => {
    await openFile(page, "bozuk.ppt", Buffer.from("bu bir sunu değil"));
    await expect(bar(page)).toContainText("Dosya açılamadı");
    await expect(thumbs(page)).toHaveCount(1);
  });
});
