import { expect, test, type Page } from "@playwright/test";
import { mkdtempSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { URL, collectErrors, interceptFileChooser } from "./ortak";

// OfficeHere Tablo dosya G/Ç Worker'ı (P4.2): büyük dosya ana iş parçacığını dondurmadan açılır
// (Worker'da okunur, kitap dilimler hâlinde kurulur), durum çubuğunda ilerleme ve İptal görünür,
// iptal açık kitabı bozmaz; kaydetme de Worker'da yapılır.

type Api = { cell(ref: string): { value: unknown } | null; cellRect(ref: string): { x: number; y: number; w: number; h: number } | null };
const api = "__ohTablo";

async function openFile(page: Page, file: string) {
  await page.locator(".oh-ribbon-file").click();
  await page.locator(".oh-bs-item", { hasText: "Aç" }).first().click();
  const [chooser] = await Promise.all([page.waitForEvent("filechooser"), page.locator(".oh-bs-action", { hasText: "Gözat" }).first().click()]);
  await chooser.setFiles(file);
}

const cellValue = (page: Page, ref: string) => page.evaluate(([k, r]) => (window as unknown as Record<string, Api>)[k].cell(r)?.value ?? null, [api, ref] as const);

/** Uzun görev gözlemcisi (sayfa yüklenmeden kurulur). */
function watchLongTasks(page: Page) {
  return page.addInitScript(() => {
    const w = window as unknown as { __lt: number[] };
    w.__lt = [];
    try {
      new PerformanceObserver((l) => {
        for (const e of l.getEntries()) w.__lt.push(Math.round(e.duration));
      }).observe({ type: "longtask", buffered: true });
    } catch {
      /* gözlemci yoksa ölçülmez */
    }
  });
}

/** Büyük CSV: Worker'da okunması birkaç saniye sürer (iptal edilecek kadar). */
function bigCsv(rows: number): string {
  const file = join(mkdtempSync(join(tmpdir(), "oh-worker-")), "buyuk-veri.csv");
  const cities = ["İzmir", "Ankara", "Bursa", "Muğla", "Adana"];
  const lines = ["No;Şehir;Tutar;Not"];
  for (let i = 1; i <= rows; i++) lines.push(`${i};${cities[i % 5]};${(i * 7) % 997};not ${i % 13}`);
  writeFileSync(file, lines.join("\n") + "\n", "utf8");
  return file;
}

test.describe("Tablo dosya G/Ç Worker'ı", () => {
  let errors: string[] = [];
  test.beforeEach(async ({ page }) => {
    errors = collectErrors(page);
    interceptFileChooser(page);
    await watchLongTasks(page);
    await page.addInitScript(() => {
      const w = window as unknown as Record<string, unknown>;
      delete w.showSaveFilePicker;
      delete w.showOpenFilePicker;
    });
    await page.goto(URL.tablo);
    await page.waitForFunction((k) => !!(window as unknown as Record<string, Api>)[k]?.cellRect("A1"), api);
  });
  test.afterEach(() => {
    expect(errors).toEqual([]);
  });

  test("büyük dosya açılırken ilerleme ve İptal görünür; iptal 1 sn içinde, açık kitap korunur", async ({ page }) => {
    test.setTimeout(120_000);
    const file = bigCsv(400_000);
    // Açık kitapta bir değer: iptalden sonra yerinde kalmalı
    const box = page.getByRole("textbox", { name: "Ad Kutusu" });
    await box.fill("A1");
    await box.press("Enter");
    await page.keyboard.type("ESKİ");
    await page.keyboard.press("Enter");
    // Kaydedilmemiş değişiklik sorusu: kaydetmeden aç (iptal edilirse kitap olduğu gibi kalır)
    await page.locator(".oh-ribbon-file").click();
    await page.locator(".oh-bs-item", { hasText: "Aç" }).first().click();
    await page.locator(".oh-bs-action", { hasText: "Gözat" }).first().click();
    const ask = page.getByRole("dialog", { name: "OfficeHere" });
    await expect(ask).toContainText("kaydetmek istiyor musunuz");
    const [chooser] = await Promise.all([page.waitForEvent("filechooser"), ask.getByRole("button", { name: "Kaydetme" }).click()]);
    await chooser.setFiles(file);
    const progress = page.locator(".oh-progress");
    await expect(progress).toBeVisible({ timeout: 30_000 });
    await expect(page.getByRole("progressbar", { name: "Açılıyor…" })).toBeVisible();
    const cancel = progress.getByRole("button", { name: "İptal" });
    const t0 = Date.now();
    await cancel.click();
    await expect(progress).toHaveCount(0, { timeout: 1000 });
    const ms = Date.now() - t0;
    await expect(page.locator(".oh-titlebar")).toContainText("Kitap1");
    await expect(page.getByRole("dialog")).toHaveCount(0);
    expect(await cellValue(page, "A1")).toBe("ESKİ");
    console.log(`iptal: ${ms} ms`);
    expect(ms).toBeLessThan(1000);
  });

  test("büyük dosya ana iş parçacığını dondurmadan açılır ve Worker'da kaydedilir", async ({ page }) => {
    test.setTimeout(180_000);
    const file = bigCsv(300_000);
    const before = await page.evaluate(() => (window as unknown as { __lt: number[] }).__lt.length);
    await openFile(page, file);
    await expect(page.locator(".oh-titlebar")).toContainText("buyuk-veri", { timeout: 120_000 });
    await expect.poll(() => cellValue(page, "B300001")).toBe("İzmir");
    expect(await cellValue(page, "C2")).toBe(7);
    // Açılış süresince en uzun görev: Worker'sız açılış tek görevde saniyeler sürüyordu
    const lt = await page.evaluate((n) => (window as unknown as { __lt: number[] }).__lt.slice(n), before);
    console.log(`açılış uzun görevleri: ${JSON.stringify(lt)}`);
    expect(Math.max(0, ...lt)).toBeLessThan(500);
    // .xlsx olarak dışa aktar: Worker'da yazılır
    await page.locator(".oh-ribbon-file").click();
    await page.locator(".oh-bs-item", { hasText: "Dışarı Aktar" }).first().click();
    const [download] = await Promise.all([page.waitForEvent("download", { timeout: 120_000 }), page.locator(".oh-bs-action", { hasText: /\.xlsx/ }).first().click()]);
    expect(download.suggestedFilename()).toBe("buyuk-veri.xlsx");
  });
});
