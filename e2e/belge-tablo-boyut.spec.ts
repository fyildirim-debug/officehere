import { expect, test, type Page } from "@playwright/test";
import { URL, collectErrors } from "./ortak";

// Belge: tablo boyutlandırma kararlı olmalı — sürüklerken zıplamaz, bırakınca sürüklenen boyutta kalır.

const body = (page: Page) => page.locator(".bl-body .ProseMirror");

async function insertTable(page: Page, rows: number, cols: number) {
  await page.getByRole("tab", { name: "Ekle", exact: true }).click();
  await page.locator(".oh-ribbon-body [data-tip='Tablo']").first().click();
  const cell = page.getByRole("button", { name: `${cols}x${rows} tablo` });
  await cell.hover();
  await cell.click();
  await expect(body(page).locator("table")).toHaveCount(1);
}

const colWidths = (page: Page) => body(page).locator("table tr:first-child > *").evaluateAll((els) => els.map((e) => Math.round(e.getBoundingClientRect().width)));
const tableWidth = async (page: Page) => Math.round((await body(page).locator("table").boundingBox())!.width);

test.describe("OfficeHere Belge — tablo boyutlandırma", () => {
  let errors: string[] = [];

  test.beforeEach(async ({ page }) => {
    errors = collectErrors(page);
    await page.goto(URL.belge);
    await expect(body(page)).toBeFocused();
  });

  test.afterEach(() => {
    expect(errors).toEqual([]);
  });

  test("sağ alt tutamaç: sürüklerken tablo düzgün büyür, bırakınca aynı boyutta kalır", async ({ page }) => {
    await insertTable(page, 3, 4);
    const before = await colWidths(page);
    const w0 = await tableWidth(page);
    // Tutamaç tablonun üzerine gelince görünür.
    await body(page).locator("table td").first().hover();
    const handle = page.locator(".bl-tbl-resize");
    await expect(handle).toBeVisible();
    const h = (await handle.boundingBox())!;
    await page.mouse.move(h.x + 3, h.y + 3);
    await page.mouse.down();
    // Sola doğru küçült: her adımda genişlik tutarlı biçimde azalmalı, zıplamamalı.
    const seen: number[] = [];
    for (let i = 1; i <= 6; i++) {
      await page.mouse.move(h.x + 3 - i * 20, h.y + 3);
      seen.push(await tableWidth(page));
    }
    for (let i = 1; i < seen.length; i++) expect(seen[i]).toBeLessThan(seen[i - 1]);
    const during = seen[seen.length - 1];
    expect(Math.abs(during - (w0 - 120))).toBeLessThanOrEqual(4);
    await page.mouse.up();
    // Bırakınca aynı boyutta kalır; sütunlar orantılı küçülür.
    const w1 = await tableWidth(page);
    expect(Math.abs(w1 - during)).toBeLessThanOrEqual(3);
    const after = await colWidths(page);
    const ratio = w1 / w0;
    after.forEach((w, i) => expect(Math.abs(w - before[i] * ratio)).toBeLessThanOrEqual(3));
    // Tek Geri Al adımı eski boyuta döndürür.
    await page.keyboard.press("Control+z");
    await expect.poll(() => tableWidth(page)).toBeGreaterThan(w0 - 4);
  });

  test("sütun kenarı: kenarın iki yanındaki sütunlar değişir, tablo genişliği ve diğer sütunlar korunur", async ({ page }) => {
    await insertTable(page, 3, 4);
    const before = await colWidths(page);
    const first = body(page).locator("table tr:first-child > *").first();
    const b = (await first.boundingBox())!;
    // İlk sütunun sağ kenarı.
    const x = b.x + b.width - 1;
    const y = b.y + b.height / 2;
    await page.mouse.move(x - 6, y);
    await page.mouse.move(x, y, { steps: 3 });
    await expect(page.locator(".bl-doc.resize-cursor, .bl-doc .resize-cursor").first()).toBeVisible();
    await page.mouse.down();
    await page.mouse.move(x + 40, y, { steps: 8 });
    await page.mouse.up();
    const after = await colWidths(page);
    expect(Math.abs(after[0] - (before[0] + 40))).toBeLessThanOrEqual(4);
    expect(Math.abs(after[1] - (before[1] - 40))).toBeLessThanOrEqual(4);
    for (let i = 2; i < before.length; i++) expect(Math.abs(after[i] - before[i])).toBeLessThanOrEqual(2);
    const sum = (xs: number[]) => xs.reduce((x, y) => x + y, 0);
    expect(Math.abs(sum(after) - sum(before))).toBeLessThanOrEqual(3);
    // Tek Geri Al adımı eski genişliklere döndürür.
    await page.keyboard.press("Control+z");
    await expect.poll(async () => (await colWidths(page))[0]).toBeLessThanOrEqual(before[0] + 2);
  });

  test("sütun kenarına basıp bırakmak tabloyu değiştirmez", async ({ page }) => {
    await insertTable(page, 2, 3);
    const before = await colWidths(page);
    const first = body(page).locator("table tr:first-child > *").first();
    const b = (await first.boundingBox())!;
    const x = b.x + b.width - 1;
    const y = b.y + b.height / 2;
    await page.mouse.move(x - 6, y);
    await page.mouse.move(x, y, { steps: 3 });
    await page.mouse.down();
    await page.mouse.up();
    expect(await colWidths(page)).toEqual(before);
  });
});
