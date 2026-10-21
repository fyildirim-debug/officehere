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
const rows = (page: Page) => body(page).locator("table tr:not(.bl-pg)");
const rowHeights = (page: Page) => rows(page).evaluateAll((els) => els.map((e) => Math.round(e.getBoundingClientRect().height)));
const tableHeight = async (page: Page) => Math.round((await body(page).locator("table").boundingBox())!.height);

/** Satırın alt kenarına gelir (imleç satır boyutlandırma imlecine döner) ve kenarı dy kadar sürükler. */
async function dragRowBorder(page: Page, index: number, dy: number, mid?: (dy: number) => Promise<void>) {
  const b = (await rows(page).nth(index).boundingBox())!;
  // İlk sütunun ortası (sütun kenarında sütun boyutlandırma önceliklidir).
  const x = b.x + b.width / 6;
  const y = b.y + b.height - 1;
  await page.mouse.move(x, y - 10);
  await page.mouse.move(x, y, { steps: 3 });
  await expect(page.locator(".bl-doc .bl-row-resize-cursor")).toHaveCount(1);
  await page.mouse.down();
  if (mid) {
    await page.mouse.move(x, y + dy / 2, { steps: 4 });
    await mid(dy / 2);
  }
  await page.mouse.move(x, y + dy, { steps: 4 });
  await page.mouse.up();
}

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

  test("son satırın alt kenarı sürüklenince tablo aşağı uzar; tek Geri Al eski yüksekliğe döndürür", async ({ page }) => {
    await insertTable(page, 3, 3);
    const before = await rowHeights(page);
    const h0 = await tableHeight(page);
    // Sürüklerken tablo uzar.
    await dragRowBorder(page, 2, 60, async (half) => expect(Math.abs((await tableHeight(page)) - (h0 + half))).toBeLessThanOrEqual(3));
    const after = await rowHeights(page);
    expect(Math.abs(after[2] - (before[2] + 60))).toBeLessThanOrEqual(3);
    expect(after.slice(0, 2)).toEqual(before.slice(0, 2));
    expect(Math.abs((await tableHeight(page)) - (h0 + 60))).toBeLessThanOrEqual(3);
    await expect(page.locator(".oh-titlebar")).toContainText("Kaydedilmedi");
    await page.keyboard.press("Control+z");
    await expect.poll(() => tableHeight(page)).toBeLessThanOrEqual(h0 + 1);
  });

  test("ara satırın alt kenarı yalnız o satırı uzatır; kenara basıp bırakmak tabloyu değiştirmez", async ({ page }) => {
    await insertTable(page, 3, 2);
    const before = await rowHeights(page);
    await dragRowBorder(page, 0, 40);
    const after = await rowHeights(page);
    expect(Math.abs(after[0] - (before[0] + 40))).toBeLessThanOrEqual(3);
    expect(after.slice(1)).toEqual(before.slice(1));
    // Tıklayıp bırakmak (sürüklemeden) yüksekliği değiştirmez.
    const now = await rowHeights(page);
    const b = (await rows(page).nth(1).boundingBox())!;
    await page.mouse.move(b.x + 20, b.y + b.height - 10);
    await page.mouse.move(b.x + 20, b.y + b.height - 1, { steps: 3 });
    await page.mouse.down();
    await page.mouse.up();
    expect(await rowHeights(page)).toEqual(now);
  });

  test("sağ alt tutamaç aşağı sürüklenince satırlar orantılı uzar, sütunlar değişmez", async ({ page }) => {
    await insertTable(page, 3, 3);
    const cols0 = await colWidths(page);
    const rows0 = await rowHeights(page);
    const h0 = await tableHeight(page);
    await body(page).locator("table td").first().hover();
    const handle = page.locator(".bl-tbl-resize");
    await expect(handle).toBeVisible();
    const hb = (await handle.boundingBox())!;
    await page.mouse.move(hb.x + 3, hb.y + 3);
    await page.mouse.down();
    await page.mouse.move(hb.x + 3, hb.y + 48, { steps: 6 });
    await page.mouse.up();
    expect(Math.abs((await tableHeight(page)) - (h0 + 45))).toBeLessThanOrEqual(4);
    expect(await colWidths(page)).toEqual(cols0);
    const rows1 = await rowHeights(page);
    rows1.forEach((h, i) => expect(Math.abs(h - (rows0[i] * (h0 + 45)) / h0)).toBeLessThanOrEqual(3));
    // Tutamaç tablonun yeni köşesine taşınır.
    const hb2 = (await handle.boundingBox())!;
    expect(Math.abs(hb2.y - (hb.y + 45))).toBeLessThanOrEqual(4);
  });

  test("Düzen › Yükseklik ve Satırları Dağıt", async ({ page }) => {
    await insertTable(page, 3, 2);
    const tableTab = page.getByRole("tab", { name: "Düzen", exact: true }).last();
    await tableTab.click();
    const spin = page.locator(".oh-ribbon-body").getByRole("textbox", { name: "Tablo Satırı Yüksekliği" });
    await spin.fill("2");
    await spin.press("Enter");
    // 2 cm = 75,6 px (yakınlaştırmaya göre ekranda ölçülür).
    const zoom = await body(page).locator("table").evaluate((t: HTMLElement) => t.getBoundingClientRect().width / t.offsetWidth);
    const h = await rowHeights(page);
    expect(Math.abs(h[0] - 75.6 * zoom)).toBeLessThanOrEqual(2);
    expect(h[1]).toBeLessThan(h[0] / 2);
    await page.locator(".oh-ribbon-body").getByRole("button", { name: "Satırları Dağıt", exact: true }).click();
    const d = await rowHeights(page);
    const avg = h.reduce((x, y) => x + y, 0) / h.length;
    d.forEach((x) => expect(Math.abs(x - avg)).toBeLessThanOrEqual(2));
  });
});
