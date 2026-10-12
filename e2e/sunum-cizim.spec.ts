import { expect, test, type Page } from "@playwright/test";
import { URL, collectErrors } from "./ortak";

// Sunum: küçük resim bölmesindeki "Yeni slayt" düğmesi ve şekil çizerken gerçek şekil önizlemesi.

const thumbs = (page: Page) => page.locator(".thumb-pane .thumb-item");

test.describe("OfficeHere Sunum — slayt ekleme ve şekil çizimi", () => {
  let errors: string[] = [];

  test.beforeEach(async ({ page }) => {
    errors = collectErrors(page);
    await page.goto(URL.sunum);
    await expect(page.locator(".stage .sv")).toBeVisible();
  });

  test.afterEach(() => {
    expect(errors).toEqual([]);
  });

  test("küçük resimlerin altındaki + düğmesi sununun sonuna slayt ekler ve onu seçer", async ({ page }) => {
    const add = page.locator(".thumb-pane .thumb-add");
    await expect(add).toBeVisible();
    await expect(add).toHaveAccessibleName("Yeni slayt ekle");
    await add.click();
    await expect(thumbs(page)).toHaveCount(2);
    // Birinci slayt seçiliyken de yeni slayt araya değil sona eklenir.
    await thumbs(page).first().click();
    await add.click();
    await expect(thumbs(page)).toHaveCount(3);
    await expect(page.locator(".thumb-item.is-current .thumb-num-text")).toHaveText("3");
    await expect(page.locator(".oh-statusbar")).toContainText("Slayt 3 / 3");
  });

  test("üçgen çizilirken önizleme dikdörtgen değil üçgendir", async ({ page }) => {
    await page.getByRole("tab", { name: "Ekle", exact: true }).click();
    await page.locator(".oh-ribbon-body [data-tip='Şekiller']").first().click();
    await page.locator(".oh-popover button[title='İkizkenar Üçgen']").first().click();
    const before = await page.locator(".stage .sv [data-shape-id]").count();
    const st = (await page.locator(".stage").boundingBox())!;
    await page.mouse.move(st.x + 400, st.y + 120);
    await page.mouse.down();
    await page.mouse.move(st.x + 560, st.y + 280, { steps: 6 });
    const path = page.locator(".insert-preview-shape path");
    await expect(path).toHaveCount(1);
    await expect(page.locator(".insert-preview")).toHaveCount(0);
    // Üçgenin yolu üç köşeden oluşur: tepe noktası üst kenarın ortasındadır.
    const d = (await path.getAttribute("d")) ?? "";
    const width = Number(await page.locator(".insert-preview-shape").getAttribute("width"));
    const [, tx, ty] = /^M\s*([\d.]+)\s+([\d.]+)/.exec(d) ?? [];
    expect(Math.abs(Number(tx) - width / 2)).toBeLessThan(1);
    expect(Number(ty)).toBe(0);
    expect(d.match(/L/g)).toHaveLength(2);
    await page.mouse.up();
    await expect(page.locator(".insert-preview-shape")).toHaveCount(0);
    await expect(page.locator(".stage .sv [data-shape-id]")).toHaveCount(before + 1);
  });
});
