import { expect, test, type Page } from "@playwright/test";
import { URL, collectErrors } from "./ortak";

// Mini araç çubuğu: Belge (mevcut davranış), Sunum (metin seçimi ve sağ tık) ve Tablo (hücre bağlam menüsü).

const bar = (page: Page) => page.locator(".oh-minibar");

/** Tablo'nun yalnız okunur test arayüzü (window.__ohTablo). */
type TabloApi = {
  cell(ref: string): { bold: boolean; numFmt: string | null } | null;
  cellRect(ref: string): { x: number; y: number; w: number; h: number } | null;
};
type TabloWindow = { __ohTablo?: TabloApi };
const contextMenu = (page: Page) => page.locator('[data-oh-popover][data-oh-level="1"]');

/** Düzenlenebilir öğedeki ilk metnin [from, to) aralığının ortası (ekran koordinatı). */
async function textPoint(page: Page, selector: string, from: number, to: number) {
  return page.evaluate(
    ([sel, a, z]) => {
      const root = document.querySelector(sel as string)!;
      const t = document.createTreeWalker(root, NodeFilter.SHOW_TEXT).nextNode()!;
      const r = document.createRange();
      r.setStart(t, a as number);
      r.setEnd(t, z as number);
      const b = r.getBoundingClientRect();
      return { x: b.x + b.width / 2, y: b.y + b.height / 2 };
    },
    [selector, from, to] as const,
  );
}

/** Çubuk bağlam menüsünün üstünde durur. */
async function expectAboveMenu(page: Page) {
  const b = (await bar(page).boundingBox())!;
  const m = (await contextMenu(page).boundingBox())!;
  expect(b.y + b.height).toBeLessThanOrEqual(m.y + 1);
}

test.describe("Mini araç çubuğu", () => {
  let errors: string[] = [];
  test.beforeEach(({ page }) => {
    errors = collectErrors(page);
  });
  test.afterEach(() => {
    expect(errors).toEqual([]);
  });

  test("Belge: seçimde ve sağ tıkta açılır, uzaklaşınca kapanır", async ({ page }) => {
    await page.goto(URL.belge);
    const body = page.locator(".bl-body .ProseMirror");
    await expect(body).toBeFocused();
    await page.keyboard.type("Mini araç çubuğu denemesi");
    const at = await textPoint(page, ".bl-body .ProseMirror p", 5, 9);
    await page.mouse.dblclick(at.x, at.y);
    await expect(page.locator(".bl-minibar.oh-minibar")).toBeVisible();
    // Fare çubuktan uzaklaştıkça saydamlaşır, yeterince uzaklaşınca çubuk kapanır.
    await page.mouse.move(at.x + 60, at.y + 160, { steps: 4 });
    await expect.poll(() => bar(page).evaluate((e) => Number(e.style.opacity || 1))).toBeLessThan(1);
    await page.mouse.move(at.x + 500, at.y + 450, { steps: 4 });
    await expect(bar(page)).toHaveCount(0);
    // Sağ tık: bağlam menüsünün üstünde; komut uygulanınca menüyle birlikte kapanır.
    await page.mouse.click(at.x, at.y, { button: "right" });
    await expect(bar(page)).toBeVisible();
    await expect(contextMenu(page)).toBeVisible();
    await expectAboveMenu(page);
    await bar(page).getByRole("button", { name: "İtalik" }).click();
    await expect(body.locator("em")).toHaveText("araç");
    await expect(bar(page)).toHaveCount(0);
    await expect(contextMenu(page)).toHaveCount(0);
  });

  test("Sunum: metin kutusunda seçimde açılır ve Kalın uygular; sağ tıkta menünün üstünde", async ({ page }) => {
    await page.goto(URL.sunum);
    await expect(page.locator(".stage .sv")).toBeVisible();
    await page.locator(".stage .sv [data-shape-id]").first().click();
    await page.keyboard.type("Merhaba dünya sunumu");
    const at = await textPoint(page, ".tb-edit", 8, 13);
    await page.mouse.dblclick(at.x, at.y);
    await expect(bar(page)).toBeVisible();
    await bar(page).getByRole("button", { name: "Kalın" }).click();
    await expect(page.locator(".tb-edit span", { hasText: /^dünya$/ })).toHaveCSS("font-weight", "700");
    // Seçim kipinde komuttan sonra çubuk açık kalır; düğme durumu güncellenir.
    await expect(bar(page).getByRole("button", { name: "Kalın" })).toHaveAttribute("aria-pressed", "true");
    // Düzenleyicide tuşa basılınca kapanır.
    await page.keyboard.press("End");
    await expect(bar(page)).toHaveCount(0);
    // Seçili sözcükte sağ tık: metin menüsüyle birlikte; komut uygulanınca ikisi de kapanır.
    await page.mouse.dblclick(at.x, at.y);
    await expect(bar(page)).toBeVisible();
    await page.mouse.click(at.x, at.y, { button: "right" });
    await expect(bar(page)).toBeVisible();
    await expectAboveMenu(page);
    await bar(page).getByRole("button", { name: "İtalik" }).click();
    await expect(bar(page)).toHaveCount(0);
    await expect(contextMenu(page)).toHaveCount(0);
    await expect(page.locator(".tb-edit span[style*='italic']").first()).toBeVisible();
  });

  test("Tablo: hücrede sağ tıkta menünün üstünde çubuk; Kalın, Yüzde ve yazı tipi listesi", async ({ page }) => {
    await page.goto(URL.tablo);
    await page.waitForFunction(() => !!(window as unknown as TabloWindow).__ohTablo?.cellRect("A1"));
    const box = page.getByRole("textbox", { name: "Ad Kutusu" });
    await box.fill("B3");
    await box.press("Enter");
    await page.keyboard.type("0,25");
    await page.keyboard.press("Enter");
    const r = (await page.evaluate(() => (window as unknown as TabloWindow).__ohTablo!.cellRect("B3")))!;
    const rightClick = () => page.mouse.click(r.x + r.w / 2, r.y + r.h / 2, { button: "right" });
    const cell = () => page.evaluate(() => (window as unknown as TabloWindow).__ohTablo!.cell("B3"));

    await rightClick();
    await expect(bar(page)).toBeVisible();
    await expect(contextMenu(page)).toBeVisible();
    await expectAboveMenu(page);
    await bar(page).getByRole("button", { name: "Kalın" }).click();
    await expect.poll(async () => (await cell())?.bold).toBe(true);
    await expect(bar(page)).toHaveCount(0);
    await expect(contextMenu(page)).toHaveCount(0);

    await rightClick();
    await bar(page).getByRole("button", { name: "Yüzde Stili" }).click();
    await expect.poll(async () => (await cell())?.numFmt).toBe("0%");

    // Yazı tipi listesi açılınca bağlam menüsü gizlenir, çubuk açık kalır; seçimden sonra ikisi de kapanır.
    await rightClick();
    await bar(page).locator(".oh-combo-arrow").first().click();
    await expect(page.locator(".oh-combo-list")).toBeVisible();
    await expect(contextMenu(page)).toBeHidden();
    await expect(bar(page)).toBeVisible();
    await page.locator(".oh-combo-option", { hasText: "Georgia" }).click();
    await expect(bar(page)).toHaveCount(0);
    await expect(page.locator("[data-oh-popover]")).toHaveCount(0);

    // Esc menüyü ve çubuğu birlikte kapatır.
    await rightClick();
    await expect(bar(page)).toBeVisible();
    await page.keyboard.press("Escape");
    await expect(bar(page)).toHaveCount(0);
    await expect(contextMenu(page)).toHaveCount(0);
  });
});
