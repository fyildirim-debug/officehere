import { expect, test, type Page } from "@playwright/test";
import { URL, collectErrors } from "./ortak";

// Belge › Görünüm › Koyu Sayfa: yalnız ekrandaki sayfa koyulaşır; seçim hatırlanır, yazdırma ve önizleme beyaz kalır.

const DARK = "rgb(31, 31, 31)";
const WHITE = "rgb(255, 255, 255)";

const body = (page: Page) => page.locator(".bl-body .ProseMirror");
const ribbon = (page: Page) => page.locator(".oh-ribbon-body");
const firstPage = (page: Page) => page.locator(".bl-pages > .bl-page").first();
const darkButton = (page: Page) => ribbon(page).getByRole("button", { name: "Koyu Sayfa", exact: true });
const bg = (page: Page, sel: string) => page.locator(sel).first().evaluate((e) => getComputedStyle(e).backgroundColor);
const fill = (page: Page, sel: string) => page.locator(sel).first().evaluate((e) => getComputedStyle(e).webkitTextFillColor);

test.describe("Belge koyu sayfa", () => {
  let errors: string[] = [];

  test.beforeEach(async ({ page }) => {
    errors = collectErrors(page);
    await page.goto(URL.belge);
    await expect(body(page)).toBeFocused();
  });

  test.afterEach(() => {
    expect(errors).toEqual([]);
  });

  test("düğmeyle sayfa koyulaşır, metin açılır; vurgulu metin koyu kalır", async ({ page }) => {
    await page.keyboard.type("Koyu sayfada okunur metin");
    // Son sözcüğü vurgula.
    await page.keyboard.press("Control+Shift+ArrowLeft");
    await ribbon(page).getByRole("button", { name: "Metin Vurgu Rengi" }).first().click();
    await expect(body(page).locator("mark")).toHaveText("metin");
    expect(await bg(page, ".bl-pages > .bl-page")).toBe(WHITE);

    await page.getByRole("tab", { name: "Görünüm", exact: true }).click();
    await darkButton(page).click();
    await expect(darkButton(page)).toHaveAttribute("aria-pressed", "true");
    await expect(firstPage(page)).toHaveCSS("background-color", DARK);
    // Belgedeki renk değişmez (otomatik = siyah), ekranda açık çizilir.
    await expect(body(page).locator("p").first()).toHaveCSS("color", "rgb(0, 0, 0)");
    expect(await fill(page, ".bl-body .ProseMirror p")).not.toBe("rgb(0, 0, 0)");
    // Vurgulu metin kâğıttaki gibi koyu kalır.
    expect(await fill(page, ".bl-body mark")).toBe("rgb(0, 0, 0)");

    await darkButton(page).click();
    await expect(darkButton(page)).toHaveAttribute("aria-pressed", "false");
    await expect(firstPage(page)).toHaveCSS("background-color", WHITE);
  });

  test("seçim yeniden yüklemede hatırlanır", async ({ page }) => {
    await page.getByRole("tab", { name: "Görünüm", exact: true }).click();
    await darkButton(page).click();
    await expect(firstPage(page)).toHaveCSS("background-color", DARK);
    await page.reload();
    await expect(body(page)).toBeVisible();
    await expect(firstPage(page)).toHaveCSS("background-color", DARK);
    await page.getByRole("tab", { name: "Görünüm", exact: true }).click();
    await expect(darkButton(page)).toHaveAttribute("aria-pressed", "true");
    // Kapatılınca da hatırlanır.
    await darkButton(page).click();
    await page.reload();
    await expect(body(page)).toBeVisible();
    await expect(firstPage(page)).toHaveCSS("background-color", WHITE);
  });

  test("Baskı Önizleme ve yazdırma beyaz sayfayla, siyah metinle", async ({ page }) => {
    await page.keyboard.type("Yazdırılacak metin");
    await page.getByRole("tab", { name: "Görünüm", exact: true }).click();
    await darkButton(page).click();
    await expect(firstPage(page)).toHaveCSS("background-color", DARK);

    await page.evaluate(() => {
      const w = window as unknown as { __printed: string[][]; print: () => void };
      w.__printed = [];
      w.print = () => {
        const root = document.getElementById("bl-print-root")!;
        const cs = (s: string) => getComputedStyle(root.querySelector(s)!);
        w.__printed.push([cs(".bl-page").backgroundColor, cs(".bl-body p").webkitTextFillColor]);
      };
    });
    await page.keyboard.press("Control+p");
    const sheet = page.locator(".bl-pr-sheet");
    await expect(sheet).toContainText("Yazdırılacak metin");
    expect(await bg(page, ".bl-pr-sheet .bl-page")).toBe(WHITE);
    expect(await fill(page, ".bl-pr-sheet .bl-body p")).toBe("rgb(0, 0, 0)");
    await page.locator(".bl-pr-print").click();
    await expect.poll(() => page.evaluate(() => (window as unknown as { __printed: string[][] }).__printed)).toEqual([[WHITE, "rgb(0, 0, 0)"]]);
  });
});
