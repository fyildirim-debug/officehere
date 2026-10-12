import { expect, test, type Page } from "@playwright/test";
import { URL, collectErrors } from "./ortak";

// Klavye erişilebilirliği: şerit sekmeleri, menü düğmeleri, bağlam menüleri ve açılan kutular
// fare kullanmadan (Alt tuş ipuçları yerine odakla) kullanılabilir.

const focusedIn = (page: Page, selector: string) => page.locator(`${selector} :focus`);
const focusedText = (page: Page) => page.evaluate(() => (document.activeElement as HTMLElement | null)?.textContent?.trim() ?? "");

test.describe("Klavye: Belge", () => {
  let errors: string[] = [];
  const body = (page: Page) => page.locator(".bl-body .ProseMirror");
  const ribbon = (page: Page) => page.locator(".oh-ribbon-body");
  const tab = (page: Page, name: string) => page.getByRole("tab", { name, exact: true });

  test.beforeEach(async ({ page }) => {
    errors = collectErrors(page);
    await page.goto(URL.belge);
    await expect(body(page)).toBeFocused();
  });

  test.afterEach(() => {
    expect(errors).toEqual([]);
  });

  test("sekmeler: ←/→ odak ve seçim, Home/End, uçlarda döner; ↓ ve Tab gövdeye iner", async ({ page }) => {
    await tab(page, "Giriş").focus();
    await page.keyboard.press("ArrowRight");
    await expect(tab(page, "Ekle")).toBeFocused();
    await expect(tab(page, "Ekle")).toHaveAttribute("aria-selected", "true");
    await expect(tab(page, "Ekle")).toHaveAttribute("tabindex", "0");
    await expect(tab(page, "Giriş")).toHaveAttribute("tabindex", "-1");
    await expect(ribbon(page).getByRole("button", { name: "Tablo", exact: true })).toBeVisible();
    await page.keyboard.press("ArrowLeft");
    await expect(tab(page, "Giriş")).toBeFocused();
    await expect(tab(page, "Giriş")).toHaveAttribute("aria-selected", "true");
    await page.keyboard.press("ArrowLeft");
    await expect(tab(page, "Yardım")).toBeFocused();
    await page.keyboard.press("ArrowRight");
    await expect(tab(page, "Giriş")).toBeFocused();
    await page.keyboard.press("End");
    await expect(tab(page, "Yardım")).toHaveAttribute("aria-selected", "true");
    await page.keyboard.press("Home");
    await expect(tab(page, "Giriş")).toHaveAttribute("aria-selected", "true");
    // ↓ gövdedeki ilk denetime iner; gövdede Tab sıradaki denetime geçer.
    await page.keyboard.press("ArrowDown");
    await expect(focusedIn(page, ".oh-ribbon-body")).toHaveCount(1);
    const first = await focusedText(page);
    await page.keyboard.press("Tab");
    await expect(focusedIn(page, ".oh-ribbon-body")).toHaveCount(1);
    expect(await page.evaluate(() => document.activeElement?.getAttribute("aria-label"))).not.toBe(first);
    // Sekmeden Tab da gövdeye iner.
    await tab(page, "Giriş").focus();
    await page.keyboard.press("Tab");
    await expect(focusedIn(page, ".oh-ribbon-body")).toHaveCount(1);
  });

  test("menü düğmesi klavyeyle açılır; ↓/↑, Home/End, harfle atlama, Enter çalıştırır, Esc odağı düğmeye verir", async ({ page }) => {
    await page.keyboard.type("merhaba dünya");
    await page.keyboard.press("Control+a");
    const btn = ribbon(page).getByRole("button", { name: "Büyük/Küçük Harf Değiştir" });
    await btn.focus();
    // ↓ açar, ilk öğeye odaklanır; Esc kapatıp odağı düğmeye verir.
    await page.keyboard.press("ArrowDown");
    const menu = page.getByRole("menu");
    await expect(menu.getByRole("menuitem", { name: "Cümle düzeni." })).toBeFocused();
    await page.keyboard.press("End");
    await expect(menu.getByRole("menuitem", { name: /^bÜYÜK/ })).toBeFocused();
    await page.keyboard.press("ArrowDown");
    await expect(menu.getByRole("menuitem", { name: "Cümle düzeni." })).toBeFocused();
    await page.keyboard.press("ArrowUp");
    await expect(menu.getByRole("menuitem", { name: /^bÜYÜK/ })).toBeFocused();
    await page.keyboard.press("Home");
    await page.keyboard.press("h");
    await expect(menu.getByRole("menuitem", { name: "Her Sözcüğün İlk Harfi Büyük" })).toBeFocused();
    await page.keyboard.press("Escape");
    await expect(menu).toHaveCount(0);
    await expect(btn).toBeFocused();
    await expect(btn).toHaveAttribute("aria-expanded", "false");
    // Enter açar; ↓↓ ve Enter "BÜYÜK HARF" uygular.
    await page.keyboard.press("Enter");
    await expect(menu.getByRole("menuitem", { name: "Cümle düzeni." })).toBeFocused();
    await page.keyboard.press("ArrowDown");
    await page.keyboard.press("ArrowDown");
    await expect(menu.getByRole("menuitem", { name: "BÜYÜK HARF" })).toBeFocused();
    await page.keyboard.press("Enter");
    await expect(menu).toHaveCount(0);
    await expect(body(page)).toHaveText("MERHABA DÜNYA");
    // Tab menüyü kapatır ve sıradaki denetime geçer.
    await btn.focus();
    await page.keyboard.press("Space");
    await expect(menu).toBeVisible();
    await page.keyboard.press("Tab");
    await expect(menu).toHaveCount(0);
    await expect(btn).not.toBeFocused();
    await expect(focusedIn(page, ".oh-ribbon-body")).toHaveCount(1);
  });

  test("fareyle açılan menüde odak menüdedir ama öğe vurgulanmaz; ↓ ilk öğeye gider", async ({ page }) => {
    const btn = ribbon(page).getByRole("button", { name: "Büyük/Küçük Harf Değiştir" });
    await btn.click();
    const menu = page.getByRole("menu");
    await expect(menu).toBeFocused();
    await expect(page.locator(".oh-menu-item:focus")).toHaveCount(0);
    await page.keyboard.press("ArrowDown");
    await expect(menu.getByRole("menuitem", { name: "Cümle düzeni." })).toBeFocused();
    await page.keyboard.press("Escape");
    await expect(menu).toHaveCount(0);
    // Fareyle açıldıysa odak açılmadan önceki yere (belgeye) döner.
    await expect(body(page)).toBeFocused();
  });

  test("bağlam menüsü: ok tuşları, alt menü →/←, Esc odağı belgeye geri verir", async ({ page }) => {
    await page.keyboard.type("Bağlam menüsü");
    const p = body(page).locator("p").first();
    const box = (await p.boundingBox())!;
    await page.mouse.click(box.x + 20, box.y + box.height / 2, { button: "right" });
    const menu = page.getByRole("menu");
    await expect(menu).toBeFocused();
    // Seçim yokken Kes/Kopyala devre dışıdır ve atlanır.
    await page.keyboard.press("ArrowDown");
    const paste = menu.getByRole("menuitem", { name: "Yapıştırma Seçenekleri:" });
    await expect(paste).toBeFocused();
    await page.keyboard.press("ArrowRight");
    await expect(page.getByRole("menuitem", { name: "Kaynak Biçimlendirmesini Koru (K)" })).toBeFocused();
    await page.keyboard.press("ArrowDown");
    await expect(page.getByRole("menuitem", { name: "Biçimlendirmeyi Birleştir (B)" })).toBeFocused();
    await page.keyboard.press("ArrowLeft");
    await expect(page.getByRole("menuitem", { name: "Kaynak Biçimlendirmesini Koru (K)" })).toHaveCount(0);
    await expect(paste).toBeFocused();
    await page.keyboard.press("ArrowDown");
    await expect(paste).not.toBeFocused();
    await page.keyboard.press("ArrowUp");
    await expect(paste).toBeFocused();
    // Alt menüde Esc yalnız alt menüyü kapatır; ikinci Esc menüyü kapatıp odağı belgeye verir.
    await page.keyboard.press("Enter");
    await expect(page.getByRole("menuitem", { name: "Kaynak Biçimlendirmesini Koru (K)" })).toBeFocused();
    await page.keyboard.press("Escape");
    await expect(paste).toBeFocused();
    await page.keyboard.press("Escape");
    await expect(menu).toHaveCount(0);
    await expect(body(page)).toBeFocused();
    // Yazılan metin belgeye (sağ tıklanan yerdeki imlece) gider.
    await page.keyboard.type("!");
    await expect(body(page)).toContainText("!");
    await expect.poll(async () => (await body(page).textContent())?.replace("!", "")).toBe("Bağlam menüsü");
  });

  test("yazı tipi kutusu: Alt+↓ açar, ↓ etkin öğeyi taşır, Enter seçer; Esc önceki değeri korur", async ({ page }) => {
    const font = ribbon(page).getByRole("combobox", { name: "Yazı Tipi", exact: true });
    await font.focus();
    const before = await font.inputValue();
    const list = page.getByRole("listbox", { name: "Yazı Tipi" });
    // Esc: liste kapanır, değer değişmez.
    await page.keyboard.press("Alt+ArrowDown");
    await expect(list).toBeVisible();
    await expect(font).toHaveAttribute("aria-expanded", "true");
    await page.keyboard.press("ArrowDown");
    await page.keyboard.press("ArrowDown");
    await expect(font).toHaveAttribute("aria-activedescendant", /.+/);
    await page.keyboard.press("Escape");
    await expect(list).toHaveCount(0);
    await expect(font).toHaveValue(before);
    // ↓ ile sonraki yazı tipi etkinleşir, Enter seçer.
    await page.keyboard.press("ArrowDown");
    await expect(list).toBeVisible();
    await page.keyboard.press("ArrowDown");
    const active = (await font.getAttribute("aria-activedescendant"))!;
    const option = page.locator(`[id="${active}"]`);
    await expect(option).toHaveAttribute("role", "option");
    await expect(option).toHaveClass(/is-active/);
    const chosen = (await option.textContent())!.trim();
    expect(chosen).not.toBe(before);
    await page.keyboard.press("Enter");
    await expect(list).toHaveCount(0);
    await expect(font).toHaveValue(chosen);
    await page.keyboard.type("Yazı");
    await expect(body(page).locator(`span[style*="${chosen}"]`)).toHaveText("Yazı");
  });

  test("galeri: ok tuşları öğeler arasında gezer, galeri tek Tab durağıdır", async ({ page }) => {
    const tiles = ribbon(page).locator(".oh-gallery-items .oh-gallery-item");
    await expect(tiles.locator('xpath=self::*[@tabindex="0"]')).toHaveCount(1);
    await tiles.first().focus();
    await page.keyboard.press("ArrowRight");
    await expect(tiles.nth(1)).toBeFocused();
    await expect(tiles.nth(1)).toHaveAttribute("tabindex", "0");
    await page.keyboard.press("ArrowLeft");
    await expect(tiles.first()).toBeFocused();
    // Alt+↓ genişletilmiş galeriyi açar; ızgarada ok tuşlarıyla gezilir.
    await page.keyboard.press("Alt+ArrowDown");
    const grid = page.locator(".oh-gallery-menu");
    await expect(grid).toBeVisible();
    await expect(grid.locator(".oh-gallery-item").first()).toBeFocused();
    await page.keyboard.press("ArrowRight");
    await expect(grid.locator(".oh-gallery-item").nth(1)).toBeFocused();
    await page.keyboard.press("Escape");
    await expect(grid).toHaveCount(0);
    await expect(tiles.first()).toBeFocused();
  });
});

test.describe("Klavye: Tablo", () => {
  test("sekmeler ←/→ ve hücre bağlam menüsünde ok tuşları; Esc odağı ızgaraya verir", async ({ page }) => {
    const errors = collectErrors(page);
    await page.goto(URL.tablo);
    await page.waitForFunction(() => !!(window as unknown as { __ohTablo?: unknown }).__ohTablo);
    const tab = (name: string) => page.getByRole("tab", { name, exact: true });
    await tab("Giriş").focus();
    await page.keyboard.press("ArrowRight");
    await expect(tab("Ekle")).toBeFocused();
    await expect(tab("Ekle")).toHaveAttribute("aria-selected", "true");
    await page.keyboard.press("Home");
    await expect(tab("Giriş")).toHaveAttribute("aria-selected", "true");

    // Izgaradayken sağ tık: Esc sonrası odak yine ızgarada (klavyeyle hücre gezilebilir).
    const r = await page.evaluate(() => (window as unknown as { __ohTablo: { cellRect(x: string): { x: number; y: number; w: number; h: number } } }).__ohTablo.cellRect("B2"));
    await page.mouse.click(r.x + r.w / 2, r.y + r.h / 2);
    await expect(page.getByRole("textbox", { name: "Hücre düzenleyici" })).toBeFocused();
    await page.mouse.click(r.x + r.w / 2, r.y + r.h / 2, { button: "right" });
    const menu = page.getByRole("menu");
    await expect(menu).toBeFocused();
    await page.keyboard.press("ArrowDown");
    await expect(menu.getByRole("menuitem", { name: "Kes" })).toBeFocused();
    await page.keyboard.press("ArrowDown");
    await expect(menu.getByRole("menuitem", { name: "Kopyala" })).toBeFocused();
    await page.keyboard.press("Escape");
    await expect(menu).toHaveCount(0);
    await expect(page.getByRole("textbox", { name: "Hücre düzenleyici" })).toBeFocused();
    expect(errors).toEqual([]);
  });
});

test.describe("Klavye: Sunum", () => {
  test("sekmeler ←/→; bölünmüş düğmede ↓ menüyü açar, Esc odağı düğmeye verir", async ({ page }) => {
    const errors = collectErrors(page);
    await page.goto(URL.sunum);
    await expect(page.locator(".stage")).toBeVisible();
    const tab = (name: string) => page.getByRole("tab", { name, exact: true });
    await tab("Giriş").focus();
    await page.keyboard.press("End");
    await expect(tab("Yardım")).toBeFocused();
    await expect(tab("Yardım")).toHaveAttribute("aria-selected", "true");
    await page.keyboard.press("ArrowRight");
    await expect(tab("Giriş")).toHaveAttribute("aria-selected", "true");

    const btn = page.locator(".oh-ribbon-body").getByRole("button", { name: "Yeni Slayt", exact: true });
    await btn.focus();
    await page.keyboard.press("ArrowDown");
    const menu = page.locator(".oh-popover");
    await expect(menu).toBeVisible();
    await expect(page.locator(".oh-popover :focus")).toHaveCount(1);
    await page.keyboard.press("ArrowDown");
    await expect(page.locator(".oh-popover :focus")).toHaveCount(1);
    await page.keyboard.press("Escape");
    await expect(menu).toHaveCount(0);
    await expect(btn).toBeFocused();
    expect(errors).toEqual([]);
  });
});
