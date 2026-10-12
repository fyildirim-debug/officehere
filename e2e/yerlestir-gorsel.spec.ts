import { expect, test, type Page } from "@playwright/test";
import { URL, collectErrors } from "./ortak";

// Yerleştir menüleri (Konum, Metni Kaydır, Öne Getir / Arkaya Gönder, Hizala): her seçenekte ne yaptığını
// gösteren kendi çizimimiz (svg.oh-arrange-icon) var ve seçimler resme uygulanır.

const PNG = Buffer.from("iVBORw0KGgoAAAANSUhEUgAAABAAAAAMCAIAAADkharWAAAAGklEQVR4nGO4HyNEEmIY1TCqYVTDqAYqagAAcKKm8RLe1JYAAAAASUVORK5CYII=", "base64");

const body = (page: Page) => page.locator(".bl-body .ProseMirror");
const ribbon = (page: Page) => page.locator(".oh-ribbon-body");
const button = (page: Page, name: string) => ribbon(page).getByRole("button", { name, exact: true }).first();
const image = (page: Page) => body(page).locator('img[src^="data:image/png"]');
const tile = (page: Page, name: RegExp | string) => page.getByRole("menuitemradio", { name: typeof name === "string" ? new RegExp(`^${name}`) : name });

test.describe("Belge › Yerleştir: görsel önizlemeli menüler", () => {
  let errors: string[] = [];

  test.beforeEach(async ({ page }) => {
    errors = collectErrors(page);
    await page.goto(URL.belge);
    await expect(body(page)).toBeFocused();
    for (let i = 0; i < 4; i++) {
      await page.keyboard.type(`Paragraf ${i + 1}. ${"Yerleşim deneme metni. ".repeat(8)}`);
      await page.keyboard.press("Enter");
    }
    await page.keyboard.press("Control+Home");
    await page.setInputFiles('input[type=file][accept="image/*"]', { name: "mavi.png", mimeType: "image/png", buffer: PNG });
    await expect(image(page)).toBeVisible();
    await page.getByRole("tab", { name: "Resim Biçimi", exact: true }).click();
  });

  test.afterEach(() => {
    expect(errors).toEqual([]);
  });

  test("Konum: 1 + 9 sayfa önizlemesi; konum seçilince resim yerleşir ve galeride seçili görünür", async ({ page }) => {
    await button(page, "Konum").click();
    const menu = page.getByRole("menu", { name: "Konum" });
    await expect(menu.getByRole("menuitemradio")).toHaveCount(10);
    await expect(menu.locator("[role=menuitemradio] svg.oh-arrange-icon")).toHaveCount(10);
    await expect(menu).toContainText("Metinle Aynı Hizada");
    await expect(menu).toContainText("Metin Kaydırmalı");
    await expect(tile(page, "Metinle Aynı Hizada")).toHaveAttribute("aria-checked", "true");
    await expect(tile(page, "Sağ Üst")).toHaveAttribute("data-tip", "Sağ Üst");

    await tile(page, "Sağ Üst").click();
    await expect(menu).toHaveCount(0);
    await expect(body(page).locator("img[data-wrap=square][data-float=right]")).toHaveCount(1);
    // Metin alanı: resmin bağlı olduğu paragrafın kutusu.
    const para = async () => image(page).evaluate((el) => {
      const r = el.closest("p")!.getBoundingClientRect();
      return { x: r.x, y: r.y, width: r.width };
    });
    const b = await para();
    const r = (await image(page).boundingBox())!;
    // Resim sayfanın sağ üstünde: sağ kenarı metin alanının sağ kenarında, metin solundan akar.
    expect(Math.abs(r.x + r.width - (b.x + b.width))).toBeLessThan(4);
    expect(r.y - b.y).toBeLessThan(40);

    await button(page, "Konum").click();
    await expect(tile(page, "Sağ Üst")).toHaveAttribute("aria-checked", "true");
    await expect(tile(page, "Metinle Aynı Hizada")).toHaveAttribute("aria-checked", "false");
    await tile(page, /^Orta, Üst ve Alt/).click();
    await expect(body(page).locator("img[data-wrap=topBottom][data-float=center]")).toHaveCount(1);
    const c = (await image(page).boundingBox())!;
    const pc = await para();
    expect(Math.abs(c.x + c.width / 2 - (pc.x + pc.width / 2))).toBeLessThan(4);

    await button(page, "Konum").click();
    await tile(page, "Metinle Aynı Hizada").click();
    await expect(body(page).locator("img[data-wrap]")).toHaveCount(0);
  });

  test("Konum galerisi klavyeyle: oklarla ızgarada gezilir, Enter seçer", async ({ page }) => {
    await button(page, "Konum").focus();
    await page.keyboard.press("Enter");
    await expect(tile(page, "Metinle Aynı Hizada")).toBeFocused();
    await page.keyboard.press("ArrowDown");
    await expect(tile(page, "Sol Üst")).toBeFocused();
    await page.keyboard.press("ArrowRight");
    await expect(tile(page, "Orta Üst")).toBeFocused();
    await page.keyboard.press("ArrowDown");
    await expect(tile(page, /^Orta, /)).toBeFocused();
    await page.keyboard.press("ArrowRight");
    await expect(tile(page, "Sağ Orta")).toBeFocused();
    await page.keyboard.press("Enter");
    await expect(body(page).locator("img[data-wrap=square][data-float=right]")).toHaveCount(1);
  });

  test("Metni Kaydır: her seçenekte öğretici çizim; seçim resme uygulanır", async ({ page }) => {
    await button(page, "Metni Kaydır").click();
    const menu = page.getByRole("menu", { name: "Metni Kaydır" });
    const items = menu.getByRole("menuitemcheckbox");
    await expect(items).toHaveText(["Metinle Aynı Hizada", "Kare", "Sıkı", "Üst ve Alt", "Metnin Arkasında", "Metnin Önünde"]);
    for (let i = 0; i < 6; i++) await expect(items.nth(i).locator("svg.oh-arrange-icon")).toHaveCount(1);
    await expect(items.first()).toHaveAttribute("aria-checked", "true");

    const pick = async (label: string, wrap: string) => {
      if (!(await page.getByRole("menu", { name: "Metni Kaydır" }).isVisible())) await button(page, "Metni Kaydır").click();
      await page.getByRole("menuitemcheckbox", { name: label, exact: true }).click();
      await expect(body(page).locator(`img[data-wrap=${wrap}]`)).toHaveCount(1);
      await button(page, "Metni Kaydır").click();
      await expect(page.getByRole("menuitemcheckbox", { name: label, exact: true })).toHaveAttribute("aria-checked", "true");
    };
    await pick("Sıkı", "tight");
    await pick("Üst ve Alt", "topBottom");
    await pick("Metnin Önünde", "front");
    await pick("Kare", "square");
  });

  test("Öne Getir / Arkaya Gönder ve Hizala: katman çizimleri; metnin önüne/arkasına ve en öne", async ({ page }) => {
    await ribbon(page).getByRole("button", { name: "Arkaya Gönder seçenekleri" }).click();
    let menu = page.getByRole("menu", { name: "Arkaya Gönder" });
    await expect(menu.getByRole("menuitem")).toHaveText(["Arkaya Gönder", "En Arkaya Gönder", "Metnin Arkasına Gönder"]);
    await expect(menu.locator("[role=menuitem] svg.oh-arrange-icon")).toHaveCount(3);
    await menu.getByRole("menuitem", { name: "Metnin Arkasına Gönder" }).click();
    await expect(body(page).locator("img[data-wrap=behind]")).toHaveCount(1);

    await ribbon(page).getByRole("button", { name: "Öne Getir seçenekleri" }).click();
    menu = page.getByRole("menu", { name: "Öne Getir" });
    await expect(menu.getByRole("menuitem")).toHaveText(["Öne Getir", "En Öne Getir", "Metnin Önüne Getir"]);
    await expect(menu.locator("[role=menuitem] svg.oh-arrange-icon")).toHaveCount(3);
    await menu.getByRole("menuitem", { name: "Metnin Önüne Getir" }).click();
    await expect(body(page).locator("img[data-wrap=front]")).toHaveCount(1);
    await ribbon(page).getByRole("button", { name: "Öne Getir seçenekleri" }).click();
    await page.getByRole("menuitem", { name: "En Öne Getir", exact: true }).click();
    await expect(body(page).locator("img[data-wrap=front][data-z='1']")).toHaveCount(1);

    // Şerit düğmeleri de kendi çizimimizle.
    for (const name of ["Konum", "Metni Kaydır"]) await expect(button(page, name).locator("svg.oh-arrange-icon")).toHaveCount(1);
    await button(page, "Hizala").click();
    menu = page.getByRole("menu", { name: "Hizala" });
    await expect(menu.getByRole("menuitem")).toHaveText(["Sola Hizala", "Ortala", "Sağa Hizala"]);
    await expect(menu.locator("[role=menuitem] svg.oh-arrange-icon")).toHaveCount(3);
  });
});

test.describe("Sunum › katman ve hizalama çizimleri", () => {
  test("sağ tık › Arkaya Gönder: çizimli alt seçenekler, şekil en arkaya geçer", async ({ page }) => {
    const errors = collectErrors(page);
    await page.goto(URL.sunum);
    const shapes = page.locator(".stage .sv [data-shape-id]");
    await expect(shapes.first()).toBeVisible();
    await page.getByRole("tab", { name: "Ekle", exact: true }).click();
    await ribbon(page).getByRole("button", { name: "Şekiller" }).click();
    await page.locator(".oh-popover [title='Dikdörtgen']").first().click();
    const st = (await page.locator(".stage").boundingBox())!;
    await page.mouse.move(st.x + 60, st.y + 60);
    await page.mouse.down();
    await page.mouse.move(st.x + 220, st.y + 160, { steps: 6 });
    await page.mouse.up();
    await expect(shapes).toHaveCount(3);
    const id = await shapes.nth(2).getAttribute("data-shape-id");
    await page.mouse.click(st.x + 140, st.y + 110, { button: "right" });
    // Alt menülü öğenin adı "›" işaretini de içerir.
    await page.getByRole("menuitem", { name: /^Arkaya Gönder/ }).hover();
    const sub = page.getByRole("menuitem", { name: "En Arkaya Gönder", exact: true });
    await expect(sub).toBeVisible();
    await expect(sub.locator("svg.oh-arrange-icon")).toHaveCount(1);
    await sub.click();
    await expect(shapes.first()).toHaveAttribute("data-shape-id", id!);
    expect(errors).toEqual([]);
  });
});
