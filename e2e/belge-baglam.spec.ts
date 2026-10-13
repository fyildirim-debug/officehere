import { expect, test, type Locator, type Page } from "@playwright/test";
import { URL, collectErrors, interceptFileChooser } from "./ortak";

// Belge: sağ tıklanan nesneye özgü bağlam menüsü (resim, grafik, tablo, denklem, İçindekiler, metin) ve
// ortalı tablonun sağ kenarından boyutlandırılınca sol kenarının yerinde kalması.

const PNG = Buffer.from("iVBORw0KGgoAAAANSUhEUgAAABAAAAAMCAIAAADkharWAAAAGklEQVR4nGO4HyNEEmIY1TCqYVTDqAYqagAAcKKm8RLe1JYAAAAASUVORK5CYII=", "base64");
const PNG2 = Buffer.from("iVBORw0KGgoAAAANSUhEUgAAAAQAAAADCAIAAAA7ljmRAAAAEElEQVR4nGM4IScHRww4OQD1xwwxB+utmQAAAABJRU5ErkJggg==", "base64");

const body = (page: Page) => page.locator(".bl-body .ProseMirror");
const ribbon = (page: Page) => page.locator(".oh-ribbon-body");
const tab = (page: Page, name: string) => page.getByRole("tab", { name, exact: true }).click();
const button = (page: Page, name: string) => ribbon(page).getByRole("button", { name, exact: true }).first();
const dialog = (page: Page) => page.locator(".oh-dialog");
const ctxMenu = (page: Page, name: string) => page.getByRole("menu", { name, exact: true });
const esc = (s: string) => s.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
/** Menü öğesi (ad kısayol ve › işaretiyle sürdüğünden başından eşlenir). */
const item = (menu: Locator, name: string, role: "menuitem" | "menuitemcheckbox" = "menuitem") => menu.getByRole(role, { name: new RegExp(`^${esc(name)}`) });
const miniBar = (page: Page) => page.locator(".bl-minibar.oh-minibar");

async function insertImage(page: Page) {
  await page.setInputFiles('input[type=file][accept="image/*"]', { name: "mavi.png", mimeType: "image/png", buffer: PNG });
  const img = body(page).locator('img[src^="data:image/png"]');
  await expect(img).toBeVisible();
  return img;
}

async function insertTable(page: Page, rows: number, cols: number) {
  await tab(page, "Ekle");
  await ribbon(page).locator("[data-tip='Tablo']").first().click();
  const cell = page.getByRole("button", { name: `${cols}x${rows} tablo` });
  await cell.hover();
  await cell.click();
  await expect(body(page).locator("table")).toHaveCount(1);
}

test.describe("OfficeHere Belge — nesneye özgü bağlam menüsü", () => {
  let errors: string[] = [];

  test.beforeEach(async ({ page }) => {
    errors = collectErrors(page);
    interceptFileChooser(page);
    await page.goto(URL.belge);
    await expect(body(page)).toBeFocused();
  });

  test.afterEach(() => {
    expect(errors).toEqual([]);
  });

  test("resim: sağ tık resmi seçer, resim menüsü açılır; Metni Kaydır › Kare uygulanır; Shift+F10 aynı menü", async ({ page }) => {
    for (let i = 0; i < 3; i++) {
      await page.keyboard.type(`Paragraf ${i + 1}. ${"Resim çevresinde akan metin. ".repeat(6)}`);
      await page.keyboard.press("Enter");
    }
    const img = await insertImage(page);
    // Seçimi bırak: metne tıkla.
    await body(page).locator("p").first().click();
    await expect(body(page).locator(".ProseMirror-selectednode")).toHaveCount(0);
    await img.click({ button: "right" });
    const menu = ctxMenu(page, "Resim");
    await expect(menu).toBeVisible();
    // Word'deki gibi sağ tık resmi seçer (tutamaçlar görünür).
    await expect(body(page).locator('.ProseMirror-selectednode img[src^="data:image/png"]')).toHaveCount(1);
    for (const name of ["Kes", "Kopyala", "Resmi Değiştir...", "Resim Olarak Kaydet...", "Metni Kaydır", "Öne Getir", "Arkaya Gönder", "Boyut ve Konum...", "Alternatif Metni Görüntüle...", "Sil"]) {
      await expect(item(menu, name)).toBeVisible();
    }
    await expect(item(menu, "Yazı Tipi...")).toHaveCount(0);
    await expect(item(menu, "Paragraf...")).toHaveCount(0);
    // Nesne menüsünde mini araç çubuğu yok.
    await expect(miniBar(page)).toHaveCount(0);
    await item(menu, "Metni Kaydır").hover();
    const wrap = ctxMenu(page, "Metni Kaydır");
    await expect(item(wrap, "Metinle Aynı Hizada", "menuitemcheckbox")).toHaveAttribute("aria-checked", "true");
    await expect(wrap.locator("svg.oh-arrange-icon")).toHaveCount(6);
    await item(wrap, "Kare", "menuitemcheckbox").click();
    await expect(menu).toHaveCount(0);
    await expect(body(page).locator("img[data-wrap=square]")).toHaveCount(1);
    // Resim seçili kalır; Shift+F10 resim menüsünü açar ve seçili kaydırma işaretlidir.
    await expect(body(page)).toBeFocused();
    await page.keyboard.press("Shift+F10");
    await expect(menu).toBeVisible();
    await item(menu, "Metni Kaydır").hover();
    await expect(item(ctxMenu(page, "Metni Kaydır"), "Kare", "menuitemcheckbox")).toHaveAttribute("aria-checked", "true");
    await page.keyboard.press("Escape");
    await page.keyboard.press("Escape");
    await expect(menu).toHaveCount(0);
    // Metne sağ tık resmi bırakır ve metin menüsünü açar.
    await body(page).locator("p").last().click({ button: "right", position: { x: 30, y: 8 } });
    const text = ctxMenu(page, "Metin");
    await expect(text).toBeVisible();
    await expect(item(text, "Yazı Tipi...")).toBeVisible();
    await expect(body(page).locator(".ProseMirror-selectednode")).toHaveCount(0);
  });

  test("resim: Resmi Değiştir seçilen dosyayı yerleştirir, Alternatif Metin iletişim kutusu açılır", async ({ page }) => {
    const img = await insertImage(page);
    const before = await img.getAttribute("src");
    await img.click({ button: "right" });
    const menu = ctxMenu(page, "Resim");
    const chooser = page.waitForEvent("filechooser");
    await item(menu, "Resmi Değiştir...").click();
    await (await chooser).setFiles({ name: "kirmizi.png", mimeType: "image/png", buffer: PNG2 });
    // Resim yerinde değişir (yeni resim eklenmez), seçili kalır.
    await expect.poll(() => img.getAttribute("src")).not.toBe(before);
    await expect(body(page).locator('img[src^="data:image/png"]')).toHaveCount(1);
    await expect(body(page).locator('.ProseMirror-selectednode img[src^="data:image/png"]')).toHaveAttribute("alt", "kirmizi");
    await img.click({ button: "right" });
    await item(menu, "Alternatif Metni Görüntüle...").click();
    await expect(dialog(page)).toContainText("Alternatif Metin");
    await dialog(page).getByRole("button", { name: "İptal" }).click();
    // Resim Yazısı Ekle: resmin altına "Şekil 1"; numara alanında metin menüsü "Alanı Güncelleştir" içerir.
    await img.click({ button: "right" });
    await item(menu, "Resim Yazısı Ekle...").click();
    await dialog(page).getByLabel("Resim yazısı", { exact: true }).fill("Deneme");
    await dialog(page).getByRole("button", { name: "Tamam" }).click();
    const seq = body(page).locator(".bl-seq");
    await expect(seq).toHaveText("1");
    await expect(seq.locator("xpath=..")).toContainText("Şekil 1 Deneme");
    await seq.click({ button: "right" });
    const text = ctxMenu(page, "Metin");
    await item(text, "Alanı Güncelleştir").click();
    await expect(seq).toHaveText("1");
  });

  test("grafik: Verileri Düzenle ve Grafik Türünü Değiştir", async ({ page }) => {
    await page.keyboard.type("Grafik öncesi");
    await page.keyboard.press("Enter");
    await tab(page, "Ekle");
    await button(page, "Grafik").click();
    await dialog(page).getByRole("button", { name: "Tamam" }).click();
    const chart = body(page).locator(".bl-chart");
    await expect(chart).toBeVisible();
    await body(page).locator("p").first().click();
    await expect(chart).not.toHaveClass(/is-selected/);
    await chart.click({ button: "right", position: { x: 120, y: 90 } });
    const menu = ctxMenu(page, "Grafik");
    await expect(menu).toBeVisible();
    await expect(chart).toHaveClass(/is-selected/);
    await expect(item(menu, "Verileri Düzenle...")).toBeVisible();
    await expect(item(menu, "Yazı Tipi...")).toHaveCount(0);
    await expect(miniBar(page)).toHaveCount(0);
    await item(menu, "Grafik Türünü Değiştir").hover();
    await item(ctxMenu(page, "Grafik Türünü Değiştir"), "Pasta", "menuitemcheckbox").click();
    await expect(chart.locator("svg path").first()).toBeVisible();
    await chart.click({ button: "right", position: { x: 120, y: 90 } });
    await item(menu, "Verileri Düzenle...").click();
    await expect(dialog(page)).toContainText("Verileri Düzenle");
    await dialog(page).getByRole("button", { name: "İptal" }).click();
  });

  test("tablo hücresi: tablo menüsü (Yazı Tipi yok); Ekle › Alta Satır Ekle satır ekler", async ({ page }) => {
    await insertTable(page, 2, 3);
    const cell = body(page).locator("table td").first();
    await cell.click();
    await page.keyboard.type("hücre");
    await cell.click({ button: "right" });
    const menu = ctxMenu(page, "Tablo");
    await expect(menu).toBeVisible();
    for (const name of ["Ekle", "Sil", "Hücreleri Böl...", "Hücre Hizalaması", "Otomatik Sığdır", "Tablo Özellikleri..."]) await expect(item(menu, name)).toBeVisible();
    await expect(item(menu, "Yazı Tipi...")).toHaveCount(0);
    await expect(item(menu, "Madde İşaretleri", "menuitemcheckbox")).toHaveCount(0);
    await item(menu, "Ekle").hover();
    await item(ctxMenu(page, "Ekle"), "Alta Satır Ekle").click();
    await expect(body(page).locator("table tr")).toHaveCount(3);
    // Hücre hizalaması: 3×3 galeri, Ortala hücreyi dikeyde ortalar.
    await cell.click({ button: "right" });
    await item(menu, "Hücre Hizalaması").hover();
    const gallery = ctxMenu(page, "Hücre Hizalaması");
    await expect(gallery.getByRole("menuitemradio")).toHaveCount(9);
    await gallery.getByRole("menuitemradio", { name: "Ortala", exact: true }).click();
    await expect(cell).toHaveCSS("vertical-align", "middle");
  });

  test("denklem ve İçindekiler menüleri", async ({ page }) => {
    await page.keyboard.type("Bölüm Bir");
    await page.keyboard.press("Control+Alt+1");
    await page.keyboard.press("End");
    await page.keyboard.press("Enter");
    await page.keyboard.press("Alt+=");
    await dialog(page).getByLabel("Denklem").fill("a^2+b^2=c^2");
    await dialog(page).getByRole("button", { name: "Ekle" }).click();
    const eq = body(page).locator(".bl-eq");
    await expect(eq).toBeVisible();
    await page.keyboard.press("End");
    await page.keyboard.press("Enter");
    await body(page).locator("h1").click();
    await eq.click({ button: "right" });
    const em = ctxMenu(page, "Denklem");
    await expect(item(em, "Denklemi Düzenle...")).toBeVisible();
    await expect(item(em, "Yazı Tipi...")).toHaveCount(0);
    await expect(eq).toHaveClass(/is-selected/);
    await item(em, "Denklemi Düzenle...").click();
    await expect(dialog(page).getByLabel("Denklem")).toHaveValue("a^2+b^2=c^2");
    await dialog(page).getByRole("button", { name: "İptal" }).click();

    // İçindekiler belgenin sonuna eklenir; yeni başlık eklenince menüden güncelleştirilir.
    await page.keyboard.press("Control+End");
    await tab(page, "Başvurular");
    await button(page, "İçindekiler").click();
    await page.getByRole("menuitem", { name: /^Otomatik Tablo 1/ }).click();
    const toc = body(page).locator(".bl-toc");
    await expect(toc.locator(".bl-toc-entry")).toHaveCount(1);
    await body(page).locator("h1").click();
    await page.keyboard.press("End");
    await page.keyboard.press("Enter");
    await page.keyboard.type("Bölüm İki");
    await page.keyboard.press("Control+Alt+1");
    await expect(body(page).locator("h1")).toHaveCount(2);
    await toc.click({ button: "right" });
    const tm = ctxMenu(page, "İçindekiler");
    await expect(item(tm, "İçindekileri Kaldır")).toBeVisible();
    await expect(item(tm, "Yazı Tipi...")).toHaveCount(0);
    await item(tm, "İçindekileri Güncelleştir").click();
    await expect(toc.locator(".bl-toc-entry")).toHaveCount(2);
  });

  test("metin: Yazı Tipi menüsü ve mini araç çubuğu", async ({ page }) => {
    await page.keyboard.type("Sıradan bir metin satırı");
    await body(page).locator("p").first().click({ button: "right", position: { x: 40, y: 8 } });
    const menu = ctxMenu(page, "Metin");
    await expect(item(menu, "Yazı Tipi...")).toBeVisible();
    await expect(item(menu, "Paragraf...")).toBeVisible();
    await expect(item(menu, "Tablo Özellikleri...")).toHaveCount(0);
    await expect(miniBar(page)).toBeVisible();
  });
});

test.describe("OfficeHere Belge — ortalı tablo sağdan boyutlandırılınca sol kenar yerinde kalır", () => {
  let errors: string[] = [];

  test.beforeEach(async ({ page }) => {
    errors = collectErrors(page);
    await page.goto(URL.belge);
    await expect(body(page)).toBeFocused();
  });

  test.afterEach(() => {
    expect(errors).toEqual([]);
  });

  const box = async (page: Page) => (await body(page).locator("table").boundingBox())!;

  test("son sütunun sağ kenarı: sol x sabit, genişlik azalır; Geri Al eski genişliğe ve ortalamaya döner", async ({ page }) => {
    await insertTable(page, 3, 3);
    // Tabloyu daralt (sola yaslıyken), sonra Tablo Özellikleri › Orta.
    await body(page).locator("table td").first().hover();
    const h = (await page.locator(".bl-tbl-resize").boundingBox())!;
    await page.mouse.move(h.x + 3, h.y + 3);
    await page.mouse.down();
    await page.mouse.move(h.x - 120, h.y + 3, { steps: 6 });
    await page.mouse.up();
    await body(page).locator("table td").first().click({ button: "right" });
    await item(ctxMenu(page, "Tablo"), "Tablo Özellikleri...").click();
    await dialog(page).getByRole("button", { name: "Orta", exact: true }).click();
    await dialog(page).getByRole("button", { name: "Tamam" }).click();
    await expect(body(page).locator("table")).toHaveAttribute("data-align", "center");
    const b0 = await box(page);
    const last = (await body(page).locator("table tr:first-child > *").last().boundingBox())!;
    const x = last.x + last.width - 1;
    const y = last.y + last.height / 2;
    await page.mouse.move(x - 6, y);
    await page.mouse.move(x, y, { steps: 3 });
    await expect(page.locator(".bl-doc.resize-cursor, .bl-doc .resize-cursor").first()).toBeVisible();
    await page.mouse.down();
    for (let i = 1; i <= 6; i++) {
      await page.mouse.move(x - i * 20, y);
      // Sürüklerken sol kenar kıpırdamaz.
      expect(Math.abs((await box(page)).x - b0.x)).toBeLessThanOrEqual(2);
    }
    await page.mouse.up();
    const b1 = await box(page);
    expect(Math.abs(b1.x - b0.x)).toBeLessThanOrEqual(2);
    expect(Math.abs(b1.width - (b0.width - 120))).toBeLessThanOrEqual(4);
    // Bırakınca sola yaslı + girinti: tablo bırakıldığı yerde kalır.
    await expect(body(page).locator("table")).not.toHaveAttribute("data-align", /.+/);
    // Tek Geri Al adımı genişliği ve ortalamayı birlikte geri alır.
    await page.keyboard.press("Control+z");
    await expect(body(page).locator("table")).toHaveAttribute("data-align", "center");
    await expect.poll(async () => Math.round((await box(page)).width)).toBeGreaterThan(b0.width - 4);
    expect(Math.abs((await box(page)).x - b0.x)).toBeLessThanOrEqual(2);
  });

  test("köşe tutamacı: HTML'den gelen ortalı tablo sağdan küçülür, sol kenar yerinde kalır", async ({ page }) => {
    const row = (a: string, b: string) => `<tr><td colwidth="160"><p>${a}</p></td><td colwidth="160"><p>${b}</p></td></tr>`;
    const html = `<!doctype html><html><body><p>Ortalı tablo:</p><table data-align="center"><tbody>${row("A1", "B1")}${row("A2", "B2")}</tbody></table><p>Son</p></body></html>`;
    await page.setInputFiles("#bl-open-input", { name: "ortali.html", mimeType: "text/html", buffer: Buffer.from(html) });
    const table = body(page).locator("table");
    await expect(table).toHaveAttribute("data-align", "center");
    const b0 = await box(page);
    await table.locator("td").first().hover();
    const handle = page.locator(".bl-tbl-resize");
    await expect(handle).toBeVisible();
    const h = (await handle.boundingBox())!;
    await page.mouse.move(h.x + 3, h.y + 3);
    await page.mouse.down();
    for (let i = 1; i <= 6; i++) {
      await page.mouse.move(h.x + 3 - i * 20, h.y + 3);
      expect(Math.abs((await box(page)).x - b0.x)).toBeLessThanOrEqual(2);
    }
    await page.mouse.up();
    const b1 = await box(page);
    expect(Math.abs(b1.x - b0.x)).toBeLessThanOrEqual(2);
    expect(Math.abs(b1.width - (b0.width - 120))).toBeLessThanOrEqual(4);
    await expect(table).toHaveAttribute("data-indent", /^\d+$/);
    await page.keyboard.press("Control+z");
    await expect(table).toHaveAttribute("data-align", "center");
    await expect.poll(async () => Math.round((await box(page)).width)).toBeGreaterThan(b0.width - 4);
    expect(Math.abs((await box(page)).x - b0.x)).toBeLessThanOrEqual(2);
  });
});
