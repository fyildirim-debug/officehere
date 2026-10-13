import { readFileSync } from "node:fs";
import { resolve } from "node:path";
import { expect, test, type Page } from "@playwright/test";
import { URL, collectErrors, interceptFileChooser } from "./ortak";

// OfficeHere Sunum: sağ tık menüsü tıklanan nesnenin türüne göre değişir
// (resim, grafik, tablo çerçevesi, metin kutusu, şekil, çizgi, grup, çoklu seçim).

const PNG = readFileSync(resolve(__dirname, "../apps/desktop/build/icons/128x128.png"));
const shapes = (page: Page) => page.locator(".stage .sv [data-shape-id]");
const ribbon = (page: Page) => page.locator(".oh-ribbon-body");
const tab = (page: Page, name: string) => page.getByRole("tab", { name, exact: true }).click();
/** Açık menüdeki öğe (onay işaretli öğeler dahil), etiketine göre. */
const item = (page: Page, name: string | RegExp) =>
  page.locator(".oh-popover [role^=menuitem]").filter({ has: page.locator(".oh-menu-label", { hasText: typeof name === "string" ? new RegExp(`^${name.replace(/[.*+?^${}()|[\]\\]/g, "\\$&")}$`) : name }) });
const menuLabels = (page: Page) => page.locator(".oh-popover [role^=menuitem] .oh-menu-label").allTextContents();

async function rightClick(page: Page, el: ReturnType<Page["locator"]>, fx = 0.5, fy = 0.5) {
  const b = (await el.boundingBox())!;
  await page.mouse.click(b.x + b.width * fx, b.y + b.height * fy, { button: "right" });
  await expect(page.locator(".oh-popover [role=menu]").first()).toBeVisible();
}

async function insertPicture(page: Page) {
  await tab(page, "Ekle");
  await ribbon(page).getByRole("button", { name: "Resimler" }).click();
  const [chooser] = await Promise.all([page.waitForEvent("filechooser"), page.getByRole("menuitem", { name: "Bu Cihaz..." }).click()]);
  await chooser.setFiles({ name: "simge.png", mimeType: "image/png", buffer: PNG });
  const pic = page.locator(".stage .sv .sh-pic").last();
  await expect(pic).toBeVisible();
  return pic;
}

async function drawShape(page: Page, title: string, x1: number, y1: number, x2: number, y2: number) {
  await tab(page, "Ekle");
  await ribbon(page).getByRole("button", { name: "Şekiller" }).click();
  await page.locator(`.oh-popover [title='${title}']`).first().click();
  const st = (await page.locator(".stage").boundingBox())!;
  await page.mouse.move(st.x + x1, st.y + y1);
  await page.mouse.down();
  await page.mouse.move(st.x + x2, st.y + y2, { steps: 6 });
  await page.mouse.up();
}

test.describe("Sunum › nesneye özgü sağ tık menüleri", () => {
  let errors: string[] = [];

  test.beforeEach(async ({ page }) => {
    errors = collectErrors(page);
    interceptFileChooser(page);
    await page.addInitScript(() => {
      delete (window as unknown as Record<string, unknown>).showSaveFilePicker;
    });
    await page.goto(URL.sunum);
    await expect(page.locator(".stage .sv")).toBeVisible();
    // Yer tutucusuz boş slayt: nesneler birbirine karışmasın.
    await page.locator(".stage").click({ position: { x: 4, y: 4 } });
    await ribbon(page).getByRole("button", { name: "Yeni Slayt" }).locator("xpath=..").locator(".oh-split-arrow").click();
    await page.locator(".layout-tile", { hasText: "Boş" }).click();
    await expect(shapes(page)).toHaveCount(0);
  });

  test.afterEach(() => {
    expect(errors).toEqual([]);
  });

  test("resim: Resmi Değiştir, Kırp, Resim Olarak Kaydet ve Alternatif Metin; grafik komutları yok", async ({ page }) => {
    const pic = await insertPicture(page);
    await rightClick(page, pic);
    const labels = await menuLabels(page);
    for (const l of ["Kes", "Kopyala", "Yapıştır", "Resmi Değiştir", "Kırp", "Grupla", "Öne Getir", "Arkaya Gönder", "Bağlantı...", "Resim Olarak Kaydet...", "Alternatif Metni Görüntüle...", "Boyut ve Konum...", "Resmi Biçimlendir...", "Sil"]) expect(labels).toContain(l);
    expect(labels).not.toContain("Verileri Düzenle...");
    expect(labels).not.toContain("Metni Düzenle");
    await item(page, /^Resmi Değiştir/).hover();
    await expect(item(page, "Dosyadan...")).toBeVisible();
    // Resim Olarak Kaydet: resmin özgün verisi indirilir.
    const [download] = await Promise.all([page.waitForEvent("download"), item(page, "Resim Olarak Kaydet...").click()]);
    expect(download.suggestedFilename()).toMatch(/\.png$/);
    expect(readFileSync(await download.path()).equals(PNG)).toBe(true);
    // Kırp: kırpma kipi açılır.
    await rightClick(page, pic);
    await item(page, "Kırp").click();
    await expect(page.locator(".canvas-hint")).toContainText("Kırpmak için");
    await page.keyboard.press("Escape");
    // Alternatif metin resme yazılır.
    await rightClick(page, pic);
    await item(page, "Alternatif Metni Görüntüle...").click();
    await page.getByRole("dialog").getByRole("textbox", { name: "Alternatif metin" }).fill("Uygulama simgesi");
    await page.getByRole("dialog").getByRole("button", { name: "Tamam" }).click();
    await expect(page.locator(".stage .sv img[alt='Uygulama simgesi']")).toHaveCount(1);
  });

  test("resim: Resmi Değiştir › Dosyadan yeni resmi aynı kutuya koyar", async ({ page }) => {
    const pic = await insertPicture(page);
    const before = await pic.locator("img").first().getAttribute("src");
    await rightClick(page, pic);
    await item(page, /^Resmi Değiştir/).hover();
    const svg = Buffer.from('<svg xmlns="http://www.w3.org/2000/svg" width="40" height="40"><rect width="40" height="40" fill="#c33"/></svg>');
    const [chooser] = await Promise.all([page.waitForEvent("filechooser"), item(page, "Dosyadan...").click()]);
    await chooser.setFiles({ name: "kare.svg", mimeType: "image/svg+xml", buffer: svg });
    await expect.poll(() => pic.locator("img").first().getAttribute("src")).not.toBe(before);
    await expect(shapes(page)).toHaveCount(1);
  });

  test("grafik: Verileri Düzenle ve Grafik Türünü Değiştir var, resim komutları yok", async ({ page }) => {
    await tab(page, "Ekle");
    await ribbon(page).getByRole("button", { name: "Grafik" }).click();
    await page.getByRole("button", { name: "Tamam" }).click();
    await page.getByRole("button", { name: "Tamam" }).click();
    const chart = page.locator(".stage .sh-chart").first();
    await expect(chart).toBeVisible();
    await rightClick(page, chart);
    const labels = await menuLabels(page);
    for (const l of ["Kes", "Kopyala", "Yapıştır", "Verileri Düzenle...", "Grafik Türünü Değiştir...", "Öne Getir", "Arkaya Gönder", "Boyut ve Konum...", "Grafik Alanını Biçimlendir...", "Sil"]) expect(labels).toContain(l);
    for (const l of ["Kırp", "Resmi Değiştir", "Metni Düzenle", "Resim Olarak Kaydet..."]) expect(labels).not.toContain(l);
    await item(page, "Verileri Düzenle...").click();
    await expect(page.getByRole("dialog")).toBeVisible();
  });

  test("tablo çerçevesi: Ekle › Satırları Alta Ekle ve Tabloyu Sil", async ({ page }) => {
    await tab(page, "Ekle");
    await ribbon(page).getByRole("button", { name: "Tablo" }).click();
    await page.locator(".oh-popover .oh-tablegrid-cells > *").nth(10 + 2).click();
    const tds = page.locator(".stage td");
    await expect(tds).toHaveCount(6);
    // Çerçeve: seçim kenarındaki taşıma şeridi.
    await rightClick(page, page.locator(".sel-move.is-t").first());
    const labels = await menuLabels(page);
    for (const l of ["Ekle", "Tabloyu Sil", "Öne Getir", "Arkaya Gönder", "Boyut ve Konum...", "Sil"]) expect(labels).toContain(l);
    expect(labels).not.toContain("Kırp");
    await item(page, /^Ekle/).hover();
    await item(page, "Satırları Alta Ekle").click();
    await expect(tds).toHaveCount(9);
    await rightClick(page, page.locator(".sel-move.is-t").first());
    await item(page, "Tabloyu Sil").click();
    await expect(tds).toHaveCount(0);
    await page.keyboard.press("Control+z");
    await expect(tds).toHaveCount(9);
  });

  test("tablo hücresi: Hücre Hizalaması › Alta Hizala hücreye uygulanır", async ({ page }) => {
    await tab(page, "Ekle");
    await ribbon(page).getByRole("button", { name: "Tablo" }).click();
    await page.locator(".oh-popover .oh-tablegrid-cells > *").nth(10 + 2).click();
    const td = page.locator(".stage td").first();
    await td.click();
    await td.click({ button: "right" });
    await item(page, /^Hücre Hizalaması/).hover();
    await item(page, "Alta Hizala").click();
    await expect(td).toHaveCSS("vertical-align", "bottom");
  });

  test("şekil ve çizgi: şekilde Metni Düzenle/Noktaları Düzenle, çizgide yok", async ({ page }) => {
    await drawShape(page, "Dikdörtgen", 60, 60, 220, 160);
    await drawShape(page, "Çizgi", 400, 300, 600, 420);
    await expect(shapes(page)).toHaveCount(2);
    await rightClick(page, shapes(page).first());
    let labels = await menuLabels(page);
    for (const l of ["Metni Düzenle", "Noktaları Düzenle", "Grupla", "Bağlantı...", "Şekli Biçimlendir..."]) expect(labels).toContain(l);
    await page.keyboard.press("Escape");
    await rightClick(page, shapes(page).nth(1));
    labels = await menuLabels(page);
    for (const l of ["Öne Getir", "Arkaya Gönder", "Boyut ve Konum...", "Şekli Biçimlendir...", "Sil"]) expect(labels).toContain(l);
    for (const l of ["Metni Düzenle", "Noktaları Düzenle", "Kırp"]) expect(labels).not.toContain(l);
  });

  test("metin kutusu (düzenlenmiyorken): Metni Düzenle, Yazı Tipi ve Paragraf", async ({ page }) => {
    await tab(page, "Ekle");
    await ribbon(page).getByRole("button", { name: "Metin Kutusu" }).click();
    await page.getByRole("menuitem", { name: "Yatay Metin Kutusu Çiz" }).click();
    const st = (await page.locator(".stage").boundingBox())!;
    await page.mouse.click(st.x + 200, st.y + 150);
    await page.keyboard.type("Not");
    await page.keyboard.press("Escape");
    await rightClick(page, shapes(page).first());
    const labels = await menuLabels(page);
    for (const l of ["Metni Düzenle", "Yazı Tipi...", "Paragraf...", "Bağlantı...", "Boyut ve Konum...", "Şekli Biçimlendir..."]) expect(labels).toContain(l);
    expect(labels).not.toContain("Noktaları Düzenle");
    await item(page, "Yazı Tipi...").click();
    await expect(page.getByRole("dialog")).toContainText("Yazı Tipi");
  });

  test("iki nesne seçili: Grupla ve Hizala; grupta Grubu Çöz", async ({ page }) => {
    await drawShape(page, "Dikdörtgen", 60, 60, 220, 160);
    await drawShape(page, "Oval", 300, 220, 420, 320);
    await shapes(page).nth(0).click();
    await shapes(page).nth(1).click({ modifiers: ["Shift"] });
    await rightClick(page, shapes(page).nth(1));
    const labels = await menuLabels(page);
    for (const l of ["Kes", "Kopyala", "Yapıştır", "Grupla", "Hizala", "Öne Getir", "Arkaya Gönder", "Sil"]) expect(labels).toContain(l);
    expect(labels).not.toContain("Metni Düzenle");
    await item(page, /^Hizala/).hover();
    await expect(item(page, "Sola Hizala").locator("svg.oh-arrange-icon")).toHaveCount(1);
    await item(page, "Sola Hizala").click();
    await expect.poll(async () => Math.abs((await shapes(page).nth(0).boundingBox())!.x - (await shapes(page).nth(1).boundingBox())!.x)).toBeLessThan(1);
    await rightClick(page, shapes(page).nth(1));
    await item(page, /^Grupla/).click();
    await expect(page.locator(".stage .sv [data-shape-id] [data-shape-id]")).toHaveCount(2);
    // Grubun bir parçasına sağ tık: grup menüsü.
    await rightClick(page, shapes(page).nth(1));
    expect(await menuLabels(page)).not.toContain("Hizala");
    await expect(item(page, /^Grubu Çöz/)).toBeVisible();
    await item(page, /^Grubu Çöz/).click();
    await expect(page.locator(".stage .sv [data-shape-id] [data-shape-id]")).toHaveCount(0);
  });

  test("Öne Getir › En Öne Getir: arkadaki şekil en öne geçer", async ({ page }) => {
    await drawShape(page, "Dikdörtgen", 60, 60, 260, 200);
    await drawShape(page, "Dikdörtgen", 160, 120, 360, 260);
    const back = await shapes(page).nth(0).getAttribute("data-shape-id");
    // İki şeklin örtüşmeyen sol üst köşesine sağ tık: alttaki şekil seçilir.
    await rightClick(page, shapes(page).nth(0), 0.15, 0.15);
    await item(page, /^Öne Getir/).hover();
    const front = item(page, "En Öne Getir");
    await expect(front.locator("svg.oh-arrange-icon")).toHaveCount(1);
    await front.click();
    await expect(shapes(page).last()).toHaveAttribute("data-shape-id", back!);
  });

  test("Shift+F10 ve Menü tuşu seçili nesneye göre menü açar", async ({ page }) => {
    const pic = await insertPicture(page);
    await pic.click();
    await page.keyboard.press("Shift+F10");
    await expect(item(page, /^Resmi Değiştir/)).toBeVisible();
    await page.keyboard.press("Escape");
    await expect(page.locator(".oh-popover")).toHaveCount(0);
    await page.keyboard.press("ContextMenu");
    await expect(item(page, "Resim Olarak Kaydet...")).toBeVisible();
    await page.keyboard.press("Escape");
    // Seçim yokken slayt menüsü.
    await page.keyboard.press("Escape");
    await page.keyboard.press("Shift+F10");
    await expect(item(page, "Arka Planı Biçimlendir...")).toBeVisible();
  });
});
