import { readFileSync } from "node:fs";
import { expect, test, type Page } from "@playwright/test";
import { URL, collectErrors } from "./ortak";

// OfficeHere Sunum: gerçek bir kullanıcı gibi uçtan uca senaryolar.

test.describe("OfficeHere Sunum senaryoları", () => {
  let errors: string[] = [];

  test.beforeEach(async ({ page }) => {
    errors = collectErrors(page);
    // Tarayıcı dosya seçicileri yerine indirme / <input type=file> yolu kullanılsın.
    await page.addInitScript(() => {
      const w = window as unknown as Record<string, unknown>;
      delete w.showSaveFilePicker;
      delete w.showOpenFilePicker;
    });
    await page.goto(URL.sunum);
    await expect(page.locator(".stage .sv")).toBeVisible();
  });

  test.afterEach(() => {
    expect(errors).toEqual([]);
  });

  const shapes = (page: Page) => page.locator(".stage .sv [data-shape-id]");
  const thumbs = (page: Page) => page.locator(".thumb-pane .thumb-item");
  const ribbon = (page: Page) => page.locator(".oh-ribbon-body");
  const tab = (page: Page, name: string) => page.getByRole("tab", { name, exact: true }).click();
  const leaveText = async (page: Page) => {
    await page.keyboard.press("Escape");
    await page.keyboard.press("Escape");
  };
  const typeTitle = async (page: Page, text: string) => {
    await shapes(page).first().click();
    await page.keyboard.type(text);
    await leaveText(page);
  };

  test("başlık yazılır, küçük resim ve başlık çubuğu güncellenir", async ({ page }) => {
    await typeTitle(page, "Yıllık Rapor İğdır");
    await expect(shapes(page).first()).toContainText("Yıllık Rapor İğdır");
    await expect(thumbs(page).first()).toContainText("Yıllık Rapor İğdır");
    await expect(page.locator(".oh-titlebar")).toContainText("Kaydedilmedi");
  });

  test("yeni slayt: Ctrl+M ve düzen galerisinden Boş düzen", async ({ page }) => {
    await page.locator(".stage").click({ position: { x: 4, y: 4 } });
    await page.keyboard.press("Control+m");
    await expect(thumbs(page)).toHaveCount(2);
    await expect(shapes(page)).toHaveCount(2);
    await ribbon(page).getByRole("button", { name: "Yeni Slayt" }).locator("xpath=..").locator(".oh-split-arrow").click();
    await page.locator(".layout-tile", { hasText: "Boş" }).click();
    await expect(thumbs(page)).toHaveCount(3);
    await expect(shapes(page)).toHaveCount(0);
    await expect(page.locator(".oh-statusbar")).toContainText("Slayt 3 / 3");
  });

  test("madde işaretli metin ve Tab ile alt düzey", async ({ page }) => {
    await page.keyboard.press("Control+m");
    await shapes(page).nth(1).click();
    await page.keyboard.type("Birinci");
    await page.keyboard.press("Enter");
    await page.keyboard.press("Tab");
    await page.keyboard.type("Alt madde");
    await leaveText(page);
    const paras = shapes(page).nth(1).locator(".tb-p");
    await expect(paras).toHaveCount(2);
    const ml = await paras.evaluateAll((els) => els.map((e) => parseFloat(getComputedStyle(e).marginLeft)));
    expect(ml[1]).toBeGreaterThan(ml[0]);
  });

  test("şekil eklenir, sürüklenerek taşınır ve geri alınır", async ({ page }) => {
    await tab(page, "Ekle");
    await ribbon(page).getByRole("button", { name: "Şekiller" }).click();
    await page.locator(".oh-popover [title='Dikdörtgen']").first().click();
    const st = (await page.locator(".stage").boundingBox())!;
    await page.mouse.move(st.x + 60, st.y + 60);
    await page.mouse.down();
    await page.mouse.move(st.x + 220, st.y + 160, { steps: 6 });
    await page.mouse.up();
    await expect(shapes(page)).toHaveCount(3);
    const rect = shapes(page).nth(2);
    const before = (await rect.boundingBox())!;
    await page.mouse.move(before.x + before.width / 2, before.y + before.height / 2);
    await page.mouse.down();
    await page.mouse.move(before.x + before.width / 2 + 200, before.y + before.height / 2 + 100, { steps: 8 });
    await page.mouse.up();
    const after = (await rect.boundingBox())!;
    expect(after.x - before.x).toBeGreaterThan(150);
    expect(after.y - before.y).toBeGreaterThan(70);
    await page.keyboard.press("Control+z");
    await expect.poll(async () => (await rect.boundingBox())!.x).toBeCloseTo(before.x, 0);
    await page.keyboard.press("Control+y");
    await expect.poll(async () => (await rect.boundingBox())!.x).toBeCloseTo(after.x, 0);
  });

  test("tema değiştirilir, slayt arka planı değişir", async ({ page }) => {
    const bg = () => page.locator(".stage .sv-bg").evaluate((e) => getComputedStyle(e).backgroundColor + getComputedStyle(e).backgroundImage);
    const initial = await bg();
    await tab(page, "Tasarım");
    await ribbon(page).locator(".oh-gallery-item").nth(3).click();
    await expect.poll(bg).not.toBe(initial);
    await page.keyboard.press("Control+z");
    await expect.poll(bg).toBe(initial);
  });

  test("slayt bağlam menüsü: çoğalt, sil ve geri al", async ({ page }) => {
    await typeTitle(page, "Kopyalanacak");
    await thumbs(page).first().click({ button: "right" });
    await expect(page.getByRole("menuitem", { name: /Slaytı Çoğalt/ })).toBeVisible();
    await page.getByRole("menuitem", { name: /Slaytı Çoğalt/ }).click();
    await expect(thumbs(page)).toHaveCount(2);
    await expect(thumbs(page).nth(1)).toContainText("Kopyalanacak");
    await thumbs(page).nth(1).click({ button: "right" });
    await page.getByRole("menuitem", { name: /Slaytı Sil/ }).click();
    await expect(thumbs(page)).toHaveCount(1);
    await page.keyboard.press("Control+z");
    await expect(thumbs(page)).toHaveCount(2);
  });

  test("şekil sağ tık menüsü ve kopyala/yapıştır", async ({ page }) => {
    await typeTitle(page, "Menü");
    await shapes(page).first().click({ button: "right" });
    for (const item of ["Kes", "Kopyala", "Yapıştır", "Bağlantı...", "Boyut ve Konum...", "Sil"]) {
      await expect(page.getByRole("menuitem", { name: new RegExp(`^${item.replace(/\./g, "\\.")}`) })).toBeVisible();
    }
    await page.getByRole("menuitem", { name: /^Kopyala/ }).click();
    await page.keyboard.press("Control+v");
    await expect(shapes(page)).toHaveCount(3);
  });

  test("slayt gösterisi: aç, ilerle, sonu, Esc ile çık", async ({ page }) => {
    await typeTitle(page, "Birinci slayt");
    await page.keyboard.press("Control+m");
    await typeTitle(page, "İkinci slayt");
    await page.keyboard.press("F5");
    const show = page.locator(".show-root");
    await expect(show).toBeVisible();
    await expect(show).toContainText("Birinci slayt");
    await page.keyboard.press("ArrowRight");
    await expect(show).toContainText("İkinci slayt");
    await page.keyboard.press("ArrowRight");
    await expect(show.locator(".show-end")).toBeVisible();
    await page.keyboard.press("Escape");
    await expect(show).toHaveCount(0);
  });

  test("gösteride kalem: çizim sonunda Tut ile slayta eklenir", async ({ page }) => {
    await page.keyboard.press("F5");
    await expect(page.locator(".show-root")).toBeVisible();
    await page.keyboard.press("Control+p");
    const st = (await page.locator(".show-stage").boundingBox())!;
    await page.mouse.move(st.x + 200, st.y + 200);
    await page.mouse.down();
    for (let i = 1; i <= 10; i++) await page.mouse.move(st.x + 200 + i * 20, st.y + 200 + (i % 2) * 15);
    await page.mouse.up();
    await expect(page.locator(".show-ink path")).toHaveCount(1);
    await page.keyboard.press("Escape"); // kalemi bırak
    await page.keyboard.press("Escape"); // gösteriyi bitir
    await expect(page.getByRole("dialog")).toContainText("Mürekkep ek açıklamalarınızı tutmak istiyor musunuz?");
    await page.getByRole("button", { name: "Tut" }).click();
    await expect(shapes(page)).toHaveCount(3);
  });

  test("notlar yazılır ve slaytlar arasında korunur", async ({ page }) => {
    const notes = page.locator(".notes-pane textarea, .notes-pane [contenteditable]").first();
    await notes.click();
    await page.keyboard.type("Konuşmacı notu: ğüşıöç");
    await page.locator(".stage").click({ position: { x: 4, y: 4 } });
    await page.keyboard.press("Control+m");
    await expect(page.locator(".notes-pane")).not.toContainText("Konuşmacı notu");
    await thumbs(page).first().click();
    await expect(page.locator(".notes-pane textarea")).toHaveValue("Konuşmacı notu: ğüşıöç");
  });

  test("PPTX kaydet (indir) → aynı dosyayı aç: slaytlar, metin, notlar ve bölümler aynı", async ({ page }) => {
    await typeTitle(page, "Gidiş Dönüş");
    await page.keyboard.press("Control+m");
    await typeTitle(page, "İkinci Başlık");
    await page.locator(".notes-pane textarea").fill("Not metni");
    await page.locator(".stage").click({ position: { x: 4, y: 4 } });
    // Bölüm ekle
    await tab(page, "Giriş");
    await ribbon(page).getByRole("button", { name: "Bölüm" }).click();
    await page.getByRole("menuitem", { name: "Bölüm Ekle" }).click();
    await page.getByLabel("Bölüm adı").fill("Ayrıntılar");
    await page.getByRole("button", { name: "Yeniden Adlandır" }).click();
    await expect(page.locator(".thumb-pane .section-head")).toHaveText(["Varsayılan Bölüm", "Ayrıntılar"]);
    const [download] = await Promise.all([page.waitForEvent("download"), page.keyboard.press("Control+s")]);
    expect(download.suggestedFilename()).toBe("Sunu1.pptx");
    // Not: test başlığındaki Türkçe karakterler çıktı yoluna girer; dosya bellekten verilir.
    const bytes = readFileSync((await download.path())!);
    expect(bytes.subarray(0, 2).toString()).toBe("PK");
    await expect(page.locator(".oh-titlebar")).not.toContainText("Kaydedilmedi");
    // Yeni sunu, sonra kaydedilen dosyayı aç.
    await page.keyboard.press("Control+n");
    await expect(thumbs(page)).toHaveCount(1);
    const [chooser] = await Promise.all([page.waitForEvent("filechooser"), page.keyboard.press("Control+o")]);
    await chooser.setFiles({ name: "Sunu1.pptx", mimeType: "application/vnd.openxmlformats-officedocument.presentationml.presentation", buffer: bytes });
    await expect(page.locator(".oh-titlebar")).toContainText("Sunu1");
    await expect(thumbs(page)).toHaveCount(2);
    await expect(thumbs(page).first()).toContainText("Gidiş Dönüş");
    await expect(thumbs(page).nth(1)).toContainText("İkinci Başlık");
    await expect(page.locator(".thumb-pane .section-head")).toHaveText(["Varsayılan Bölüm", "Ayrıntılar"]);
    await thumbs(page).nth(1).click();
    await expect(page.locator(".notes-pane textarea")).toHaveValue("Not metni");
  });

  test("köprü: Ctrl+K ile metne bağlantı eklenir", async ({ page }) => {
    await shapes(page).first().click();
    await page.keyboard.type("Web sitemiz");
    await page.keyboard.press("Shift+Home");
    await page.keyboard.press("Control+k");
    await expect(page.getByRole("dialog")).toContainText("Köprü Ekle");
    await page.getByLabel("Adres").fill("ornek.com.tr");
    await page.getByRole("button", { name: "Tamam" }).click();
    await leaveText(page);
    await expect(page.locator(".stage span[data-link]")).toHaveAttribute("data-link", "https://ornek.com.tr");
  });

  test("Asıl Slayt Görünümü: açılır, asıl slayt düzenlenir, kapatılınca slayta yansır", async ({ page }) => {
    await typeTitle(page, "Asıl test");
    await tab(page, "Görünüm");
    await ribbon(page).getByRole("button", { name: "Asıl Slayt" }).click();
    await expect(page.getByRole("tab", { name: "Asıl Slayt", exact: true })).toBeVisible();
    await expect(page.locator(".oh-statusbar")).toContainText("Asıl Slayt");
    await page.locator(".master-item").first().click();
    // Asıl başlığı kalın yap.
    await shapes(page).first().click();
    await page.keyboard.press("Control+a");
    await page.keyboard.press("Control+b");
    await leaveText(page);
    await page.getByRole("tab", { name: "Asıl Slayt", exact: true }).click();
    await ribbon(page).getByRole("button", { name: "Asıl Görünümü Kapat" }).click();
    await expect(page.getByRole("tab", { name: "Asıl Slayt", exact: true })).toHaveCount(0);
    const weight = await shapes(page).first().locator("span[data-r]").first().evaluate((e) => getComputedStyle(e).fontWeight);
    expect(Number(weight)).toBeGreaterThanOrEqual(700);
  });

  test("grafik eklenir, verisi düzenlenir ve türü değiştirilir", async ({ page }) => {
    await page.keyboard.press("Control+m");
    await tab(page, "Ekle");
    await ribbon(page).getByRole("button", { name: "Grafik" }).click();
    await page.getByRole("option", { name: "Sütun" }).click();
    await page.getByRole("button", { name: "Tamam" }).click();
    await expect(page.getByRole("dialog")).toContainText("Verileri Düzenle");
    await page.getByLabel("B2", { exact: true }).fill("12,5");
    await page.getByRole("button", { name: "Tamam" }).click();
    const chart = page.locator(".stage .sh-chart");
    await expect(chart).toHaveCount(1);
    await expect(chart.locator("rect")).not.toHaveCount(0);
    await page.getByRole("tab", { name: "Grafik Tasarımı" }).click();
    await ribbon(page).getByRole("button", { name: "Verileri Düzenle" }).click();
    await expect(page.getByLabel("B2", { exact: true })).toHaveValue("12,5");
    await page.getByRole("button", { name: "İptal" }).click();
    await ribbon(page).getByRole("button", { name: "Grafik Türünü Değiştir" }).click();
    await page.getByRole("option", { name: "Pasta" }).click();
    await page.getByRole("button", { name: "Tamam" }).click();
    await expect(chart.locator("path").first()).toBeVisible();
  });

  test("dar pencerede (900 px) yatay kaydırma yok, şerit daralır", async ({ page }) => {
    await page.setViewportSize({ width: 900, height: 700 });
    await expect.poll(() => page.evaluate(() => document.documentElement.scrollWidth)).toBeLessThanOrEqual(900);
    await expect(page.locator(".stage .sv")).toBeVisible();
  });
});
