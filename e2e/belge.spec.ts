import { readFileSync } from "node:fs";
import { expect, test, type Page } from "@playwright/test";
import { URL, collectErrors } from "./ortak";

// OfficeHere Belge: gerçek bir kullanıcı gibi uçtan uca senaryolar (derlenmiş uygulama, yalnız arayüz).

test.describe("OfficeHere Belge senaryoları", () => {
  let errors: string[] = [];

  test.beforeEach(async ({ page }) => {
    errors = collectErrors(page);
    // Tarayıcı dosya seçicileri yerine indirme / <input type=file> yolu kullanılsın.
    await page.addInitScript(() => {
      const w = window as unknown as Record<string, unknown>;
      delete w.showSaveFilePicker;
      delete w.showOpenFilePicker;
    });
    await page.goto(URL.belge);
    await expect(body(page)).toBeVisible();
    // Belge açılışta odaklanır: hemen yazılabilir.
    await expect(body(page)).toBeFocused();
  });

  test.afterEach(() => {
    expect(errors).toEqual([]);
  });

  const body = (page: Page) => page.locator(".bl-body .ProseMirror");
  const ribbon = (page: Page) => page.locator(".oh-ribbon-body");
  const tab = (page: Page, name: string) => page.getByRole("tab", { name, exact: true }).click();
  const button = (page: Page, name: string) => ribbon(page).getByRole("button", { name, exact: true }).first();
  /** Menü öğesi (açıklamalı öğelerin erişilebilir adı açıklamayı da içerir). */
  const menuItem = (page: Page, name: string) => page.getByRole("menuitem", { name: new RegExp(`^${name}(\\s|$)`) }).first();
  const pages = (page: Page) => page.locator(".bl-pages > .bl-page");
  const status = (page: Page) => page.locator(".oh-statusbar");

  test("yazma ve kalın: Ctrl+B ve şerit düğmesi, başlıkta kaydedilmedi işareti", async ({ page }) => {
    await page.keyboard.type("Merhaba ");
    await page.keyboard.press("Control+b");
    await page.keyboard.type("kalın");
    await page.keyboard.press("Control+b");
    await page.keyboard.type(" dünya");
    await expect(body(page).locator("strong")).toHaveText("kalın");
    await expect(body(page)).toContainText("Merhaba kalın dünya");
    // Sözcüğü seçip şeritteki İtalik düğmesiyle biçimle.
    await page.keyboard.press("Control+Shift+ArrowLeft");
    await expect.poll(() => page.evaluate(() => window.getSelection()!.toString())).toBe("dünya");
    await button(page, "İtalik").click();
    await expect(body(page).locator("em")).toHaveText("dünya");
    await expect(button(page, "İtalik")).toHaveAttribute("aria-pressed", "true");
    await expect(page.locator(".oh-titlebar")).toContainText("Kaydedilmedi");
    await expect(status(page)).toContainText("3 sözcük");
  });

  test("madde işaretli ve numaralı liste; boş öğede Enter listeyi bitirir", async ({ page }) => {
    await button(page, "Madde İşaretleri").click();
    await page.keyboard.type("Elma");
    await page.keyboard.press("Enter");
    await page.keyboard.type("Armut");
    await page.keyboard.press("Enter");
    await page.keyboard.press("Enter");
    await page.keyboard.type("Liste dışı");
    await expect(body(page).locator("ul > li")).toHaveCount(2);
    await expect(body(page).locator("> section > p").last()).toHaveText("Liste dışı");
    // Tümünü seçip numaralandırmaya çevir.
    await page.keyboard.press("Control+a");
    await button(page, "Numaralandırma").click();
    await expect(body(page).locator("ol > li")).toHaveCount(3);
    await expect(body(page).locator("ul")).toHaveCount(0);
  });

  test("tablo ekleme, Tab ile gezinme ve bağlamsal sekmeler", async ({ page }) => {
    await tab(page, "Ekle");
    await button(page, "Tablo").click();
    await page.getByRole("button", { name: "3x2 tablo" }).click();
    const cells = body(page).locator("table td");
    await expect(cells).toHaveCount(6);
    await expect(page.getByRole("tab", { name: "Tablo Tasarımı", exact: true })).toBeVisible();
    for (const t of ["A1", "B1", "C1", "A2", "B2", "C2"]) {
      await page.keyboard.type(t);
      await page.keyboard.press("Tab");
    }
    // Son hücrede Tab yeni satır ekler (Word'deki gibi).
    await expect(cells).toHaveCount(9);
    await expect(cells.nth(4)).toHaveText("B2");
    // Bağlamsal Düzen sekmesinden satır silinir.
    await page.getByRole("tab", { name: "Düzen", exact: true }).last().click();
    await button(page, "Sil").click();
    await menuItem(page, "Satırları Sil").click();
    await expect(cells).toHaveCount(6);
    // Hücreleri Böl: B1 üç sütuna bölünür, diğer satırdaki hücre genişler.
    await cells.nth(1).click();
    await button(page, "Hücreleri Böl").click();
    const dlg = page.locator(".oh-dialog");
    await dlg.getByLabel("Sütun sayısı:").fill("3");
    await dlg.getByRole("button", { name: "Tamam" }).click();
    await expect(cells).toHaveCount(8);
    await expect(body(page).locator("tr").nth(1).locator("td").nth(1)).toHaveAttribute("colspan", "3");
  });

  test("stil galerisi: Başlık 1 ve gezinti bölmesinde başlık", async ({ page }) => {
    await page.keyboard.type("Giriş Bölümü");
    await ribbon(page).locator('.oh-gallery-item[data-tip="Başlık 1"]').first().click();
    await expect(body(page).locator("h1")).toHaveText("Giriş Bölümü");
    await page.keyboard.press("End");
    await page.keyboard.press("Enter");
    await page.keyboard.type("Gövde metni");
    // Başlıktan sonra Enter normal paragrafa geçer.
    await expect(body(page).locator("> section > p").last()).toHaveText("Gövde metni");
    await tab(page, "Görünüm");
    await ribbon(page).getByLabel("Gezinti Bölmesi").check();
    await expect(page.locator(".bl-nav")).toContainText("Giriş Bölümü");
  });

  test("sayfa sonu iki sayfa oluşturur; bölüm sonu ve yatay bölüm", async ({ page }) => {
    await page.keyboard.type("Birinci sayfa");
    await page.keyboard.press("Control+Enter");
    await page.keyboard.type("İkinci sayfa");
    await expect(pages(page)).toHaveCount(2);
    await expect(status(page)).toContainText("Sayfa 2 / 2");
    // İmleç ikinci sayfanın içinde (gövde alanında) durur.
    const second = (await pages(page).nth(1).boundingBox())!;
    const caret = await page.evaluate(() => window.getSelection()!.getRangeAt(0).getBoundingClientRect().top);
    expect(caret).toBeGreaterThan(second.y);
    expect(caret).toBeLessThan(second.y + second.height);
    // Bölüm sonu (Sonraki Sayfa) ve yalnız yeni bölüme yatay yönlendirme.
    await tab(page, "Düzen");
    await button(page, "Kesmeler").click();
    await menuItem(page, "Sonraki Sayfa").click();
    await page.keyboard.type("Yatay bölüm");
    await expect(pages(page)).toHaveCount(3);
    await button(page, "Yönlendirme").click();
    await menuItem(page, "Yatay").click();
    const third = (await pages(page).nth(2).boundingBox())!;
    const first = (await pages(page).nth(0).boundingBox())!;
    expect(third.width).toBeGreaterThan(third.height);
    expect(first.width).toBeLessThan(first.height);
  });

  test("bul ve değiştir: Ctrl+H, Tümünü Değiştir büyük harfi korur, geri al", async ({ page }) => {
    // Otomatik Düzelt cümle başlarını büyük harfe çevirir (Word'deki gibi).
    await page.keyboard.type("elma armut Elma. elmalı turta");
    await expect(body(page)).toHaveText("Elma armut Elma. Elmalı turta");
    await page.keyboard.press("Control+h");
    const dlg = page.locator(".oh-dialog");
    await expect(dlg).toBeVisible();
    await dlg.getByLabel("Aranan:").fill("elma");
    await dlg.getByLabel("Yeni değer:").fill("ayva");
    await expect(dlg).toContainText("3 eşleşme");
    await dlg.getByRole("button", { name: "Tümünü Değiştir" }).click();
    await expect(body(page)).toHaveText("Ayva armut Ayva. Ayvalı turta");
    await dlg.getByRole("button", { name: "İptal" }).click();
    await expect(dlg).toHaveCount(0);
    await page.keyboard.press("Control+z");
    await expect(body(page)).toHaveText("Elma armut Elma. Elmalı turta");
  });

  test("DOCX olarak kaydet, yeniden aç: içerik aynı", async ({ page, browser }, testInfo) => {
    await ribbon(page).locator('.oh-gallery-item[data-tip="Başlık 1"]').first().click();
    await page.keyboard.type("Rapor Başlığı");
    await page.keyboard.press("Enter");
    await page.keyboard.type("Düz ve ");
    await page.keyboard.press("Control+b");
    await page.keyboard.type("kalın");
    await page.keyboard.press("Control+b");
    await page.keyboard.type(" metin, Türkçe İıĞğŞş.");
    await page.keyboard.press("Enter");
    await button(page, "Madde İşaretleri").click();
    await page.keyboard.type("Birinci madde");
    await page.keyboard.press("Enter");
    await page.keyboard.type("İkinci madde");
    const before = await body(page).innerText();

    const download = page.waitForEvent("download");
    await page.keyboard.press("Control+s");
    const file = await download;
    expect(file.suggestedFilename()).toMatch(/\.docx$/);
    const path = testInfo.outputPath(file.suggestedFilename());
    await file.saveAs(path);
    await expect(page.locator(".oh-titlebar")).not.toContainText("Kaydedilmedi");

    // Yeni bir pencerede aç.
    const other = await browser.newPage();
    const otherErrors = collectErrors(other);
    await other.goto(URL.belge);
    // Dosya içeriği doğrudan verilir (yol Türkçe karakter içerebilir).
    await other.setInputFiles("#bl-open-input", { name: file.suggestedFilename(), mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document", buffer: readFileSync(path) });
    await expect(other.locator(".oh-titlebar")).toContainText(file.suggestedFilename().replace(/\.docx$/, ""));
    const ob = other.locator(".bl-body .ProseMirror");
    await expect(ob.locator("h1")).toHaveText("Rapor Başlığı");
    await expect(ob.locator("strong")).toHaveText("kalın");
    await expect(ob.locator("ul > li")).toHaveCount(2);
    expect(await ob.innerText()).toBe(before);
    expect(otherErrors).toEqual([]);
    await other.close();
  });

  test("geri al ve yinele: yazma, biçim ve şerit düğmeleri", async ({ page }) => {
    await page.keyboard.type("bir");
    // Geri alma geçmişi 500 ms'den uzun aradan sonra yeni adım başlatır (Word'deki gibi sözcük sözcük).
    await page.waitForTimeout(600);
    await page.keyboard.type(" iki");
    // Otomatik Düzelt cümle başını büyütür; düzeltme ayrı bir geri alma adımıdır.
    await expect(body(page)).toHaveText("Bir iki");
    await page.keyboard.press("Control+z");
    await expect(body(page)).toHaveText("Bir");
    await page.keyboard.press("Control+z");
    await expect(body(page)).toHaveText("bir");
    await page.keyboard.press("Control+y");
    await page.keyboard.press("Control+y");
    await expect(body(page)).toHaveText("Bir iki");
    await page.keyboard.press("Control+a");
    await page.keyboard.press("Control+u");
    await expect(body(page).locator("u")).toHaveText("Bir iki");
    // Hızlı erişim araç çubuğundaki Geri Al düğmesi
    await page.locator(".oh-titlebar").getByRole("button", { name: /^Geri Al/ }).click();
    await expect(body(page).locator("u")).toHaveCount(0);
    await expect(body(page)).toHaveText("Bir iki");
  });

  test("yakınlaştırma: durum çubuğu, Ctrl+tekerlek ve Birden Çok Sayfa", async ({ page }) => {
    const canvas = page.locator(".bl-canvas");
    const zoom = () => canvas.evaluate((el) => parseFloat(getComputedStyle(el).zoom));
    await expect(status(page)).toContainText("%100");
    await status(page).getByRole("button", { name: "Yakınlaştır" }).click();
    await expect(status(page)).toContainText("%110");
    expect(await zoom()).toBeCloseTo(1.1, 2);
    await status(page).getByRole("button", { name: "Uzaklaştır" }).click();
    await status(page).getByRole("button", { name: "Uzaklaştır" }).click();
    await expect(status(page)).toContainText("%90");
    // Yakınlaştır iletişim kutusu: %200
    await tab(page, "Görünüm");
    await button(page, "Yakınlaştır").click();
    const dlg = page.locator(".oh-dialog");
    await dlg.getByLabel("%200").check();
    await dlg.getByRole("button", { name: "Tamam" }).click();
    await expect(status(page)).toContainText("%200");
    // Birden Çok Sayfa: iki sayfa yan yana
    await page.keyboard.press("Control+Enter");
    await button(page, "Birden Çok Sayfa").click();
    await expect(page.locator(".bl-workspace")).toHaveClass(/is-tiled/);
    const a = (await pages(page).nth(0).boundingBox())!;
    const b = (await pages(page).nth(1).boundingBox())!;
    expect(Math.abs(a.y - b.y)).toBeLessThan(2);
    expect(b.x).toBeGreaterThan(a.x + a.width);
  });

  test("sağ tık menüsü: Kes/Kopyala, Yazı Tipi iletişim kutusu", async ({ page }) => {
    await page.keyboard.type("Bağlam menüsü deneme");
    await page.keyboard.press("Control+Shift+ArrowLeft");
    await expect.poll(() => page.evaluate(() => window.getSelection()!.toString())).toBe("deneme");
    const word = body(page).locator("p").first();
    const box = (await word.boundingBox())!;
    await page.mouse.click(box.x + box.width - 20, box.y + box.height / 2, { button: "right" });
    const menu = page.getByRole("menu");
    await expect(menu).toBeVisible();
    for (const item of ["Kes", "Kopyala", "Yazı Tipi...", "Paragraf...", "Yeni Açıklama"]) {
      await expect(menu.getByRole("menuitem", { name: new RegExp(`^${item.replace(/\./g, "\\.")}`) }).first()).toBeVisible();
    }
    // Seçim sağ tıkta korunur.
    expect(await page.evaluate(() => window.getSelection()!.toString())).toBe("deneme");
    await menu.getByRole("menuitem", { name: /^Yazı Tipi\.\.\./ }).click();
    const dlg = page.locator(".oh-dialog");
    await expect(dlg).toContainText("Yazı Tipi");
    await dlg.getByRole("button", { name: "İptal" }).click();
    await expect(dlg).toHaveCount(0);
    await expect(body(page)).toBeFocused();
  });

  test("değişiklikleri izle: ekleme ve biçim değişikliği, tümünü reddet", async ({ page }) => {
    await page.keyboard.type("Önceki metin");
    await tab(page, "Gözden Geçir");
    await button(page, "Değişiklikleri İzle").click();
    await expect(status(page)).toContainText("Değişiklikleri İzle: Açık");
    await page.keyboard.type(" eklendi");
    await expect(body(page).locator("ins")).toHaveText(" eklendi");
    await page.keyboard.press("Control+Home");
    await page.keyboard.press("Control+Shift+ArrowRight");
    // Tarayıcının seçim değişikliği olayı düzenleyiciye ulaşsın.
    await expect.poll(() => page.evaluate(() => window.getSelection()!.toString())).toMatch(/^Önceki ?$/);
    await page.waitForTimeout(50);
    await page.keyboard.press("Control+b");
    await expect(body(page).locator(".bl-fmtchg")).toHaveCount(1);
    await ribbon(page).getByRole("button", { name: "Reddet seçenekleri" }).click();
    await menuItem(page, "Tüm Değişiklikleri Reddet").click();
    await expect(body(page)).toHaveText("Önceki metin");
    await expect(body(page).locator("strong, ins, .bl-fmtchg")).toHaveCount(0);
  });

  test("metin kutusu, dipnot ve üst bilgi", async ({ page }) => {
    await page.keyboard.type("Gövde paragrafı");
    // Dipnot: başvuru işareti ve sayfa altındaki not alanı
    await tab(page, "Başvurular");
    await button(page, "Dipnot Ekle").click();
    const note = page.locator('.ProseMirror[data-note-id]').first();
    await expect(note).toBeFocused();
    await page.keyboard.type("Kaynak bilgisi");
    await expect(page.locator(".bl-page").first()).toContainText("Kaynak bilgisi");
    await expect(body(page).locator(".bl-noteref")).toHaveAttribute("data-num", "1");
    // Metin kutusu
    await body(page).click();
    await tab(page, "Ekle");
    await button(page, "Metin Kutusu").click();
    await menuItem(page, "Basit Metin Kutusu").click();
    await page.keyboard.type("Kutudaki yazı");
    await expect(body(page).locator(".bl-textbox")).toContainText("Kutudaki yazı");
    // Üst bilgi: Ekle > Üst Bilgi > Boş
    await button(page, "Üst Bilgi").click();
    await menuItem(page, "Boş").click();
    await expect(page.getByRole("tab", { name: "Üst Bilgi ve Alt Bilgi", exact: true })).toBeVisible();
    // Hazır yer tutucu seçili gelir; yazılan metin yerine geçer.
    const hf = page.locator(".bl-hf-edit .ProseMirror");
    await expect(hf).toBeFocused();
    expect(await page.evaluate(() => window.getSelection()!.toString())).toBe("[Buraya yazın]");
    await page.keyboard.type("Şirket Adı");
    await expect(hf).toHaveText("Şirket Adı");
    await page.keyboard.press("Escape");
    await expect(page.getByRole("tab", { name: "Üst Bilgi ve Alt Bilgi", exact: true })).toHaveCount(0);
    await expect(page.locator(".bl-page").first()).toContainText("Şirket Adı");
  });
});
