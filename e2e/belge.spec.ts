import { readFileSync } from "node:fs";
import { resolve } from "node:path";
import { expect, test, type Page } from "@playwright/test";
import { URL, collectErrors, interceptFileChooser } from "./ortak";
import { belgeTur4Tests } from "./belge-tur4";

// OfficeHere Belge: gerçek bir kullanıcı gibi uçtan uca senaryolar (derlenmiş uygulama, yalnız arayüz).

test.describe("OfficeHere Belge senaryoları", () => {
  let errors: string[] = [];

  test.beforeEach(async ({ page }) => {
    errors = collectErrors(page);
    interceptFileChooser(page);
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
  const menuItem = (page: Page, name: string) => {
    // İşaretlenebilir öğeler (checked) menuitemcheckbox rolündedir.
    const re = new RegExp(`^${name}(\\s|$)`);
    return page.getByRole("menuitem", { name: re }).or(page.getByRole("menuitemcheckbox", { name: re })).first();
  };
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

  /* ---------------------------------------------------------------- */
  /* 3. tur: kabul bulguları ve yeni özellikler                          */
  /* ---------------------------------------------------------------- */

  /** Panodan yapıştırır (uzun metinleri hızlıca girmek için). */
  const paste = (page: Page, text: string) =>
    page.evaluate((t) => {
      const dt = new DataTransfer();
      dt.setData("text/plain", t);
      document.querySelector(".bl-body .ProseMirror")!.dispatchEvent(new ClipboardEvent("paste", { clipboardData: dt, bubbles: true, cancelable: true }));
    }, text);
  const dialog = (page: Page) => page.locator(".oh-dialog");
  const PNG = Buffer.from("iVBORw0KGgoAAAANSUhEUgAAABAAAAAMCAIAAADkharWAAAAGklEQVR4nGO4HyNEEmIY1TCqYVTDqAYqagAAcKKm8RLe1JYAAAAASUVORK5CYII=", "base64");

  test("#1 sayfadan uzun ilk paragraf satır sınırından bölünür, taşma olmaz", async ({ page }) => {
    await paste(page, "Uzun ilk paragraf metni sayfalar boyunca akar. ".repeat(260).trim());
    // Sayfa sayısı sistem yazı tiplerine göre değişir (CI'da 3, yerelde 4); önemli olan birden çok sayfaya bölünmesi.
    await expect.poll(() => pages(page).count()).toBeGreaterThanOrEqual(3);
    // Her metin satırı bir sayfanın içinde (sayfa arası boşlukta ya da alt kenar boşluğunda satır yok).
    const bad = await page.evaluate(() => {
      const pgs = Array.from(document.querySelectorAll(".bl-pages > .bl-page")).map((e) => e.getBoundingClientRect());
      // Yalnız metin satırları (sayfa sonu ara parçaları hariç).
      const walk = document.createTreeWalker(document.querySelector(".bl-body .ProseMirror section > p")!, NodeFilter.SHOW_TEXT);
      let n = 0;
      for (let t = walk.nextNode(); t; t = walk.nextNode()) {
        if (t.parentElement?.closest(".bl-pg, .ProseMirror-widget")) continue;
        const range = document.createRange();
        range.selectNodeContents(t);
        for (const r of Array.from(range.getClientRects())) {
          if (r.height < 4 || r.width < 4) continue;
          if (!pgs.some((x) => r.top >= x.top && r.bottom <= x.bottom - 60)) n++;
        }
      }
      return n;
    });
    expect(bad).toBe(0);
    await expect(body(page).locator("section > p")).toHaveCount(1);
  });

  test("#2 sürekli bölüm sonundan sonra yatay yönlendirme yeni sayfada başlar", async ({ page }) => {
    await page.keyboard.type("Dikey sayfa metni");
    await tab(page, "Düzen");
    await button(page, "Kesmeler").click();
    await menuItem(page, "Sürekli").click();
    await page.keyboard.type("Yatay bölüm metni");
    await expect(pages(page)).toHaveCount(1);
    await button(page, "Yönlendirme").click();
    await menuItem(page, "Yatay").click();
    await expect(pages(page)).toHaveCount(2);
    const [first, second] = [(await pages(page).nth(0).boundingBox())!, (await pages(page).nth(1).boundingBox())!];
    expect(first.width).toBeLessThan(first.height);
    expect(second.width).toBeGreaterThan(second.height);
    const text = (await body(page).getByText("Yatay bölüm metni").boundingBox())!;
    expect(text.y).toBeGreaterThan(second.y);
  });

  test("hata sınırı: çöken belge bölmesi kurtarılır, içerik kaybolmaz", async ({ page }) => {
    await page.keyboard.type("Kaybolmaması gereken metin");
    await page.evaluate(() => window.dispatchEvent(new CustomEvent("officehere:crash-test", { detail: "document" })));
    await expect(page.getByText("Bir sorun oluştu — belgeyi kurtar/kaydet")).toBeVisible();
    await expect(page.getByRole("button", { name: "Kaydet" }).last()).toBeVisible();
    await page.getByRole("button", { name: "Belgeyi Kurtar" }).click();
    await expect(body(page)).toContainText("Kaybolmaması gereken metin");
  });

  test("Dosya › Yazdır: önizleme, özel sayfa aralığı ve kopyalar gerçekten uygulanır", async ({ page }) => {
    await page.evaluate(() => {
      const w = window as unknown as { __printed: string[][]; print: () => void };
      w.__printed = [];
      w.print = () => w.__printed.push(Array.from(document.getElementById("bl-print-root")!.children).map((p) => /Sayfa (bir|iki|üç)/.exec(p.textContent ?? "")?.[0] ?? "?"));
    });
    for (const t of ["Sayfa bir", "Sayfa iki", "Sayfa üç"]) {
      await page.keyboard.type(t);
      if (t !== "Sayfa üç") await page.keyboard.press("Control+Enter");
    }
    await page.keyboard.press("Control+p");
    await expect(page.locator(".bl-pr-count")).toHaveText("/ 3");
    await page.getByLabel("Sayfalar", { exact: true }).fill("1, 3");
    await expect(page.locator(".bl-pr-count")).toHaveText("/ 2");
    await page.getByLabel("Geçerli sayfa").fill("1");
    await expect(page.locator(".bl-pr-sheet")).toContainText("Sayfa bir");
    await page.getByRole("button", { name: "Sonraki Sayfa" }).click();
    await expect(page.locator(".bl-pr-sheet")).toContainText("Sayfa üç");
    await page.getByLabel("Kopyalar").fill("2");
    await page.locator(".bl-pr-print").click();
    await expect.poll(() => page.evaluate(() => (window as unknown as { __printed: string[][] }).__printed)).toEqual([["Sayfa bir", "Sayfa üç", "Sayfa bir", "Sayfa üç"]]);
  });

  test("otomatik düzelt: simgeler, kesir, tire, Türkçe cümle başı ve geri alma", async ({ page }) => {
    await page.keyboard.type("istanbul (c) 1/2 bir -- iki. ılık :) ");
    await expect(body(page)).toHaveText("İstanbul © ½ bir – iki. Ilık ☺ ");
    // Ctrl+Z yalnız son düzeltmeyi geri alır.
    await page.keyboard.press("Control+z");
    await expect(body(page)).toHaveText("İstanbul © ½ bir – iki. Ilık :) ");
    // Seçenekler: Dosya › Seçenekler › Otomatik Düzeltme Seçenekleri listesi
    await page.locator(".oh-ribbon-file").click();
    await page.locator(".oh-bs-item", { hasText: "Seçenekler" }).first().click();
    await page.getByRole("button", { name: "Otomatik Düzeltme Seçenekleri..." }).click();
    await expect(dialog(page)).toContainText("(c)");
    await expect(dialog(page)).toContainText("☺");
  });

  test("mini araç çubuğu seçimde açılır ve biçim uygular", async ({ page }) => {
    await page.keyboard.type("Mini araç çubuğu denemesi");
    // "araç" sözcüğüne çift tıklama
    const at = await page.evaluate(() => {
      const range = document.createRange();
      const text = document.querySelector(".bl-body .ProseMirror p")!.firstChild!;
      range.setStart(text, 5);
      range.setEnd(text, 9);
      const r = range.getBoundingClientRect();
      return { x: r.x + r.width / 2, y: r.y + r.height / 2 };
    });
    await page.mouse.dblclick(at.x, at.y);
    const bar = page.locator(".bl-minibar");
    await expect(bar).toBeVisible();
    await bar.getByRole("button", { name: "Kalın" }).click();
    await expect(body(page).locator("strong")).toHaveText("araç");
  });

  test("mini araç çubuğu yazı tipi listesi açılınca kapanmaz (liste ortasındaki yazı tipi)", async ({ page }) => {
    // Listenin ortasındaki bir yazı tipi seçiliyken liste kendi içinde kaydırılır; bu kaydırma çubuğu kapatmamalı.
    const fontBox = page.locator(".oh-ribbon .oh-combo-input").first();
    await fontBox.click();
    await fontBox.fill("Cambria");
    await fontBox.press("Enter");
    await body(page).click();
    await page.keyboard.type("Selam dünya");
    const at = await page.evaluate(() => {
      const w = document.createTreeWalker(document.querySelector(".bl-body .ProseMirror p")!, NodeFilter.SHOW_TEXT);
      const text = w.nextNode()!;
      const range = document.createRange();
      range.setStart(text, 0);
      range.setEnd(text, 5);
      const r = range.getBoundingClientRect();
      return { x: r.x + r.width / 2, y: r.y + r.height / 2 };
    });
    await page.mouse.dblclick(at.x, at.y);
    const bar = page.locator(".bl-minibar");
    await expect(bar).toBeVisible();
    await expect(bar.locator(".oh-combo-input").first()).toHaveValue("Cambria");
    await bar.locator(".oh-combo-arrow").first().click();
    await expect(page.locator(".oh-combo-list")).toBeVisible();
    await expect(bar).toBeVisible();
    await page.locator(".oh-combo-option", { hasText: "Georgia" }).click();
    await expect(body(page).locator('span[style*="Georgia"]')).toHaveText("Selam");
  });

  test("grafik tutamaçlarla büyütülüp küçültülür, her boyutlandırma ayrı geri alınır", async ({ page }) => {
    await tab(page, "Ekle");
    await button(page, "Grafik").click();
    await dialog(page).getByRole("button", { name: "Tamam" }).click();
    const svg = body(page).locator(".bl-chart svg");
    await expect(svg).toHaveAttribute("width", "576");
    // Grafik yazıları metin seçimine katılmaz.
    expect(await body(page).locator(".bl-chart").evaluate((el) => getComputedStyle(el).userSelect)).toBe("none");
    const box = (await svg.boundingBox())!;
    await page.mouse.click(box.x + box.width / 2, box.y + box.height / 2);
    await expect(page.locator(".bl-chart-h")).toHaveCount(8);
    const drag = async (cls: string, dx: number, dy: number) => {
      const h = (await page.locator(`.bl-chart-h.${cls}`).boundingBox())!;
      await page.mouse.move(h.x + h.width / 2, h.y + h.height / 2);
      await page.mouse.down();
      await page.mouse.move(h.x + h.width / 2 + dx, h.y + h.height / 2 + dy, { steps: 8 });
      await page.mouse.up();
    };
    await drag("is-se", 100, 60);
    await expect(svg).toHaveAttribute("width", "676");
    await expect(svg).toHaveAttribute("height", "396");
    await drag("is-w", 200, 0);
    await expect(svg).toHaveAttribute("width", "476");
    await page.keyboard.press("Control+z");
    await expect(svg).toHaveAttribute("width", "676");
    await page.keyboard.press("Control+z");
    await expect(svg).toHaveAttribute("width", "576");
    await expect(svg).toHaveCount(1);
  });

  test("metin alanından geniş tablo kenar boşluğuna taşar; tabloda kaydırma çubuğu çıkmaz", async ({ page }) => {
    const docx = readFileSync(resolve(__dirname, "..", "apps/belge/src/io/__fixtures__/genis-tablo.docx"));
    await page.setInputFiles("#bl-open-input", { name: "genis-tablo.docx", mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document", buffer: docx });
    const card = body(page).locator("td", { hasText: "CARD11" });
    await card.click();
    const m = await body(page)
      .locator(".tableWrapper")
      .first()
      .evaluate((w: HTMLElement) => {
        const cs = getComputedStyle(w);
        return { ox: cs.overflowX, oy: cs.overflowY, sbH: w.offsetHeight - w.clientHeight, sbV: w.offsetWidth - w.clientWidth, table: w.querySelector("table")!.getBoundingClientRect().width, wrapper: w.getBoundingClientRect().width };
      });
    expect(m).toMatchObject({ ox: "visible", oy: "visible", sbH: 0, sbV: 0 });
    expect(m.table).toBeGreaterThan(m.wrapper + 50);
    // Son sütun kırpılmadan görünür.
    const last = (await body(page).locator("td", { hasText: ">2015>200.412" }).first().boundingBox())!;
    expect(last.width).toBeGreaterThan(100);
  });

  test("satır içi resmin boyutlandırma tutamaçları resmin köşelerindedir", async ({ page }) => {
    await page.keyboard.type("Önce ");
    await tab(page, "Ekle");
    await button(page, "Resimler").click();
    const [fc] = await Promise.all([page.waitForEvent("filechooser"), menuItem(page, "Bu Cihaz...").click()]);
    await fc.setFiles({ name: "kare.png", mimeType: "image/png", buffer: PNG });
    const img = body(page).locator("[data-resize-wrapper] img");
    await expect(img).toBeVisible();
    await img.click();
    await expect(body(page).locator('[data-resize-handle="top-right"]')).toBeVisible();
    const ib = (await img.boundingBox())!;
    const corners = { "top-left": [ib.x, ib.y], "top-right": [ib.x + ib.width, ib.y], "bottom-left": [ib.x, ib.y + ib.height], "bottom-right": [ib.x + ib.width, ib.y + ib.height] };
    for (const [dir, [x, y]] of Object.entries(corners)) {
      const hb = (await body(page).locator(`[data-resize-handle="${dir}"]`).boundingBox())!;
      expect(Math.abs(hb.x + hb.width / 2 - x), dir).toBeLessThan(3);
      expect(Math.abs(hb.y + hb.height / 2 - y), dir).toBeLessThan(3);
    }
    // Resim metinle aynı satırda kalır (kapsayıcı satır içi).
    const kap = await body(page).locator("[data-resize-container]").evaluate((el) => getComputedStyle(el).display);
    expect(kap).toBe("inline-flex");
  });

  test("sekme durakları: Paragraf › Sekmeler, ondalık ve noktalı öncü", async ({ page }) => {
    await page.keyboard.type("Ürün\t12,50");
    await page.locator('.oh-ribbon-launcher[data-tip="Paragraf Ayarları"]').first().click();
    await dialog(page).getByRole("button", { name: "Sekmeler..." }).click();
    await dialog(page).getByLabel("Sekme durak konumu").fill("12");
    await dialog(page).getByLabel("Ondalık").check();
    await dialog(page).getByLabel("2 .......").check();
    await dialog(page).getByRole("button", { name: "Ayarla" }).click();
    await dialog(page).getByRole("button", { name: "Tamam" }).click();
    await expect(body(page).locator(".bl-tab.is-dot")).toHaveCount(1);
    await expect(page.locator(".bl-ruler-tab.is-decimal")).toHaveCount(1);
  });

  test("resim: metni kaydır (kare, metnin arkasında) ve resim yazısı + çapraz başvuru", async ({ page }) => {
    await tab(page, "Ekle");
    await button(page, "Resimler").click();
    const [fc] = await Promise.all([page.waitForEvent("filechooser"), menuItem(page, "Bu Cihaz...").click()]);
    await fc.setFiles({ name: "kare.png", mimeType: "image/png", buffer: PNG });
    await expect(body(page).locator('img[src^="data:"]')).toHaveCount(1);
    await tab(page, "Resim Biçimi");
    await button(page, "Metni Kaydır").click();
    await menuItem(page, "Kare").click();
    await expect(body(page).locator("img[data-wrap=square]")).toHaveCount(1);
    await button(page, "Metni Kaydır").click();
    await menuItem(page, "Metnin Arkasında").click();
    await expect(body(page).locator("img[data-wrap=behind]")).toHaveCount(1);
    await tab(page, "Başvurular");
    await button(page, "Resim Yazısı Ekle").click();
    await dialog(page).getByLabel("Resim yazısı", { exact: true }).fill(": Grafik");
    await dialog(page).getByRole("button", { name: "Tamam" }).click();
    await expect(body(page)).toContainText("Şekil 1: Grafik");
    await page.keyboard.press("Control+End");
    await page.keyboard.press("Enter");
    await page.keyboard.type("Bkz. ");
    await button(page, "Çapraz Başvuru").click();
    await dialog(page).getByRole("button", { name: "Ekle" }).click();
    await dialog(page).locator(".oh-btn", { hasText: "Kapat" }).click();
    await expect(body(page).locator(".bl-xref")).toHaveText("Şekil 1");
  });

  test("alıntı ve kaynakça (APA), stil IEEE'ye geçince alıntı numaralanır", async ({ page }) => {
    await page.keyboard.type("Bir alıntı ");
    await tab(page, "Başvurular");
    await button(page, "Alıntı Ekle").click();
    await menuItem(page, "Yeni Kaynak Ekle...").click();
    await dialog(page).getByLabel("Yazar").fill("Kaya, Ali");
    await dialog(page).getByLabel("Başlık").fill("Türkçe Kitap");
    await dialog(page).getByLabel("Yıl").fill("2019");
    await dialog(page).getByLabel("Yayımcı").fill("Yayınevi");
    await dialog(page).getByRole("button", { name: "Tamam" }).click();
    await expect(body(page)).toContainText("Bir alıntı (Kaya, 2019)");
    await page.keyboard.press("Enter");
    await button(page, "Kaynakça").click();
    await menuItem(page, "Kaynakça").click();
    await expect(body(page).locator(".bl-bib")).toContainText("Kaya, A. (2019). Türkçe Kitap. Yayınevi.");
    await button(page, "Stil: APA").click();
    await menuItem(page, "IEEE").click();
    await expect(body(page)).toContainText("Bir alıntı [1]");
    await expect(body(page).locator(".bl-bib")).toContainText("[1] A. Kaya, Türkçe Kitap. Yayınevi, 2019.");
  });

  test("stiller bölmesi ve yeni stil oluşturma", async ({ page }) => {
    await page.keyboard.type("Özel biçimli paragraf");
    await page.keyboard.press("Control+Alt+Shift+s");
    const pane = page.locator(".bl-styles-pane");
    await expect(pane).toBeVisible();
    await pane.getByRole("button", { name: "Yeni Stil" }).click();
    await dialog(page).getByLabel("Ad").fill("Vurgum");
    await dialog(page).getByRole("button", { name: "Kalın" }).click();
    await dialog(page).getByRole("button", { name: "Tamam" }).click();
    await expect(pane).toContainText("Vurgum");
    await expect(body(page).locator("p[data-style=Vurgum]")).toHaveCount(1);
  });

  test("tablo: metni tabloya dönüştür, sırala, başlık satırlarını yinele, metne dönüştür", async ({ page }) => {
    for (const r of ["Ad\tPuan", "Zeynep\t85", "Ali\t92", "Çiğdem\t7"]) {
      await page.keyboard.type(r);
      if (r !== "Çiğdem\t7") await page.keyboard.press("Enter");
    }
    await page.keyboard.press("Control+a");
    await tab(page, "Ekle");
    await button(page, "Tablo").click();
    await menuItem(page, "Metni Tabloya Dönüştür...").click();
    await expect(dialog(page).getByLabel("Satır sayısı")).toHaveValue("4");
    await dialog(page).getByRole("button", { name: "Tamam" }).click();
    const rows = body(page).locator("table tr:not(.bl-pg)");
    await expect(rows).toHaveCount(4);
    await page.getByRole("tab", { name: "Düzen", exact: true }).last().click();
    await button(page, "Sırala").click();
    await dialog(page).getByLabel("Üst bilgi satırı var").check();
    await dialog(page).getByLabel("Sıralama türü").selectOption("number");
    await dialog(page).getByLabel("Sıralama sütunu").selectOption({ label: "Puan" });
    await dialog(page).getByRole("button", { name: "Tamam" }).click();
    await expect(rows.locator("td:first-child")).toHaveText(["Ad", "Çiğdem", "Zeynep", "Ali"]);
    await button(page, "Başlık Satırlarını Yinele").click();
    await expect(button(page, "Başlık Satırlarını Yinele")).toHaveAttribute("aria-pressed", "true");
    await button(page, "Metne Dönüştür").click();
    await dialog(page).getByRole("button", { name: "Tamam" }).click();
    await expect(body(page).locator("table")).toHaveCount(0);
    await expect(body(page).locator("section > p").nth(1)).toHaveText("Çiğdem\t7");
  });

  test("Postalar: CSV alıcı listesi, birleştirme alanı ve sonuç önizleme", async ({ page }) => {
    await page.keyboard.type("Sayın ");
    await tab(page, "Postalar");
    await page.getByTestId("bl-merge-input").setInputFiles({ name: "alicilar.csv", mimeType: "text/csv", buffer: Buffer.from("Ad;Şehir\nAyşe Yılmaz;İzmir\nAli Kaya;Ankara\n") });
    await button(page, "Birleştirme Alanı Ekle").click();
    await menuItem(page, "Ad").click();
    await expect(body(page).locator(".bl-merge")).toHaveText("«Ad»");
    await button(page, "Sonuçları Önizle").click();
    await expect(body(page)).toContainText("Sayın Ayşe Yılmaz");
    await page.getByRole("button", { name: "Sonraki Kayıt" }).click();
    await expect(body(page)).toContainText("Sayın Ali Kaya");
  });

  test("Anahat ve Taslak görünümleri: düzey süzgeci, yükselt/düşür", async ({ page }) => {
    await button(page, "Başlık 1").click();
    await page.keyboard.type("Giriş");
    await page.keyboard.press("Enter");
    await page.keyboard.type("Gövde metni");
    await tab(page, "Görünüm");
    await button(page, "Anahat").click();
    await expect(page.getByRole("tab", { name: "Anahat", exact: true })).toHaveAttribute("aria-selected", "true");
    await expect(body(page).locator(".bl-ol-h")).toHaveCount(1);
    await page.getByLabel("Düzey Göster").selectOption("1");
    await expect(body(page).getByText("Gövde metni")).toBeHidden();
    await page.getByLabel("Düzey Göster").selectOption("0");
    await body(page).getByText("Gövde metni").click();
    await page.keyboard.press("Alt+Shift+ArrowLeft");
    await expect(body(page).locator("h1")).toHaveCount(2);
    await button(page, "Anahatı Kapat").click();
    await tab(page, "Görünüm");
    await button(page, "Taslak").click();
    await expect(pages(page)).toHaveCount(0);
    await expect(body(page)).toContainText("Gövde metni");
  });

  test("belge özellikleri: Dosya › Bilgi'de başlık ve yazar düzenlenir", async ({ page }) => {
    const info = async () => {
      await page.locator(".oh-ribbon-file").click();
      await page.locator(".oh-bs-item", { hasText: "Bilgi" }).first().click();
    };
    await info();
    await page.getByLabel("Başlık").fill("Yıllık Rapor");
    await page.getByLabel("Yazar").fill("Ayşe Yılmaz");
    await page.locator(".oh-bs-back").click();
    await expect(page.locator(".oh-backstage")).toHaveCount(0);
    await info();
    await expect(page.getByLabel("Başlık")).toHaveValue("Yıllık Rapor");
    await expect(page.getByLabel("Yazar")).toHaveValue("Ayşe Yılmaz");
  });
});

// 4. tur: kabul testi düzeltmeleri ve yeni özellikler.
belgeTur4Tests();
