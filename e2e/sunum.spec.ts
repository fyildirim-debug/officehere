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

// 3. tur: çökme reproları, bağlam menüleri, otomatik kurtarma, gösteri ayarları, komut araması.
test.describe("OfficeHere Sunum — sağlamlık ve yeni özellikler", () => {
  let errors: string[] = [];

  test.beforeEach(async ({ page }) => {
    errors = collectErrors(page);
    await page.addInitScript(() => {
      const w = window as unknown as Record<string, unknown>;
      delete w.showSaveFilePicker;
      delete w.showOpenFilePicker;
    });
    await page.goto(URL.sunum);
    await expect(page.locator(".stage .sv")).toBeVisible();
  });

  test.afterEach(() => {
    // Bilinçli çökme testinin beklenen kayıtları dışında hata olmamalı.
    expect(errors.filter((e) => !e.includes("Tanılama amaçlı çökme"))).toEqual([]);
  });

  const shapes = (page: Page) => page.locator(".stage .sv [data-shape-id]");
  const thumbs = (page: Page) => page.locator(".thumb-pane .thumb-item");
  const ribbon = (page: Page) => page.locator(".oh-ribbon-body");
  const tab = (page: Page, name: string) => page.getByRole("tab", { name, exact: true }).click();
  const leaveText = async (page: Page) => {
    await page.keyboard.press("Escape");
    await page.keyboard.press("Escape");
  };
  const alive = async (page: Page) => {
    await expect(page.locator(".oh-statusbar")).toBeVisible();
    await expect(page.locator(".stage .sv")).toBeVisible();
  };
  const newSlideWithLayout = async (page: Page, layout: string) => {
    await page.locator(".stage").click({ position: { x: 4, y: 4 } });
    await ribbon(page).getByRole("button", { name: "Yeni Slayt" }).locator("xpath=..").locator(".oh-split-arrow").click();
    await page.locator(".layout-tile", { hasText: layout }).first().click();
  };

  test("çökme reprosu A: madde düzeyli metinde Ctrl+B ve iki kez Ctrl+Shift+> uygulamayı çökertmez", async ({ page }) => {
    await newSlideWithLayout(page, "Başlık ve İçerik");
    await shapes(page).nth(1).click();
    await page.keyboard.type("Birinci madde");
    await page.keyboard.press("Enter");
    await page.keyboard.press("Tab");
    await page.keyboard.type("İkinci düzey");
    await leaveText(page);
    await shapes(page).nth(1).dblclick();
    await page.keyboard.press("Control+a");
    await page.keyboard.press("Control+b");
    for (let i = 0; i < 4; i++) {
      await page.keyboard.press("Control+Shift+Period");
      await page.waitForTimeout(120);
    }
    await alive(page);
    await leaveText(page);
    await expect(shapes(page).nth(1)).toContainText("İkinci düzey");
    // Küçültme kararlıdır: metin kutuya sığar (taşma yok) ve ölçek tekrar tekrar değişmez.
    const box = shapes(page).nth(1).locator(".tb").first();
    const fit = await box.evaluate((el) => (el.firstElementChild as HTMLElement).scrollHeight <= (el as HTMLElement).clientHeight + 2);
    expect(fit).toBe(true);
  });

  test("çökme reprosu B: boşluksuz uzun kelime ve uzun adres kutu genişliğinde kırılır", async ({ page }) => {
    await shapes(page).first().click();
    for (let i = 0; i < 6; i++) {
      await page.keyboard.type("A".repeat(20));
      await page.waitForTimeout(60);
    }
    await alive(page);
    const tb = shapes(page).first().locator(".tb");
    const fits = await tb.evaluate((el) => {
      const inner = el.firstElementChild as HTMLElement;
      return inner.scrollWidth <= (el as HTMLElement).clientWidth + 2;
    });
    expect(fits).toBe(true);
    await leaveText(page);
    await expect(shapes(page).first()).toContainText("AAAAAAAAAA");
    await shapes(page).nth(1).click();
    await page.keyboard.insertText("https://www.ornek.com.tr/cok/uzun/bir/adres/segment1/segment2/segment3/segment4/dosya.html");
    await page.waitForTimeout(200);
    await alive(page);
  });

  test("metin düzenlerken sağ tık menüsü Kes/Kopyala/Yapıştır, Yazı Tipi, Paragraf, listeler ve Bağlantı içerir", async ({ page }) => {
    await shapes(page).first().click();
    await page.keyboard.type("Merhaba dünya");
    await page.keyboard.press("Shift+Home");
    await shapes(page).first().locator(".tb-edit").click({ button: "right" });
    for (const item of ["Kes", "Kopyala", "Yapıştırma Seçenekleri:", "Yazı Tipi...", "Paragraf...", "Madde İşaretleri", "Numaralandırma", "Bağlantı..."]) {
      await expect(page.getByRole("menuitem", { name: new RegExp(`^${item.replace(/[.]/g, "\\.")}`) })).toBeVisible();
    }
    await page.getByRole("menuitem", { name: /^Yazı Tipi/ }).click();
    await expect(page.getByRole("dialog")).toContainText("Yazı Tipi");
    await page.getByRole("button", { name: "İptal" }).click();
  });

  test("tablo hücresinde sağ tık: Ekle › Satırları Alta Ekle ve Sil › Sütunları Sil", async ({ page }) => {
    await page.keyboard.press("Control+m");
    await tab(page, "Ekle");
    await ribbon(page).getByRole("button", { name: "Tablo" }).click();
    await page.locator(".oh-popover .oh-tablegrid-cells > *").nth(10 + 2).click();
    const tds = page.locator(".stage td");
    await expect(tds).toHaveCount(6);
    await tds.first().click();
    await tds.first().click({ button: "right" });
    await page.getByRole("menuitem", { name: /^Ekle/ }).hover();
    await page.getByRole("menuitem", { name: "Satırları Alta Ekle" }).click();
    await expect(tds).toHaveCount(9);
    await tds.first().click({ button: "right" });
    await page.getByRole("menuitem", { name: /^Sil/ }).hover();
    await page.getByRole("menuitem", { name: "Sütunları Sil" }).click();
    await expect(tds).toHaveCount(6);
  });

  test("Otomatik Kurtarma: kaydedilmemiş sunu yenilemeden sonra önerilir ve geri yüklenir", async ({ page }) => {
    await shapes(page).first().click();
    await page.keyboard.type("Kurtarılacak başlık");
    await leaveText(page);
    // Taslak birkaç saniyelik boşluktan sonra yazılır.
    await expect.poll(() => page.evaluate(async () => (await indexedDB.databases()).some((d) => d.name === "officehere-sunum")), { timeout: 8000 }).toBe(true);
    await page.waitForTimeout(3500);
    page.on("dialog", (d) => void d.accept());
    await page.reload();
    await expect(page.locator(".stage .sv")).toBeVisible();
    const bar = page.locator(".oh-messagebar", { hasText: "Kurtarılan sunu" });
    await expect(bar).toBeVisible();
    await bar.getByRole("button", { name: "Geri Yükle" }).click();
    await expect(shapes(page).first()).toContainText("Kurtarılacak başlık");
    await expect(page.locator(".oh-titlebar")).toContainText("Kaydedilmedi");
  });

  test("yazı tipi boyutu adımları sunu listesini izler (44→48→54→60)", async ({ page }) => {
    await page.keyboard.press("Control+m");
    await shapes(page).first().click();
    await page.keyboard.type("Boyut");
    await page.keyboard.press("Control+a");
    const sizeBox = page.getByRole("combobox", { name: "Yazı Tipi Boyutu" }).or(ribbon(page).locator("input").nth(1)).first();
    await expect(sizeBox).toHaveValue("44");
    for (const v of ["48", "54", "60"]) {
      await page.keyboard.press("Control+Shift+Period");
      await expect(sizeBox).toHaveValue(v);
    }
    await page.keyboard.press("Control+Shift+Comma");
    await expect(sizeBox).toHaveValue("54");
  });

  test("hata sınırı: slayt alanı çökerse belge korunur, Yeniden Dene ile geri gelir", async ({ page }) => {
    await shapes(page).first().click();
    await page.keyboard.type("Korunan metin");
    await leaveText(page);
    await page.evaluate(() => {
      (window as unknown as Record<string, string>).__ohSunumCrash = "Slayt alanı";
    });
    await page.keyboard.press("Control+m");
    const box = page.locator(".crash-box");
    await expect(box).toContainText("Bir sorun oluştu");
    await expect(box.getByRole("button", { name: "Belgeyi Kaydet" })).toBeVisible();
    await expect(thumbs(page).first()).toContainText("Korunan metin");
    await box.getByRole("button", { name: "Yeniden Dene", exact: true }).click();
    await expect(page.locator(".crash-box")).toHaveCount(0);
    await alive(page);
  });

  test("Ctrl+M not bölmesinde de yeni slayt ekler", async ({ page }) => {
    await page.locator(".notes-pane textarea").click();
    await page.keyboard.press("Control+m");
    await expect(thumbs(page)).toHaveCount(2);
  });

  test("Slayt Boyutu › Standart (4:3) ölçekleme sorusunu sorar", async ({ page }) => {
    await tab(page, "Tasarım");
    await ribbon(page).getByRole("button", { name: "Slayt Boyutu" }).click();
    await page.getByRole("menuitem", { name: "Standart (4:3)" }).click();
    await expect(page.getByRole("dialog")).toContainText("Slayt Boyutu");
    await page.getByRole("button", { name: "Sığdırmayı Sağla" }).click();
    const ratio = await page.locator(".stage").evaluate((e) => e.clientWidth / e.clientHeight);
    expect(ratio).toBeCloseTo(4 / 3, 1);
  });

  test("Gösteriyi Ayarla: slayt aralığı yalnız seçili slaytları oynatır, G tuşu ızgarayı açar", async ({ page }) => {
    for (const t of ["Bir", "İki", "Üç"]) {
      await page.keyboard.press("Control+m");
      await shapes(page).first().click();
      await page.keyboard.type(t);
      await leaveText(page);
    }
    await tab(page, "Slayt Gösterisi");
    await expect(ribbon(page).getByRole("button", { name: "Baştan" })).toBeVisible();
    await expect(ribbon(page).locator("[role=group][aria-label='Monitörler']").getByRole("button", { name: "Sunucu Görünümü" })).toBeVisible();
    await ribbon(page).getByRole("button", { name: "Slayt Gösterisini Ayarla" }).click();
    const dlg = page.getByRole("dialog");
    await dlg.getByRole("radio", { name: "Slayt aralığı" }).check();
    await dlg.getByRole("spinbutton", { name: "Başlangıç" }).fill("3");
    await dlg.getByRole("spinbutton", { name: "Bitiş" }).fill("4");
    await dlg.getByRole("button", { name: "Tamam" }).click();
    await ribbon(page).getByRole("button", { name: "Baştan" }).click();
    const show = page.locator(".show-root");
    await expect(show).toContainText("İki");
    await page.keyboard.press("ArrowRight");
    await expect(show).toContainText("Üç");
    await page.keyboard.press("g");
    await expect(page.locator(".show-grid")).toBeVisible();
    await page.keyboard.press("Escape");
    await page.keyboard.press("ArrowRight");
    await expect(show.locator(".show-end")).toBeVisible();
    await page.keyboard.press("Escape");
    await expect(show).toHaveCount(0);
  });

  test("Sunucu Görünümü izleyici gösterisini ikinci pencerede açar ve eşitler", async ({ page, context }) => {
    await shapes(page).first().click();
    await page.keyboard.type("Birinci");
    await leaveText(page);
    await page.keyboard.press("Control+m");
    await shapes(page).first().click();
    await page.keyboard.type("İkinci");
    await leaveText(page);
    await thumbs(page).first().click();
    await tab(page, "Slayt Gösterisi");
    const [popup] = await Promise.all([context.waitForEvent("page"), ribbon(page).getByRole("button", { name: "Sunucu Görünümü" }).click()]);
    await expect(popup.locator(".show-stage")).toContainText("Birinci");
    await expect(page.locator(".presenter")).toContainText("İzleyici penceresi bağlı");
    await page.keyboard.press("ArrowRight");
    await expect(popup.locator(".show-stage")).toContainText("İkinci");
    await page.keyboard.press("Escape");
    await expect(page.locator(".show-root")).toHaveCount(0);
    await expect.poll(() => popup.isClosed()).toBe(true);
  });

  test("komut araması şerit komutlarını bulur ve çalıştırır; Bul satırı Bul kutusunu açar", async ({ page }) => {
    const search = page.getByRole("combobox", { name: "Komut ara" });
    await page.keyboard.press("Alt+q");
    await expect(search).toBeFocused();
    await page.keyboard.type("slaytı gizle");
    const list = page.getByRole("listbox", { name: "Arama sonuçları" });
    await expect(list.getByRole("option").first()).toContainText("Slaytı Gizle");
    await page.keyboard.press("Enter");
    await expect(thumbs(page).first()).toHaveClass(/is-hidden|hidden/);
    await search.click();
    await page.keyboard.type("başlık");
    await list.getByRole("option", { name: /için Bul/ }).click();
    await expect(page.getByRole("dialog")).toContainText("Bul");
    await expect(page.getByRole("dialog").locator("input").first()).toHaveValue("başlık");
  });

  const titledSlides = async (page: Page, names: string[]) => {
    for (const t of names) {
      await page.keyboard.press("Control+m");
      await shapes(page).first().click();
      await page.keyboard.type(t);
      await leaveText(page);
    }
  };

  test("Zamanlama Provası: gösteri sonunda süreler slayt zamanlaması olarak kaydedilir", async ({ page }) => {
    await titledSlides(page, ["Bir"]);
    await tab(page, "Slayt Gösterisi");
    await ribbon(page).getByRole("button", { name: "Zamanlama Provası" }).click();
    await expect(page.getByRole("toolbar", { name: "Kayıt" })).toBeVisible();
    await page.waitForTimeout(1200);
    await page.keyboard.press("ArrowRight");
    await page.waitForTimeout(1200);
    await page.keyboard.press("ArrowRight");
    await expect(page.locator(".show-end")).toBeVisible();
    await page.keyboard.press("ArrowRight");
    await expect(page.getByRole("dialog")).toContainText("slayt zamanlamaları olarak kaydedilsin mi");
    await page.getByRole("button", { name: "Evet" }).click();
    // Slayt Sıralayıcısı'nda her slaydın altında kaydedilen süre görünür.
    await expect(page.locator(".sorter-item .sorter-meta").first()).toContainText(/00:0\d/);
    await page.locator(".sorter-item").nth(1).click();
    await page.getByRole("tab", { name: "Geçişler", exact: true }).click();
    await expect(ribbon(page).locator("label", { hasText: "Sonra:" }).locator("input")).toBeChecked();
  });

  test("Özel Slayt Gösterisi tanımlanır ve yalnız seçilen slaytlar oynatılır", async ({ page }) => {
    await titledSlides(page, ["Bir", "İki", "Üç"]);
    await tab(page, "Slayt Gösterisi");
    await ribbon(page).getByRole("button", { name: "Özel Slayt Gösterisi" }).click();
    await page.getByRole("menuitem", { name: "Özel Gösteriler..." }).click();
    await page.getByRole("button", { name: "Yeni..." }).click();
    const left = page.getByRole("listbox", { name: "Sunudaki slaytlar" });
    await left.locator("label", { hasText: "Üç" }).locator("input").check();
    await left.locator("label", { hasText: "Bir" }).locator("input").check();
    await page.getByRole("button", { name: "Ekle >>" }).click();
    await page.getByLabel("Slayt gösterisi adı").fill("Kısa Tur");
    await page.getByRole("button", { name: "Tamam" }).click();
    await expect(page.getByRole("listbox", { name: "Özel gösteriler" })).toContainText("Kısa Tur (2 slayt)");
    await page.getByRole("button", { name: "Göster", exact: true }).click();
    const show = page.locator(".show-root");
    await expect(show).toContainText("Bir");
    await page.keyboard.press("ArrowRight");
    await expect(show).toContainText("Üç");
    await page.keyboard.press("ArrowRight");
    await expect(show.locator(".show-end")).toBeVisible();
    await page.keyboard.press("Escape");
    await ribbon(page).getByRole("button", { name: "Özel Slayt Gösterisi" }).click();
    await expect(page.getByRole("menuitem", { name: "Kısa Tur" })).toBeVisible();
  });

  test("Şekilleri Birleştir: iki şekil tek bir serbest biçime dönüşür", async ({ page }) => {
    await newSlideWithLayout(page, "Boş");
    const st = (await page.locator(".stage").boundingBox())!;
    const draw = async (title: string, x1: number, y1: number, x2: number, y2: number) => {
      await tab(page, "Ekle");
      await ribbon(page).getByRole("button", { name: "Şekiller" }).click();
      await page.locator(`.oh-popover [title='${title}']`).first().click();
      await page.mouse.move(st.x + x1, st.y + y1);
      await page.mouse.down();
      await page.mouse.move(st.x + x2, st.y + y2, { steps: 5 });
      await page.mouse.up();
    };
    await draw("Dikdörtgen", 60, 100, 260, 280);
    await draw("Oval", 160, 180, 340, 360);
    await expect(shapes(page)).toHaveCount(2);
    await page.keyboard.press("Control+a");
    await page.getByRole("tab", { name: "Şekil Biçimi" }).click();
    await ribbon(page).getByRole("button", { name: "Şekilleri Birleştir" }).click();
    await page.getByRole("menuitem", { name: "Birleşim" }).click();
    await expect(shapes(page)).toHaveCount(1);
    await page.keyboard.press("Control+z");
    await expect(shapes(page)).toHaveCount(2);
  });

  test("tema galerisinde sağ tık: Varsayılan Tema Olarak Ayarla yeni sunuya uygulanır", async ({ page }) => {
    await tab(page, "Tasarım");
    const tile = ribbon(page).locator(".oh-gallery-item").nth(2);
    const name = (await tile.getAttribute("aria-label"))!;
    await tile.click({ button: "right" });
    await expect(page.getByRole("menuitem", { name: "Tüm Slaytlara Uygula" })).toBeVisible();
    await page.getByRole("menuitem", { name: "Varsayılan Tema Olarak Ayarla" }).click();
    await page.keyboard.press("Control+n");
    await tab(page, "Tasarım");
    await expect(ribbon(page).getByRole("group", { name: "Temalar" }).locator(".oh-gallery-item.is-selected")).toHaveAttribute("aria-label", name);
  });

  test("Yeni Slayt menüsü tema adını gösterir; Anahattan Slaytlar metin ana hattından slayt üretir", async ({ page }) => {
    await page.locator(".stage").click({ position: { x: 4, y: 4 } });
    await ribbon(page).getByRole("button", { name: "Yeni Slayt" }).locator("xpath=..").locator(".oh-split-arrow").click();
    await expect(page.locator(".oh-popover .oh-menu-header").first()).not.toHaveText("Tema");
    const [chooser] = await Promise.all([page.waitForEvent("filechooser"), page.getByRole("menuitem", { name: "Anahattan Slaytlar..." }).click()]);
    await chooser.setFiles({ name: "anahat.txt", mimeType: "text/plain", buffer: Buffer.from("Gündem\n\tBütçe\n\t\tAyrıntı\nSonuç\n") });
    await expect(thumbs(page)).toHaveCount(3);
    await expect(thumbs(page).nth(1)).toContainText("Gündem");
    await expect(thumbs(page).nth(2)).toContainText("Sonuç");
  });

  test("grafik: yığılmış sütun eklenir, Satır/Sütun Değiştir ve eksen başlığı uygulanır", async ({ page }) => {
    await page.keyboard.press("Control+m");
    await tab(page, "Ekle");
    await ribbon(page).getByRole("button", { name: "Grafik" }).click();
    await page.getByRole("option", { name: "Sütun" }).click();
    await page.getByRole("radio", { name: "Yığılmış Sütun", exact: true }).click();
    await page.getByRole("button", { name: "Tamam" }).click();
    await page.getByRole("button", { name: "Tamam" }).click();
    const chart = page.locator(".stage .sh-chart");
    await expect(chart).toContainText("Kategori 1");
    await page.getByRole("tab", { name: "Grafik Tasarımı" }).click();
    await ribbon(page).getByRole("button", { name: "Satır/Sütun Değiştir" }).click();
    await expect(chart).toContainText("Seri 1");
    await ribbon(page).getByRole("button", { name: "Grafik Öğesi Ekle" }).click();
    await page.getByRole("menuitem", { name: "Eksen Başlıkları" }).hover();
    await page.getByRole("menuitem", { name: "Birincil Dikey" }).click();
    await expect(chart).toContainText("Eksen Başlığı");
  });
});
