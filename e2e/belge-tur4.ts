import { readFileSync } from "node:fs";
import { resolve } from "node:path";
import { expect, test, type Page } from "@playwright/test";
import { URL, collectErrors } from "./ortak";

// OfficeHere Belge 4. tur: kabul testi bulgularının düzeltmeleri ve yeni özellikler (belge.spec.ts içinden kaydedilir).

const FIXTURE = resolve(__dirname, "..", "apps/belge/src/io/__fixtures__/cok-bolumlu.docx");
const PNG = Buffer.from(
  "iVBORw0KGgoAAAANSUhEUgAAAHgAAABaCAIAAAD8YgW4AAAAjklEQVR42u3QMQ0AMAgAsDFxCEMiilBBwtFKaGT1Y99XIFo0okWLRrRoRIsWjWjRiBYtGtGiES1aNKJFI1q0aESLRrRo0YgWjWjRohEtGtGiRSNaNKJFi0a0aESLFo1o0YgWLRrRohEtWjSiRSNatGhEi0a0aNGIFo1o0aIRLRrRokUjWjSiRYtGtGhEXzFqugJOfKSx3QAAAABJRU5ErkJggg==",
  "base64",
);
const DOCX_MIME = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";

export function belgeTur4Tests() {
  test.describe("OfficeHere Belge 4. tur", () => {
    let errors: string[] = [];

    test.beforeEach(async ({ page }) => {
      errors = collectErrors(page);
      await page.addInitScript(() => {
        const w = window as unknown as Record<string, unknown>;
        delete w.showSaveFilePicker;
        delete w.showOpenFilePicker;
      });
      await page.goto(URL.belge);
      await expect(body(page)).toBeFocused();
    });

    test.afterEach(() => {
      expect(errors).toEqual([]);
    });

    const body = (page: Page) => page.locator(".bl-body .ProseMirror");
    const ribbon = (page: Page) => page.locator(".oh-ribbon-body");
    const tab = (page: Page, name: string) => page.getByRole("tab", { name, exact: true }).click();
    const button = (page: Page, name: string) => ribbon(page).getByRole("button", { name, exact: true }).first();
    const menuItem = (page: Page, name: string) => page.getByRole("menuitem", { name: new RegExp(`^${name}(\\s|$)`) }).first();
    const dialog = (page: Page) => page.getByRole("dialog");
    const autosave = (page: Page) => page.evaluate(() => window.dispatchEvent(new Event("officehere:autosave")));

    test("#1 Birden Çok Sayfa: DOCX'ten açılan çok bölümlü belge satırlara sarılır, metin sayfa dışına taşmaz", async ({ page }) => {
      await page.setInputFiles("#bl-open-input", { name: "cok-bolumlu.docx", mimeType: DOCX_MIME, buffer: readFileSync(FIXTURE) });
      await expect(page.locator(".oh-titlebar")).toContainText("cok-bolumlu");
      await expect.poll(() => page.locator(".bl-pages > .bl-page").count(), { timeout: 20_000 }).toBeGreaterThan(10);
      await tab(page, "Görünüm");
      await button(page, "Birden Çok Sayfa").click();
      const layout = () =>
        page.evaluate(() => {
          const pages = Array.from(document.querySelectorAll(".bl-pages > .bl-page")).map((e) => e.getBoundingClientRect());
          const tw = document.createTreeWalker(document.querySelector(".bl-body .ProseMirror")!, NodeFilter.SHOW_TEXT);
          const rg = document.createRange();
          let outside = 0;
          for (let n = tw.nextNode(); n; n = tw.nextNode()) {
            if (!n.textContent?.trim() || n.parentElement?.closest(".bl-pg, [contenteditable=false]")) continue;
            rg.selectNodeContents(n);
            for (const r of Array.from(rg.getClientRects())) {
              const x = r.left + 2;
              const y = r.top + r.height / 2;
              if (r.height > 0 && !pages.some((p) => x >= p.left - 2 && x <= p.right + 2 && y >= p.top - 2 && y <= p.bottom + 2)) outside++;
            }
          }
          const sc = document.querySelector(".bl-scroll")!;
          return { rows: new Set(pages.map((p) => Math.round(p.top))).size, outside, wide: sc.scrollWidth / sc.clientWidth };
        });
      await expect.poll(async () => (await layout()).rows, { timeout: 15_000 }).toBeGreaterThan(2);
      await expect.poll(async () => (await layout()).outside, { timeout: 15_000 }).toBe(0);
      expect((await layout()).wide).toBeLessThan(2.5);
    });

    test("#2 otomatik kurtarma taslağı pencereye özgü; yalnız kapanan pencerenin taslağı önerilir", async ({ page, context }) => {
      await page.keyboard.type("ANA PENCERE METNİ");
      await autosave(page);
      const second = await context.newPage();
      await second.goto(URL.belge);
      await expect(second.locator(".bl-body .ProseMirror")).toBeVisible();
      await second.waitForTimeout(800);
      // Canlı pencerenin taslağı önerilmez.
      await expect(second.locator(".oh-messagebar", { hasText: "Kurtarılan" })).toHaveCount(0);
      await second.locator(".bl-body .ProseMirror").click();
      await second.keyboard.type("İKİNCİ PENCERE");
      await autosave(second);
      const drafts = await second.evaluate(() =>
        Object.keys(localStorage)
          .filter((k) => k.startsWith("officehere.belge.draft."))
          .map((k) => localStorage.getItem(k) ?? ""),
      );
      expect(drafts).toHaveLength(2);
      expect(drafts.join()).toContain("ANA PENCERE METNİ");
      expect(drafts.join()).toContain("İKİNCİ PENCERE");
      await page.close();
      const third = await context.newPage();
      const thirdErrors = collectErrors(third);
      await third.goto(URL.belge);
      const bar = third.locator(".oh-messagebar", { hasText: "Kurtarılan belge" });
      await expect(bar).toBeVisible();
      await bar.getByRole("button", { name: "Geri Yükle" }).click();
      await expect(third.locator(".bl-body .ProseMirror")).toContainText("ANA PENCERE METNİ");
      await expect(third.locator(".bl-body .ProseMirror")).not.toContainText("İKİNCİ");
      // Devralınan taslak artık bu pencerenin anahtarında (eski anahtar silindi).
      const after = await third.evaluate(() => Object.keys(localStorage).filter((k) => k.startsWith("officehere.belge.draft.")).length);
      expect(after).toBe(2);
      expect(thirdErrors).toEqual([]);
    });

    test("#3 Anahat: Yalnızca İlk Satırı Göster açılır ve kapanır", async ({ page }) => {
      await button(page, "Başlık 1").click();
      await page.keyboard.type("Başlık");
      await page.keyboard.press("Enter");
      await page.keyboard.type("Uzun gövde metni ".repeat(20));
      await tab(page, "Görünüm");
      await button(page, "Anahat").click();
      const cb = page.getByLabel("Yalnızca İlk Satırı Göster");
      await cb.check();
      await expect(cb).toBeChecked();
      await expect(body(page).locator(".bl-ol-first")).toHaveCount(1);
      await cb.uncheck();
      await expect(cb).not.toBeChecked();
      await expect(body(page).locator(".bl-ol-first")).toHaveCount(0);
    });

    test("#4 900 px: Postalar şeridi taşmaz, gruplar açılır düğmeye daralır", async ({ page }) => {
      await page.setViewportSize({ width: 900, height: 700 });
      await tab(page, "Postalar");
      await expect(page.locator(".oh-ribbon-group.is-collapsed").first()).toBeVisible();
      const fits = await ribbon(page).evaluate((el) => (el.firstElementChild as HTMLElement).scrollWidth <= el.clientWidth + 1);
      expect(fits).toBe(true);
      await page.locator(".oh-ribbon-group.is-collapsed").last().getByRole("button").first().click();
      await expect(page.getByRole("button", { name: "Son ve Birleştir" })).toBeVisible();
    });

    test("#5 kare kaydırmalı resim sürüklenerek taşınır (çapa ve konum güncellenir)", async ({ page }) => {
      for (let i = 0; i < 5; i++) {
        await page.keyboard.type(`Paragraf ${i + 1}. ${"Kaydırma deneme metni. ".repeat(8)}`);
        await page.keyboard.press("Enter");
      }
      await page.keyboard.press("Control+Home");
      await page.setInputFiles('input[type=file][accept="image/*"]', { name: "mavi.png", mimeType: "image/png", buffer: PNG });
      const img = body(page).locator('img[src^="data:image/png"]');
      await expect(img).toBeVisible();
      await button(page, "Metni Kaydır").click();
      await menuItem(page, "Kare").click();
      const b0 = (await img.boundingBox())!;
      await page.mouse.move(b0.x + 40, b0.y + 30);
      await page.mouse.down();
      await page.mouse.move(b0.x + 140, b0.y + 110, { steps: 4 });
      await page.mouse.move(b0.x + 260, b0.y + 190, { steps: 4 });
      await page.mouse.up();
      await expect.poll(async () => Math.round((await img.boundingBox())!.y - b0.y), { timeout: 5_000 }).toBeGreaterThan(120);
      const b1 = (await img.boundingBox())!;
      expect(Math.abs(b1.x - (b0.x + 220))).toBeLessThan(8);
      expect(Math.abs(b1.y - (b0.y + 160))).toBeLessThan(8);
      // Tek Geri Al adımı resmi eski yerine döndürür.
      await page.keyboard.press("Control+z");
      await expect.poll(async () => Math.round((await img.boundingBox())!.y - b0.y)).toBeLessThan(4);
    });

    test("#6 Bul/Değiştir: Türkçe i/İ ve ı/I eşleşmesi; sayaç, vurgu ve Tümünü Değiştir tutarlı", async ({ page }) => {
      await page.keyboard.type("İstanbul ve istanbul; ISI ve ısı ve isi; Iğdır ve ığdır; iğdır.");
      await page.keyboard.press("Control+h");
      const dlg = dialog(page);
      const query = dlg.getByLabel("Aranan:");
      const expectCount = async (q: string, n: number) => {
        await query.fill(q);
        await expect(dlg.locator(".bl-dlg-msg")).toHaveText(`${n} eşleşme`);
        await expect(body(page).locator(".bl-hit")).toHaveCount(n);
      };
      await expectCount("istanbul", 2);
      await expectCount("İSTANBUL", 2);
      await expectCount("iğdır", 1);
      await expectCount("ığdır", 2);
      await expectCount("ısı", 2);
      await expectCount("isi", 1);
      await query.fill("istanbul");
      await dlg.getByLabel("Yeni değer:").fill("ankara");
      await dlg.getByRole("button", { name: "Tümünü Değiştir" }).click();
      await expect(body(page)).toContainText("Ankara ve ankara;");
      await expect(dlg.locator(".bl-dlg-msg")).toHaveText("2 değiştirme yapıldı.");
    });

    test("F4 son eylemi yineler, Alt+Shift+D tarih alanı, alt/üst simge dışlaması, bitişik (c)", async ({ page }) => {
      await page.keyboard.type("birinci");
      await page.keyboard.press("Enter");
      await page.keyboard.type("ikinci");
      const selected = () => page.evaluate(() => window.getSelection()!.toString());
      await page.keyboard.press("Control+Home");
      await page.keyboard.press("Shift+End");
      await expect.poll(selected).toBe("Birinci");
      await page.keyboard.press("Control+b");
      await expect(body(page).locator("strong")).toHaveText(["Birinci"]);
      await page.keyboard.press("Control+End");
      await page.keyboard.press("Shift+Home");
      await expect.poll(selected).toBe("ikinci");
      await page.keyboard.press("F4");
      await expect(body(page).locator("strong")).toHaveText(["Birinci", "ikinci"]);
      await page.keyboard.press("End");
      await page.keyboard.press("Enter");
      await page.keyboard.press("Alt+Shift+D");
      await expect(body(page).locator(".bl-field").last()).toHaveText(/^\d{2}\.\d{2}\.\d{4}$/);
      await page.keyboard.type(" Marka(c) ");
      await expect(body(page)).toContainText("Marka©");
      await page.keyboard.press("Control+a");
      await page.keyboard.press("Control+Shift+=");
      await page.keyboard.press("Control+=");
      await expect(body(page).locator("sup")).toHaveCount(0);
      await expect(body(page).locator("sub").first()).toBeVisible();
    });

    test("üst bilgi: Sonraki bölüme geçiş, Öncekine Bağla ile bölüme özgü üst bilgi", async ({ page }) => {
      await page.keyboard.type("Birinci bölüm");
      await tab(page, "Düzen");
      await button(page, "Kesmeler").click();
      await menuItem(page, "Sonraki Sayfa").click();
      await page.keyboard.type("İkinci bölüm");
      await page.locator(".bl-scroll").evaluate((el) => (el.scrollTop = 0));
      const first = (await page.locator(".bl-pages > .bl-page").first().boundingBox())!;
      await page.mouse.dblclick(first.x + first.width / 2, first.y + 40);
      await expect(page.getByRole("tab", { name: "Üst Bilgi ve Alt Bilgi" })).toHaveAttribute("aria-selected", "true");
      await page.keyboard.type("ORTAK ÜST");
      await button(page, "Sonraki").click();
      await expect(button(page, "Öncekine Bağla")).toHaveAttribute("aria-pressed", "true");
      await expect(page.locator(".bl-hf-same").first()).toBeVisible();
      await button(page, "Öncekine Bağla").click();
      await expect(button(page, "Öncekine Bağla")).toHaveAttribute("aria-pressed", "false");
      await page.locator(".bl-hf-edit .ProseMirror").click();
      await page.keyboard.press("Control+a");
      await page.keyboard.type("İKİNCİ ÜST");
      await button(page, "Üst Bilgi ve Alt Bilgiyi Kapat").click();
      await expect(page.locator(".bl-hf-header")).toHaveText(["ORTAK ÜST", "İKİNCİ ÜST"]);
    });

    test("yeni özellikler: satır numaraları, denklem ve grafik", async ({ page }) => {
      await page.keyboard.type("Birinci satır metni.");
      await page.keyboard.press("Enter");
      await page.keyboard.type("İkinci satır metni.");
      await tab(page, "Düzen");
      await button(page, "Satır Numaraları").click();
      await menuItem(page, "Sürekli").click();
      await expect(page.locator(".bl-lnum")).toHaveText(["1", "2"]);
      await page.keyboard.press("End");
      await page.keyboard.press("Enter");
      await page.keyboard.press("Alt+=");
      await dialog(page).getByLabel("Denklem").fill("a^2+b^2=c^2");
      await dialog(page).getByRole("button", { name: "Ekle" }).click();
      await expect(body(page).locator(".bl-eq math msup")).toHaveCount(3);
      await page.keyboard.press("End");
      await page.keyboard.press("Enter");
      await tab(page, "Ekle");
      await button(page, "Grafik").click();
      await dialog(page).getByLabel("Grafik başlığı").fill("Satışlar");
      await dialog(page).getByRole("button", { name: "Tamam" }).click();
      await expect(body(page).locator(".bl-chart svg")).toContainText("Satışlar");
    });

    test("Koru: salt okunur korumada yazılamaz, koruma durdurulunca yazılır", async ({ page }) => {
      await page.keyboard.type("Korunan metin");
      await tab(page, "Gözden Geçir");
      await button(page, "Düzenlemeyi Kısıtla").click();
      await dialog(page).getByRole("button", { name: "Evet, Korumayı Başlat" }).click();
      const bar = page.locator(".oh-messagebar", { hasText: "Kısıtlanmış Düzenleme" });
      await expect(bar).toBeVisible();
      await body(page).click();
      await page.keyboard.press("End");
      await page.keyboard.type(" EK");
      await page.keyboard.press("Enter");
      await expect(body(page)).toHaveText("Korunan metin");
      await bar.getByRole("button", { name: "Korumayı Durdur" }).click();
      await dialog(page).getByRole("button", { name: "Korumayı Durdur" }).click();
      await expect(bar).toHaveCount(0);
      await body(page).click();
      await page.keyboard.press("End");
      await page.keyboard.type(" EK");
      await expect(body(page)).toHaveText("Korunan metin EK");
    });

    test("Postalar: veri kaynağı belgeyle saklanır (kaydet → yeniden aç → önizleme)", async ({ page, browser }, testInfo) => {
      await page.keyboard.type("Sayın ");
      await tab(page, "Postalar");
      await page.getByTestId("bl-merge-input").setInputFiles({ name: "alicilar.csv", mimeType: "text/csv", buffer: Buffer.from("Ad;Şehir\nAyşe Yılmaz;İzmir\nAli Kaya;Ankara\n") });
      await button(page, "Birleştirme Alanı Ekle").click();
      await menuItem(page, "Ad").click();
      const download = page.waitForEvent("download");
      await page.keyboard.press("Control+s");
      const file = await download;
      const path = testInfo.outputPath(file.suggestedFilename());
      await file.saveAs(path);
      const other = await browser.newPage();
      const otherErrors = collectErrors(other);
      await other.goto(URL.belge);
      await other.setInputFiles("#bl-open-input", { name: file.suggestedFilename(), mimeType: DOCX_MIME, buffer: readFileSync(path) });
      await expect(other.locator(".bl-body .bl-merge")).toHaveText("«Ad»");
      await other.getByRole("tab", { name: "Postalar", exact: true }).click();
      await expect(other.locator(".oh-ribbon-body").getByRole("button", { name: "Alıcı Listesini Düzenle", exact: true })).toBeEnabled();
      await other.locator(".oh-ribbon-body").getByRole("button", { name: "Sonuçları Önizle", exact: true }).click();
      await expect(other.locator(".bl-body .ProseMirror")).toContainText("Sayın Ayşe Yılmaz");
      expect(otherErrors).toEqual([]);
      await other.close();
    });

    test("Karşılaştır: özgün dosya ile geçerli belge farkları yeni pencerede izlenen değişiklik", async ({ page, context }, testInfo) => {
      await page.keyboard.type("Kırmızı elma ağaçta durur.");
      const download = page.waitForEvent("download");
      await page.keyboard.press("Control+s");
      const file = await download;
      const path = testInfo.outputPath("ozgun.docx");
      await file.saveAs(path);
      await page.keyboard.press("Control+a");
      await page.keyboard.type("Yeşil elma ağaçta durur.");
      await tab(page, "Gözden Geçir");
      await button(page, "Karşılaştır").click();
      await menuItem(page, "Karşılaştır...").click();
      const dlg = dialog(page);
      // Dosya içeriği doğrudan verilir (çıktı yolu Türkçe karakter içerir).
      await dlg.locator('input[type=file][aria-label="Özgün belge dosyası"]').setInputFiles({ name: "ozgun.docx", mimeType: DOCX_MIME, buffer: readFileSync(path) });
      await expect(dlg.locator('input[name="cmp-orig"]').nth(1)).toBeChecked();
      await dlg.locator('input[name="cmp-rev"]').first().check();
      const popup = context.waitForEvent("page");
      await dlg.getByRole("button", { name: "Tamam" }).click();
      const result = await popup;
      const resultErrors = collectErrors(result);
      await expect(result.locator(".oh-titlebar")).toContainText("Karşılaştırma Sonucu1");
      await expect(result.locator(".bl-body del")).toContainText("Kırmızı");
      await expect(result.locator(".bl-body ins")).toContainText("Yeşil");
      await expect(result.locator(".oh-messagebar", { hasText: "Kurtarılan" })).toHaveCount(0);
      expect(resultErrors).toEqual([]);
    });
  });
}
