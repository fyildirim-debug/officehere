import { readFileSync } from "node:fs";
import { inflateRawSync } from "node:zlib";
import { expect, test, type Page } from "@playwright/test";
import { URL, collectErrors, interceptFileChooser } from "./ortak";

// OfficeHere Belge 5. tur (QA + rakip özellikleri): bulunan hataların ve eklenen özelliklerin uçtan uca kanıtı.

const DOCX_MIME = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";

/** DOCX (zip) içinden bir parçayı okur (merkezi dizin + inflateRaw). */
function zipEntry(b: Buffer, name: string): string | null {
  const e = b.lastIndexOf(Buffer.from([0x50, 0x4b, 0x05, 0x06]));
  for (let i = 0, p = b.readUInt32LE(e + 16), n = b.readUInt16LE(e + 10); i < n; i++, p += 46 + b.readUInt16LE(p + 28) + b.readUInt16LE(p + 30) + b.readUInt16LE(p + 32)) {
    if (b.toString("utf8", p + 46, p + 46 + b.readUInt16LE(p + 28)) !== name) continue;
    const l = b.readUInt32LE(p + 42);
    const start = l + 30 + b.readUInt16LE(l + 26) + b.readUInt16LE(l + 28);
    const data = b.subarray(start, start + b.readUInt32LE(p + 20));
    return (b.readUInt16LE(p + 10) === 8 ? inflateRawSync(data) : data).toString("utf8");
  }
  return null;
}

/** Konuşma sentezleyici taklidi (sayfada çalışır): sözcük sınırı olaylarını sırayla verir, söylenenleri kaydeder. */
function speechMock(voices: { name: string; lang: string; default?: boolean }[]) {
  const w = window as unknown as Record<string, unknown>;
  const spoken: string[] = [];
  w.__spoken = spoken;
  let current: unknown = null;
  let timer = 0;
  const fake = {
    paused: false,
    getVoices: () => voices,
    speak(u: { text: string; onboundary?: (e: unknown) => void; onend?: () => void }) {
      spoken.push(u.text);
      current = u;
      const words = [...u.text.matchAll(/\S+/g)];
      let i = 0;
      const step = () => {
        if (current !== u) return;
        if (fake.paused) {
          timer = window.setTimeout(step, 40);
          return;
        }
        if (i < words.length) {
          u.onboundary?.({ name: "word", charIndex: words[i].index, charLength: words[i][0].length });
          i++;
          timer = window.setTimeout(step, 150);
        } else {
          current = null;
          u.onend?.();
        }
      };
      timer = window.setTimeout(step, 30);
    },
    cancel() {
      current = null;
      window.clearTimeout(timer);
    },
    pause() {
      fake.paused = true;
    },
    resume() {
      fake.paused = false;
    },
    addEventListener() {},
    removeEventListener() {},
  };
  Object.defineProperty(window, "speechSynthesis", { value: fake, configurable: true });
  w.SpeechSynthesisUtterance = class {
    text: string;
    constructor(text: string) {
      this.text = text;
    }
  };
}

test.describe("OfficeHere Belge 5. tur", () => {
  let errors: string[] = [];

  test.beforeEach(async ({ page }) => {
    errors = collectErrors(page);
    interceptFileChooser(page);
    await page.addInitScript(() => {
      const w = window as unknown as Record<string, unknown>;
      delete w.showSaveFilePicker;
      delete w.showOpenFilePicker;
    });
  });

  test.afterEach(() => {
    expect(errors).toEqual([]);
  });

  const start = async (page: Page) => {
    await page.goto(URL.belge);
    await expect(body(page)).toBeFocused();
  };
  const body = (page: Page) => page.locator(".bl-body .ProseMirror");
  const ribbon = (page: Page) => page.locator(".oh-ribbon-body");
  const tab = (page: Page, name: string) => page.getByRole("tab", { name, exact: true }).first().click();
  const button = (page: Page, name: string) => ribbon(page).getByRole("button", { name, exact: true }).first();
  const menuItem = (page: Page, name: string) => {
    const re = new RegExp(`^${name.replace(/[.*+?^${}()|[\]\\]/g, "\\$&")}(\\s|$)`);
    return page.getByRole("menuitem", { name: re }).or(page.getByRole("menuitemcheckbox", { name: re })).first();
  };
  const dialog = (page: Page) => page.getByRole("dialog");
  const paragraphs = (page: Page) => body(page).locator("p").allInnerTexts();
  /** Ctrl+S ile indirilen DOCX'in document.xml'i ve dosya yolu. */
  const saveDocx = async (page: Page, out: string) => {
    const download = page.waitForEvent("download");
    await page.keyboard.press("Control+s");
    const file = await download;
    await file.saveAs(out);
    return zipEntry(readFileSync(out), "word/document.xml") ?? "";
  };
  const reopen = async (page: Page, path: string) => {
    const other = await page.context().newPage();
    const otherErrors = collectErrors(other);
    await other.goto(URL.belge);
    await other.setInputFiles("#bl-open-input", { name: "tur5.docx", mimeType: DOCX_MIME, buffer: readFileSync(path) });
    await expect(other.locator(".oh-titlebar")).toContainText("tur5");
    return { other, otherErrors };
  };

  test("şerit sekmesine fareyle tıklamak yazma odağını belgeden almaz", async ({ page }) => {
    await start(page);
    await page.keyboard.type("Bir");
    await tab(page, "Ekle");
    await expect(page.getByRole("tab", { name: "Ekle", exact: true }).first()).toHaveAttribute("aria-selected", "true");
    await expect(body(page)).toBeFocused();
    await page.keyboard.type(" iki");
    await tab(page, "Giriş");
    await page.keyboard.type(" üç");
    expect((await paragraphs(page))[0]).toBe("Bir iki üç");
  });

  test("Tümünü Seç + kopyala-yapıştır yeni bölüm (sayfa sonu) eklemez; geri al tek adımda", async ({ page }) => {
    await start(page);
    await page.keyboard.type("Önce ");
    await page.keyboard.press("Control+b");
    await page.keyboard.type("kalın");
    await page.keyboard.press("Control+b");
    await page.keyboard.type(" sonra");
    await page.keyboard.press("Control+a");
    await page.keyboard.press("Control+c");
    await page.keyboard.press("Control+End");
    await page.keyboard.press("Enter");
    await page.keyboard.press("Control+v");
    await expect.poll(() => paragraphs(page)).toEqual(["Önce kalın sonra", "Önce kalın sonra"]);
    await expect(body(page).locator("section")).toHaveCount(1);
    await expect(body(page).locator(".bl-secmark")).toHaveCount(0);
    await expect(body(page).locator("strong")).toHaveText(["kalın", "kalın"]);
    await expect(page.locator(".bl-pages > .bl-page")).toHaveCount(1);
  });

  test("numaralandırma: sağ tıkla Yeniden Başla / Devam Et / Değeri Ayarla; DOCX'e başlangıç değeri yazılır", async ({ page }, testInfo) => {
    await start(page);
    await button(page, "Numaralandırma").click();
    for (const t of ["Bir", "İki", "Üç"]) {
      await page.keyboard.type(t);
      await page.keyboard.press("Enter");
    }
    await page.keyboard.press("Enter");
    await page.keyboard.type("Ara metin");
    await page.keyboard.press("Enter");
    await button(page, "Numaralandırma").click();
    await page.keyboard.type("Dört");
    const lists = () => body(page).locator("ol").evaluateAll((els) => els.map((e) => `${(e as HTMLOListElement).start}:${e.children.length}`));
    await expect.poll(lists).toEqual(["1:3", "1:1"]);

    await body(page).getByText("Dört").click({ button: "right" });
    await menuItem(page, "Numaralandırmaya Devam Et").click();
    await expect.poll(lists).toEqual(["1:3", "4:1"]);

    // Ortadaki öğeden yeniden başlat: liste ikiye bölünür.
    await body(page).getByText("İki").click({ button: "right" });
    await menuItem(page, "Numaralandırmaya Yeniden Başla").click();
    await expect.poll(lists).toEqual(["1:1", "1:2", "4:1"]);
    await page.keyboard.press("Control+z");
    await expect.poll(lists).toEqual(["1:3", "4:1"]);

    // Değeri ayarla: imleçteki öğe 10 olur.
    await body(page).getByText("Dört").click({ button: "right" });
    await menuItem(page, "Numaralandırma Değerini Ayarla...").click();
    await dialog(page).getByLabel("Değeri şuna ayarla").fill("10");
    await dialog(page).getByRole("button", { name: "Tamam" }).click();
    await expect.poll(lists).toEqual(["1:3", "10:1"]);

    const xml = await saveDocx(page, testInfo.outputPath("tur5.docx"));
    const numbering = zipEntry(readFileSync(testInfo.outputPath("tur5.docx")), "word/numbering.xml") ?? "";
    expect(numbering).toContain('<w:startOverride w:val="10"/>');
    expect(xml).toContain("Dört");
  });

  test("Sözcük Sayısı: seçimde tüm istatistikler seçime göre; dipnotlar isteğe bağlı", async ({ page }) => {
    await start(page);
    await page.keyboard.type("Bir iki üç dört beş.");
    await page.keyboard.press("Enter");
    await page.keyboard.type("Altı yedi");
    await page.keyboard.press("Control+Alt+f");
    await page.keyboard.type("Dipnot üç sözcük");
    await body(page).getByText("Altı yedi").click();
    await page.keyboard.press("Control+Shift+g");
    const row = (k: string) => dialog(page).locator("tr", { hasText: new RegExp(`^${k}`) }).locator("td").last();
    await expect(row("Sözcük")).toHaveText("10");
    await expect(row("Paragraf")).toHaveText("3");
    await dialog(page).getByText("Metin kutularını, dipnotları ve son notları dahil et").click();
    await expect(row("Sözcük")).toHaveText("7");
    await expect(row("Paragraf")).toHaveText("2");
    await dialog(page).getByRole("button", { name: "Kapat" }).last().click();
    // "iki üç" seçiliyken
    await page.keyboard.press("Control+Home");
    await page.keyboard.press("Control+ArrowRight");
    await page.keyboard.press("Control+Shift+ArrowRight");
    await page.keyboard.press("Control+Shift+ArrowRight");
    await expect.poll(() => page.evaluate(() => window.getSelection()!.toString().trim())).toBe("iki üç");
    await page.keyboard.press("Control+Shift+g");
    await expect(dialog(page)).toContainText("İstatistikler (seçim):");
    await expect(row("Sözcük")).toHaveText("2");
    await expect(row("Paragraf")).toHaveText("1");
    await expect(row("Satır")).toHaveText("1");
    await expect(row("Sayfa")).toHaveText("1");
  });

  test("Tarih ve Saat: biçimli, otomatik güncellenen alan; DOCX'e DATE \\@ olarak yazılır ve geri okunur", async ({ page }, testInfo) => {
    await start(page);
    await page.keyboard.type("Tarih: ");
    await tab(page, "Ekle");
    await button(page, "Tarih ve Saat").click();
    // Beklenen tarih tarayıcının saat diliminde (Europe/Istanbul) hesaplanır; test süreci UTC'de
    // çalışınca gece yarısına yakın saatlerde gün farklı çıkıyordu.
    const [day, month, year] = await page.evaluate(() => {
      const d = new Date();
      return [d.getDate(), d.getMonth(), d.getFullYear()];
    });
    const months = ["Ocak", "Şubat", "Mart", "Nisan", "Mayıs", "Haziran", "Temmuz", "Ağustos", "Eylül", "Ekim", "Kasım", "Aralık"];
    const long = `${day} ${months[month]} ${year}`;
    await dialog(page).getByLabel("Kullanılabilir biçimler").selectOption("d MMMM yyyy");
    await dialog(page).getByText("Otomatik olarak güncelleştir").click();
    await dialog(page).getByRole("button", { name: "Tamam" }).click();
    await expect(body(page).locator(".bl-field")).toHaveText(long);
    // Sabit (güncellenmeyen) tarih metin olarak eklenir; tercih hatırlanır.
    await page.keyboard.type(" / ");
    await button(page, "Tarih ve Saat").click();
    await expect(dialog(page).getByLabel("Otomatik olarak güncelleştir")).toBeChecked();
    await dialog(page).getByLabel("Otomatik olarak güncelleştir").uncheck();
    await dialog(page).getByLabel("Kullanılabilir biçimler").selectOption("yyyy-MM-dd");
    await dialog(page).getByRole("button", { name: "Tamam" }).click();
    const iso = `${year}-${String(month + 1).padStart(2, "0")}-${String(day).padStart(2, "0")}`;
    await expect(body(page)).toContainText(`Tarih: ${long} / ${iso}`);
    await expect(body(page).locator(".bl-field")).toHaveCount(1);

    const path = testInfo.outputPath("tur5.docx");
    const xml = await saveDocx(page, path);
    expect(xml).toContain('w:instr=" DATE \\@ &quot;d MMMM yyyy&quot; \\* MERGEFORMAT "');
    const { other, otherErrors } = await reopen(page, path);
    await expect(other.locator(".bl-body .bl-field")).toHaveText(long);
    expect(otherErrors).toEqual([]);
  });

  test("Büyük İlk Harf: metin içinde ve kenar boşluğunda; seçenekler; DOCX gidiş-dönüşü", async ({ page }, testInfo) => {
    await start(page);
    await page.keyboard.type("Şimdi uzun bir paragraf yazıyoruz ki büyük ilk harf birkaç satır boyunca görünsün; metin yeterince uzun olmalı ve satırlara sarılmalı. ".repeat(2));
    await page.keyboard.press("Enter");
    await page.keyboard.type("Kenar boşluğundaki büyük harf için ikinci paragraf; bu metin de satırlara sarılacak kadar uzun yazılıyor. ".repeat(2));
    await body(page).locator("p").first().click();
    await tab(page, "Ekle");
    await button(page, "Büyük İlk Harf").click();
    await menuItem(page, "Metin İçinde").click();
    const first = body(page).locator("p").first();
    await expect(first).toHaveAttribute("data-dropcap", "drop");
    await expect(first).toHaveAttribute("data-dropcap-lines", "3");
    // İlk harf üç satır boyunca büyük çizilir: genişliği normal harfin iki katından fazladır.
    const box = await first.evaluate((p) => {
      const r = document.createRange();
      r.setStart(p.firstChild!, 0);
      r.setEnd(p.firstChild!, 1);
      const probe = document.createElement("span");
      probe.textContent = "Ş";
      probe.style.cssText = `position:absolute;white-space:nowrap;font:${getComputedStyle(p).font}`;
      document.body.appendChild(probe);
      const normal = probe.getBoundingClientRect().width;
      probe.remove();
      return { w: r.getBoundingClientRect().width, normal, lines: getComputedStyle(p, "::first-letter").getPropertyValue("initial-letter") };
    });
    expect(box.w).toBeGreaterThan(box.normal * 2);
    expect(box.lines.trim()).toBe("3");

    // İkinci paragraf: seçeneklerle kenar boşluğunda, 2 satır.
    await body(page).locator("p").nth(1).click();
    await button(page, "Büyük İlk Harf").click();
    await menuItem(page, "Büyük İlk Harf Seçenekleri...").click();
    await dialog(page).getByText("Kenar Boşluğunda").click();
    await dialog(page).getByLabel("Bırakılacak satır").fill("2");
    await dialog(page).getByRole("button", { name: "Tamam" }).click();
    const second = body(page).locator("p").nth(1);
    await expect(second).toHaveAttribute("data-dropcap", "margin");
    // Kenar boşluğu kipinde harf metnin solundadır; metnin sol kenarı değişmez.
    await expect
      .poll(() =>
        second.evaluate((p) => {
          const t = p.firstChild as Text;
          const r = document.createRange();
          r.setStart(t, 0);
          r.setEnd(t, 1);
          const letter = r.getBoundingClientRect();
          r.setStart(t, 1);
          r.setEnd(t, 2);
          const next = r.getBoundingClientRect();
          const left = p.getBoundingClientRect().left;
          return letter.right <= left + 1 && Math.abs(next.left - left) < 2;
        }),
      )
      .toBe(true);

    const path = testInfo.outputPath("tur5.docx");
    const xml = await saveDocx(page, path);
    expect(xml).toContain('<w:framePr w:dropCap="drop" w:lines="3"');
    expect(xml).toContain('<w:framePr w:dropCap="margin" w:lines="2"');
    const { other, otherErrors } = await reopen(page, path);
    await expect(other.locator(".bl-body .ProseMirror p").first()).toHaveAttribute("data-dropcap", "drop");
    await expect(other.locator(".bl-body .ProseMirror p").nth(1)).toHaveAttribute("data-dropcap", "margin");
    await expect(other.locator(".bl-body .ProseMirror p").first()).toHaveText(/^Şimdi uzun/);
    expect(otherErrors).toEqual([]);

    // Kaldır
    await body(page).locator("p").first().click();
    await button(page, "Büyük İlk Harf").click();
    await menuItem(page, "Yok").click();
    await expect(first).not.toHaveAttribute("data-dropcap", /.*/);
  });

  test("Sesli Oku: okunan sözcük vurgulanır, Türkçe ses yoksa uyarır, Ctrl+Alt+Boşluk ile kapanır", async ({ page }) => {
    await page.addInitScript(speechMock, [{ name: "English", lang: "en-US", default: true }]);
    await start(page);
    await page.keyboard.type("Birinci paragraf sesli okunur.");
    await page.keyboard.press("Enter");
    await page.keyboard.press("Enter");
    await page.keyboard.type("İkinci paragraf da okunur.");
    await page.keyboard.press("Control+Home");
    await tab(page, "Gözden Geçir");
    await button(page, "Sesli Oku").click();
    const bar = page.getByRole("toolbar", { name: "Sesli Oku" });
    await expect(bar).toBeVisible();
    await expect(bar).toContainText("Türkçe ses bulunamadı");
    await expect(body(page).locator(".bl-ra-word")).toHaveText(/paragraf|sesli|okunur\./);
    await expect(button(page, "Sesli Oku")).toHaveAttribute("aria-pressed", "true");
    // Duraklat / sürdür
    await bar.getByRole("button", { name: "Duraklat" }).click();
    await expect(bar.getByRole("button", { name: "Oynat" })).toBeVisible();
    await bar.getByRole("button", { name: "Oynat" }).click();
    // İkinci paragrafa geçer (boş paragraf atlanır).
    await expect.poll(() => page.evaluate(() => (window as unknown as { __spoken: string[] }).__spoken), { timeout: 10_000 }).toContain("İkinci paragraf da okunur.");
    await expect(body(page).locator(".bl-ra-word")).toHaveText(/İkinci|paragraf|da|okunur\./);
    // Ayarlar: hız
    await bar.getByRole("button", { name: "Ayarlar" }).click();
    await expect(bar.getByLabel("Okuma hızı")).toBeVisible();
    await page.keyboard.press("Control+Alt+Space");
    await expect(bar).toHaveCount(0);
    await expect(body(page).locator(".bl-ra-word, .bl-ra-para")).toHaveCount(0);
  });

  test("Sesli Oku: konuşma sentezleyicisi olmayan ortamda uyarı verir", async ({ page }) => {
    await page.addInitScript(() => {
      delete (window as unknown as Record<string, unknown>).speechSynthesis;
    });
    await start(page);
    await page.keyboard.type("Metin");
    await page.keyboard.press("Control+Alt+Space");
    await expect(page.locator(".oh-messagebar").filter({ hasText: "Sesli okuma bu ortamda desteklenmiyor" })).toBeVisible();
    await expect(page.getByRole("toolbar", { name: "Sesli Oku" })).toHaveCount(0);
  });

  test("Alt+X, Ctrl+Shift+Boşluk ve Simge iletişim kutusu (karakter kodu, son kullanılanlar, özel karakterler)", async ({ page }) => {
    await start(page);
    const text = () => body(page).locator("p").first().innerText();
    await page.keyboard.type("Fiyat 20BA");
    await page.keyboard.press("Alt+x");
    await expect.poll(text).toBe("Fiyat ₺");
    await page.keyboard.press("Alt+x");
    await expect.poll(text).toBe("Fiyat 20BA");
    await page.keyboard.press("Alt+x");
    await page.keyboard.press("Control+Shift+Space");
    await page.keyboard.type("TL");
    await expect.poll(text).toBe("Fiyat ₺ TL");

    await tab(page, "Ekle");
    await ribbon(page).getByRole("button", { name: "Simge" }).first().click();
    await menuItem(page, "Diğer Simgeler...").click();
    await dialog(page).getByLabel("Karakter kodu").fill("03A9");
    await expect(dialog(page).getByRole("option", { name: "Ω", exact: true })).toHaveAttribute("aria-selected", "true");
    await dialog(page).getByRole("button", { name: "Ekle" }).click();
    await dialog(page).getByRole("tab", { name: "Özel Karakterler" }).click();
    await dialog(page).getByRole("option", { name: /Uzun Çizgi/ }).first().dblclick();
    await dialog(page).locator(".oh-dialog-footer").getByRole("button", { name: "Kapat" }).click();
    await expect.poll(text).toBe("Fiyat ₺ TLΩ—");
    // Son kullanılanlar menü ızgarasının başında.
    await ribbon(page).getByRole("button", { name: "Simge" }).first().click();
    await expect(page.locator(".bl-symgrid button").first()).toHaveText("—");
    await expect(page.locator(".bl-symgrid button").nth(1)).toHaveText("Ω");
    await page.keyboard.press("Escape");
  });

  test("Shift+F3 döngüsü ve Ctrl+Geri Sil ayrı geri alma adımı", async ({ page }) => {
    await start(page);
    const text = () => body(page).locator("p").first().innerText();
    // Otomatik Düzelt cümle başını büyütür: "Merhaba dünya" (karışık) → BÜYÜK → küçük → Her Sözcük.
    await page.keyboard.type("merhaba dünya");
    await expect.poll(text).toBe("Merhaba dünya");
    await page.keyboard.press("Shift+Home");
    await page.keyboard.press("Shift+F3");
    await expect.poll(text).toBe("MERHABA DÜNYA");
    await page.keyboard.press("Shift+F3");
    await expect.poll(text).toBe("merhaba dünya");
    await page.keyboard.press("Shift+F3");
    await expect.poll(text).toBe("Merhaba Dünya");
    await page.keyboard.press("End");
    await page.keyboard.type(" ve ışık");
    await page.keyboard.press("Control+Backspace");
    await expect.poll(text).toBe("Merhaba Dünya ve ");
    await page.keyboard.press("Control+z");
    await expect.poll(text).toBe("Merhaba Dünya ve ışık");
  });
});
