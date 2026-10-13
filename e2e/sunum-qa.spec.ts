import { readFileSync } from "node:fs";
import { inflateRawSync } from "node:zlib";
import { expect, test, type Page } from "@playwright/test";
import { URL, collectErrors, interceptFileChooser } from "./ortak";

// OfficeHere Sunum 5. tur: bulunan hataların yeniden üretimleri ve eklenen özellikler
// (yazdırma düzenleri ve dinleyici notları, damlalık, eşit aralık kılavuzları, videoya aktarma).

const shapes = (page: Page) => page.locator(".stage .sv [data-shape-id]");
const thumbs = (page: Page) => page.locator(".thumb-pane .thumb-item");
const ribbon = (page: Page) => page.locator(".oh-ribbon-body");
const tab = (page: Page, name: string) => page.getByRole("tab", { name, exact: true }).click();
const leaveText = async (page: Page) => {
  await page.keyboard.press("Escape");
  await page.keyboard.press("Escape");
};
/** Düzenlenen metnin paragrafları ve sol girintileri. */
const editParas = (page: Page) => page.locator(".tb-edit .tb-p").evaluateAll((els) => els.map((e) => ({ text: e.textContent ?? "", ml: parseFloat(getComputedStyle(e).marginLeft) })));

async function openApp(page: Page) {
  await page.addInitScript(() => {
    const w = window as unknown as Record<string, unknown>;
    delete w.showSaveFilePicker;
    delete w.showOpenFilePicker;
  });
  await page.goto(URL.sunum);
  await expect(page.locator(".stage .sv")).toBeVisible();
}

/** Yeni boş düzenli slayt. */
async function blankSlide(page: Page) {
  await page.keyboard.press("Control+m");
  await ribbon(page).getByRole("button", { name: "Düzen" }).click();
  await page.locator(".layout-tile", { hasText: "Boş" }).click();
}

/** Slayt koordinatlarında (px) dikdörtgen çizer. */
async function drawRect(page: Page, x: number, y: number, w: number, h: number) {
  const st = (await page.locator(".stage .sv").boundingBox())!;
  const k = st.width / 1280;
  await tab(page, "Ekle");
  await ribbon(page).getByRole("button", { name: "Şekiller" }).click();
  await page.locator(".oh-popover [title='Dikdörtgen']").first().click();
  await page.mouse.move(st.x + x * k, st.y + y * k);
  await page.mouse.down();
  await page.mouse.move(st.x + (x + w) * k, st.y + (y + h) * k, { steps: 5 });
  await page.mouse.up();
}

/** Seçili şekle Animasyonlar sekmesinden efekt uygular (isteğe bağlı başlangıç ve gecikme). */
async function animate(page: Page, effect: string, trigger?: string, delay?: string) {
  await tab(page, "Animasyonlar");
  await ribbon(page).getByRole("button", { name: effect }).first().click();
  if (trigger) {
    await ribbon(page).locator(".oh-spinner-wrap", { hasText: "Başlangıç" }).locator("input, button").first().click();
    await page.getByRole("option", { name: trigger }).click();
  }
  if (delay) {
    const d = ribbon(page).locator(".oh-spinner-wrap", { hasText: "Gecikme" }).locator("input");
    await d.fill(delay);
    await d.press("Enter");
  }
}

/** Önizleme (animasyon/geçiş; galeriden 50 ms sonra başlar) bitene kadar bekler. */
const noPreview = async (page: Page) => {
  await page.waitForTimeout(200);
  await expect(page.locator(".show-root.is-preview")).toHaveCount(0, { timeout: 15_000 });
};

/** Gösterideki şekillerin görünürlüğü ("görünür" = visibility visible ve opaklık > 0,5). */
const showVisible = (page: Page) =>
  page
    .locator(".show-root .show-layer")
    .last()
    .locator("[data-shape-id]")
    .evaluateAll((els) => els.map((e) => getComputedStyle(e).visibility === "visible" && Number(getComputedStyle(e).opacity) > 0.5));

/** PPTX (zip) içinden bir parçayı okur. */
function zipEntry(b: Buffer, name: string): string | null {
  const e = b.lastIndexOf(Buffer.from([0x50, 0x4b, 0x05, 0x06]));
  for (let i = 0, p = b.readUInt32LE(e + 16), n = b.readUInt16LE(e + 10); i < n; i++, p += 46 + b.readUInt16LE(p + 28) + b.readUInt16LE(p + 30) + b.readUInt16LE(p + 32)) {
    if (b.toString("utf8", p + 46, p + 46 + b.readUInt16LE(p + 28)) !== name) continue;
    const l = b.readUInt32LE(p + 42),
      start = l + 30 + b.readUInt16LE(l + 26) + b.readUInt16LE(l + 28),
      data = b.subarray(start, start + b.readUInt32LE(p + 20));
    return (b.readUInt16LE(p + 10) === 8 ? inflateRawSync(data) : data).toString("utf8");
  }
  return null;
}

async function saveAndReopen(page: Page): Promise<Buffer> {
  const [download] = await Promise.all([page.waitForEvent("download"), page.keyboard.press("Control+s")]);
  const bytes = readFileSync((await download.path())!);
  await page.keyboard.press("Control+n");
  await expect(thumbs(page)).toHaveCount(1);
  const [chooser] = await Promise.all([page.waitForEvent("filechooser"), page.keyboard.press("Control+o")]);
  await chooser.setFiles({ name: "Sunu1.pptx", mimeType: "application/vnd.openxmlformats-officedocument.presentationml.presentation", buffer: bytes });
  await expect(page.locator(".oh-titlebar")).toContainText("Sunu1");
  return bytes;
}

async function openBackstage(page: Page, item: string) {
  await page.locator(".oh-ribbon-file").click();
  await page.getByText(item, { exact: true }).first().click();
}

test.describe("OfficeHere Sunum 5. tur", () => {
  let errors: string[] = [];

  test.beforeEach(async ({ page }) => {
    errors = collectErrors(page);
    interceptFileChooser(page);
    await openApp(page);
  });

  test.afterEach(() => {
    expect(errors).toEqual([]);
  });

  test("H1: animasyon önizlemesi sürerken slayta tıklama kaybolmaz (önizleme durur, tıklama uygulanır)", async ({ page }) => {
    await blankSlide(page);
    await drawRect(page, 100, 100, 200, 150);
    await tab(page, "Animasyonlar");
    await ribbon(page).getByRole("button", { name: "Soluklaştır (Giriş)" }).click();
    await expect(page.locator(".show-root.is-preview")).toHaveCount(1);
    await expect(page.locator(".sel-frame")).toHaveCount(1);
    // Önizleme oynarken boş alana tıklama seçimi kaldırır ve önizlemeyi durdurur.
    const st = (await page.locator(".stage .sv").boundingBox())!;
    await page.mouse.click(st.x + st.width - 20, st.y + st.height - 20);
    await expect(page.locator(".show-root.is-preview")).toHaveCount(0);
    await expect(page.locator(".sel-frame")).toHaveCount(0);
    // Önizleme sürerken şekil seçilebilir.
    await ribbon(page).getByRole("button", { name: "Önizleme" }).click();
    await expect(page.locator(".show-root.is-preview")).toHaveCount(1);
    await shapes(page).first().click();
    await expect(page.locator(".sel-frame")).toHaveCount(1);
  });

  test("H2/H3: gecikmeli giriş efekti gecikme boyunca görünmez; 'Öncekinden Sonra' Belir şekli gösterir", async ({ page }) => {
    await blankSlide(page);
    await drawRect(page, 80, 80, 150, 100);
    await animate(page, "Soluklaştır (Giriş)", undefined, "1,5");
    await noPreview(page);
    await drawRect(page, 400, 80, 150, 100);
    await animate(page, "Belir (Giriş)");
    await noPreview(page);
    await drawRect(page, 700, 80, 150, 100);
    await animate(page, "Belir (Giriş)", "Öncekinden Sonra");
    await noPreview(page);
    await page.keyboard.press("Shift+F5");
    await expect(page.locator(".show-root")).toBeVisible();
    await expect.poll(() => showVisible(page)).toEqual([false, false, false]);
    // 1. tıklama: soluklaştırma 1,5 sn gecikmeli; hemen ardından hâlâ görünmez.
    await page.keyboard.press(" ");
    await page.waitForTimeout(300);
    expect((await showVisible(page))[0]).toBe(false);
    await expect.poll(() => showVisible(page), { timeout: 5000 }).toEqual([true, false, false]);
    // 2. tıklama: Belir ve ardından (Öncekinden Sonra) Belir.
    await page.keyboard.press(" ");
    await expect.poll(() => showVisible(page)).toEqual([true, true, true]);
    await page.keyboard.press("Escape");
  });

  test("H8: önizleme sürerken başlatılan gösteri, eski önizlemenin bitişiyle kapanmaz", async ({ page }) => {
    await blankSlide(page);
    await drawRect(page, 100, 100, 200, 150);
    await animate(page, "Soluklaştır (Giriş)", undefined, "1");
    await noPreview(page);
    // Galeriden yeni efekt: önizleme başlar; bitmeden gösteri açılır.
    await ribbon(page).getByRole("button", { name: "Yüzerek Gir (Giriş)" }).click();
    await expect(page.locator(".show-root.is-preview")).toHaveCount(1);
    await page.keyboard.press("Shift+F5");
    await expect(page.locator(".show-root:not(.is-preview)")).toBeVisible();
    await page.waitForTimeout(3000);
    await expect(page.locator(".show-root:not(.is-preview)")).toBeVisible();
    await page.keyboard.press("Escape");
    await expect(page.locator(".show-root")).toHaveCount(0);
  });

  test("H4/H5: imleç taşındıktan hemen sonra Tab imleç yerine uygulanır; geri alma yalnız son yazımı geri alır", async ({ page }) => {
    await page.keyboard.press("Control+m");
    await shapes(page).nth(1).click();
    await page.keyboard.type("Birinci paragraf");
    await page.keyboard.press("Enter");
    await page.keyboard.type("İkinci ortada");
    for (let i = 0; i < 6; i++) await page.keyboard.press("ArrowLeft");
    await page.keyboard.press("Tab");
    expect((await editParas(page))[1].text).toBe("İkinci \tortada");
    // Satır başında Tab düzeyi artırır (sona sekme eklemez).
    await page.keyboard.press("Home");
    await page.keyboard.press("Tab");
    const p = await editParas(page);
    expect(p[1].text).toBe("İkinci \tortada");
    expect(p[1].ml).toBeGreaterThan(p[0].ml);
    // Yazım: imleç taşındıktan sonra yazılan "X" tek başına geri alınır.
    await page.keyboard.press("End");
    await page.keyboard.press("ArrowLeft");
    await page.keyboard.type("X");
    expect((await editParas(page))[1].text).toBe("İkinci \tortadXa");
    await page.keyboard.press("Control+z");
    expect((await editParas(page))[1].text).toBe("İkinci \tortada");
    expect((await editParas(page))[0].text).toBe("Birinci paragraf");
  });

  test("Alt+Shift+Sol/Sağ liste düzeyini, Alt+Shift+Yukarı/Aşağı paragraf sırasını değiştirir", async ({ page }) => {
    await page.keyboard.press("Control+m");
    await shapes(page).nth(1).click();
    await page.keyboard.type("Bir");
    await page.keyboard.press("Enter");
    await page.keyboard.type("İki");
    await page.keyboard.press("Enter");
    await page.keyboard.type("Üç");
    await page.keyboard.press("Alt+Shift+ArrowRight");
    let p = await editParas(page);
    expect(p[2].ml).toBeGreaterThan(p[1].ml);
    await page.keyboard.press("Alt+Shift+ArrowLeft");
    p = await editParas(page);
    expect(p[2].ml).toBe(p[1].ml);
    await page.keyboard.press("Alt+Shift+ArrowUp");
    expect((await editParas(page)).map((x) => x.text)).toEqual(["Bir", "Üç", "İki"]);
    await page.keyboard.press("Alt+Shift+ArrowUp");
    expect((await editParas(page)).map((x) => x.text)).toEqual(["Üç", "Bir", "İki"]);
    // İmleç taşınan paragrafla gider: yazılan metin "Üç"e eklenir.
    await page.keyboard.type("!");
    expect((await editParas(page))[0].text).toBe("Üç!");
    await page.keyboard.press("Alt+Shift+ArrowDown");
    expect((await editParas(page)).map((x) => x.text)).toEqual(["Bir", "Üç!", "İki"]);
    // Düzenlemeden çıkınca seçili kutunun tüm paragrafları düşürülür.
    await leaveText(page);
    await shapes(page).nth(1).click();
    await page.keyboard.press("Escape");
    await page.keyboard.press("Alt+Shift+ArrowRight");
    const levels = await shapes(page).nth(1).locator(".tb-p").evaluateAll((els) => els.map((e) => parseFloat(getComputedStyle(e).marginLeft)));
    expect(levels.every((m) => m > 30)).toBe(true);
  });

  test("H7: Gösteriyi Ayarla seçenekleri kaydedilen PPTX'te korunur (p:showPr)", async ({ page }) => {
    for (const t of ["Bir", "İki"]) {
      await page.keyboard.press("Control+m");
      await shapes(page).first().click();
      await page.keyboard.type(t);
      await leaveText(page);
    }
    await tab(page, "Slayt Gösterisi");
    await ribbon(page).getByRole("button", { name: "Slayt Gösterisini Ayarla" }).click();
    let dlg = page.getByRole("dialog");
    await dlg.getByRole("radio", { name: "Bilgi noktasında göz atılan (tam ekran)" }).check();
    await dlg.getByRole("checkbox", { name: "Animasyonsuz göster" }).check();
    await dlg.getByRole("radio", { name: "Slayt aralığı" }).check();
    await dlg.getByRole("spinbutton", { name: "Başlangıç" }).fill("2");
    await dlg.getByRole("spinbutton", { name: "Bitiş" }).fill("3");
    await dlg.getByRole("button", { name: "Tamam" }).click();
    await expect(page.locator(".oh-titlebar")).toContainText("Kaydedilmedi");
    const bytes = await saveAndReopen(page);
    expect(zipEntry(bytes, "ppt/presProps.xml")).toContain('<p:kiosk/><p:sldRg st="2" end="3"/>');
    await tab(page, "Slayt Gösterisi");
    await ribbon(page).getByRole("button", { name: "Slayt Gösterisini Ayarla" }).click();
    dlg = page.getByRole("dialog");
    await expect(dlg.getByRole("radio", { name: "Bilgi noktasında göz atılan (tam ekran)" })).toBeChecked();
    await expect(dlg.getByRole("checkbox", { name: "Animasyonsuz göster" })).toBeChecked();
    await expect(dlg.getByRole("spinbutton", { name: "Başlangıç" })).toHaveValue("2");
  });

  test("Damlalık: başka bir şeklin rengini seçili şeklin dolgusuna uygular; Esc vazgeçer", async ({ page }) => {
    await blankSlide(page);
    await drawRect(page, 100, 100, 200, 150);
    await tab(page, "Şekil Biçimi");
    await ribbon(page).getByRole("button", { name: "Şekil Dolgusu seçenekleri" }).click();
    await page.locator(".oh-popover [aria-label='#FF0000']").first().click();
    await drawRect(page, 600, 300, 200, 150);
    const fill = (i: number) => shapes(page).nth(i).locator("path.sh-path").first().getAttribute("fill");
    expect(await fill(1)).not.toBe("#FF0000");
    // Esc ile vazgeçme: renk değişmez, seçim korunur.
    await tab(page, "Şekil Biçimi");
    await ribbon(page).getByRole("button", { name: "Şekil Dolgusu seçenekleri" }).click();
    await page.getByRole("menuitem", { name: "Damlalık" }).click();
    await expect(page.locator(".eyedropper-layer.is-ready")).toBeVisible();
    await page.keyboard.press("Escape");
    await expect(page.locator(".eyedropper-layer")).toHaveCount(0);
    // Kırmızı şeklin üzerinden renk alınır.
    await ribbon(page).getByRole("button", { name: "Şekil Dolgusu seçenekleri" }).click();
    await page.getByRole("menuitem", { name: "Damlalık" }).click();
    await expect(page.locator(".eyedropper-layer.is-ready")).toBeVisible();
    const red = (await shapes(page).nth(0).boundingBox())!;
    await page.mouse.move(red.x + red.width / 2, red.y + red.height / 2);
    await expect(page.locator(".eyedropper-tip")).toHaveAttribute("data-color", "#FF0000");
    await expect(page.locator(".eyedropper-tip")).toContainText("RGB(255, 0, 0)");
    await page.mouse.down();
    await page.mouse.up();
    await expect.poll(() => fill(1)).toBe("#FF0000");
    await expect(page.locator(".eyedropper-layer")).toHaveCount(0);
    await page.keyboard.press("Control+z");
    await expect.poll(() => fill(1)).not.toBe("#FF0000");
  });

  test("eşit aralık kılavuzu: iki şeklin arasına sürüklenen şekil ortaya yapışır ve aralık işaretleri görünür", async ({ page }) => {
    await blankSlide(page);
    await drawRect(page, 100, 300, 100, 100);
    await drawRect(page, 500, 300, 100, 100);
    await drawRect(page, 300, 560, 80, 80);
    const st = (await page.locator(".stage .sv").boundingBox())!;
    const k = st.width / 1280;
    const c = (await shapes(page).nth(2).boundingBox())!;
    await page.mouse.move(c.x + c.width / 2, c.y + c.height / 2);
    await page.mouse.down();
    // Hedef: x = (200 + 500 - 80) / 2 = 310 (3 px yakına bırakılır).
    await page.mouse.move(st.x + (313 + 40) * k, st.y + 350 * k, { steps: 10 });
    await expect(page.locator(".smart-gap")).toHaveCount(2);
    await page.mouse.up();
    await expect(page.locator(".smart-gap")).toHaveCount(0);
    expect(await shapes(page).nth(2).evaluate((e) => parseFloat((e as HTMLElement).style.left))).toBeCloseTo(310, 0);
  });

  test("Yazdır: dinleyici notları (6 slayt), özel aralık, not sayfası üst/alt bilgisi ve ayarların PPTX'e yazılması", async ({ page }) => {
    await shapes(page).first().click();
    await page.keyboard.type("Kapak");
    await leaveText(page);
    for (let i = 2; i <= 8; i++) {
      await page.keyboard.press("Control+m");
      await shapes(page).first().click();
      await page.keyboard.type(`Konu ${i}`);
      await leaveText(page);
    }
    await page.locator(".notes-pane textarea").fill("Konuşmacı notu");
    await page.locator(".stage").click({ position: { x: 4, y: 4 } });
    // Ctrl+P hemen yazdırmaz; Dosya › Yazdır sayfasını açar.
    await page.keyboard.press("Control+p");
    await expect(page.getByRole("heading", { name: "Yazdır" })).toBeVisible();
    const preview = page.locator(".su-printpage-preview");
    await expect(preview.locator(".pp-page.is-slide")).toHaveCount(1);
    await expect(preview.locator(".print-prev-nav")).toContainText("1 / 8");
    await page.getByLabel("Yazdırma düzeni").selectOption("handouts6:horz");
    await expect(preview.locator(".pp-page.is-handout .pp-slide")).toHaveCount(6);
    await expect(preview.locator(".print-prev-nav")).toContainText("1 / 2");
    await preview.getByRole("button", { name: "Sonraki Sayfa" }).click();
    await expect(preview.locator(".pp-page.is-handout .pp-slide")).toHaveCount(2);
    // Özel aralık: 2-4 → 3 slayt tek sayfada.
    await page.getByLabel("Slayt aralığı").fill("2-4");
    await expect(page.getByLabel("Yazdırılacak slaytlar")).toHaveValue("custom");
    await expect(preview.locator(".print-prev-nav")).toContainText("1 / 1");
    await expect(preview.locator(".pp-slide")).toHaveCount(3);
    await page.getByLabel("Slayt aralığı").fill("2-x");
    await expect(page.locator(".su-printpage-error")).toBeVisible();
    await page.getByLabel("Slayt aralığı").fill("");
    // Not sayfaları: üst bilgi metni Üst Bilgi ve Alt Bilgi > Notlar sekmesinden.
    await page.getByLabel("Yazdırma düzeni").selectOption("notes");
    await page.getByRole("button", { name: "Üst Bilgi ve Alt Bilgiyi Düzenle" }).click();
    const dlg = page.getByRole("dialog", { name: "Üst Bilgi ve Alt Bilgi" });
    await expect(dlg.getByRole("tab", { name: "Notlar ve Dinleyici Notları" })).toHaveAttribute("aria-selected", "true");
    await dlg.getByRole("checkbox", { name: "Üst Bilgi" }).check();
    await dlg.getByLabel("Üst bilgi metni").fill("Eğitim Kitapçığı");
    await dlg.getByRole("button", { name: "Tümüne Uygula" }).click();
    await openBackstage(page, "Yazdır");
    await expect(preview.locator(".pp-page.is-notes")).toContainText("Eğitim Kitapçığı");
    await expect(preview.locator(".pp-page.is-notes .pp-hf.is-bottom")).toContainText("1");
    await page.locator(".oh-bs-back").click();
    const bytes = await saveAndReopen(page);
    expect(zipEntry(bytes, "ppt/presProps.xml")).toContain('<p:prnPr prnWhat="notes"/>');
    expect(zipEntry(bytes, "ppt/notesMasters/notesMaster1.xml")).toContain("Eğitim Kitapçığı");
    await openBackstage(page, "Yazdır");
    await expect(page.getByLabel("Yazdırma düzeni")).toHaveValue("notes");
    await expect(preview.locator(".pp-page.is-notes")).toContainText("Eğitim Kitapçığı");
  });

  test("Dışarı Aktar › Video Oluştur: oynatılabilir WebM, süre zamanlamalara göre", async ({ page }) => {
    await shapes(page).first().click();
    await page.keyboard.type("Video");
    await leaveText(page);
    await blankSlide(page);
    await drawRect(page, 100, 100, 300, 200);
    await tab(page, "Geçişler");
    await ribbon(page).getByRole("button", { name: "İt" }).click();
    await noPreview(page);
    await openBackstage(page, "Dışarı Aktar");
    await page.locator(".oh-bs-action", { hasText: "Video Oluştur" }).click();
    const dlg = page.getByRole("dialog", { name: "Video Oluştur" });
    await dlg.getByLabel("Video kalitesi").selectOption("480");
    await dlg.getByLabel("Her slaytta geçirilecek saniye").fill("1,5");
    // 1,5 + (1 geçiş + 1,5) = 4 sn
    await expect(dlg.locator(".video-info")).toContainText("2 slayt · Video süresi: 00:04");
    const [download] = await Promise.all([page.waitForEvent("download", { timeout: 60_000 }), dlg.getByRole("button", { name: "Video Oluştur" }).click()]);
    expect(download.suggestedFilename()).toBe("Sunu1.webm");
    const bytes = readFileSync((await download.path())!);
    expect(bytes.subarray(0, 4).toString("hex")).toBe("1a45dfa3");
    const meta = await page.evaluate(async (b64) => {
      const bin = Uint8Array.from(atob(b64), (c) => c.charCodeAt(0));
      const v = document.createElement("video");
      v.muted = true;
      v.src = self.URL.createObjectURL(new Blob([bin], { type: "video/webm" }));
      await new Promise((res, rej) => ((v.onloadedmetadata = res), (v.onerror = () => rej(new Error(`video ${v.error?.code}`)))));
      return { duration: v.duration, w: v.videoWidth, h: v.videoHeight };
    }, bytes.toString("base64"));
    expect(meta.h).toBe(480);
    expect(meta.w).toBe(854);
    expect(meta.duration).toBeCloseTo(4, 1);
    await expect(page.locator(".oh-messagebar")).toContainText("Video oluşturuldu");
  });

  test("telefon genişliği (390 px): yatay sayfa kaydırması yok; Dosya › Yazdır sayfası sığar", async ({ page }) => {
    await page.setViewportSize({ width: 390, height: 844 });
    const scrollW = () => page.evaluate(() => document.documentElement.scrollWidth);
    await expect.poll(scrollW).toBeLessThanOrEqual(390);
    // Küçük resim bölmesi daralır; slayt alanı görünür kalır.
    expect((await page.locator(".thumb-pane").boundingBox())!.width).toBeLessThanOrEqual(110);
    expect((await page.locator(".stage").boundingBox())!.width).toBeGreaterThan(150);
    await openBackstage(page, "Yazdır");
    await expect(page.locator(".su-printpage-paper")).toBeVisible();
    await expect.poll(scrollW).toBeLessThanOrEqual(390);
    const paper = (await page.locator(".su-printpage-paper").boundingBox())!;
    expect(paper.x + paper.width).toBeLessThanOrEqual(390);
  });

  test("Ana Hat: başlıkta Enter yeni slayt, gövdede Enter düzeyi korur, Shift+Tab satırı slayt yapar, başlıkta Tab indirir, Ctrl+Z geri alır", async ({ page }) => {
    await shapes(page).first().click();
    await page.keyboard.type("Bir");
    await leaveText(page);
    await page.keyboard.press("Control+m");
    await shapes(page).first().click();
    await page.keyboard.type("İki");
    await leaveText(page);
    await tab(page, "Görünüm");
    await ribbon(page).getByRole("button", { name: "Ana Hat Görünümü" }).click();
    const outline = () => page.locator(".outline-pane .outline-slide").evaluateAll((els) => els.map((e) => [(e.querySelector(".outline-title") as HTMLInputElement).value, (e.querySelector(".outline-body") as HTMLTextAreaElement | null)?.value ?? null]));
    await page.locator(".outline-pane .outline-title").nth(1).click();
    await page.keyboard.press("End");
    await page.keyboard.press("Enter");
    await page.keyboard.type("Üç");
    await expect.poll(outline).toEqual([
      ["Bir", ""],
      ["İki", ""],
      ["Üç", ""],
    ]);
    await expect(page.locator(".oh-statusbar")).toContainText("Slayt 3 / 3");
    await page.keyboard.press("ArrowDown");
    await page.keyboard.type("Madde A");
    await page.keyboard.press("Enter");
    await page.keyboard.press("Tab");
    await page.keyboard.type("Alt madde");
    await page.keyboard.press("Enter");
    await page.keyboard.type("Alt madde 2");
    await page.keyboard.press("Enter");
    await page.keyboard.press("Shift+Tab");
    await page.keyboard.type("Dört");
    await expect.poll(outline).toEqual([
      ["Bir", ""],
      ["İki", ""],
      ["Üç", "Madde A\n\tAlt madde\n\tAlt madde 2\nDört"],
    ]);
    // İlk düzeydeki "Dört" satırında Shift+Tab: yeni slayt.
    await page.keyboard.press("Shift+Tab");
    await expect.poll(outline).toEqual([
      ["Bir", ""],
      ["İki", ""],
      ["Üç", "Madde A\n\tAlt madde\n\tAlt madde 2"],
      ["Dört", ""],
    ]);
    await expect(page.locator(".stage .sv")).toContainText("Dört");
    // Yeni slaydın başlığında Tab: önceki slayda indirilir.
    await page.keyboard.press("Tab");
    await expect.poll(outline).toEqual([
      ["Bir", ""],
      ["İki", ""],
      ["Üç", "Madde A\n\tAlt madde\n\tAlt madde 2\nDört"],
    ]);
    await page.keyboard.press("Control+z");
    await expect.poll(outline).toHaveLength(4);
    await page.keyboard.press("Control+y");
    await expect.poll(outline).toHaveLength(3);
  });

  test("Yazdır › dinleyici notları: baskı kökü her sayfayı ayrı A4 sayfası olarak basar (PDF sayfa sayısı)", async ({ page }) => {
    // Sistem yazdırma penceresi açılmasın; yazdırma kökü PDF olarak alınır. Kök, yazdırmadan
    // 0,5 sn sonra kaldırılır (afterprint göndermeyen tarayıcılar için): PDF bitene dek tutulur.
    await page.evaluate(() => {
      (window as unknown as { __printed: number }).__printed = 0;
      window.print = () => {
        (window as unknown as { __printed: number }).__printed++;
      };
      const st = window.setTimeout;
      window.setTimeout = ((f: TimerHandler, ms?: number, ...rest: unknown[]) => st(f, ms === 500 ? 600_000 : ms, ...rest)) as typeof window.setTimeout;
    });
    await shapes(page).first().click();
    await page.keyboard.type("Kapak");
    await leaveText(page);
    for (let i = 2; i <= 7; i++) {
      await page.keyboard.press("Control+m");
      await shapes(page).first().click();
      await page.keyboard.type(`Konu ${i}`);
      await leaveText(page);
    }
    await page.keyboard.press("Control+p");
    await page.getByLabel("Yazdırma düzeni").selectOption("handouts4:horz");
    await page.getByLabel("Renk").selectOption("gray");
    await page.emulateMedia({ media: "print" });
    await page.locator(".su-printpage-btn", { hasText: "Yazdır" }).click();
    await expect(page.locator(".print-root .pp-page.is-handout.is-gray")).toHaveCount(2);
    const pdf = await page.pdf({ preferCSSPageSize: true, printBackground: true });
    expect(await page.evaluate(() => (window as unknown as { __printed: number }).__printed)).toBe(1);
    expect(pdf.toString("latin1").match(/\/Type\s*\/Page[^s]/g)?.length).toBe(2);
  });

  test("H10: Ctrl+sürükle ile bırakılan kopya tek geri alma adımıdır", async ({ page }) => {
    await blankSlide(page);
    await drawRect(page, 100, 100, 200, 150);
    const b = (await shapes(page).first().boundingBox())!;
    await page.keyboard.down("Control");
    await page.mouse.move(b.x + 20, b.y + 20);
    await page.mouse.down();
    await page.mouse.move(b.x + 320, b.y + 180, { steps: 8 });
    await page.mouse.up();
    await page.keyboard.up("Control");
    await expect(shapes(page)).toHaveCount(2);
    const copy = (await shapes(page).nth(1).boundingBox())!;
    expect(copy.x - b.x).toBeGreaterThan(250);
    await page.keyboard.press("Control+z");
    await expect(shapes(page)).toHaveCount(1);
    expect((await shapes(page).first().boundingBox())!.x).toBeCloseTo(b.x, 0);
    await page.keyboard.press("Control+y");
    await expect(shapes(page)).toHaveCount(2);
  });

  test("Dışarı Aktar › PNG: geçerli slayt PNG, tüm slaytlar ZIP olarak iner", async ({ page }) => {
    await shapes(page).first().click();
    await page.keyboard.type("PNG");
    await leaveText(page);
    await page.keyboard.press("Control+m");
    await openBackstage(page, "Dışarı Aktar");
    const [one] = await Promise.all([page.waitForEvent("download"), page.locator(".oh-bs-action", { hasText: "Geçerli Slaytı Resim Olarak Kaydet" }).click()]);
    expect(one.suggestedFilename()).toBe("Sunu1-Slayt2.png");
    expect(readFileSync((await one.path())!).subarray(1, 4).toString()).toBe("PNG");
    await openBackstage(page, "Dışarı Aktar");
    const [all] = await Promise.all([page.waitForEvent("download"), page.locator(".oh-bs-action", { hasText: "Tüm Slaytları Resim Olarak Kaydet" }).click()]);
    const zip = readFileSync((await all.path())!);
    expect(zip.subarray(0, 2).toString()).toBe("PK");
    expect(zip.includes(Buffer.from("Slayt2.png"))).toBe(true);
  });
});
