import { basename, resolve } from "node:path";
import { expect, type Locator, type Page } from "@playwright/test";
import { URL } from "../ortak";

// Görsel regresyon ortak yardımcıları. Uygulamalar tarayıcı kipinde (yerel köprü yok) açılır, dosyalar
// derlemden (test-derlem/) dosya seçiciyle yüklenir; görüntü yalnız sayfa/slayt/ızgara öğesinden alınır.
// Karşılaştırma ayarları playwright.gorsel.config.ts'de (maxDiffPixelRatio 0.002, maxDiffPixels 40, threshold 0.1).

export type Uygulama = "belge" | "tablo" | "sunum";
export type Tema = "colorful" | "dark";

export const DERLEM = resolve(__dirname, "../../test-derlem");

/** Belge sayfası (A4, 1123 px) kaydırma alanına bütünüyle sığsın diye uzun pencere. */
export const UZUN_PENCERE = { width: 1400, height: 1500 };

/**
 * Kasıtlı 1 px kaydırma (OH_GORSEL_KAYDIR=1): ekran görüntüsünden hemen önce içerik kendi kutusu içinde
 * 1 px sağa kaydırılır. Karşılaştırıcının bunu yakaladığını (içeriği olan her görüntünün kırmızıya döndüğünü)
 * yerelde denetlemek içindir; CI'da kullanılmaz.
 */
export const KAYDIR_CSS =
  "html, .bl-body .ProseMirror, .bl-hf, .tb-canvas-wrap > canvas, .stage-slide, .oh-ribbon-tablist, .oh-ribbon-content { translate: 1px 0 !important; }";

/** Sayfa hatalarını toplar (konsol iletileri değil: derlemde eksik resim gibi beklenen ağ hataları var). */
export function sayfaHatalari(page: Page): string[] {
  const hatalar: string[] = [];
  page.on("pageerror", (e) => hatalar.push(e.message));
  return hatalar;
}

/** Uygulamayı boş belgeyle açar; tema localStorage'dan, tarayıcı dosya API'leri kapalı (dosya girişi kullanılır). */
export async function uygulamaAc(page: Page, app: Uygulama, tema: Tema = "colorful"): Promise<void> {
  await page.addInitScript((t) => {
    const w = window as unknown as Record<string, unknown>;
    delete w.showSaveFilePicker;
    delete w.showOpenFilePicker;
    localStorage.setItem("officehere.uiTheme", t);
  }, tema);
  await page.goto(URL[app]);
  if (app === "belge") await expect(page.locator(".bl-body .ProseMirror")).toBeFocused();
  else if (app === "tablo") await expect.poll(() => page.evaluate(() => !!(window as unknown as { __ohTablo?: { cellRect(r: string): unknown } }).__ohTablo?.cellRect("A1"))).toBe(true);
  else await expect(page.locator(".stage .sv")).toBeVisible();
}

/** Derlem dosyasını (test-derlem/ altındaki göreli yol) uygulamada açar ve başlıkta adını bekler. */
export async function dosyaAc(page: Page, app: Uygulama, yol: string): Promise<void> {
  const dosya = resolve(DERLEM, yol);
  if (app === "belge") {
    await page.setInputFiles("#bl-open-input", dosya);
  } else if (app === "tablo") {
    await page.locator(".oh-ribbon-file").click();
    await page.locator(".oh-bs-item", { hasText: "Aç" }).first().click();
    const [secici] = await Promise.all([page.waitForEvent("filechooser"), page.locator(".oh-bs-action", { hasText: "Gözat" }).first().click()]);
    await secici.setFiles(dosya);
  } else {
    await page.locator(".stage").click({ position: { x: 4, y: 4 } });
    const [secici] = await Promise.all([page.waitForEvent("filechooser"), page.keyboard.press("Control+o")]);
    await secici.setFiles(dosya);
  }
  await expect(page.locator(".oh-titlebar")).toContainText(basename(yol).replace(/\.[^.]+$/, ""), { timeout: 30_000 });
}

/**
 * Yazı tipleri ve resimler yüklenene, ardından iki kare çizilene dek bekler. Resimler img.decode() ile
 * zorlanmaz: tam boyutta önden çözme, çizimdeki ölçekli çözmeyle ±1 renk düzeyi farklı sonuç verebiliyor.
 */
export async function oturt(page: Page): Promise<void> {
  await page.evaluate(async () => {
    await document.fonts.ready;
    const yuklendi = (i: HTMLImageElement) =>
      new Promise((r) => {
        i.addEventListener("load", r, { once: true });
        i.addEventListener("error", r, { once: true });
      });
    await Promise.all([...document.images].filter((i) => !i.complete).map(yuklendi));
    await new Promise((r) => requestAnimationFrame(() => requestAnimationFrame(r)));
  });
}

/** Fareyi ipucu ve üzerine gelme vurgusu oluşturmayacak bir yere çeker. */
export async function fareyiCek(page: Page): Promise<void> {
  await page.mouse.move(1, (page.viewportSize()?.height ?? 880) - 1);
  await expect(page.locator(".oh-tooltip")).toHaveCount(0);
}

/** Öğenin (ya da sayfanın) görüntüsünü temel görüntüyle karşılaştırır: e2e/gorsel/__ekran__/linux/<ad>.png */
export async function karsilastir(page: Page, hedef: Locator | Page, ad: string): Promise<void> {
  await oturt(page);
  if (!process.env.OH_GORSEL_KAYDIR) return expect(hedef).toHaveScreenshot(`${ad}.png`);
  // Kaydırma denemesinde test ilk farkta durmasın: her görüntünün ayrı ayrı yakalandığı görülsün.
  await page.addStyleTag({ content: KAYDIR_CSS });
  await expect.soft(hedef).toHaveScreenshot(`${ad}.png`);
}

/** Dosya yolundan temel görüntü adı: "poi/SampleShow.pptx" → "SampleShow-pptx". */
export function dosyaAdi(yol: string): string {
  return basename(yol).replace(/\.([^.]+)$/, "-$1");
}
