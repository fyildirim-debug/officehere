import { expect, test } from "@playwright/test";
import { dosyaAc, dosyaAdi, karsilastir, sayfaHatalari, uygulamaAc } from "./ortak";

// Sunum görsel regresyonu: derlem dosyalarının ilk 1-3 slaydı (düzenleme sahnesi). SmartArt içeren POI
// dosyaları bugünkü (SmartArt'sız) çizimi kayda geçirir; geçiş katmanı (dalga 3) temeli bilerek günceller.

const DOSYALAR: [yol: string, slayt: number][] = [
  ["lo/lo-sunu.pptx", 3],
  ["lo/lo-sunu.odp", 2],
  ["lo/lo-sunu24.pptx", 2],
  ["py/py-sunu.pptx", 3],
  ["poi/SampleShow.pptx", 2],
  ["poi/SmartArt.pptx", 1],
  ["poi/smartart-simple.pptx", 1],
  ["poi/backgrounds.pptx", 3],
  ["poi/bar-chart.pptx", 1],
  ["poi/layouts.pptx", 3],
  ["poi/pie-chart.pptx", 1],
  ["poi/shapes.pptx", 3],
  ["poi/table_test.pptx", 1],
];

for (const [yol, slayt] of DOSYALAR) {
  test(`Sunum: ${yol} ilk ${slayt} slayt`, async ({ page }) => {
    const hatalar = sayfaHatalari(page);
    await uygulamaAc(page, "sunum");
    await dosyaAc(page, "sunum", yol);
    const kucukler = page.locator(".thumb-pane .thumb-item");
    await expect.poll(() => kucukler.count()).toBeGreaterThanOrEqual(slayt);
    for (let i = 0; i < slayt; i++) {
      if (i > 0) await kucukler.nth(i).click();
      await expect(kucukler.nth(i)).toHaveClass(/\bis-current\b/);
      await karsilastir(page, page.locator(".stage"), `sunum-${dosyaAdi(yol)}-s${i + 1}`);
    }
    expect(hatalar).toEqual([]);
  });
}
