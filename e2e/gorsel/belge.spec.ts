import { expect, test } from "@playwright/test";
import { UZUN_PENCERE, dosyaAc, dosyaAdi, karsilastir, sayfaHatalari, uygulamaAc } from "./ortak";

// Belge görsel regresyonu: derlem dosyalarının ilk 1-3 sayfası. Sayfa sayısı sabit (sayfalama tamamlanınca
// en az bu kadar sayfa olmalı), böylece her koşu aynı temel görüntüleri üretir.

const DOSYALAR: [yol: string, sayfa: number][] = [
  ["lo/lo-rapor.docx", 3],
  ["lo/lo-karma.docx", 3],
  ["lo/lo-uzun.docx", 2],
  ["lo/lo-karma.odt", 2],
  ["lo/lo-rapor.rtf", 2],
  ["py/py-belge.docx", 3],
  ["poi/Headers.docx", 3],
  ["poi/Numbering.docx", 1],
  ["poi/TestTableColumns.docx", 1],
  ["poi/WithGIF.docx", 1],
  ["poi/checkboxes.docx", 1],
  ["poi/footnotes.docx", 1],
  ["lo/lo-karma.doc", 1],
  ["poi/delins.docx", 1],
];

test.use({ viewport: UZUN_PENCERE });

for (const [yol, sayfa] of DOSYALAR) {
  test(`Belge: ${yol} ilk ${sayfa} sayfa`, async ({ page }) => {
    const hatalar = sayfaHatalari(page);
    await uygulamaAc(page, "belge");
    await dosyaAc(page, "belge", yol);
    await expect.poll(() => page.locator(".bl-page").count()).toBeGreaterThanOrEqual(sayfa);
    for (let i = 0; i < sayfa; i++) {
      await karsilastir(page, page.locator(`.bl-page[data-page="${i}"]`), `belge-${dosyaAdi(yol)}-s${i + 1}`);
    }
    expect(hatalar).toEqual([]);
  });
}
