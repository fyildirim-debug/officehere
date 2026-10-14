import { expect, test } from "@playwright/test";
import { dosyaAc, dosyaAdi, karsilastir, sayfaHatalari, uygulamaAc } from "./ortak";

// Tablo görsel regresyonu: derlem dosyalarının ızgarası (hücreler, kenarlıklar, koşullu biçim, grafik ve
// şekil katmanı dahil). Çok sayfalı iki dosyada ikinci sayfa da (sağdan sola sayfa dahil) çekilir.

const DOSYALAR: [yol: string, sayfalar: string[]][] = [
  ["lo/lo-butce.xlsx", ["Satışlar", "Özet"]],
  ["lo/lo-butce.ods", ["Satışlar"]],
  ["lo/lo-butce.csv", ["lo-butce"]],
  ["py/py-tablo.xlsx", ["Veri"]],
  ["poi/ExcelTables.xlsx", ["ExcelTable"]],
  ["poi/NewStyleConditionalFormattings.xlsx", ["CF"]],
  ["poi/SimpleScatterChart.xlsx", ["Sheet1"]],
  ["poi/WithThreeCharts.xlsx", ["Sheet1"]],
  ["poi/comments.xlsx", ["Sheet1"]],
  ["poi/right-to-left.xlsx", ["English", "عربى"]],
  ["poi/StructuredReferences.xlsx", ["Formulas"]],
  ["poi/ExcelPivotTableSample.xlsx", ["Tabelle1"]],
];

for (const [yol, sayfalar] of DOSYALAR) {
  test(`Tablo: ${yol} (${sayfalar.join(", ")})`, async ({ page }) => {
    const hatalar = sayfaHatalari(page);
    await uygulamaAc(page, "tablo");
    await dosyaAc(page, "tablo", yol);
    for (const [i, ad] of sayfalar.entries()) {
      // Dosyanın kaydedilmiş etkin sayfası ilk sayfa olmayabilir (ExcelPivotTableSample).
      const sekme = page.locator(".tb-tabs-strip").getByRole("tab", { name: ad, exact: true });
      if ((await sekme.getAttribute("aria-selected")) !== "true") await sekme.click();
      await expect(sekme).toHaveAttribute("aria-selected", "true");
      await karsilastir(page, page.locator(".tb-canvas-wrap"), `tablo-${dosyaAdi(yol)}-${i + 1}`);
    }
    expect(hatalar).toEqual([]);
  });
}
