import { expect, test } from "@playwright/test";
import { UZUN_PENCERE, dosyaAc, karsilastir, sayfaHatalari, uygulamaAc } from "./ortak";

// Karşılaştırıcının öz denetimi: içerik sayfa kutusu içinde 1 px kaydırılınca görüntü temelden ayrışmalı.
// Eşikler (maxDiffPixelRatio, threshold) gevşetilip gerçek kaymaları kaçıracak hale gelirse bu test kırılır.

test.use({ viewport: UZUN_PENCERE });

test("1 px kaydırma temel görüntüden ayrışır", async ({ page }) => {
  const hatalar = sayfaHatalari(page);
  await uygulamaAc(page, "belge");
  await dosyaAc(page, "belge", "lo/lo-rapor.docx");
  const sayfa = page.locator('.bl-page[data-page="0"]');
  await karsilastir(page, sayfa, "ozdenetim-belge-kaydirma");
  await page.addStyleTag({ content: ".bl-body .ProseMirror, .bl-hf { translate: 1px 0; }" });
  await expect(sayfa).not.toHaveScreenshot("ozdenetim-belge-kaydirma.png", { timeout: 5_000 });
  expect(hatalar).toEqual([]);
});
