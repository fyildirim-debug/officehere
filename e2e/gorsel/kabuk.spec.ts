import { expect, test } from "@playwright/test";
import { fareyiCek, karsilastir, sayfaHatalari, uygulamaAc, type Tema, type Uygulama } from "./ortak";

// Kabuk görsel regresyonu: her uygulamada boş belgenin tam penceresi (açık ve koyu tema) ve üç şerit
// durumu (Giriş, Ekle, 900 px dar pencerede daralmış gruplar).
// Fare ipucu ve üzerine gelme vurgusu oluşturmasın diye her görüntüden önce köşeye çekilir.

const UYGULAMALAR: Uygulama[] = ["belge", "tablo", "sunum"];
const TEMALAR: [Tema, string][] = [
  ["colorful", "acik"],
  ["dark", "koyu"],
];

for (const app of UYGULAMALAR) {
  for (const [tema, ad] of TEMALAR) {
    test(`${app}: boş belge, ${ad} tema`, async ({ page }) => {
      const hatalar = sayfaHatalari(page);
      await uygulamaAc(page, app, tema);
      await fareyiCek(page);
      await karsilastir(page, page, `kabuk-${app}-${ad}`);
      expect(hatalar).toEqual([]);
    });
  }

  test(`${app}: şerit, Giriş ve Ekle sekmesi`, async ({ page }) => {
    const hatalar = sayfaHatalari(page);
    const serit = page.locator(".oh-ribbon");
    const sekme = (ad: string) => serit.getByRole("tab", { name: ad, exact: true });
    await uygulamaAc(page, app);
    await fareyiCek(page);
    await expect(sekme("Giriş")).toHaveAttribute("aria-selected", "true");
    await karsilastir(page, serit, `serit-${app}-giris`);

    await sekme("Ekle").click();
    await expect(sekme("Ekle")).toHaveAttribute("aria-selected", "true");
    await fareyiCek(page);
    await karsilastir(page, serit, `serit-${app}-ekle`);
    expect(hatalar).toEqual([]);
  });

  test.describe(() => {
    // Pencere baştan dar açılır: açıkken boyutlandırmak yuvarlak köşelerde bayt düzeyinde çizim farkı bırakıyor.
    test.use({ viewport: { width: 900, height: 700 } });
    test(`${app}: şerit, 900 px dar pencere (daralmış gruplar)`, async ({ page }) => {
      const hatalar = sayfaHatalari(page);
      const serit = page.locator(".oh-ribbon");
      await uygulamaAc(page, app);
      await expect(serit.locator(".oh-ribbon-group.is-collapsed").first()).toBeVisible();
      await fareyiCek(page);
      await karsilastir(page, serit, `serit-${app}-dar`);
      expect(hatalar).toEqual([]);
    });
  });
}
