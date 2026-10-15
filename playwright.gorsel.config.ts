import { defineConfig, devices } from "@playwright/test";

// Görsel regresyon (e2e/gorsel): derlenmiş uygulamaların ekran görüntüleri temel görüntülerle karşılaştırılır.
// Temeller yalnız Linux'ta, CI ile aynı imajda (mcr.microsoft.com/playwright:v1.64.0-noble + Carlito, Caladea,
// Liberation, DejaVu) üretilir: e2e/gorsel/__ekran__/linux/. Doğrudan değil `pnpm gorsel` /
// `pnpm gorsel:guncelle` ile çalıştırın (e2e/gorsel/calistir.mjs; ayrıntı: e2e/gorsel/__ekran__/README.md).
// OH_GORSEL_EKRAN: temel dizinini değiştirir (yerel belirlenimcilik denemesi, depoya yazılmaz).
const only = process.env.OH_E2E_APPS?.split(",").map((s) => s.trim()).filter(Boolean);
const base = Number(process.env.OH_E2E_PORT ?? 4173);
const apps = [
  { name: "belge", port: base },
  { name: "tablo", port: base + 1 },
  { name: "sunum", port: base + 2 },
].filter((a) => !only?.length || only.includes(a.name));
const ekran = process.env.OH_GORSEL_EKRAN ?? "e2e/gorsel/__ekran__";

export default defineConfig({
  testDir: "e2e/gorsel",
  snapshotPathTemplate: `${ekran}/{platform}/{arg}{ext}`,
  outputDir: "test-results/gorsel",
  timeout: 60_000,
  expect: {
    // Kararlı görüntü bekleme (iki ardışık aynı kare) yük altında uzayabilir.
    timeout: 15_000,
    // Etkin sınır min(40 piksel, %0,2): seyrek sayfa/slaytta (bir başlık ve alt başlık) 1 px kayma eşik üstü
    // yalnız ~100 piksel değiştirir ve yalnız oranla kaçardı. Çizim aynı imajda bayt bayt belirlenimci olduğu için
    // sıkı sınır yanlış alarm üretmez (threshold 0.1: ~26 gri düzeyinden küçük farklar zaten sayılmaz).
    toHaveScreenshot: { maxDiffPixelRatio: 0.002, maxDiffPixels: 40, threshold: 0.1, animations: "disabled", caret: "hide", scale: "css" },
  },
  fullyParallel: true,
  workers: process.env.CI ? 3 : undefined,
  // Görsel farkta yeniden deneme yanlış alarmı gizler; kararsızlık hata olarak görünmeli.
  retries: 0,
  reporter: process.env.CI ? [["list"], ["html", { open: "never", outputFolder: "playwright-report/gorsel" }]] : "list",
  use: {
    ...devices["Desktop Chrome"],
    viewport: { width: 1400, height: 880 },
    deviceScaleFactor: 1,
    locale: "tr-TR",
    timezoneId: "Europe/Istanbul",
    colorScheme: "light",
    contextOptions: { reducedMotion: "reduce" },
    launchOptions: {
      // OH_CHROMIUM: yalnız yerel deneme (temel görüntüler sabit imajdaki tarayıcıyla üretilir)
      ...(process.env.OH_CHROMIUM ? { executablePath: process.env.OH_CHROMIUM } : {}),
      args: [
        "--font-render-hinting=none",
        "--disable-lcd-text",
        "--force-color-profile=srgb",
        // Belirlenimci çizim: resimler eşzamanlı çözülür, kısmi ve çok iş parçacıklı raster yok.
        "--disable-checker-imaging",
        "--disable-image-animation-resync",
        "--disable-partial-raster",
        "--disable-threaded-animation",
        "--num-raster-threads=1",
      ],
    },
  },
  webServer: apps.map((a) => ({
    command: `pnpm --filter @officehere/${a.name} exec vite preview --port ${a.port} --strictPort`,
    url: `http://localhost:${a.port}/`,
    reuseExistingServer: !process.env.CI,
    timeout: 60_000,
  })),
});
