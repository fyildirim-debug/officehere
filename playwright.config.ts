import { defineConfig, devices } from "@playwright/test";

// Uçtan uca testler derlenmiş uygulamalar üzerinde çalışır (pnpm build sonrası).
// Önizleme sunucuları geliştirme portlarıyla çakışmasın diye 417x portlarını kullanır.
// OH_E2E_APPS=belge,tablo ile yalnız istenen uygulamaların sunucuları başlatılır.
const only = process.env.OH_E2E_APPS?.split(",").map((s) => s.trim()).filter(Boolean);
const apps = [
  { name: "belge", port: 4173 },
  { name: "tablo", port: 4174 },
  { name: "sunum", port: 4175 },
].filter((a) => !only?.length || only.includes(a.name));

export default defineConfig({
  testDir: "e2e",
  timeout: 60_000,
  expect: { timeout: 10_000 },
  fullyParallel: true,
  retries: process.env.CI ? 1 : 0,
  reporter: process.env.CI ? [["list"], ["html", { open: "never" }]] : "list",
  use: {
    ...devices["Desktop Chrome"],
    viewport: { width: 1400, height: 880 },
    locale: "tr-TR",
    timezoneId: "Europe/Istanbul",
    acceptDownloads: true,
  },
  webServer: apps.map((a) => ({
    command: `pnpm --filter @officehere/${a.name} exec vite preview --port ${a.port} --strictPort`,
    url: `http://localhost:${a.port}/`,
    reuseExistingServer: !process.env.CI,
    timeout: 60_000,
  })),
});
