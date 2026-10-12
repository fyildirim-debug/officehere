import { defineConfig } from "@playwright/test";

// Masaüstü (Electron) uçtan uca testleri. Önce: pnpm web && pnpm --filter @officehere/desktop build:electron
// Linux'ta ekransız ortamda: xvfb-run -a pnpm e2e:masaustu
export default defineConfig({
  testDir: "e2e-masaustu",
  timeout: 90_000,
  expect: { timeout: 15_000 },
  workers: 1,
  retries: process.env.CI ? 1 : 0,
  reporter: "list",
});
