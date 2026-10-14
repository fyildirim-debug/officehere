import { configDefaults, defineConfig } from "vitest/config";
import react from "@vitejs/plugin-react";
import { lisansBildirimi } from "../../scripts/lisans-bildirimi.mjs";

export default defineConfig({
  plugins: [react(), lisansBildirimi()],
  // Masaüstü kabuğu dosyaları file:// ile yüklediği için göreli yol.
  base: "./",
  server: { port: 5174, strictPort: true },
  preview: { port: 5174, strictPort: true },
  build: { target: "es2022", chunkSizeWarningLimit: 4000 },
  // Performans testi süreye bağlıdır: paralel koşuda yanlış alarm vermesin diye burada dışarıda bırakılır,
  // ayrı ve sıralı olarak `pnpm test:perf` (vitest.perf.config.ts) ile koşar.
  test: { globals: true, environment: "node", exclude: [...configDefaults.exclude, "test/perf.test.ts"] },
});
