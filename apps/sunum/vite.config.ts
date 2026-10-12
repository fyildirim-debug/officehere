import { defineConfig } from "vitest/config";
import react from "@vitejs/plugin-react";
import { lisansBildirimi } from "../../scripts/lisans-bildirimi.mjs";

export default defineConfig({
  plugins: [react(), lisansBildirimi()],
  // Masaüstü kabuğu dosyaları file:// ile yüklediği için göreli yol.
  base: "./",
  server: { port: 5175, strictPort: true },
  preview: { port: 5175, strictPort: true },
  build: { target: "es2022", chunkSizeWarningLimit: 4000 },
  test: { globals: true, environment: "node" },
});
