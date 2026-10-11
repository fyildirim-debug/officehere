import { defineConfig } from "vitest/config";
import react from "@vitejs/plugin-react";

export default defineConfig({
  plugins: [react()],
  // Masaüstü kabuğu dosyaları file:// ile yüklediği için göreli yol.
  base: "./",
  server: { port: 5174, strictPort: true },
  preview: { port: 5174, strictPort: true },
  build: { target: "es2022", chunkSizeWarningLimit: 4000 },
  test: { globals: true, environment: "node" },
});
