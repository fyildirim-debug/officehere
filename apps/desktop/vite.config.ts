import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";
import { lisansBildirimi } from "../../scripts/lisans-bildirimi.mjs";

// Başlangıç ekranı (launcher). Masaüstünde dist-electron/launcher'dan, web'de site kökünden sunulur.
export default defineConfig({
  root: "src/launcher",
  plugins: [react(), lisansBildirimi()],
  base: "./",
  server: { port: 5176, strictPort: true },
  build: { outDir: "../../dist-electron/launcher", emptyOutDir: true, target: "es2022" },
});
