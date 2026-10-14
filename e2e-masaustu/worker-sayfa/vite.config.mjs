// Masaüstü Worker duman testinin sayfası (worker.spec.ts derler). Uygulamalarla aynı ayarlar:
// file:// için göreli yol, ES modül Worker'ı, küçültülmüş çıktı (sınıf adları değişir).
import path from "node:path";
import { fileURLToPath } from "node:url";
import react from "@vitejs/plugin-react";

const dir = path.dirname(fileURLToPath(import.meta.url));
// React kökte kurulu değil; ui paketinin kopyası kullanılır (tek örnek).
const modules = path.resolve(dir, "../../packages/ui/node_modules");

export default {
  root: dir,
  base: "./",
  logLevel: "warn",
  plugins: [react()],
  resolve: { alias: { react: path.join(modules, "react"), "react-dom": path.join(modules, "react-dom") } },
  worker: { format: "es" },
  build: { target: "es2022", emptyOutDir: true },
};
