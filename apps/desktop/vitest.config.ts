import { defineConfig } from "vitest/config";

// Ana süreç birim testleri (Vite kökü başlangıç ekranı olduğu için ayrı yapılandırma).
export default defineConfig({
  test: { include: ["test/**/*.test.ts"], environment: "node" },
});
