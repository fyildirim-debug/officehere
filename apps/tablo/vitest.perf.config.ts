import { configDefaults, defineConfig } from "vitest/config";

// Performans testleri (pnpm test:perf): yalnız test/perf.test.ts, tek iş parçacığında ve sıralı.
// Normal `pnpm test` bunları vite.config.ts'deki exclude ile atlar.
// `--mode temel` ile koşulunca ölçümler test/perf-temel.json'a yazılır (test:perf:temel).
export default defineConfig({
  test: {
    globals: true,
    environment: "node",
    include: ["test/perf.test.ts"],
    exclude: [...configDefaults.exclude],
    fileParallelism: false,
    maxWorkers: 1,
    testTimeout: 120_000,
  },
});
