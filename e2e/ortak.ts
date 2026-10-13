import type { Page } from "@playwright/test";

const base = Number(process.env.OH_E2E_PORT ?? 4173);

export const URL = {
  belge: `http://localhost:${base}/`,
  tablo: `http://localhost:${base + 1}/`,
  sunum: `http://localhost:${base + 2}/`,
} as const;

/** Sayfa hatalarını ve konsol hatalarını toplar. */
export function collectErrors(page: Page): string[] {
  const errors: string[] = [];
  page.on("pageerror", (e) => errors.push(`pageerror: ${e.message}`));
  page.on("console", (m) => {
    if (m.type() === "error") errors.push(`console: ${m.text()}`);
  });
  return errors;
}

/**
 * Dosya seçici yakalamasını sayfanın başında açar. Yakalama tetikleyiciyle
 * (Ctrl+O, Aç düğmesi) aynı anda istenirse yük altında seçici yakalanmadan
 * açılabilir ve test takılır.
 */
export function interceptFileChooser(page: Page): void {
  page.on("filechooser", () => {});
}
