import type { Page } from "@playwright/test";

export const URL = {
  belge: "http://localhost:4173/",
  tablo: "http://localhost:4174/",
  sunum: "http://localhost:4175/",
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
