import { inflateRawSync } from "node:zlib";
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

/** ZIP paketinden (DOCX/XLSX/PPTX) bir parçayı metin olarak okur (merkezi dizin + inflateRaw); yoksa null. */
export function zipEntry(b: Buffer, name: string): string | null {
  const e = b.lastIndexOf(Buffer.from([0x50, 0x4b, 0x05, 0x06]));
  for (let i = 0, p = b.readUInt32LE(e + 16), n = b.readUInt16LE(e + 10); i < n; i++, p += 46 + b.readUInt16LE(p + 28) + b.readUInt16LE(p + 30) + b.readUInt16LE(p + 32)) {
    if (b.toString("utf8", p + 46, p + 46 + b.readUInt16LE(p + 28)) !== name) continue;
    const l = b.readUInt32LE(p + 42);
    const start = l + 30 + b.readUInt16LE(l + 26) + b.readUInt16LE(l + 28);
    const data = b.subarray(start, start + b.readUInt32LE(p + 20));
    return (b.readUInt16LE(p + 10) === 8 ? inflateRawSync(data) : data).toString("utf8");
  }
  return null;
}
