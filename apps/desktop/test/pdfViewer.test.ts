import { describe, expect, it, vi } from "vitest";
import { readFileSync } from "node:fs";
import path from "node:path";

// PDF görüntüleyicisinin yan etkisiz kuralları (Electron'suz).
vi.mock("electron", () => ({ BrowserWindow: class {}, WebContentsView: class {}, dialog: {}, ipcMain: { on: vi.fn() }, session: {}, shell: {} }));

const { hasPdfHeader, isPdfPath } = await import("../src/main/pdfViewer");

const ascii = (s: string) => new TextEncoder().encode(s);

describe("PDF görüntüleyici", () => {
  it("yalnız .pdf uzantısı (büyük-küçük harf fark etmez)", () => {
    expect(isPdfPath("/belgeler/Fatura.PDF")).toBe(true);
    expect(isPdfPath("C:\\Belgeler\\örnek belge.pdf")).toBe(true);
    expect(isPdfPath("/belgeler/rapor.docx")).toBe(false);
    expect(isPdfPath("/belgeler/pdf")).toBe(false);
    expect(isPdfPath("/belgeler/rapor.pdf.exe")).toBe(false);
  });

  it("PDF imzası ilk 1024 bayt içinde aranır", () => {
    expect(hasPdfHeader(ascii("%PDF-1.7\n%âãÏÓ"))).toBe(true);
    const junk = new Uint8Array(1024);
    junk.set(ascii("%PDF-1.4"), 600);
    expect(hasPdfHeader(junk)).toBe(true);
    const late = new Uint8Array(1100);
    late.set(ascii("%PDF-1.4"), 1020);
    expect(hasPdfHeader(late)).toBe(false);
    expect(hasPdfHeader(ascii("PK\u0003\u0004 docx"))).toBe(false);
    expect(hasPdfHeader(ascii("%PDF"))).toBe(false);
    expect(hasPdfHeader(new Uint8Array(0))).toBe(false);
  });

  it(".pdf için dosya ilişkilendirmesi yok (varsayılan PDF uygulaması olmaya aday değiliz)", () => {
    const yml = readFileSync(path.join(__dirname, "..", "electron-builder.yml"), "utf8");
    const exts = [...yml.matchAll(/^\s*-\s*ext:\s*(\S+)/gm)].map((m) => m[1].toLowerCase());
    expect(exts).toContain("docx");
    expect(exts).not.toContain("pdf");
    expect(yml).not.toMatch(/application\/pdf/i);
  });
});
