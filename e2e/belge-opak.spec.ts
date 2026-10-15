import { readFileSync } from "node:fs";
import { resolve } from "node:path";
import { expect, test, type Page } from "@playwright/test";
import { URL, collectErrors, zipEntry } from "./ortak";

// OfficeHere Belge geçiş katmanı (öğe düzeyi): gömülü OLE nesnesi gibi desteklenmeyen öğeler yer tutucu kart
// olarak görünür ve .docx'e kaydederken korunur; kopyala/yapıştır ve kes/yapıştır ile çoğaltılan öğe kendi
// parça kopyalarını alır (kimlikler yazarken yenilenir); .odt'ye dışarı aktarmada yer tutucu metin yazılır ve
// uyarılır.

const POI = resolve(__dirname, "../test-derlem/poi");
const poi = (name: string) => readFileSync(resolve(POI, name));
const CARD = "Desteklenmeyen öğe: Gömülü nesne (hesap tablosu) — kaydederken korunur";

test.describe("OfficeHere Belge desteklenmeyen öğeler", () => {
  let errors: string[] = [];

  test.beforeEach(async ({ page }) => {
    errors = collectErrors(page);
    await page.addInitScript(() => {
      const w = window as unknown as Record<string, unknown>;
      delete w.showSaveFilePicker;
      delete w.showOpenFilePicker;
    });
    await page.goto(URL.belge);
    await expect(body(page)).toBeVisible();
  });

  test.afterEach(() => {
    expect(errors).toEqual([]);
  });

  const body = (page: Page) => page.locator(".bl-body .ProseMirror");
  const cards = (page: Page) => body(page).locator(".bl-opaque");
  const open = async (page: Page, name: string, buffer: Buffer) => {
    await page.setInputFiles("#bl-open-input", { name, mimeType: "application/octet-stream", buffer });
    await expect(page.locator(".oh-titlebar")).toContainText(name.replace(/\.[^.]+$/, ""));
  };
  /** Ctrl+S ile indirilen .docx'in baytları. */
  const save = async (page: Page, out: string) => {
    const download = page.waitForEvent("download");
    await page.keyboard.press("Control+s");
    const file = await download;
    await file.saveAs(out);
    return readFileSync(out);
  };
  /** Kaydedilen belgedeki gömülü nesnelerin parça hedefleri (her biri pakette olmalı). */
  const objectTargets = (bytes: Buffer) => {
    const doc = zipEntry(bytes, "word/document.xml") ?? "";
    const rels = zipEntry(bytes, "word/_rels/document.xml.rels") ?? "";
    const target = (id: string) => new RegExp(`Id="${id}"[^>]*Target="([^"]+)"`).exec(rels)?.[1] ?? null;
    return [...doc.matchAll(/<o:OLEObject\b[^>]*\sr:id="([^"]+)"/g)].map((m) => target(m[1]!));
  };

  test("gömülü nesne kart olarak görünür; kopyala/yapıştırla çoğaltılır, her kopya kendi nesnesiyle kaydedilir", async ({ page }, testInfo) => {
    await open(page, "Gomulu.docx", poi("EmbeddedDocument.docx"));
    await expect(cards(page)).toHaveCount(1);
    await expect(cards(page).first()).toHaveAttribute("role", "img");
    await expect(cards(page).first()).toHaveAttribute("aria-label", CARD);
    // Nesne korunduğu için "kaydederseniz kaybolur" uyarısı yok.
    await expect(page.locator(".oh-messagebar", { hasText: "gömülü nesneler" })).toHaveCount(0);

    await cards(page).first().click();
    await expect(cards(page).first()).toHaveClass(/ProseMirror-selectednode/);
    await page.keyboard.press("Control+c");
    await page.keyboard.press("Control+End");
    await page.keyboard.press("Enter");
    await page.keyboard.press("Control+v");
    await page.keyboard.type(" ara ");
    await page.keyboard.press("Control+v");
    await expect(cards(page)).toHaveCount(3);
    for (const c of await cards(page).all()) await expect(c).toHaveAttribute("aria-label", CARD);

    const bytes = await save(page, testInfo.outputPath("Gomulu.docx"));
    const targets = objectTargets(bytes);
    expect(targets).toHaveLength(3);
    expect(new Set(targets).size).toBe(3);
    for (const t of targets) expect(zipEntry(bytes, `word/${t}`)).not.toBeNull();
    const doc = zipEntry(bytes, "word/document.xml")!;
    // VML şekil kimlikleri belgede benzersiz.
    const shapes = [...doc.matchAll(/<v:shape id="([^"]+)"/g)].map((m) => m[1]);
    expect(new Set(shapes).size).toBe(3);
    expect(doc).toContain(" ara ");
  });

  test("kes/yapıştır nesneyi taşır; geri al öğeyi geri getirir", async ({ page }, testInfo) => {
    await open(page, "Tasi.docx", poi("EmbeddedDocument.docx"));
    await expect(cards(page)).toHaveCount(1);
    await cards(page).first().click();
    await page.keyboard.press("Control+x");
    await expect(cards(page)).toHaveCount(0);
    await page.keyboard.press("Control+Home");
    await page.keyboard.press("Control+v");
    await expect(cards(page)).toHaveCount(1);
    // İlk paragrafa taşındı.
    await expect(body(page).locator("p").first().locator(".bl-opaque")).toHaveCount(1);
    const bytes = await save(page, testInfo.outputPath("Tasi.docx"));
    expect(objectTargets(bytes)).toHaveLength(1);
    await page.keyboard.press("Control+z");
    await page.keyboard.press("Control+z");
    await expect(cards(page)).toHaveCount(1);
  });

  test("Dışarı Aktar › .odt: nesnenin yerine yer tutucu metin yazılır ve uyarılır", async ({ page }, testInfo) => {
    await open(page, "Aktar.docx", poi("EmbeddedDocument.docx"));
    await expect(cards(page)).toHaveCount(1);
    await page.locator(".oh-ribbon-file").click();
    await page.locator(".oh-bs-item", { hasText: "Dışarı Aktar" }).first().click();
    const download = page.waitForEvent("download");
    await page.locator(".oh-bs-action", { hasText: "OpenDocument Metin (.odt)" }).click();
    const file = await download;
    const path = testInfo.outputPath(file.suggestedFilename());
    await file.saveAs(path);
    expect(zipEntry(readFileSync(path), "content.xml")).toContain("[Desteklenmeyen öğe: Gömülü nesne (hesap tablosu)]");
    await expect(page.locator(".oh-messagebar")).toContainText("yer tutucu metin yazıldı");
  });
});
