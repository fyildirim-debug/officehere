import { expect, test, type Page } from "@playwright/test";
import { mkdtempSync, readFileSync, writeFileSync } from "node:fs";
import { createRequire } from "node:module";
import { tmpdir } from "node:os";
import { join, resolve } from "node:path";
import { URL, collectErrors, interceptFileChooser } from "./ortak";

// OfficeHere Tablo dosya G/Ç sağlamlığı: tablolu (ListObject) XLSX'i düzenleyip kaydetme ve
// yeniden açma, bozuk dosyada Türkçe ileti + ayrıntılar, eksik yapılı dosyanın kurtarılması,
// dev birleştirmeli dosyanın takılmadan açılıp kaydedilmesi.

const FIX = resolve(__dirname, "..", "apps/tablo/test/fixtures");
type Zip = { file(name: string): { async(t: "string"): Promise<string> } | null; file(name: string, data: string): void; remove(name: string): void; generateAsync(o: { type: "nodebuffer"; compression: "DEFLATE" }): Promise<Buffer> };
const JSZip = createRequire(resolve(__dirname, "..", "apps/tablo/package.json"))("jszip") as { loadAsync(data: Buffer): Promise<Zip> };

type Cell = { value: string | number | boolean | null; formula: string | null };
type Api = {
  cell(ref: string, sheet?: string): Cell | null;
  cellRect(ref: string): unknown;
  sheetNames(): string[];
  tables(sheet?: string): { name: string; ref: string; columns: string[] }[];
  loadWarnings(): string[];
};
const ready = (page: Page) => page.evaluate(() => !!(window as unknown as { __ohTablo?: Api }).__ohTablo?.cellRect("A1"));
const cell = (page: Page, ref: string, sheet?: string) => page.evaluate(([r, s]) => (window as unknown as { __ohTablo: Api }).__ohTablo.cell(r!, s ?? undefined), [ref, sheet ?? null] as const);
const tables = (page: Page) => page.evaluate(() => (window as unknown as { __ohTablo: Api }).__ohTablo.tables());
const sheetNames = (page: Page) => page.evaluate(() => (window as unknown as { __ohTablo: Api }).__ohTablo.sheetNames());

async function openFile(page: Page, file: string) {
  await page.locator(".oh-ribbon-file").click();
  await page.locator(".oh-bs-item", { hasText: "Aç" }).first().click();
  const [chooser] = await Promise.all([page.waitForEvent("filechooser"), page.locator(".oh-bs-action", { hasText: "Gözat" }).first().click()]);
  await chooser.setFiles(file);
}

async function focusGrid(page: Page) {
  await page.locator("canvas").first().click({ position: { x: 400, y: 400 } });
}

/** Fikstürün değiştirilmiş kopyasını geçici klasöre yazar. */
async function variant(name: string, out: string, edits: Record<string, (xml: string) => string | null>): Promise<string> {
  const z = await JSZip.loadAsync(readFileSync(join(FIX, name)));
  for (const [path, fn] of Object.entries(edits)) {
    const next = fn((await z.file(path)?.async("string")) ?? "");
    if (next === null) z.remove(path);
    else z.file(path, next);
  }
  const file = join(mkdtempSync(join(tmpdir(), "oh-tablo-io-")), out);
  writeFileSync(file, await z.generateAsync({ type: "nodebuffer", compression: "DEFLATE" }));
  return file;
}

test.describe("OfficeHere Tablo dosya G/Ç sağlamlığı", () => {
  let errors: string[] = [];
  test.beforeEach(async ({ page }) => {
    errors = collectErrors(page);
    interceptFileChooser(page);
    await page.addInitScript(() => {
      const w = window as unknown as Record<string, unknown>;
      delete w.showSaveFilePicker;
      delete w.showOpenFilePicker;
    });
    await page.goto(URL.tablo);
    await expect.poll(() => ready(page)).toBe(true);
  });
  test.afterEach(() => {
    expect(errors).toEqual([]);
  });

  test("tablolu LibreOffice XLSX: hücre değiştir → kaydet → yeniden açılır, tablo sütunları korunur", async ({ page }) => {
    await openFile(page, join(FIX, "lo-tablo.xlsx"));
    await expect.poll(() => cell(page, "A2").then((c) => c?.value)).toBe("Ürün");
    expect((await tables(page))[0]).toMatchObject({ name: "Veri", ref: "A2:H52" });
    expect((await tables(page))[0].columns).toHaveLength(8);
    const box = page.getByRole("textbox", { name: "Ad Kutusu" });
    await box.fill("C5");
    await box.press("Enter");
    await page.keyboard.type("99");
    await page.keyboard.press("Enter");
    await expect.poll(() => cell(page, "C5").then((c) => c?.value)).toBe(99);
    await focusGrid(page);
    const [dl] = await Promise.all([page.waitForEvent("download"), page.keyboard.press("Control+s")]);
    expect(dl.suggestedFilename()).toBe("lo-tablo.xlsx");
    const file = join(mkdtempSync(join(tmpdir(), "oh-tablo-io-")), "lo-tablo.xlsx");
    await dl.saveAs(file);
    await page.reload();
    await expect.poll(() => ready(page)).toBe(true);
    await openFile(page, file);
    await expect.poll(() => cell(page, "C5").then((c) => c?.value)).toBe(99);
    await expect(page.getByRole("dialog")).toHaveCount(0);
    const t = (await tables(page))[0];
    expect(t.columns).toEqual(["Ürün", "Şehir", "Adet", "Fiyat", "Tarih", "Tutar", "KDV", "Durum"]);
    expect(await sheetNames(page)).toEqual(["Satışlar", "Özet", "Zincir"]);
  });

  test("bozuk (yarım indirilmiş) dosya: Türkçe ileti, teknik ayrıntı ayrı ve kapalı", async ({ page }) => {
    const b = readFileSync(join(FIX, "lo-tablo.xlsx"));
    const file = join(mkdtempSync(join(tmpdir(), "oh-tablo-io-")), "yarim.xlsx");
    writeFileSync(file, b.subarray(0, b.length >> 1));
    await openFile(page, file);
    const dialog = page.getByRole("dialog");
    await expect(dialog).toContainText('"yarim.xlsx" dosyası açılamadı.');
    await expect(dialog).toContainText("Dosya bozuk ya da eksik indirilmiş");
    const message = dialog.locator(".oh-dialog-message");
    await expect(message).not.toContainText(/zip|Corrupted|undefined/i);
    const detail = dialog.getByTestId("alert-detail");
    await expect(detail.locator("pre")).toBeHidden();
    await detail.locator("summary", { hasText: "Ayrıntılar" }).click();
    await expect(detail.locator("pre")).toBeVisible();
    await expect(detail.locator("pre")).not.toBeEmpty();
    await dialog.getByRole("button", { name: "Tamam" }).click();
    await expect(dialog).toHaveCount(0);
  });

  test("kitap yapısı (workbook.xml) eksik dosya boş açılmaz: sayfalar kurtarılır ve uyarılır", async ({ page }) => {
    const file = await variant("lo-tablo.xlsx", "yapisiz.xlsx", { "xl/workbook.xml": () => null });
    await openFile(page, file);
    await expect.poll(() => cell(page, "A2").then((c) => c?.value)).toBe("Ürün");
    expect(await sheetNames(page)).toEqual(["Sayfa1", "Sayfa2", "Sayfa3"]);
    await expect(page.getByText(/kurtarıldı/).first()).toBeVisible();
    expect((await page.evaluate(() => (window as unknown as { __ohTablo: Api }).__ohTablo.loadWarnings())).length).toBeGreaterThan(0);
  });

  test("tüm sayfayı kaplayan birleştirme (A1:XFD1048576) takılmadan açılır ve kaydedilir", async ({ page }) => {
    const file = await variant("lo-tablo.xlsx", "birlesik.xlsx", { "xl/worksheets/sheet3.xml": (x) => x.replace("</sheetData>", '</sheetData><mergeCells count="1"><mergeCell ref="A1:XFD1048576"/></mergeCells>') });
    const t0 = Date.now();
    await openFile(page, file);
    await expect.poll(() => cell(page, "A2").then((c) => c?.value), { timeout: 15_000 }).toBe("Ürün");
    expect(Date.now() - t0).toBeLessThan(15_000);
    await focusGrid(page);
    const [dl] = await Promise.all([page.waitForEvent("download", { timeout: 15_000 }), page.keyboard.press("Control+s")]);
    expect(dl.suggestedFilename()).toBe("birlesik.xlsx");
  });
});
