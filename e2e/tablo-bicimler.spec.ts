import { expect, test, type Page } from "@playwright/test";
import { mkdtempSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join, resolve } from "node:path";
import { URL, collectErrors } from "./ortak";

// OfficeHere Tablo dosya biçimleri: .ods/.xls/.xltx açma, biçim sorusu, .ods dışarı aktarma
// ve indirilen dosyayı yeniden açma. Değerler window.__ohTablo test arayüzünden okunur.

const FIX = resolve(__dirname, "..", "apps/tablo/src/io/__fixtures__");

type Cell = { value: string | number | boolean | null; formula: string | null; bold: boolean };
type Api = { cell(ref: string, sheet?: string): Cell | null; cellRect(ref: string): unknown; sheetNames(): string[] };
const api = (page: Page) => page.evaluate(() => !!(window as unknown as { __ohTablo?: Api }).__ohTablo?.cellRect("A1"));
const cell = (page: Page, ref: string, sheet?: string) => page.evaluate(([r, s]) => (window as unknown as { __ohTablo: Api }).__ohTablo.cell(r!, s ?? undefined), [ref, sheet ?? null] as const);
const sheetNames = (page: Page) => page.evaluate(() => (window as unknown as { __ohTablo: Api }).__ohTablo.sheetNames());

/** Dosya › Aç › Gözat ile dosya açar. */
async function openFile(page: Page, file: string) {
  await page.locator(".oh-ribbon-file").click();
  await page.locator(".oh-bs-item", { hasText: "Aç" }).first().click();
  const [chooser] = await Promise.all([page.waitForEvent("filechooser"), page.locator(".oh-bs-action", { hasText: "Gözat" }).first().click()]);
  await chooser.setFiles(file);
}

/** Izgaraya odak verir (klavye kısayolları için). */
async function focusGrid(page: Page) {
  await page.locator("canvas").first().click({ position: { x: 400, y: 400 } });
}

test.describe("OfficeHere Tablo dosya biçimleri", () => {
  let errors: string[] = [];
  test.beforeEach(async ({ page }) => {
    errors = collectErrors(page);
    await page.addInitScript(() => {
      const w = window as unknown as Record<string, unknown>;
      delete w.showSaveFilePicker;
      delete w.showOpenFilePicker;
    });
    await page.goto(URL.tablo);
    await expect.poll(() => api(page)).toBe(true);
  });
  test.afterEach(() => {
    expect(errors).toEqual([]);
  });

  test("ODS açılır: sayfalar, değerler ve formül sonuçları görünür", async ({ page }) => {
    await openFile(page, join(FIX, "ornek.ods"));
    await expect.poll(() => cell(page, "A1").then((c) => c?.value)).toBe("Ürün");
    expect(await sheetNames(page)).toEqual(["Satışlar", "Özet", "Uzun", "Gizli"]);
    expect(await cell(page, "D5")).toMatchObject({ value: 234.25, formula: "=SUM(D2:D4)" });
    expect(await cell(page, "C6")).toMatchObject({ value: "Büyük", formula: '=IF(D5>100,"Büyük","Küçük")' });
    expect(await cell(page, "B1", "Özet")).toMatchObject({ value: 234.25, formula: "=Satışlar!D5" });
    expect((await cell(page, "A1"))!.bold).toBe(true);
    await expect(page.locator(".oh-titlebar")).toContainText("ornek");
  });

  test("XLS açılır: uyumluluk iletisi; Kaydet .xlsx olarak Farklı Kaydet yapar", async ({ page }) => {
    await openFile(page, join(FIX, "ornek.xls"));
    await expect.poll(() => cell(page, "A1").then((c) => c?.value)).toBe("Ürün");
    expect(await cell(page, "D2")).toMatchObject({ value: 25, formula: "=B2*C2" });
    expect(await cell(page, "E7")).toMatchObject({ value: -1, formula: "=IFERROR(1/0,-1)" });
    expect(await cell(page, "A300", "Uzun")).toMatchObject({ value: "Satır 300 — çğıöşü ÇĞİÖŞÜ uzun metin 2100" });
    await expect(page.getByText(/Uyumluluk modu/)).toBeVisible();
    await focusGrid(page);
    const [dl] = await Promise.all([page.waitForEvent("download"), page.keyboard.press("Control+s")]);
    expect(dl.suggestedFilename()).toBe("ornek.xlsx");
    await expect(page.getByText(/Uyumluluk modu/)).toHaveCount(0);
  });

  test("ODS açıkken Kaydet: biçim bir kez sorulur, ODS olarak kaydedilir ve yeniden açılır", async ({ page }) => {
    await openFile(page, join(FIX, "ornek.ods"));
    await expect.poll(() => cell(page, "A1").then((c) => c?.value)).toBe("Ürün");
    await focusGrid(page);
    await page.keyboard.press("Control+s");
    const dialog = page.getByRole("dialog");
    await expect(dialog).toContainText("OpenDocument");
    const [dl] = await Promise.all([page.waitForEvent("download"), dialog.getByRole("button", { name: "ODS olarak kaydet" }).click()]);
    expect(dl.suggestedFilename()).toBe("ornek.ods");
    const file = join(mkdtempSync(join(tmpdir(), "oh-ods-")), "ornek.ods");
    await dl.saveAs(file);
    // İkinci kayıtta soru yok
    await focusGrid(page);
    const [dl2] = await Promise.all([page.waitForEvent("download"), page.keyboard.press("Control+s")]);
    expect(dl2.suggestedFilename()).toBe("ornek.ods");
    await expect(page.getByRole("dialog")).toHaveCount(0);
    // İndirilen dosya yeniden açılır
    await page.reload();
    await expect.poll(() => api(page)).toBe(true);
    await openFile(page, file);
    await expect.poll(() => cell(page, "D5").then((c) => c?.value)).toBe(234.25);
    expect((await cell(page, "D5"))!.formula).toBe("=SUM(D2:D4)");
    expect(await sheetNames(page)).toEqual(["Satışlar", "Özet", "Uzun", "Gizli"]);
  });

  test("Dışarı Aktar › ODS: yeni kitap indirilir ve yeniden açılınca aynıdır", async ({ page }) => {
    const box = page.getByRole("textbox", { name: "Ad Kutusu" });
    for (const [ref, text] of [["A1", "Şehir"], ["B1", "Tutar"], ["A2", "İzmir"], ["B2", "12,5"], ["B3", "=TOPLA(B2;7,5)"]]) {
      await box.fill(ref);
      await box.press("Enter");
      await page.keyboard.type(text);
      await page.keyboard.press("Enter");
    }
    await expect.poll(() => cell(page, "B3").then((c) => c?.value)).toBe(20);
    await page.locator(".oh-ribbon-file").click();
    await page.locator(".oh-bs-item", { hasText: "Dışarı Aktar" }).first().click();
    const [dl] = await Promise.all([page.waitForEvent("download"), page.getByText("OpenDocument Hesap Tablosu (.ods)").click()]);
    expect(dl.suggestedFilename()).toBe("Kitap1.ods");
    const file = join(mkdtempSync(join(tmpdir(), "oh-ods-")), "Kitap1.ods");
    await dl.saveAs(file);
    // Dışarı aktarma kitabı kaydedilmiş yapmaz: Aç önce kaydetme sorusunu getirir
    await page.locator(".oh-bs-item", { hasText: "Aç" }).first().click();
    await page.locator(".oh-bs-action", { hasText: "Gözat" }).first().click();
    const [chooser] = await Promise.all([page.waitForEvent("filechooser"), page.getByRole("button", { name: "Kaydetme", exact: true }).click()]);
    await chooser.setFiles(file);
    await expect(page.locator(".oh-titlebar")).toContainText("Kitap1");
    await expect.poll(() => cell(page, "B3").then((c) => c?.formula)).toBe("=SUM(B2,7.5)");
    expect(await cell(page, "A2")).toMatchObject({ value: "İzmir" });
    expect(await cell(page, "B3")).toMatchObject({ value: 20 });
  });

  test("Şablon (.xltx) adsız kitap olarak açılır; Kaydet .xlsx önerir", async ({ page }) => {
    await openFile(page, join(FIX, "sablon.xltx"));
    await expect.poll(() => cell(page, "A1").then((c) => c?.value)).toBe("Fatura");
    await expect(page.locator(".oh-titlebar")).toContainText("sablon1");
    await focusGrid(page);
    const [dl] = await Promise.all([page.waitForEvent("download"), page.keyboard.press("Control+s")]);
    expect(dl.suggestedFilename()).toBe("sablon1.xlsx");
  });

  test("CSV açıkken Kaydet: yalnız etkin sayfa uyarısı; İptal kaydetmez", async ({ page }) => {
    const csv = join(mkdtempSync(join(tmpdir(), "oh-csv-")), "liste.csv");
    writeFileSync(csv, "Ad,Tutar\nÇay,12.5\n", "utf8");
    await openFile(page, csv);
    await expect.poll(() => cell(page, "B2").then((c) => c?.value)).toBe(12.5);
    await focusGrid(page);
    await page.keyboard.press("Control+s");
    const dialog = page.getByRole("dialog");
    await expect(dialog).toContainText("yalnız etkin sayfa");
    await dialog.getByRole("button", { name: "İptal" }).click();
    await expect(page.getByRole("dialog")).toHaveCount(0);
  });
});
