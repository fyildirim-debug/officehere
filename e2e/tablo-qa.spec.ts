import { expect, test, type Page } from "@playwright/test";
import { mkdtempSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { URL, collectErrors, interceptFileChooser } from "./ortak";

// Tablo 5. tur QA: şerit sekmesi odağı, Ctrl ile doldurma, Otomatik Doldurma Seçenekleri,
// hücre içi onay kutusu, Özel Yapıştır, Birleştir, yeni işlevler ve dinamik dizi kaydet-aç.

type Cell = { value: string | number | boolean | null; formula: string | null; edit: string; numFmt: string | null; bold: boolean; fill: string | null };
type Api = {
  cell(ref: string, sheet?: string): Cell | null;
  cellRect(ref: string): { x: number; y: number; w: number; h: number } | null;
  selection(): string;
  undoCount(): number;
  sheetNames(): string[];
};
const api = "__ohTablo";

async function open(page: Page) {
  await page.goto(URL.tablo);
  await page.waitForFunction((k) => !!(window as unknown as Record<string, Api | undefined>)[k]?.cellRect("A1"), api);
}
const cell = (page: Page, ref: string, sheet?: string) => page.evaluate(([r, s]) => (window as unknown as { __ohTablo: Api }).__ohTablo.cell(r!, s ?? undefined), [ref, sheet ?? null] as const);
const rectOf = async (page: Page, ref: string) => (await page.evaluate((r) => (window as unknown as { __ohTablo: Api }).__ohTablo.cellRect(r), ref))!;
const selection = (page: Page) => page.evaluate(() => (window as unknown as { __ohTablo: Api }).__ohTablo.selection());

async function go(page: Page, ref: string) {
  const box = page.getByRole("textbox", { name: "Ad Kutusu" });
  await box.fill(ref);
  await box.press("Enter");
}
async function enter(page: Page, ref: string, text: string) {
  await go(page, ref);
  await page.keyboard.type(text);
  await page.keyboard.press("Enter");
}
async function clickCell(page: Page, ref: string) {
  const r = await rectOf(page, ref);
  await page.mouse.click(r.x + r.w / 2, r.y + r.h / 2);
}
/** Seçimin sağ alt köşesindeki tutamaçtan hedef hücreye sürükler. */
async function dragFill(page: Page, from: string, to: string, ctrl = false) {
  const a = await rectOf(page, from);
  const b = await rectOf(page, to);
  await page.mouse.move(a.x + a.w - 1, a.y + a.h - 1);
  await page.mouse.down();
  if (ctrl) await page.keyboard.down("Control");
  await page.mouse.move(b.x + b.w / 2, b.y + b.h / 2, { steps: 8 });
  const tip = await page.locator(".tb-tip").textContent().catch(() => null);
  await page.mouse.up();
  if (ctrl) await page.keyboard.up("Control");
  return tip;
}
const ribbon = (page: Page, name: string) => page.locator(`.oh-ribbon-body button[aria-label="${name}"]`).first();
const tab = (page: Page, name: string) => page.getByRole("tab", { name, exact: true });

test.describe("Tablo 5. tur QA", () => {
  let errors: string[] = [];
  test.beforeEach(({ page }) => {
    errors = collectErrors(page);
    interceptFileChooser(page);
  });
  test.afterEach(() => {
    expect(errors).toEqual([]);
  });

  test("şerit sekmesine tıklamak klavye odağını hücreden almaz", async ({ page }) => {
    await open(page);
    await clickCell(page, "B2");
    await tab(page, "Veri").click();
    await page.keyboard.press("ArrowDown");
    expect(await selection(page)).toBe("B3");
    await expect(tab(page, "Veri")).toHaveAttribute("aria-selected", "true");
    await page.keyboard.type("yazı");
    await page.keyboard.press("Enter");
    expect((await cell(page, "B3"))!.value).toBe("yazı");
    await tab(page, "Formüller").click();
    await clickCell(page, "B3");
    await tab(page, "Giriş").click();
    await page.keyboard.press("Delete");
    expect((await cell(page, "B3"))!.value).toBeNull();
  });

  test("Ctrl ile sürüklenen tek sayı birer artar; ipucu sonuçla aynı", async ({ page }) => {
    await open(page);
    await enter(page, "A1", "5");
    await clickCell(page, "A1");
    expect(await dragFill(page, "A1", "A4", true)).toBe("8");
    expect([(await cell(page, "A2"))!.value, (await cell(page, "A4"))!.value]).toEqual([6, 8]);
    await enter(page, "C1", "Ürün 1");
    await clickCell(page, "C1");
    expect(await dragFill(page, "C1", "C3", true)).toBe("Ürün 1");
    expect((await cell(page, "C3"))!.value).toBe("Ürün 1");
  });

  test("Otomatik Doldurma Seçenekleri: tarihte Ayları Doldur, sayıda Seri Doldur", async ({ page }) => {
    await open(page);
    await enter(page, "A1", "31.01.2024");
    await clickCell(page, "A1");
    await dragFill(page, "A1", "A4");
    expect((await cell(page, "A2"))!.edit).toBe("01.02.2024");
    const tag = page.getByRole("button", { name: "Otomatik Doldurma Seçenekleri" });
    await tag.click();
    await expect(page.getByRole("menuitemcheckbox", { name: "Seri Doldur" })).toHaveAttribute("aria-checked", "true");
    await page.getByRole("menuitemcheckbox", { name: "Ayları Doldur" }).click();
    expect([(await cell(page, "A2"))!.edit, (await cell(page, "A3"))!.edit, (await cell(page, "A4"))!.edit]).toEqual(["29.02.2024", "31.03.2024", "30.04.2024"]);
    const undo = await page.evaluate(() => (window as unknown as { __ohTablo: Api }).__ohTablo.undoCount());
    await tag.click();
    await page.getByRole("menuitemcheckbox", { name: "Hücreleri Kopyala" }).click();
    expect((await cell(page, "A4"))!.edit).toBe("31.01.2024");
    // Seçenek değiştirmek yeni geri alma adımı eklemez
    expect(await page.evaluate(() => (window as unknown as { __ohTablo: Api }).__ohTablo.undoCount())).toBe(undo);
    await enter(page, "C1", "7");
    await clickCell(page, "C1");
    await dragFill(page, "C1", "C3");
    await tag.click();
    await page.getByRole("menuitemcheckbox", { name: "Seri Doldur" }).click();
    expect((await cell(page, "C3"))!.value).toBe(9);
    // Başka bir değişiklikten sonra düğme kaybolur
    await enter(page, "E1", "x");
    await expect(tag).toHaveCount(0);
  });

  test("hücre içi onay kutusu: Ekle, tıkla, Ara Çubuğu, Delete ve formül", async ({ page }) => {
    await open(page);
    await enter(page, "A1", "Süt al");
    await enter(page, "A2", "Rapor yaz");
    await go(page, "B1:B2");
    await tab(page, "Ekle").click();
    await ribbon(page, "Onay Kutusu").click();
    expect((await cell(page, "B1"))!.value).toBe(false);
    await enter(page, "C1", "=EĞERSAY(B1:B2;DOĞRU)");
    await clickCell(page, "B2");
    expect((await cell(page, "B2"))!.value).toBe(true);
    expect((await cell(page, "C1"))!.value).toBe(1);
    await go(page, "B1");
    await page.keyboard.press("Space");
    expect((await cell(page, "B1"))!.value).toBe(true);
    expect((await cell(page, "C1"))!.value).toBe(2);
    await page.keyboard.press("Delete");
    expect((await cell(page, "B1"))!.value).toBe(false);
    await page.keyboard.press("Control+z");
    expect((await cell(page, "B1"))!.value).toBe(true);
    // Formül çubuğu DOĞRU/YANLIŞ gösterir
    await expect(page.getByRole("textbox", { name: "Formül Çubuğu" })).toHaveValue("DOĞRU");
  });

  test("Özel Yapıştır: değerler ve sayı biçimleri yalnız sayı biçimini getirir; Bağlantıyı Yapıştır", async ({ page }) => {
    await open(page);
    await enter(page, "A1", "1.234,50 ₺");
    await go(page, "A1");
    await page.keyboard.press("Control+b");
    const fmt = (await cell(page, "A1"))!.numFmt;
    expect(fmt).toContain("₺");
    await page.keyboard.press("Control+c");
    await go(page, "C1");
    await page.keyboard.press("Control+Alt+v");
    const dlg = page.getByRole("dialog", { name: "Özel Yapıştır" });
    await expect(dlg.getByText("Devrik dönüşüm")).toBeVisible();
    await dlg.getByLabel("Değerler ve sayı biçimleri").check();
    await dlg.getByRole("button", { name: "Tamam" }).click();
    // Yalnız sayı biçimi gelir, kalın yazı gelmez
    expect(await cell(page, "C1")).toMatchObject({ value: 1234.5, bold: false, numFmt: fmt });
    await go(page, "A1");
    await page.keyboard.press("Control+c");
    await go(page, "E1");
    await page.keyboard.press("Control+Alt+v");
    await dlg.getByRole("button", { name: "Bağlantıyı Yapıştır" }).click();
    expect(await cell(page, "E1")).toMatchObject({ value: 1234.5, formula: "=$A$1" });
    await enter(page, "A1", "10");
    expect((await cell(page, "E1"))!.value).toBe(10);
  });

  test("Birleştir: iki aralık etiketlerle toplanır", async ({ page }) => {
    await open(page);
    const rows: [string, string][] = [["A1", "Bölge"], ["B1", "Satış"], ["A2", "Ege"], ["B2", "10"], ["A3", "Marmara"], ["B3", "20"], ["D1", "Bölge"], ["E1", "Satış"], ["D2", "Marmara"], ["E2", "5"], ["D3", "Akdeniz"], ["E3", "7"]];
    for (const [r, v] of rows) await enter(page, r, v);
    await go(page, "G1");
    await tab(page, "Veri").click();
    await ribbon(page, "Birleştir").click();
    const dlg = page.getByRole("dialog", { name: "Birleştir" });
    const ref = dlg.getByRole("textbox", { name: "Başvuru" });
    await ref.fill("A1:B3");
    await dlg.getByRole("button", { name: "Ekle" }).click();
    await ref.fill("D1:E3");
    await dlg.getByRole("button", { name: "Ekle" }).click();
    await dlg.getByLabel("Üst satır").check();
    await dlg.getByLabel("Sol sütun").check();
    await dlg.getByRole("button", { name: "Tamam" }).click();
    expect((await cell(page, "H1"))!.value).toBe("Satış");
    expect([(await cell(page, "G2"))!.value, (await cell(page, "G3"))!.value, (await cell(page, "G4"))!.value]).toEqual(["Ege", "Marmara", "Akdeniz"]);
    expect([(await cell(page, "H2"))!.value, (await cell(page, "H3"))!.value, (await cell(page, "H4"))!.value]).toEqual([10, 25, 7]);
  });

  test("yeni işlevler Türkçe adlarıyla yazılır ve otomatik tamamlanır", async ({ page }) => {
    await open(page);
    await enter(page, "A1", '=REGEXTEST("abc123";"[0-9]+")');
    await enter(page, "A2", "=ONDAİKİ(10;8)");
    await enter(page, "A3", '=ÇEVİR(1;"mi";"km")');
    await enter(page, "A4", "=LAMBDA(x;x*2)(21)");
    expect((await cell(page, "A1"))!.value).toBe(true);
    expect((await cell(page, "A2"))!.value).toBe("00001010");
    expect((await cell(page, "A3"))!.value).toBeCloseTo(1.609344, 9);
    expect((await cell(page, "A4"))!.value).toBe(42);
    await go(page, "A5");
    await page.keyboard.type("=VSEÇ");
    await expect(page.getByRole("option", { name: /VSEÇTOPLA/ })).toBeVisible();
    await page.keyboard.press("Escape");
  });

  test("taşan dinamik dizi XLSX kaydet-aç sonrası yeniden taşar", async ({ page }) => {
    await open(page);
    await enter(page, "A1", "=SIRALI(3;2)");
    await enter(page, "D1", "=FİLTRE(A1:A3;A1:A3>1)");
    expect((await cell(page, "B3"))!.value).toBe(6);
    await page.locator(".oh-ribbon-file").click();
    await page.locator(".oh-bs-item", { hasText: "Dışarı Aktar" }).first().click();
    const [download] = await Promise.all([page.waitForEvent("download"), page.getByText("Çalışma kitabını indir (.xlsx)").click()]);
    const file = join(mkdtempSync(join(tmpdir(), "oh-tablo-qa-")), "dizi.xlsx");
    await download.saveAs(file);
    await page.evaluate(() => {
      delete (window as unknown as { showOpenFilePicker?: unknown }).showOpenFilePicker;
    });
    await page.locator(".oh-bs-item", { hasText: "Aç" }).first().click();
    await page.locator(".oh-bs-action", { hasText: "Gözat" }).first().click();
    const [chooser] = await Promise.all([page.waitForEvent("filechooser"), page.getByRole("button", { name: "Kaydetme", exact: true }).click()]);
    await chooser.setFiles(file);
    await expect(page.locator(".oh-titlebar")).toContainText("dizi");
    await page.waitForFunction(() => (window as unknown as { __ohTablo: Api }).__ohTablo.cell("B3")?.value === 6);
    expect(await cell(page, "A1")).toMatchObject({ value: 1, formula: "=SEQUENCE(3,2)" });
    expect((await cell(page, "B3"))!.value).toBe(6);
    expect((await cell(page, "D2"))!.value).toBe(5);
    // Kaynak değişince dizi yeniden boyutlanır (taşma alanı sabit değer değil)
    await enter(page, "A1", "=SIRALI(2;2)");
    expect((await cell(page, "B3"))!.value).toBeNull();
    expect((await cell(page, "A1"))!.value).toBe(1);
  });

  test("dokunmatik ekranda parmakla kaydırma sayfayı kaydırır, dokunma hücre seçer", async ({ browser }) => {
    const ctx = await browser.newContext({ viewport: { width: 900, height: 700 }, hasTouch: true, locale: "tr-TR" });
    const page = await ctx.newPage();
    const errs = collectErrors(page);
    await open(page);
    const cdp = await ctx.newCDPSession(page);
    const swipe = async (x1: number, y1: number, x2: number, y2: number) => {
      await cdp.send("Input.dispatchTouchEvent", { type: "touchStart", touchPoints: [{ x: x1, y: y1 }] });
      for (let i = 1; i <= 10; i++) await cdp.send("Input.dispatchTouchEvent", { type: "touchMove", touchPoints: [{ x: x1 + ((x2 - x1) * i) / 10, y: y1 + ((y2 - y1) * i) / 10 }] });
      await cdp.send("Input.dispatchTouchEvent", { type: "touchEnd", touchPoints: [] });
    };
    const before = await rectOf(page, "A10");
    await swipe(400, 550, 400, 350);
    await expect.poll(async () => (await rectOf(page, "A10")).y).toBeLessThan(before.y - 150);
    expect(await selection(page)).toBe("A1");
    const d5 = await rectOf(page, "D20");
    await page.touchscreen.tap(d5.x + d5.w / 2, d5.y + d5.h / 2);
    await expect.poll(() => selection(page)).toBe("D20");
    expect(errs).toEqual([]);
    await ctx.close();
  });

  test("Bul ve Değiştir Türkçe büyük/küçük harfi eşler (izmir → İzmir, ısparta → ISPARTA)", async ({ page }) => {
    await open(page);
    await enter(page, "A1", "İzmir");
    await enter(page, "A2", "ISPARTA");
    await enter(page, "A3", "Ankara");
    await enter(page, "B1", '=EĞERSAY(A1:A3;"izm*")');
    expect((await cell(page, "B1"))!.value).toBe(1);
    await go(page, "C5");
    await page.keyboard.press("Control+f");
    const dlg = page.getByRole("dialog", { name: "Bul ve Değiştir" });
    const q = dlg.locator(".tb-find input[type=text]").first();
    await q.fill("izmir");
    await dlg.getByRole("button", { name: "Sonrakini Bul" }).click();
    await expect.poll(() => selection(page)).toBe("A1");
    await q.fill("ısparta");
    await dlg.getByRole("button", { name: "Sonrakini Bul" }).click();
    await expect.poll(() => selection(page)).toBe("A2");
    await dlg.getByRole("tab", { name: "Değiştir" }).click();
    await q.fill("ısparta");
    await dlg.locator(".tb-find input[type=text]").nth(1).fill("Burdur");
    await dlg.getByRole("button", { name: "Tümünü Değiştir" }).click();
    await expect(page.getByRole("dialog").filter({ hasText: "1 değiştirme yapıldı" })).toBeVisible();
    expect((await cell(page, "A2"))!.value).toBe("Burdur");
  });
});
