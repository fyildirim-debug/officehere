import { expect, test, type Page } from "@playwright/test";
import { mkdtempSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join, resolve } from "node:path";
import { URL, collectErrors } from "./ortak";
import { tabloTur4Tests } from "./tablo-tur4";

// OfficeHere Tablo uçtan uca senaryoları. Izgara canvas olduğu için hücre konumları
// ve değerleri uygulamanın yalnız okunur test arayüzünden (window.__ohTablo) alınır.

interface CellInfo {
  value: string | number | boolean | null;
  formula: string | null;
  edit: string;
  numFmt: string | null;
  bold: boolean;
  fill: string | null;
}

type Api = {
  cell(ref: string, sheet?: string): CellInfo | null;
  cellRect(ref: string): { x: number; y: number; w: number; h: number } | null;
  activeSheet(): string;
  sheetNames(): string[];
  selection(): string;
  chartCount(sheet?: string): number;
  pivotCount(sheet?: string): number;
  sparklines(sheet?: string): { type: string; cells: string[] }[];
  isRowHidden(row: number, sheet?: string): boolean;
  undoCount(): number;
};

declare global {
  interface Window {
    __ohTablo?: Api;
  }
}

async function open(page: Page) {
  await page.goto(URL.tablo);
  await page.waitForFunction(() => !!window.__ohTablo && !!window.__ohTablo.cellRect("A1"));
}

const cell = (page: Page, ref: string, sheet?: string) => page.evaluate(([r, s]) => window.__ohTablo!.cell(r!, s ?? undefined), [ref, sheet ?? null] as const);
const selection = (page: Page) => page.evaluate(() => window.__ohTablo!.selection());

async function clickCell(page: Page, ref: string, opts?: { button?: "left" | "right" }) {
  const r = await page.evaluate((x) => window.__ohTablo!.cellRect(x), ref);
  if (!r) throw new Error(`Hücre görünür değil: ${ref}`);
  await page.mouse.click(r.x + r.w / 2, r.y + r.h / 2, opts);
}

/** Ad Kutusu ile hücreye gidip yazar ve Enter'a basar. */
async function enter(page: Page, ref: string, text: string) {
  const box = page.getByRole("textbox", { name: "Ad Kutusu" });
  await box.fill(ref);
  await box.press("Enter");
  await page.keyboard.type(text);
  await page.keyboard.press("Enter");
}

async function fillData(page: Page, rows: (string | number)[][], start = 0) {
  for (let i = 0; i < rows.length; i++)
    for (let j = 0; j < rows[i].length; j++) {
      const v = rows[i][j];
      if (v === "") continue;
      await enter(page, `${String.fromCharCode(65 + j)}${i + 1 + start}`, String(v));
    }
}

const ribbon = (page: Page, name: string) => page.locator(`.oh-ribbon-body button[aria-label="${name}"]`).first();
const tab = (page: Page, name: string) => page.getByRole("tab", { name, exact: true });

test.describe("OfficeHere Tablo uçtan uca", () => {
  // Her senaryoda sayfa/konsol hatası olmamalı
  let errors: string[] = [];
  test.beforeEach(({ page }) => {
    errors = collectErrors(page);
  });
  test.afterEach(() => {
    expect(errors).toEqual([]);
  });

  test("değer yazma, Enter ile aşağı inme ve sayı/tarih/yüzde tanıma", async ({ page }) => {
    await open(page);
    await clickCell(page, "A1");
    await page.keyboard.type("Merhaba");
    await page.keyboard.press("Enter");
    expect(await selection(page)).toBe("A2");
    await page.keyboard.type("1.234,5");
    await page.keyboard.press("Enter");
    await page.keyboard.type("12.10.2026");
    await page.keyboard.press("Enter");
    await page.keyboard.type("%15");
    await page.keyboard.press("Enter");
    expect((await cell(page, "A1"))!.value).toBe("Merhaba");
    expect((await cell(page, "A2"))!.value).toBe(1234.5);
    expect((await cell(page, "A3"))!.numFmt).toBe("dd.mm.yyyy");
    expect((await cell(page, "A3"))!.edit).toBe("12.10.2026");
    expect((await cell(page, "A4"))!.value).toBe(0.15);
    await expect(page.getByRole("textbox", { name: "Formül Çubuğu" })).toHaveValue("");
  });

  test("=TOPLA ve =SUM aynı sonucu verir; Türkçe gösterilir", async ({ page }) => {
    await open(page);
    await fillData(page, [["10"], ["20"], ["30"]]);
    await enter(page, "B1", "=TOPLA(A1:A3)");
    await enter(page, "B2", "=SUM(A1;A3)");
    await enter(page, "B3", "=ortalama(a1:a3)");
    expect(await cell(page, "B1")).toMatchObject({ value: 60, formula: "=SUM(A1:A3)", edit: "=TOPLA(A1:A3)" });
    expect((await cell(page, "B2"))!.value).toBe(40);
    expect((await cell(page, "B3"))!.value).toBe(20);
    await clickCell(page, "B1");
    await expect(page.getByRole("textbox", { name: "Formül Çubuğu" })).toHaveValue("=TOPLA(A1:A3)");
  });

  test("doldurma tutamacıyla seri (1, 2 → 6 ve Pazartesi → Çarşamba)", async ({ page }) => {
    await open(page);
    await fillData(page, [["1", "Pazartesi"], ["2"]]);
    await clickCell(page, "A1");
    await page.keyboard.press("Shift+ArrowDown");
    const a2 = (await page.evaluate(() => window.__ohTablo!.cellRect("A2")))!;
    const a6 = (await page.evaluate(() => window.__ohTablo!.cellRect("A6")))!;
    await page.mouse.move(a2.x + a2.w - 1, a2.y + a2.h - 1);
    await page.mouse.down();
    await page.mouse.move(a6.x + a6.w / 2, a6.y + a6.h / 2, { steps: 8 });
    await page.mouse.up();
    expect((await cell(page, "A6"))!.value).toBe(6);
    await clickCell(page, "B1");
    const b1 = (await page.evaluate(() => window.__ohTablo!.cellRect("B1")))!;
    const b3 = (await page.evaluate(() => window.__ohTablo!.cellRect("B3")))!;
    await page.mouse.move(b1.x + b1.w - 1, b1.y + b1.h - 1);
    await page.mouse.down();
    await page.mouse.move(b3.x + b3.w / 2, b3.y + b3.h / 2, { steps: 8 });
    await page.mouse.up();
    expect((await cell(page, "B3"))!.value).toBe("Çarşamba");
  });

  test("kalın ve dolgu rengi (şerit düğmeleri ve Ctrl+B)", async ({ page }) => {
    await open(page);
    await enter(page, "A1", "Başlık");
    await clickCell(page, "A1");
    await page.locator('.oh-ribbon-body [aria-label="Kalın"]').first().click();
    await page.locator('.oh-ribbon-body [aria-label="Dolgu Rengi"]').first().click();
    let c = (await cell(page, "A1"))!;
    expect(c.bold).toBe(true);
    expect(c.fill).toBe("#FFFF00");
    await clickCell(page, "A1");
    await page.keyboard.press("Control+b");
    c = (await cell(page, "A1"))!;
    expect(c.bold).toBe(false);
  });

  test("satır ekleme ve silme formül başvurularını kaydırır", async ({ page }) => {
    await open(page);
    await fillData(page, [["1"], ["2"], ["3"]]);
    await enter(page, "A4", "=TOPLA(A1:A3)");
    await enter(page, "B1", "=$A$2*A3");
    // 2. satır başlığına sağ tık → Ekle
    const r2 = (await page.evaluate(() => window.__ohTablo!.cellRect("A2")))!;
    await page.mouse.click(r2.x - 10, r2.y + r2.h / 2, { button: "right" });
    await page.getByRole("menuitem", { name: "Ekle", exact: true }).click();
    expect((await cell(page, "A5"))!.formula).toBe("=SUM(A1:A4)");
    expect((await cell(page, "B1"))!.formula).toBe("=$A$3*A4");
    expect((await cell(page, "A5"))!.value).toBe(6);
    await page.mouse.click(r2.x - 10, r2.y + r2.h / 2, { button: "right" });
    await page.getByRole("menuitem", { name: "Sil", exact: true }).click();
    expect((await cell(page, "A4"))!.formula).toBe("=SUM(A1:A3)");
  });

  test("Veri > A'dan Z'ye sırala (Türkçe harf sırası)", async ({ page }) => {
    await open(page);
    await fillData(page, [["Ad", "Puan"], ["Şule", "3"], ["ali", "5"], ["Çağrı", "1"], ["Ilgın", "2"]]);
    await clickCell(page, "A2");
    await tab(page, "Veri").click();
    await page.locator(`.oh-ribbon-body [aria-label="A'dan Z'ye Sırala"]`).first().click();
    const names = [];
    for (let r = 2; r <= 5; r++) names.push((await cell(page, `A${r}`))!.value);
    expect(names).toEqual(["ali", "Çağrı", "Ilgın", "Şule"]);
    expect((await cell(page, "A1"))!.value).toBe("Ad");
    expect((await cell(page, "B3"))!.value).toBe(1);
  });

  test("otomatik filtre: değere göre süzme ve temizleme", async ({ page }) => {
    await open(page);
    await fillData(page, [["Şehir", "Tutar"], ["İzmir", "10"], ["Ankara", "20"], ["izmir", "30"], ["Bursa", "40"]]);
    await clickCell(page, "A1");
    await tab(page, "Veri").click();
    await ribbon(page, "Filtre").click();
    const a1 = (await page.evaluate(() => window.__ohTablo!.cellRect("A1")))!;
    await page.mouse.click(a1.x + a1.w - 7, a1.y + a1.h / 2);
    await page.locator("label", { hasText: "(Tümünü Seç)" }).click();
    await page.locator("label", { hasText: /^İzmir$/ }).click();
    await page.getByRole("button", { name: "Tamam" }).click();
    const hidden = await page.evaluate(() => [1, 2, 3, 4].map((r) => window.__ohTablo!.isRowHidden(r)));
    expect(hidden).toEqual([false, true, false, true]);
    await expect(page.locator(".oh-statusbar")).toContainText("4 kayıttan 2 tanesi bulundu");
    await ribbon(page, "Temizle").click();
    expect(await page.evaluate(() => [1, 2, 3, 4].some((r) => window.__ohTablo!.isRowHidden(r)))).toBe(false);
  });

  test("yeni sayfa, yeniden adlandırma ve sayfalar arası formül", async ({ page }) => {
    await open(page);
    await enter(page, "A1", "42");
    await page.getByRole("button", { name: "Yeni sayfa" }).click();
    expect(await page.evaluate(() => window.__ohTablo!.activeSheet())).toBe("Sayfa2");
    await enter(page, "A1", "=Sayfa1!A1*2");
    expect((await cell(page, "A1", "Sayfa2"))!.value).toBe(84);
    // = yazıp diğer sayfadaki hücreye tıklayarak başvuru
    await clickCell(page, "B1");
    await page.keyboard.type("=");
    await page.locator(".tb-tab", { hasText: "Sayfa1" }).click();
    await clickCell(page, "A1");
    await page.keyboard.press("Enter");
    expect((await cell(page, "B1", "Sayfa2"))!.formula).toBe("=Sayfa1!A1");
    // Sayfa1'i yeniden adlandır: formüller izler
    await page.locator(".tb-tab", { hasText: "Sayfa1" }).dblclick();
    await page.keyboard.type("Veriler");
    await page.keyboard.press("Enter");
    expect(await page.evaluate(() => window.__ohTablo!.sheetNames())).toEqual(["Veriler", "Sayfa2"]);
    expect((await cell(page, "A1", "Sayfa2"))!.formula).toBe("=Veriler!A1*2");
  });

  test("geri al / yinele zinciri", async ({ page }) => {
    await open(page);
    await enter(page, "A1", "1");
    await enter(page, "A2", "2");
    await enter(page, "A3", "=A1+A2");
    await clickCell(page, "A1");
    await page.keyboard.press("Control+b");
    expect(await page.evaluate(() => window.__ohTablo!.undoCount())).toBe(4);
    for (let i = 0; i < 4; i++) await page.keyboard.press("Control+z");
    expect((await cell(page, "A1"))!.value).toBe(null);
    expect((await cell(page, "A3"))!.value).toBe(null);
    for (let i = 0; i < 4; i++) await page.keyboard.press("Control+y");
    expect((await cell(page, "A3"))!.value).toBe(3);
    expect((await cell(page, "A1"))!.bold).toBe(true);
  });

  test("grafik ekleme ve bağlamsal Grafik Tasarımı sekmesi", async ({ page }) => {
    await open(page);
    await fillData(page, [["Ay", "Satış"], ["Ocak", "100"], ["Şubat", "150"], ["Mart", "120"]]);
    await clickCell(page, "A1");
    await tab(page, "Ekle").click();
    await ribbon(page, "Sütun veya Çubuk").click();
    await page.locator(".oh-menu-item", { hasText: "Kümelenmiş Sütun" }).first().click();
    expect(await page.evaluate(() => window.__ohTablo!.chartCount())).toBe(1);
    await expect(tab(page, "Grafik Tasarımı")).toBeVisible();
    await expect(page.locator(".tb-float svg").first()).toBeVisible();
  });

  test("XLSX indir → aynı dosyayı aç: değerler ve formüller korunur", async ({ page }) => {
    await open(page);
    await fillData(page, [["Ürün", "Adet", "Fiyat", "Tutar"], ["Çay", "10", "2,5", "=B2*C2"], ["Şeker", "4", "30", "=B3*C3"]]);
    await enter(page, "D4", "=TOPLA(D2:D3)");
    await clickCell(page, "A1");
    await page.keyboard.press("Control+b");
    await page.locator(".oh-ribbon-file").click();
    await page.locator(".oh-bs-item", { hasText: "Dışarı Aktar" }).first().click();
    const [download] = await Promise.all([page.waitForEvent("download"), page.getByText("Çalışma kitabını indir (.xlsx)").click()]);
    expect(download.suggestedFilename()).toBe("Kitap1.xlsx");
    const file = join(mkdtempSync(join(tmpdir(), "oh-tablo-")), "Kitap1.xlsx");
    await download.saveAs(file);
    // Yeni kitap aç, sonra dosyayı yükle
    await page.evaluate(() => {
      delete (window as unknown as { showOpenFilePicker?: unknown }).showOpenFilePicker;
    });
    // Dışarı aktarma kitabı "kaydedilmiş" yapmaz: Aç önce kaydetme sorusunu getirir
    await page.locator(".oh-bs-item", { hasText: "Aç" }).first().click();
    await page.locator(".oh-bs-action", { hasText: "Gözat" }).first().click();
    const [chooser] = await Promise.all([page.waitForEvent("filechooser"), page.getByRole("button", { name: "Kaydetme", exact: true }).click()]);
    await chooser.setFiles(file);
    await expect(page.locator(".oh-titlebar")).toContainText("Kitap1");
    await page.waitForFunction(() => window.__ohTablo!.cell("D4")?.value === 145);
    expect(await cell(page, "D2")).toMatchObject({ value: 25, formula: "=B2*C2" });
    expect(await cell(page, "D4")).toMatchObject({ value: 145, formula: "=SUM(D2:D3)" });
    expect((await cell(page, "A1"))!.bold).toBe(true);
    expect((await cell(page, "A3"))!.value).toBe("Şeker");
  });

  test("CSV içe aktarma (noktalı virgül, Türkçe karakter, ondalık virgül)", async ({ page }) => {
    await open(page);
    const dir = mkdtempSync(join(tmpdir(), "oh-csv-"));
    const file = join(dir, "satis.csv");
    writeFileSync(file, "Ad;Şehir;Tutar\nÇağrı;İzmir;1.234,50\nŞule;Muğla;99,9\n", "utf8");
    await page.evaluate(() => {
      delete (window as unknown as { showOpenFilePicker?: unknown }).showOpenFilePicker;
    });
    await page.locator(".oh-ribbon-file").click();
    await page.locator(".oh-bs-item", { hasText: "Aç" }).first().click();
    const [chooser] = await Promise.all([page.waitForEvent("filechooser"), page.locator(".oh-bs-action", { hasText: "Gözat" }).first().click()]);
    await chooser.setFiles(file);
    await page.waitForFunction(() => window.__ohTablo!.cell("B2")?.value === "İzmir");
    expect((await cell(page, "C2"))!.value).toBe(1234.5);
    expect((await cell(page, "A3"))!.value).toBe("Şule");
    expect(await page.evaluate(() => window.__ohTablo!.sheetNames())).toEqual(["satis"]);
  });

  test("hücre bağlam menüsü: Hücreleri Biçimlendir ve İçeriği Temizle", async ({ page }) => {
    await open(page);
    await enter(page, "B2", "123");
    await clickCell(page, "B2", { button: "right" });
    const menu = page.getByRole("menu");
    await expect(menu.getByRole("menuitem", { name: "Kes" })).toBeVisible();
    await expect(menu.getByRole("menuitem", { name: "Kopyala" })).toBeVisible();
    await menu.getByRole("menuitem", { name: "Hücreleri Biçimlendir..." }).click();
    await expect(page.getByRole("dialog", { name: "Hücreleri Biçimlendir" })).toBeVisible();
    await page.keyboard.press("Escape");
    await clickCell(page, "B2", { button: "right" });
    await page.getByRole("menuitem", { name: "İçeriği Temizle" }).click();
    expect((await cell(page, "B2"))!.value).toBe(null);
  });

  test("Özet Tablo: oluştur, alan seç, sütuna sürükle, düzenleme engeli", async ({ page }) => {
    await open(page);
    await fillData(page, [["Bölge", "Yıl", "Satış"], ["Ege", "2024", "100"], ["Ege", "2025", "50"], ["Marmara", "2024", "200"]]);
    await clickCell(page, "A1");
    await tab(page, "Ekle").click();
    await ribbon(page, "Özet Tablo").click();
    await page.getByRole("dialog", { name: "Özet Tablo Oluştur" }).getByRole("button", { name: "Tamam" }).click();
    await expect(page.getByRole("complementary", { name: "Özet Tablo Alanları" })).toBeVisible();
    expect(await page.evaluate(() => window.__ohTablo!.activeSheet())).toBe("Sayfa2");
    await page.locator(".tb-pv-field", { hasText: "Bölge" }).locator("input").check();
    await page.locator(".tb-pv-field", { hasText: "Satış" }).locator("input").check();
    await expect.poll(() => page.evaluate(() => window.__ohTablo!.cell("B6")?.value)).toBe(350);
    expect((await cell(page, "A3"))!.value).toBe("Satır Etiketleri");
    expect((await cell(page, "A4"))!.value).toBe("Ege");
    await page.locator(".tb-pv-field", { hasText: "Yıl" }).dragTo(page.locator('.tb-pv-area[data-area="cols"]'));
    await expect.poll(() => page.evaluate(() => window.__ohTablo!.cell("B4")?.value)).toBe(2024);
    expect((await cell(page, "D7"))!.value).toBe(350);
    await expect(tab(page, "Özet Tablo Analizi")).toBeVisible();
    await clickCell(page, "B5");
    await page.keyboard.type("9");
    await expect(page.getByRole("dialog").getByText("Özet tablonun bu bölümü doğrudan düzenlenemez", { exact: false })).toBeVisible();
  });

  test("Formül Denetimi: etkileyen okları ve Formülü Değerlendir", async ({ page }) => {
    await open(page);
    await fillData(page, [["2"], ["3"]]);
    await enter(page, "B1", "=(A1+A2)*10");
    await clickCell(page, "B1");
    await tab(page, "Formüller").click();
    await ribbon(page, "Etkileyenleri İzle").click();
    await ribbon(page, "Formülü Değerlendir").click();
    const dlg = page.getByRole("dialog", { name: "Formülü Değerlendir" });
    await expect(dlg.getByLabel("Değerlendirme").last()).toHaveText("=(A1+A2)*10");
    await dlg.getByRole("button", { name: "Değerlendir" }).click();
    await expect(dlg.getByLabel("Değerlendirme").last()).toHaveText("=(2+A2)*10");
    for (let i = 0; i < 3; i++) await dlg.getByRole("button", { name: "Değerlendir" }).click();
    await expect(dlg.getByLabel("Değerlendirme").last()).toHaveText("=50");
  });

  test("Bul ve Değiştir: 'Tümü tamamlandı' uyarısı bul penceresinin üstünde açılır", async ({ page }) => {
    await open(page);
    await fillData(page, [["elma"], ["armut elma"]]);
    await clickCell(page, "A1");
    await page.keyboard.press("Control+h");
    const find = page.getByRole("dialog", { name: "Bul ve Değiştir" });
    await find.locator("input").first().fill("elma");
    await find.locator("input").nth(1).fill("ELMA");
    await find.getByRole("button", { name: "Tümünü Değiştir" }).click();
    const alert = page.getByRole("dialog", { name: "OfficeHere Tablo" });
    await expect(alert).toContainText("2 değiştirme yapıldı");
    await alert.getByRole("button", { name: "Tamam" }).click();
    await expect(find).toBeVisible();
    expect((await cell(page, "A2"))!.value).toBe("armut ELMA");
  });
  test("Mini grafikler: ekle, türünü değiştir, doldurma ile kopyala, temizle", async ({ page }) => {
    await open(page);
    await fillData(page, [["Ürün", "Oca", "Şub", "Mar"], ["Kalem", "10", "14", "9"], ["Defter", "-5", "3", "-2"]]);
    await clickCell(page, "B2");
    await page.keyboard.press("Shift+ArrowRight");
    await page.keyboard.press("Shift+ArrowRight");
    await tab(page, "Ekle").click();
    await ribbon(page, "Çizgi").click();
    const dlg = page.getByRole("dialog", { name: "Mini Grafikler Oluştur" });
    await expect(dlg.getByLabel("Veri Aralığı")).toHaveValue("B2:D2");
    await dlg.getByLabel("Konum Aralığı").fill("E2");
    await dlg.getByRole("button", { name: "Tamam" }).click();
    expect(await page.evaluate(() => window.__ohTablo!.sparklines())).toEqual([{ type: "line", cells: ["E2"] }]);
    await expect(tab(page, "Mini Grafik")).toBeVisible();
    await ribbon(page, "Sütun").click();
    expect((await page.evaluate(() => window.__ohTablo!.sparklines()))[0].type).toBe("column");
    // Doldurma tutamacıyla E3'e kopyala (aynı gruba, göreli veri aralığıyla)
    const r = (await page.evaluate(() => window.__ohTablo!.cellRect("E2")))!;
    const r3 = (await page.evaluate(() => window.__ohTablo!.cellRect("E3")))!;
    await page.mouse.move(r.x + r.w - 1, r.y + r.h - 1);
    await page.mouse.down();
    await page.mouse.move(r3.x + r3.w / 2, r3.y + r3.h - 2, { steps: 5 });
    await page.mouse.up();
    expect(await page.evaluate(() => window.__ohTablo!.sparklines())).toEqual([{ type: "column", cells: ["E2", "E3"] }]);
    // Temizle ve geri al
    await clickCell(page, "E3");
    await ribbon(page, "Temizle").click();
    expect(await page.evaluate(() => window.__ohTablo!.sparklines())).toEqual([{ type: "column", cells: ["E2"] }]);
    await page.keyboard.press("Control+z");
    expect((await page.evaluate(() => window.__ohTablo!.sparklines()))[0].cells).toEqual(["E2", "E3"]);
  });
});

// 3. tur: kabul testi bulgularının düzeltmeleri
const FIXTURES = resolve(__dirname, "..", "apps/tablo/test/fixtures/openpyxl");

/** Dosya › Aç › Gözat ile dosya açar (kitapta kaydedilmemiş değişiklik olmamalı). */
async function openFile(page: Page, file: string) {
  await page.evaluate(() => {
    delete (window as unknown as { showOpenFilePicker?: unknown }).showOpenFilePicker;
  });
  await page.locator(".oh-ribbon-file").click();
  await page.locator(".oh-bs-item", { hasText: "Aç" }).first().click();
  const [chooser] = await Promise.all([page.waitForEvent("filechooser"), page.locator(".oh-bs-action", { hasText: "Gözat" }).first().click()]);
  await chooser.setFiles(file);
}

/** Düğmeye tıklayıp işlemin ve ardından gelen iki çizim karesinin süresini (ms) ölçer. */
async function timedClick(selector: ReturnType<Page["locator"]>): Promise<number> {
  return selector.evaluate(async (b: HTMLElement) => {
    const t = performance.now();
    b.click();
    await new Promise((r) => requestAnimationFrame(() => requestAnimationFrame(r)));
    return performance.now() - t;
  });
}

test.describe("OfficeHere Tablo 3. tur", () => {
  let errors: string[] = [];
  test.beforeEach(({ page }) => {
    errors = collectErrors(page);
  });
  test.afterEach(() => {
    expect(errors).toEqual([]);
  });

  test("openpyxl ile üretilmiş XLSX açılır (mutlak ilişki hedefleri, notlar, tablo, grafikler)", async ({ page }) => {
    await open(page);
    await openFile(page, join(FIXTURES, "karmasik.xlsx"));
    await page.waitForFunction(() => window.__ohTablo!.sheetNames().includes("Bütçe"));
    expect(await page.evaluate(() => window.__ohTablo!.sheetNames())).toEqual(expect.arrayContaining(["Bütçe", "Satışlar"]));
    expect(await page.evaluate(() => window.__ohTablo!.chartCount("Bütçe"))).toBe(2);
    expect(await page.evaluate(() => window.__ohTablo!.chartCount("Satışlar"))).toBe(1);
    await expect(page.locator(".tb-float.is-chart svg").first()).toBeVisible();
    await expect(page.locator(".oh-messagebar")).toHaveCount(0);
  });

  test("okunamayan parça dosyayı reddettirmez; bilgi çubuğunda bildirilir", async ({ page }) => {
    await open(page);
    await openFile(page, join(FIXTURES, "bozuk_aciklama.xlsx"));
    await expect(page.locator(".oh-messagebar")).toContainText("okunamadı ve atlandı");
    await expect(page.locator(".oh-messagebar")).toContainText("açıklamalar");
    expect((await cell(page, "B2"))!.value).toBe(100);
    expect(await page.evaluate(() => window.__ohTablo!.chartCount())).toBe(1);
    await page.locator(".oh-messagebar").getByRole("button", { name: "Kapat" }).click();
    await expect(page.locator(".oh-messagebar")).toHaveCount(0);
  });

  test("50.000 satırda filtre uygulama ve temizleme 1 saniyenin altında", async ({ page }) => {
    await open(page);
    const file = join(mkdtempSync(join(tmpdir(), "oh-big-")), "buyuk.csv");
    const cities = ["İzmir", "Ankara", "Bursa", "Muğla", "Adana"];
    const lines = ["Şehir;Tutar"];
    for (let i = 0; i < 50000; i++) lines.push(`${cities[i % 5]};${i % 997}`);
    writeFileSync(file, lines.join("\n") + "\n", "utf8");
    await openFile(page, file);
    await page.waitForFunction(() => window.__ohTablo!.cell("A50001")?.value === "Adana", null, { timeout: 30_000 });
    await clickCell(page, "A1");
    await tab(page, "Veri").click();
    await ribbon(page, "Filtre").click();
    const a1 = (await page.evaluate(() => window.__ohTablo!.cellRect("A1")))!;
    await page.mouse.click(a1.x + a1.w - 7, a1.y + a1.h / 2);
    await page.locator("label", { hasText: "(Tümünü Seç)" }).click();
    await page.locator("label", { hasText: /^İzmir$/ }).click();
    const apply = await timedClick(page.getByRole("button", { name: "Tamam" }));
    expect(await page.evaluate(() => [1, 2, 6, 49996].map((r) => window.__ohTablo!.isRowHidden(r)))).toEqual([false, true, false, false]);
    await expect(page.locator(".oh-statusbar")).toContainText("50.000 kayıttan 10.000 tanesi bulundu");
    const clear = await timedClick(ribbon(page, "Temizle"));
    expect(await page.evaluate(() => [2, 3, 49999].some((r) => window.__ohTablo!.isRowHidden(r)))).toBe(false);
    await expect(page.locator(".oh-statusbar")).not.toContainText("tanesi bulundu");
    console.log(`50.000 satır: filtre uygula ${Math.round(apply)} ms, temizle ${Math.round(clear)} ms`);
    expect(apply).toBeLessThan(1000);
    expect(clear).toBeLessThan(1000);
  });

  test("tablo alttaki ve sağdaki girişle genişler; hesaplanan sütun ve yapılandırılmış başvuru", async ({ page }) => {
    await open(page);
    await fillData(page, [["Ürün", "Adet", "Fiyat"], ["Çay", "10", "2,5"], ["Şeker", "4", "30"]]);
    await clickCell(page, "A1");
    await page.keyboard.press("Control+t");
    await page.getByRole("dialog", { name: "Tablo Oluştur" }).getByRole("button", { name: "Tamam" }).click();
    await enter(page, "D1", "Tutar");
    await enter(page, "D2", "=[@Adet]*[@Fiyat]");
    expect(await cell(page, "D2")).toMatchObject({ value: 25, edit: "=[@Adet]*[@Fiyat]" });
    // Hesaplanan sütun tüm satırlara yayılır
    expect(await cell(page, "D3")).toMatchObject({ value: 120, edit: "=[@Adet]*[@Fiyat]" });
    // Alttaki satıra yazınca tablo genişler, formül yeni satıra gelir
    await enter(page, "A4", "Un");
    await enter(page, "B4", "2");
    await enter(page, "C4", "50");
    expect((await cell(page, "D4"))!.value).toBe(100);
    await enter(page, "F1", "=TOPLA(Tablo1[Tutar])");
    expect(await cell(page, "F1")).toMatchObject({ value: 245, formula: "=SUM(Tablo1[Tutar])", edit: "=TOPLA(Tablo1[Tutar])" });
  });

  test("F4 son biçim işlemini yineler", async ({ page }) => {
    await open(page);
    await enter(page, "B2", "1");
    await enter(page, "B5", "2");
    await clickCell(page, "B2");
    await page.keyboard.press("Control+b");
    await clickCell(page, "B5");
    await page.keyboard.press("F4");
    expect((await cell(page, "B5"))!.bold).toBe(true);
  });

  test("grafik: çift tık biçim bölmesi, Ctrl+C/V kopya, Delete yalnız grafiği siler", async ({ page }) => {
    await open(page);
    await fillData(page, [["Ay", "Satış"], ["Ocak", "10"], ["Şubat", "14"], ["Mart", "9"]]);
    await clickCell(page, "B2");
    await page.keyboard.press("Alt+F1");
    expect(await page.evaluate(() => window.__ohTablo!.chartCount())).toBe(1);
    const chart = page.locator(".tb-float.is-chart").first();
    // Çift tık tıklanan öğenin bölmesini açar: kenar boşluğu grafik alanı, sütun veri serisi
    await chart.dblclick({ position: { x: 6, y: 40 } });
    await expect(page.getByRole("complementary", { name: "Grafik Alanını Biçimlendir" })).toBeVisible();
    await page.getByRole("complementary", { name: "Grafik Alanını Biçimlendir" }).getByRole("button", { name: "Kapat" }).click();
    const box = (await chart.boundingBox())!;
    await page.mouse.click(box.x + box.width / 2, box.y + 12);
    await page.keyboard.press("Control+c");
    await clickCell(page, "A12");
    await page.keyboard.press("Control+v");
    await expect.poll(() => page.evaluate(() => window.__ohTablo!.chartCount())).toBe(2);
    // Yapıştırılan grafik seçili: Delete onu siler, hücreler yerinde kalır
    await page.keyboard.press("Delete");
    await expect.poll(() => page.evaluate(() => window.__ohTablo!.chartCount())).toBe(1);
    expect((await cell(page, "B2"))!.value).toBe(10);
    await page.keyboard.press("Control+z");
    await expect.poll(() => page.evaluate(() => window.__ohTablo!.chartCount())).toBe(2);
  });

  test("Alt Toplam anahat düzeyleriyle; 2 düğmesi ayrıntıyı gizler", async ({ page }) => {
    await open(page);
    await fillData(page, [["Bölge", "Tutar"], ["Doğu", "10"], ["Doğu", "20"], ["Batı", "5"]]);
    await clickCell(page, "A2");
    await tab(page, "Veri").click();
    await ribbon(page, "Alt Toplam").click();
    await page.getByRole("dialog", { name: "Alt Toplam" }).getByRole("button", { name: "Tamam" }).click();
    expect(await cell(page, "B4")).toMatchObject({ value: 30, formula: "=SUBTOTAL(9,B2:B3)" });
    expect(await cell(page, "A7")).toMatchObject({ value: "Genel Toplam" });
    expect((await cell(page, "B7"))!.value).toBe(35);
    // Satır anahattı düzey düğmesi "2" (sütun başlığı hizasında, satır numaralarının solunda)
    const canvas = (await page.locator(".tb-canvas").boundingBox())!;
    const a1 = (await page.evaluate(() => window.__ohTablo!.cellRect("A1")))!;
    await page.mouse.click(canvas.x + 2 + 1.5 * 14, a1.y - 10);
    expect(await page.evaluate(() => [1, 2, 3, 4].map((r) => window.__ohTablo!.isRowHidden(r)))).toEqual([true, true, false, true]);
  });

  test("Sayfa Düzeni görünümünde hücre düzenlenir; Sonlar > Sayfa Sonu Ekle sayfaları böler", async ({ page }) => {
    await open(page);
    await fillData(page, [["Ürün", "Adet"], ["Çay", "10"], ["Şeker", "4"]]);
    await tab(page, "Görünüm").click();
    await ribbon(page, "Sayfa Düzeni").click();
    await expect(page.locator(".tb-pl-page")).toHaveCount(1);
    await page.locator('.tb-pl td[data-r="1"][data-c="1"]').click();
    expect(await selection(page)).toBe("B2");
    await page.keyboard.type("77");
    await page.keyboard.press("Enter");
    expect((await cell(page, "B2"))!.value).toBe(77);
    expect(await selection(page)).toBe("B3");
    await expect(page.locator('.tb-pl td[data-r="1"][data-c="1"]')).toHaveText("77");
    // B3'ün üstüne ve soluna sayfa sonu: 2 x 2 sayfa
    await tab(page, "Sayfa Düzeni").click();
    await ribbon(page, "Sonlar").click();
    await page.getByRole("menuitem", { name: "Sayfa Sonu Ekle" }).click();
    await expect(page.locator(".tb-pl-page")).toHaveCount(4);
    await expect(page.locator(".tb-pl-num").first()).toHaveText("Sayfa 1 / 4");
    await page.keyboard.press("Control+z");
    await expect(page.locator(".tb-pl-page")).toHaveCount(1);
  });
});

tabloTur4Tests();
