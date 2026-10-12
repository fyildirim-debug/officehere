import { expect, test, type Page } from "@playwright/test";
import { mkdtempSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { URL, collectErrors } from "./ortak";

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
  test("değer yazma, Enter ile aşağı inme ve sayı/tarih/yüzde tanıma", async ({ page }) => {
    const errors = collectErrors(page);
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
    expect(errors).toEqual([]);
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

  test("PivotTable: oluştur, alan seç, sütuna sürükle, düzenleme engeli", async ({ page }) => {
    await open(page);
    await fillData(page, [["Bölge", "Yıl", "Satış"], ["Ege", "2024", "100"], ["Ege", "2025", "50"], ["Marmara", "2024", "200"]]);
    await clickCell(page, "A1");
    await tab(page, "Ekle").click();
    await ribbon(page, "PivotTable").click();
    await page.getByRole("dialog", { name: "PivotTable Oluştur" }).getByRole("button", { name: "Tamam" }).click();
    await expect(page.getByRole("complementary", { name: "PivotTable Alanları" })).toBeVisible();
    expect(await page.evaluate(() => window.__ohTablo!.activeSheet())).toBe("Sayfa2");
    await page.locator(".tb-pv-field", { hasText: "Bölge" }).locator("input").check();
    await page.locator(".tb-pv-field", { hasText: "Satış" }).locator("input").check();
    await expect.poll(() => page.evaluate(() => window.__ohTablo!.cell("B6")?.value)).toBe(350);
    expect((await cell(page, "A3"))!.value).toBe("Satır Etiketleri");
    expect((await cell(page, "A4"))!.value).toBe("Ege");
    await page.locator(".tb-pv-field", { hasText: "Yıl" }).dragTo(page.locator('.tb-pv-area[data-area="cols"]'));
    await expect.poll(() => page.evaluate(() => window.__ohTablo!.cell("B4")?.value)).toBe(2024);
    expect((await cell(page, "D7"))!.value).toBe(350);
    await expect(tab(page, "PivotTable Analiz")).toBeVisible();
    await clickCell(page, "B5");
    await page.keyboard.type("9");
    await expect(page.getByRole("dialog").getByText("PivotTable raporunun bu bölümü değiştirilemez", { exact: false })).toBeVisible();
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
});
