import { readFileSync } from "node:fs";
import { createRequire } from "node:module";
import { resolve } from "node:path";
import { expect, test, type Page } from "@playwright/test";
import { URL, collectErrors } from "./ortak";

// OfficeHere Tablo 4. tur: XLSX öğe sırası (K1), şerit girişleri (Y1), Aralığa
// Dönüştür (Y2), kurtarma taslakları (Y3) ve tabloda Tab (O1). tablo.spec.ts
// içinden kaydedilir; ızgara canvas olduğu için değerler window.__ohTablo'dan okunur.

// jszip uygulamanın bağımlılığıdır; pnpm yerleşiminde uygulama dizininden çözülür
interface ZipFile {
  async(type: "string"): Promise<string>;
}
interface Zip {
  file(path: string): ZipFile | null;
  files: Record<string, unknown>;
}
const JSZip = createRequire(resolve(__dirname, "..", "apps/tablo/package.json"))("jszip") as { loadAsync(data: Buffer): Promise<Zip> };

/** ECMA-376 CT_Worksheet alt öğe sırası (§18.3.1.99). */
const WORKSHEET_ORDER = [
  "sheetPr", "dimension", "sheetViews", "sheetFormatPr", "cols", "sheetData", "sheetCalcPr", "sheetProtection", "protectedRanges",
  "scenarios", "autoFilter", "sortState", "dataConsolidate", "customSheetViews", "mergeCells", "phoneticPr", "conditionalFormatting",
  "dataValidations", "hyperlinks", "printOptions", "pageMargins", "pageSetup", "headerFooter", "rowBreaks", "colBreaks",
  "customProperties", "cellWatches", "ignoredErrors", "smartTags", "drawing", "legacyDrawing", "legacyDrawingHF", "drawingHF",
  "picture", "oleObjects", "controls", "webPublishItems", "tableParts", "extLst",
];

/** Kök öğenin doğrudan alt öğelerinin adları (önek atılmış). */
function topLevel(xml: string): string[] {
  const out: string[] = [];
  let depth = 0;
  for (const m of xml.replace(/<\?[\s\S]*?\?>|<!--[\s\S]*?-->/g, "").matchAll(/<(\/?)([\w:.-]+)[^>]*?(\/?)>/g)) {
    const name = m[2].replace(/^\w+:/, "");
    if (m[1]) depth--;
    else {
      if (depth === 1) out.push(name);
      if (!m[3]) depth++;
    }
  }
  return out;
}

interface CellInfo {
  value: string | number | boolean | null;
  formula: string | null;
  edit: string;
}
type Api = { cell(ref: string): CellInfo | null; cellRect(ref: string): { x: number; y: number; w: number; h: number } | null; selection(): string; chartCount(): number };
const api = "__ohTablo";

export function tabloTur4Tests() {
  test.describe("OfficeHere Tablo 4. tur", () => {
    let errors: string[] = [];
    test.beforeEach(async ({ page }) => {
      errors = collectErrors(page);
      await page.addInitScript(() => {
        const w = window as unknown as Record<string, unknown>;
        delete w.showSaveFilePicker;
        delete w.showOpenFilePicker;
      });
      await open(page);
    });
    test.afterEach(() => {
      expect(errors).toEqual([]);
    });

    const t = (page: Page) => page.evaluate((k) => !!(window as unknown as Record<string, Api | undefined>)[k]?.cellRect("A1"), api);
    async function open(page: Page) {
      await page.goto(URL.tablo);
      await expect.poll(() => t(page)).toBe(true);
    }
    const cell = (page: Page, ref: string) => page.evaluate(([k, r]) => (window as unknown as Record<string, Api>)[k].cell(r), [api, ref] as const);
    const selection = (page: Page) => page.evaluate((k) => (window as unknown as Record<string, Api>)[k].selection(), api);
    const tab = (page: Page, name: string) => page.getByRole("tab", { name, exact: true });
    const ribbon = (page: Page, name: string) => page.locator(`.oh-ribbon-body button[aria-label="${name}"]`).first();
    const editorOpen = (page: Page) => page.locator(".tb-editor.is-visible");
    /** Odak çalışma sayfasında (ızgaranın klavye girişi). */
    const gridFocused = (page: Page) => expect(page.getByRole("textbox", { name: "Hücre düzenleyici" })).toBeFocused();
    const recoverBar = (page: Page) => page.locator(".oh-messagebar", { hasText: "Kurtarılan çalışma kitabı" });

    async function goTo(page: Page, ref: string) {
      const box = page.getByRole("textbox", { name: "Ad Kutusu" });
      await box.fill(ref);
      await box.press("Enter");
    }
    async function enter(page: Page, ref: string, text: string) {
      await goTo(page, ref);
      await page.keyboard.type(text);
      await page.keyboard.press("Enter");
    }
    async function fillData(page: Page, rows: string[][]) {
      for (let i = 0; i < rows.length; i++)
        for (let j = 0; j < rows[i].length; j++) if (rows[i][j] !== "") await enter(page, `${String.fromCharCode(65 + j)}${i + 1}`, rows[i][j]);
    }
    async function makeTable(page: Page, at = "A1") {
      await goTo(page, at);
      await page.keyboard.press("Control+t");
      await page.getByRole("dialog", { name: "Tablo Oluştur" }).getByRole("button", { name: "Tamam" }).click();
    }
    /** IndexedDB'deki kurtarma taslağı sayısı (veritabanı yoksa oluşturmadan 0). */
    const draftCount = (page: Page) =>
      page.evaluate(async () => {
        const dbs = await indexedDB.databases();
        if (!dbs.some((d) => d.name === "officehere-tablo")) return 0;
        return new Promise<number>((res) => {
          const r = indexedDB.open("officehere-tablo");
          r.onerror = () => res(-1);
          r.onsuccess = () => {
            const db = r.result;
            if (!db.objectStoreNames.contains("drafts")) {
              db.close();
              return res(0);
            }
            const q = db.transaction("drafts", "readonly").objectStore("drafts").count();
            q.onsuccess = () => {
              db.close();
              res(q.result);
            };
          };
        });
      });

    test("K1: tüm özellikler aynı sayfadayken indirilen XLSX'te sayfa öğeleri şema sırasında", async ({ page }) => {
      await fillData(page, [["Ay", "Satış", "Hedef"], ["Ocak", "100", "90"], ["Şubat", "150", "120"], ["Mart", "120", "130"]]);
      // Veri çubuğu ve simge kümesi
      await goTo(page, "B2:B4");
      await ribbon(page, "Koşullu Biçimlendirme").click();
      await page.getByRole("menuitem", { name: "Veri Çubukları" }).hover();
      await page.locator(".oh-menu .oh-gallery-item").first().click();
      await goTo(page, "C2:C4");
      await ribbon(page, "Koşullu Biçimlendirme").click();
      await page.getByRole("menuitem", { name: "Simge Kümeleri" }).hover();
      await page.locator(".oh-menu button[title]").first().click();
      // Birleştirme, grafik, mini grafik, sayfa sonu, tablo
      await goTo(page, "G1:H1");
      await ribbon(page, "Birleştir ve Ortala").click();
      await goTo(page, "A1");
      await tab(page, "Ekle").click();
      await ribbon(page, "Sütun veya Çubuk").click();
      await page.locator(".oh-menu-item", { hasText: "Kümelenmiş Sütun" }).first().click();
      expect(await page.evaluate((k) => (window as unknown as Record<string, Api>)[k].chartCount(), api)).toBe(1);
      await goTo(page, "B2:C2");
      await tab(page, "Ekle").click();
      await ribbon(page, "Çizgi").click();
      const sp = page.getByRole("dialog", { name: "Mini Grafikler Oluştur" });
      await sp.getByLabel("Konum Aralığı").fill("E2");
      await sp.getByRole("button", { name: "Tamam" }).click();
      await goTo(page, "A3");
      await tab(page, "Sayfa Düzeni").click();
      await ribbon(page, "Sonlar").click();
      await page.getByRole("menuitem", { name: "Sayfa Sonu Ekle" }).click();
      await makeTable(page);
      // İndir
      await page.locator(".oh-ribbon-file").click();
      await page.locator(".oh-bs-item", { hasText: "Dışarı Aktar" }).first().click();
      const [download] = await Promise.all([page.waitForEvent("download"), page.getByText("Çalışma kitabını indir (.xlsx)").click()]);
      const zip = await JSZip.loadAsync(readFileSync((await download.path())!));
      const sheet = await zip.file("xl/worksheets/sheet1.xml")!.async("string");
      const names = topLevel(sheet);
      for (const n of ["sheetData", "mergeCells", "conditionalFormatting", "pageSetup", "rowBreaks", "drawing", "tableParts", "extLst"]) expect(names, n).toContain(n);
      // Her öğe şemada tanımlı ve sıra bozulmuyor
      const idx = names.map((n) => WORKSHEET_ORDER.indexOf(n));
      expect(idx.filter((i) => i < 0), names.join(",")).toEqual([]);
      for (let i = 1; i < idx.length; i++) expect(idx[i] >= idx[i - 1], `${names[i - 1]} → ${names[i]}`).toBe(true);
      // Veri çubuğu: boş x14:id yok; her kimliğin x14 kuralı sayfa düzeyindeki extLst'te
      const ids = [...sheet.matchAll(/<x14:id>([^<]*)<\/x14:id>/g)].map((m) => m[1]);
      expect(ids.length).toBeGreaterThan(0);
      for (const id of ids) {
        expect(id).toMatch(/^\{[0-9A-F-]{36}\}$/);
        expect(sheet).toContain(`id="${id}"`);
      }
      expect(sheet).toContain("<x14:conditionalFormattings>");
      expect(sheet).toContain("<x14:sparklineGroups");
      expect(sheet.indexOf("<x14:conditionalFormattings>")).toBeGreaterThan(sheet.lastIndexOf("</conditionalFormatting>"));
      // Çalışma kitabı XML'i de şema sırasında
      const wbNames = topLevel(await zip.file("xl/workbook.xml")!.async("string"));
      const WB = ["fileVersion", "fileSharing", "workbookPr", "workbookProtection", "bookViews", "sheets", "functionGroups", "externalReferences", "definedNames", "calcPr", "oleSize", "customWorkbookViews", "pivotCaches", "smartTagPr", "smartTagTypes", "webPublishing", "fileRecoveryPr", "webPublishObjects", "extLst"];
      const wi = wbNames.map((n) => WB.indexOf(n));
      expect(wi.filter((i) => i < 0), wbNames.join(",")).toEqual([]);
      for (let i = 1; i < wi.length; i++) expect(wi[i] >= wi[i - 1], `${wbNames[i - 1]} → ${wbNames[i]}`).toBe(true);
    });

    test("Y1: Tablo Adı kutusunda Enter hücre düzenlemesi başlatmaz; boşluklu ad uyarı verir", async ({ page }) => {
      await fillData(page, [["Bölge", "Tutar"], ["Ege", "10"], ["Akdeniz", "20"]]);
      await makeTable(page, "B2");
      await tab(page, "Tablo Tasarımı").click();
      const name = page.getByRole("textbox", { name: "Tablo Adı" });
      await name.fill("Satis");
      await name.press("Enter");
      await expect(editorOpen(page)).toHaveCount(0);
      await gridFocused(page);
      // Odak ızgaraya döndü: ok tuşu seçimi taşır, hücre değeri değişmez
      await page.keyboard.press("ArrowDown");
      expect(await selection(page)).toBe("B3");
      expect((await cell(page, "B2"))!.value).toBe(10);
      await enter(page, "D1", "=TOPLA(Satis[Tutar])");
      expect((await cell(page, "D1"))!.value).toBe(30);
      // Boşluklu ad reddedilir, eski ad kalır
      await goTo(page, "B2");
      await tab(page, "Tablo Tasarımı").click();
      await name.fill("Satış Tablosu");
      await name.press("Enter");
      const dlg = page.getByRole("dialog");
      await expect(dlg).toContainText("boşluk");
      await dlg.getByRole("button", { name: "Tamam" }).click();
      await expect(name).toHaveValue("Satis");
      await expect(editorOpen(page)).toHaveCount(0);
      expect((await cell(page, "B2"))!.value).toBe(10);
      // Yazı tipi boyutu kutusunda da Enter düzenleme başlatmaz
      await tab(page, "Giriş").click();
      const size = page.locator('.oh-ribbon-body input[aria-label="Yazı Tipi Boyutu"]').first();
      await size.fill("14");
      await size.press("Enter");
      await gridFocused(page);
      await expect(editorOpen(page)).toHaveCount(0);
      await page.keyboard.press("ArrowRight");
      expect(await selection(page)).toBe("C2");
    });

    test("Y2: Aralığa Dönüştür yapısal başvuruları A1 başvurusuna çevirir", async ({ page }) => {
      await fillData(page, [["Ürün", "Adet", "Fiyat"], ["Çay", "10", "2,5"], ["Şeker", "4", "30"]]);
      await makeTable(page);
      await enter(page, "D1", "Tutar");
      await enter(page, "D2", "=[@Adet]*[@Fiyat]");
      await enter(page, "F1", "=TOPLA(Tablo1[Tutar])");
      expect((await cell(page, "F1"))!.value).toBe(145);
      await goTo(page, "B2");
      await tab(page, "Tablo Tasarımı").click();
      await ribbon(page, "Aralığa Dönüştür").click();
      await page.getByRole("dialog").getByRole("button", { name: "Evet" }).click();
      await expect(tab(page, "Tablo Tasarımı")).toHaveCount(0);
      expect(await cell(page, "F1")).toMatchObject({ value: 145, formula: "=SUM($D$2:$D$3)" });
      expect(await cell(page, "D3")).toMatchObject({ value: 120, formula: "=$B3*$C3" });
      // Geri alınca tablo ve yapısal başvuru geri gelir
      await goTo(page, "H10");
      await page.keyboard.press("Control+z");
      expect((await cell(page, "F1"))!.formula).toBe("=SUM(Tablo1[Tutar])");
    });

    test("Y3: kaydedince taslak silinir; kapanan pencerenin taslağı önerilir ve Sil ile kalıcı silinir", async ({ page }) => {
      await enter(page, "A1", "Merhaba");
      await expect.poll(() => draftCount(page), { timeout: 20_000 }).toBe(1);
      const [dl] = await Promise.all([page.waitForEvent("download"), page.keyboard.press("Control+s")]);
      expect(dl.suggestedFilename()).toBe("Kitap1.xlsx");
      await expect.poll(() => draftCount(page)).toBe(0);
      await page.reload();
      await expect.poll(() => t(page)).toBe(true);
      await page.waitForTimeout(1200);
      await expect(recoverBar(page)).toHaveCount(0);
      // Kaydedilmemiş değişiklik → yenile → çubuk → Sil
      await enter(page, "A2", "Dünya");
      await expect.poll(() => draftCount(page), { timeout: 20_000 }).toBe(1);
      await page.reload();
      await expect(recoverBar(page)).toBeVisible();
      await recoverBar(page).getByRole("button", { name: "Sil", exact: true }).click();
      await expect(recoverBar(page)).toHaveCount(0);
      await expect.poll(() => draftCount(page)).toBe(0);
      await page.reload();
      await expect.poll(() => t(page)).toBe(true);
      await page.waitForTimeout(1200);
      await expect(recoverBar(page)).toHaveCount(0);
    });

    test("Y3: açık bir pencerenin taslağı diğer pencerede önerilmez", async ({ page, context }) => {
      await enter(page, "B1", "Pencere1");
      await expect.poll(() => draftCount(page), { timeout: 20_000 }).toBe(1);
      const p2 = await context.newPage();
      await p2.goto(URL.tablo);
      await expect.poll(() => t(p2)).toBe(true);
      await p2.waitForTimeout(1500);
      await expect(recoverBar(p2)).toHaveCount(0);
      await p2.close();
    });

    test("O1: tablonun son hücresinde Tab yeni satır ekler, hesaplanan sütun yeni satıra gelir", async ({ page }) => {
      await fillData(page, [["Ürün", "Adet"], ["Çay", "10"], ["Şeker", "4"]]);
      await makeTable(page);
      await enter(page, "C1", "Tutar");
      await enter(page, "C2", "=[@Adet]*2");
      expect(await cell(page, "C3")).toMatchObject({ value: 8, edit: "=[@Adet]*2" });
      // Satır sonunda Tab sonraki satırın ilk hücresine geçer
      await goTo(page, "C2");
      await page.keyboard.press("Tab");
      expect(await selection(page)).toBe("A3");
      // Son hücrede Tab yeni tablo satırı açar
      await goTo(page, "C3");
      await page.keyboard.press("Tab");
      expect(await selection(page)).toBe("A4");
      await expect(tab(page, "Tablo Tasarımı")).toBeVisible();
      expect((await cell(page, "C4"))!.edit).toBe("=[@Adet]*2");
      await page.keyboard.type("Un");
      await page.keyboard.press("Tab");
      await page.keyboard.type("3");
      await page.keyboard.press("Enter");
      expect((await cell(page, "C4"))!.value).toBe(6);
    });
  });
}
