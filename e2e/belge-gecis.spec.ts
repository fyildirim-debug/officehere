import { readFileSync } from "node:fs";
import { resolve } from "node:path";
import { expect, test, type Page } from "@playwright/test";
import { URL, collectErrors, zipEntry } from "./ortak";

// OfficeHere Belge geçiş katmanı (paket düzeyi): dijital imzalı belgede açılış uyarısı ve imzanın kaydedilen
// dosyadan düşmesi, .docm'nin .docm olarak makrolarıyla kaydedilmesi, kurtarılan taslakta özgün paketin
// kurtarma deposundan kullanılması ve paket bulunamazsa uyarı.

const POI = resolve(__dirname, "../test-derlem/poi");
const poi = (name: string) => readFileSync(resolve(POI, name));

test.describe("OfficeHere Belge geçiş katmanı", () => {
  let errors: string[] = [];

  test.beforeEach(async ({ page, context }) => {
    errors = collectErrors(page);
    await context.addInitScript(() => {
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
  const bar = (page: Page) => page.locator(".oh-messagebar");
  const open = (page: Page, name: string, buffer: Buffer) => page.setInputFiles("#bl-open-input", { name, mimeType: "application/octet-stream", buffer });
  /** Ctrl+S ile indirilen dosya: adı ve baytları. */
  const save = async (page: Page, out: string) => {
    const download = page.waitForEvent("download");
    await page.keyboard.press("Control+s");
    const file = await download;
    await file.saveAs(out);
    return { name: file.suggestedFilename(), bytes: readFileSync(out) };
  };
  /** Kurtarma deposundaki (IndexedDB "kaynaklar") özgün paket sayısı; `clear` ile boşaltır. */
  const sources = (page: Page, clear = false) =>
    page.evaluate(
      (clear) =>
        new Promise<number>((res) => {
          const r = indexedDB.open("officehere-belge");
          r.onerror = () => res(-1);
          r.onsuccess = () => {
            const db = r.result;
            if (!db.objectStoreNames.contains("kaynaklar")) {
              db.close();
              return res(0);
            }
            const tx = db.transaction("kaynaklar", "readwrite");
            const s = tx.objectStore("kaynaklar");
            const q = s.count();
            q.onsuccess = () => {
              if (clear) s.clear();
            };
            tx.oncomplete = () => {
              db.close();
              res(q.result);
            };
          };
        }),
      clear,
    );
  /** Belgeyi değiştirip taslağı hemen yazdırır, pencereyi kapatır; yeni pencerede kurtarma çubuğunu döner. */
  const crashAndReopen = async (page: Page, text: string) => {
    await body(page).click();
    await page.keyboard.press("Control+End");
    await page.keyboard.type(text);
    await page.evaluate(() => window.dispatchEvent(new Event("officehere:autosave")));
    await expect.poll(() => sources(page)).toBe(1);
    const next = await page.context().newPage();
    errors.push(...collectErrors(next));
    await page.close();
    await next.goto(URL.belge);
    const recovered = next.locator(".oh-messagebar", { hasText: "Kurtarılan belge" });
    await expect(recovered).toBeVisible();
    return { next, recovered };
  };

  test("dijital imzalı belge: açılışta uyarılır; kaydedilen dosyada imza yok, tema ve yazı tipi tablosu korunur", async ({ page }, testInfo) => {
    const src = poi("hello-world-signed.docx");
    expect(zipEntry(src, "_rels/.rels")).toContain("digital-signature/origin");
    await open(page, "imzali.docx", src);
    await expect(page.locator(".oh-titlebar")).toContainText("imzali");
    await expect(bar(page)).toContainText("Kaydederseniz dijital imza kaldırılır");
    const { bytes } = await save(page, testInfo.outputPath("imzali.docx"));
    expect(zipEntry(bytes, "_rels/.rels")).not.toContain("digital-signature");
    expect(zipEntry(bytes, "[Content_Types].xml")).not.toContain("digital-signature");
    expect(zipEntry(bytes, "word/theme/theme1.xml")).toContain("a:theme");
    expect(zipEntry(bytes, "word/fontTable.xml")).toContain("w:fonts");
  });

  test(".docm: makroların çalıştırılmadığı bildirilir; Ctrl+S .docm olarak makro projesiyle kaydeder", async ({ page }, testInfo) => {
    const src = poi("SimpleMacro.docm");
    await open(page, "SimpleMacro.docm", src);
    await expect(page.locator(".oh-titlebar")).toContainText("SimpleMacro");
    await expect(bar(page)).toContainText("makrolar çalıştırılmaz");
    const { name, bytes } = await save(page, testInfo.outputPath("SimpleMacro.docm"));
    expect(name).toBe("SimpleMacro.docm");
    expect(zipEntry(bytes, "word/vbaProject.bin")).not.toBeNull();
    expect(zipEntry(bytes, "word/_rels/document.xml.rels")).toContain("relationships/vbaProject");
    expect(zipEntry(bytes, "[Content_Types].xml")).toContain('PartName="/word/document.xml" ContentType="application/vnd.ms-word.document.macroEnabled.main+xml"');
  });

  test("kurtarılan taslak: özgün paket kurtarma deposundan kullanılır, kaydedilen dosyada tema korunur", async ({ page }, testInfo) => {
    const src = poi("Headers.docx");
    expect(zipEntry(src, "word/theme/theme1.xml")).not.toBeNull();
    await open(page, "Basliklar.docx", src);
    await expect(page.locator(".oh-titlebar")).toContainText("Basliklar");
    const { next, recovered } = await crashAndReopen(page, " KURTARILACAK");
    await recovered.getByRole("button", { name: "Geri Yükle" }).click();
    await expect(body(next)).toContainText("KURTARILACAK");
    await expect(next.locator(".oh-messagebar", { hasText: "kurtarma deposunda bulunamadı" })).toHaveCount(0);
    const { bytes } = await save(next, testInfo.outputPath("kurtarilan.docx"));
    expect(zipEntry(bytes, "word/document.xml")).toContain("KURTARILACAK");
    expect(zipEntry(bytes, "word/theme/theme1.xml")).toContain("a:theme");
    expect(zipEntry(bytes, "word/_rels/document.xml.rels")).toContain("relationships/theme");
  });

  test("kurtarılan taslağın özgün paketi bulunamazsa belge taslak içeriğiyle açılır ve uyarılır", async ({ page }, testInfo) => {
    await open(page, "Basliklar.docx", poi("Headers.docx"));
    await expect(page.locator(".oh-titlebar")).toContainText("Basliklar");
    const { next, recovered } = await crashAndReopen(page, " YEDEK GÖRÜNÜM");
    // Paket depoda yok (kota ya da silinmiş depo): boşaltılır, kapanan pencerenin son yazımı da beklenir.
    expect(await sources(next, true)).toBe(1);
    await expect.poll(() => sources(next, true)).toBe(0);
    await recovered.getByRole("button", { name: "Geri Yükle" }).click();
    await expect(body(next)).toContainText("YEDEK GÖRÜNÜM");
    await expect(next.locator(".oh-messagebar", { hasText: "kurtarma deposunda bulunamadı" })).toBeVisible();
    // Kaydetmek yine çalışır (yalnız yazıcının ürettiği parçalarla).
    const { bytes } = await save(next, testInfo.outputPath("yedek.docx"));
    expect(zipEntry(bytes, "word/document.xml")).toContain("YEDEK GÖRÜNÜM");
    expect(zipEntry(bytes, "word/theme/theme1.xml")).toBeNull();
  });
});
