import { readFileSync } from "node:fs";
import { createRequire } from "node:module";
import { resolve } from "node:path";
import { expect, test, type Page } from "@playwright/test";
import { URL, collectErrors, interceptFileChooser } from "./ortak";

// OfficeHere Sunum dosya G/Ç sağlamlığı: hasarlı dosyalar sessiz kayıp olmadan
// (uyarı + yerine boş slayt ya da paket içinden kurtarma) açılır; bozuk ya da
// zararlı dosyalar Türkçe, sınıflandırılmış bir iletiyle reddedilir; dev
// parçalar (zip bombası) takılmadan atlanır. Örnekler testte LibreOffice ile
// üretilmiş kendi fikstürlerimizden türetilir.

/** Testte kullanılan JSZip yüzeyi (paket uygulamanın bağımlılığından yüklenir). */
interface Zip {
  file(path: string): { async(type: "string"): Promise<string> } | null;
  file(path: string, data: string): unknown;
  remove(path: string): unknown;
  generateAsync(opts: { type: "nodebuffer"; compression: "DEFLATE" }): Promise<Buffer>;
}
const JSZip = createRequire(resolve(__dirname, "..", "apps/sunum/package.json"))("jszip") as { loadAsync(data: Buffer): Promise<Zip> };
const fixture = (n: string) => readFileSync(resolve(__dirname, "..", "apps/sunum/src/io/__fixtures__", n));
const poi = (n: string) => readFileSync(resolve(__dirname, "..", "test-derlem/poi", n));
const thumbs = (page: Page) => page.locator(".thumb-pane .thumb-item");
const stage = (page: Page) => page.locator(".stage .sv");
const bar = (page: Page) => page.locator(".oh-messagebar");

async function mutate(name: string, fn: (z: Zip) => void | Promise<void>): Promise<Buffer> {
  const z = await JSZip.loadAsync(fixture(name));
  await fn(z);
  return z.generateAsync({ type: "nodebuffer", compression: "DEFLATE" });
}

/** ZIP girdisinin sıkıştırılmış verisini bozar (bayt bozulması). */
function corruptEntry(zip: Buffer, name: string): Buffer {
  const out = Buffer.from(zip);
  for (let i = 0; i + 30 < out.length; i++) {
    if (out.readUInt32LE(i) !== 0x04034b50) continue;
    const nameLen = out.readUInt16LE(i + 26);
    const extraLen = out.readUInt16LE(i + 28);
    if (out.toString("utf8", i + 30, i + 30 + nameLen) !== name) continue;
    const start = i + 30 + nameLen + extraLen;
    const size = out.readUInt32LE(i + 18);
    for (let k = 2; k < Math.min(size, 600); k++) out[start + k] = 0xff;
    return out;
  }
  throw new Error(`girdi yok: ${name}`);
}

async function openApp(page: Page) {
  await page.addInitScript(() => {
    const w = window as unknown as Record<string, unknown>;
    delete w.showSaveFilePicker;
    delete w.showOpenFilePicker;
  });
  await page.goto(URL.sunum);
  await expect(stage(page)).toBeVisible();
  await page.locator(".stage").click({ position: { x: 4, y: 4 } });
}

async function openFile(page: Page, name: string, buffer: Buffer) {
  const [chooser] = await Promise.all([page.waitForEvent("filechooser"), page.keyboard.press("Control+o")]);
  await chooser.setFiles({ name, mimeType: "application/octet-stream", buffer });
}

test.describe("OfficeHere Sunum dosya G/Ç sağlamlığı", () => {
  let errors: string[] = [];

  test.beforeEach(async ({ page }) => {
    errors = collectErrors(page);
    interceptFileChooser(page);
    await openApp(page);
  });

  test.afterEach(() => {
    expect(errors).toEqual([]);
  });

  test("sağlam dosya uyarısız açılır", async ({ page }) => {
    await openFile(page, "lo-proje.pptx", fixture("lo-proje.pptx"));
    await expect(thumbs(page)).toHaveCount(7);
    await expect(page.locator(".oh-messagebar.is-warning, .oh-messagebar.is-error")).toHaveCount(0);
  });

  test("silinen slayt parçası: yerine boş slayt konur, uyarı ve ayrıntı gösterilir", async ({ page }) => {
    const bytes = await mutate("lo-proje.pptx", (z) => void z.remove("ppt/slides/slide3.xml"));
    await openFile(page, "hasarli.pptx", bytes);
    await expect(thumbs(page)).toHaveCount(7);
    await expect(bar(page)).toHaveClass(/is-warning/);
    await expect(bar(page)).toContainText("Dosya hasarlı");
    await expect(bar(page)).toContainText("3. slayt dosyada bulunamadı; yerine boş slayt kondu.");
    await expect(bar(page).locator("pre")).toBeHidden();
    await bar(page).locator("summary", { hasText: "Ayrıntılar" }).click();
    await expect(bar(page).locator("pre")).toBeVisible();
    await expect(bar(page).locator("pre")).toContainText("ppt/slides/slide3.xml");
  });

  test("bozuk sunu ilişkileri: tüm slaytlar paket içinden kurtarılır", async ({ page }) => {
    const bytes = await mutate("lo-proje.pptx", async (z) => {
      const rels = await z.file("ppt/_rels/presentation.xml.rels")!.async("string");
      z.file("ppt/_rels/presentation.xml.rels", rels.replace(/Target="[^"]+"/g, 'Target="../../yok/olmayan.xml"'));
    });
    await openFile(page, "iliskiler.pptx", bytes);
    await expect(thumbs(page)).toHaveCount(7);
    await expect(bar(page)).toContainText("paket içinden bulunarak açıldı");
    await expect(stage(page)).toContainText("Proje");
  });

  test("bozuk sıkıştırılmış veri: ham teknik ileti yerine Türkçe hata", async ({ page }) => {
    const odp = await (await JSZip.loadAsync(fixture("lo-bicim.odp"))).generateAsync({ type: "nodebuffer", compression: "DEFLATE" });
    await openFile(page, "bozuk.odp", corruptEntry(odp, "content.xml"));
    await expect(bar(page)).toHaveClass(/is-error/);
    await expect(bar(page)).toContainText("Dosya açılamadı");
    // Görünen ileti Türkçe; özgün teknik ileti yalnız (kapalı) "Ayrıntılar" içinde.
    await expect(bar(page)).not.toContainText(/Bug|mismatch|invalid|undefined/i, { useInnerText: true });
    await bar(page).locator("summary", { hasText: "Ayrıntılar" }).click();
    await expect(bar(page).locator("pre")).toContainText("size mismatch");
    await expect(thumbs(page)).toHaveCount(1);
  });

  test("dev slayt parçası (zip bombası) takılmadan atlanır", async ({ page }) => {
    test.setTimeout(120_000);
    const bytes = await mutate("lo-proje.pptx", async (z) => {
      const x = await z.file("ppt/slides/slide2.xml")!.async("string");
      z.file("ppt/slides/slide2.xml", x + " ".repeat(140 * 1024 * 1024));
    });
    const t0 = Date.now();
    await openFile(page, "bomba.pptx", bytes);
    await expect(bar(page)).toContainText("2. slayt olağandışı büyük olduğu için okunmadı", { timeout: 15_000 });
    expect(Date.now() - t0).toBeLessThan(15_000);
    await expect(thumbs(page)).toHaveCount(7);
  });

  test("makrolu sunu .pptm olarak kaydedilir: makro projesi korunur, ileti çalıştırılmadığını söyler", async ({ page }) => {
    await openFile(page, "SimpleMacro.pptm", poi("SimpleMacro.pptm"));
    await expect(bar(page)).toContainText("makrolar kaydederken korunur, ancak OfficeHere'da çalıştırılmaz");
    const [download] = await Promise.all([page.waitForEvent("download"), page.keyboard.press("Control+s")]);
    expect(download.suggestedFilename()).toBe("SimpleMacro.pptm");
    const z = await JSZip.loadAsync(readFileSync((await download.path())!));
    expect(z.file("ppt/vbaProject.bin")).not.toBeNull();
    expect(await z.file("[Content_Types].xml")!.async("string")).toContain("application/vnd.ms-powerpoint.presentation.macroEnabled.main+xml");
  });

  test("dijital imzalı sunu: açılışta bildirilir, kaydedince imzanın kaldırıldığı söylenir", async ({ page }) => {
    await openFile(page, "hello-world-signed.pptx", poi("hello-world-signed.pptx"));
    await expect(bar(page)).toContainText("dijital olarak imzalı");
    const [download] = await Promise.all([page.waitForEvent("download"), page.keyboard.press("Control+s")]);
    expect(download.suggestedFilename()).toBe("hello-world-signed.pptx");
    await expect(bar(page)).toContainText("Sunu kaydedildi.");
    await expect(bar(page)).toContainText(/imza/);
    const z = await JSZip.loadAsync(readFileSync((await download.path())!));
    expect(z.file("_xmlsignatures/origin.sigs")).toBeNull();
  });

  test("kod bölme: okuyucu/yazıcı parçası ayrı yüklenir, açma ve kaydetme çalışır", async ({ page }) => {
    // İlk yükleme yalnız ana parça(lar)la; okuyucu/yazıcı parçası boşta ya da ilk açmada gelir.
    const loaded = await page.evaluate(() => performance.getEntriesByType("resource").map((r) => r.name.split("/").pop() ?? ""));
    expect(loaded.some((n) => n.startsWith("index-") && n.endsWith(".js"))).toBe(true);
    await openFile(page, "lo-bicim.odp", fixture("lo-bicim.odp"));
    await expect(thumbs(page)).toHaveCount(4);
    const after = await page.evaluate(() => performance.getEntriesByType("resource").map((r) => r.name.split("/").pop() ?? ""));
    expect(after.some((n) => n.startsWith("formatsImpl-"))).toBe(true);
    // Kaydet (Ctrl+S → .odp biçim sorusu → indirme) tembel yazıcıyla çalışır.
    await page.keyboard.press("Control+s");
    const dlg = page.getByRole("dialog");
    const [download] = await Promise.all([page.waitForEvent("download"), dlg.getByRole("button", { name: ".odp olarak kaydet" }).click()]);
    expect(download.suggestedFilename()).toBe("lo-bicim.odp");
  });
});
