import { _electron as electron, expect, test, type ElectronApplication, type Page } from "@playwright/test";
import type {} from "../packages/core/src/bridge";
import { copyFileSync, existsSync, mkdtempSync } from "node:fs";
import { tmpdir } from "node:os";
import path from "node:path";
import { pdfBilgi } from "./pdf-incele";

// PDF dışa aktarma (masaüstü): ortak seçenek iletişim kutusu, PDF/A-2b ve PDF/A-1b, sayfa aralığı,
// başlıklardan yer imleri, belge bilgileri ve "dışa aktardıktan sonra". Konum ve ileti kutuları ana
// süreçte sahteyle değiştirilir (otomasyonla tıklanamaz); PDF dosyası diskteki çıktıdan denetlenir.

const root = path.resolve(__dirname, "..");
const desktop = path.join(root, "apps/desktop");
const appsDir = process.env.OH_E2E_APPS_DIR ?? path.join(root, "dist-web");
const TR = ["ı", "İ", "ş", "ğ", "ç", "ö", "ü"];

function electronPath(): string {
  return require(path.join(desktop, "node_modules/electron")) as unknown as string;
}

function fixture(rel: string): string {
  const dir = mkdtempSync(path.join(tmpdir(), "oh-pdf-"));
  const dest = path.join(dir, path.basename(rel));
  copyFileSync(path.join(root, rel), dest);
  return dest;
}

async function launch(args: string[]): Promise<{ app: ElectronApplication; win: Page }> {
  const userData = mkdtempSync(path.join(tmpdir(), "oh-pdf-veri-"));
  const app = await electron.launch({
    executablePath: electronPath(),
    args: [desktop, "--no-sandbox", ...args],
    env: { ...process.env, OH_APPS_DIR: appsDir, OH_USER_DATA: userData, ELECTRON_DISABLE_SECURITY_WARNINGS: "1" },
  });
  const win = await app.firstWindow();
  await win.waitForLoadState("domcontentloaded");
  await win.waitForSelector(".oh-titlebar");
  return { app, win };
}

type Kayit = { save: { defaultPath?: string; filters?: { name: string }[] }[]; messages: { message: string; detail?: string }[]; opened: string[]; shown: string[]; response: number };

/** Kaydet ve ileti kutuları, dosya açma / klasör gösterme sahteyle değiştirilir; PDF verilen yola yazılır. */
async function stub(app: ElectronApplication, out: string, response = 0) {
  await app.evaluate(
    ({ dialog, shell }, a) => {
      const g = globalThis as unknown as { __pdf: Kayit };
      g.__pdf = { save: [], messages: [], opened: [], shown: [], response: a.response };
      (dialog as unknown as { showSaveDialog: unknown }).showSaveDialog = async (...args: unknown[]) => {
        g.__pdf.save.push((args.length > 1 ? args[1] : args[0]) as Kayit["save"][number]);
        return { canceled: false, filePath: a.out };
      };
      (dialog as unknown as { showMessageBox: unknown }).showMessageBox = async (...args: unknown[]) => {
        const o = (args.length > 1 ? args[1] : args[0]) as { message: string; detail?: string };
        g.__pdf.messages.push({ message: o.message, detail: o.detail });
        return { response: g.__pdf.response, checkboxChecked: false };
      };
      (shell as unknown as { openPath: unknown }).openPath = async (p: string) => {
        g.__pdf.opened.push(p);
        return "";
      };
      (shell as unknown as { showItemInFolder: unknown }).showItemInFolder = (p: string) => {
        g.__pdf.shown.push(p);
      };
    },
    { out, response },
  );
}

/** Düzenlenmiş belge kapatılırken "kaydet?" sorusu beklenmesin: pencereler doğrudan kapatılır. */
async function close(app: ElectronApplication) {
  await app.evaluate(({ BrowserWindow }) => BrowserWindow.getAllWindows().forEach((w) => w.destroy()));
  await app.close();
}

const kayit = (app: ElectronApplication) => app.evaluate(() => (globalThis as unknown as { __pdf: Kayit }).__pdf);

/** Dosya › Dışarı Aktar › PDF; seçenek kutusunda biçim, aralık ve sonrası seçilir. */
async function exportPdf(win: Page, o: { standard: RegExp; range?: string; after?: "none" | "open" | "folder"; title?: string }) {
  await win.locator(".oh-ribbon-file").click();
  await win.locator(".oh-bs-item", { hasText: "Dışarı Aktar" }).first().click();
  await win.locator(".oh-bs-action", { hasText: /PDF.*Oluştur/ }).click();
  const dlg = win.getByRole("dialog", { name: "PDF Olarak Dışa Aktar" });
  await expect(dlg).toBeVisible();
  await dlg.getByRole("radio", { name: o.standard }).check();
  if (o.range) await dlg.getByRole("textbox", { name: "Sayfa aralığı" }).fill(o.range);
  if (o.title !== undefined) await dlg.getByRole("textbox", { name: "Başlık" }).fill(o.title);
  await dlg.getByRole("combobox", { name: "Dışa aktardıktan sonra" }).selectOption(o.after ?? "none");
  await dlg.getByRole("button", { name: "Dışa aktar" }).click();
  await expect(dlg).toHaveCount(0);
}

test("Belge: PDF/A-2b — yer imleri başlık sayısı kadar, Türkçe metin geri okunur, dışa aktarınca açılır", async () => {
  const file = fixture("apps/belge/src/io/__fixtures__/libreoffice-resaved-report.docx");
  const out = path.join(path.dirname(file), "rapor.pdf");
  const { app, win } = await launch([file]);
  await win.waitForSelector(".ProseMirror");
  await win.locator(".bl-body .ProseMirror").first().click();
  await win.keyboard.press("Control+End");
  await win.keyboard.press("Enter");
  await win.keyboard.type("Çalışma ığüşöç İĞÜŞÖÇ sonu.");
  const headings = await win.locator(".bl-body .ProseMirror :is(h1, h2, h3, h4, h5, h6)").allTextContents();
  expect(headings.length).toBe(5);
  await stub(app, out);
  await exportPdf(win, { standard: /^PDF\/A-2b/, after: "open" });
  await expect.poll(() => existsSync(out)).toBe(true);
  await expect(win.locator(".oh-messagebar, [role=status]").filter({ hasText: "rapor.pdf" }).first()).toBeVisible();
  await expect.poll(async () => (await kayit(app)).opened).toEqual([out]);
  const k = await kayit(app);
  expect(k.save[0].filters?.[0].name).toBe("PDF/A");
  expect(k.messages).toEqual([]);
  const pdf = await pdfBilgi(out);
  expect(pdf.yerImleri).toEqual(headings.map((h) => h.trim()));
  expect(pdf.pdfa).toBe("2B");
  expect(pdf.outputIntent).toBe(true);
  expect(pdf.id).toBe(true);
  expect(pdf.etiketli).toBe(true);
  expect(pdf.olusturan).toBe("OfficeHere Belge");
  expect(pdf.yaziTipleri.length).toBeGreaterThan(0);
  expect(pdf.yaziTipleri.filter((f) => !f.gomulu)).toEqual([]);
  const text = pdf.metin.join("\n");
  expect(text).toContain("Yönetici Özeti");
  expect(text).toContain("ığüşöç İĞÜŞÖÇ");
  for (const ch of TR) expect(text).toContain(ch);
  await close(app);
});

test("Belge: sayfa aralığı yalnız seçilen sayfaları basar; saydamlıksız belge PDF/A-1b olur", async () => {
  const { app, win } = await launch([]);
  const [belge] = await Promise.all([app.waitForEvent("window"), win.locator(".ohl-app", { hasText: "Belge" }).click()]);
  await belge.waitForSelector(".bl-body .ProseMirror");
  await belge.locator(".bl-body .ProseMirror").first().click();
  await belge.keyboard.press("Control+Alt+1");
  await belge.keyboard.type("Birinci sayfa başlığı");
  await belge.keyboard.press("Enter");
  await belge.keyboard.type("Birinci sayfa metni.");
  await belge.keyboard.press("Control+Enter");
  await belge.keyboard.press("Control+Alt+1");
  await belge.keyboard.type("İkinci sayfa başlığı");
  await belge.keyboard.press("Enter");
  await belge.keyboard.type("Ağaç gölgesinde ılık şiir.");
  await belge.keyboard.press("Control+Enter");
  await belge.keyboard.type("Üçüncü sayfa.");
  await expect(belge.locator(".bl-page")).toHaveCount(3);
  const out = path.join(mkdtempSync(path.join(tmpdir(), "oh-pdf-")), "aralik.pdf");
  await stub(app, out);
  await exportPdf(belge, { standard: /^PDF\/A-1b/, range: "2-3", after: "folder", title: "Deneme yazısı ığüş" });
  await expect.poll(() => existsSync(out)).toBe(true);
  await expect.poll(async () => (await kayit(app)).shown).toEqual([out]);
  // Saydamlık yok: soru sorulmadan 1b olarak kaydedilir.
  expect((await kayit(app)).messages).toEqual([]);
  const pdf = await pdfBilgi(out);
  expect(pdf.sayfa).toBe(2);
  expect(pdf.metin[0]).toContain("İkinci sayfa başlığı");
  expect(pdf.metin[0]).toContain("Ağaç gölgesinde ılık şiir.");
  expect(pdf.metin[1]).toContain("Üçüncü sayfa.");
  expect(pdf.metin.join("\n")).not.toContain("Birinci");
  expect(pdf.yerImleri).toEqual(["İkinci sayfa başlığı"]);
  expect(pdf.pdfa).toBe("1B");
  expect(pdf.baslik).toBe("Deneme yazısı ığüş");
  expect(pdf.yaziTipleri.length).toBeGreaterThan(0);
  expect(pdf.yaziTipleri.every((f) => f.gomulu && (f.cidSet || f.ad === "/Type3"))).toBe(true);
  await close(app);
});

test("Belge: 'önce sayfa sonu' başlıkları yazdırma kopyasında sayfayı boşaltmaz (her başlık yer imi)", async () => {
  const file = fixture("test-derlem/lo/lo-rapor.doc");
  const out = path.join(path.dirname(file), "rapor-doc.pdf");
  const { app, win } = await launch([file]);
  await win.locator(".oh-titlebar").filter({ hasText: "lo-rapor" }).waitFor();
  await win.waitForSelector(".bl-body .ProseMirror");
  const headings = await win.locator(".bl-body .ProseMirror :is(h1, h2, h3, h4, h5, h6)").allTextContents();
  expect(headings.length).toBe(16);
  await stub(app, out);
  await exportPdf(win, { standard: /^PDF\/A-2b/ });
  await expect.poll(() => existsSync(out)).toBe(true);
  const pdf = await pdfBilgi(out);
  expect(pdf.yerImleri).toEqual(headings.map((h) => h.trim()));
  // Bölüm başlıkları ve ardından gelen metin kendi sayfalarında (önceden yalnız üst/alt bilgi kalıyordu).
  for (const n of [2, 5, 7]) expect(pdf.metin.some((t) => t.includes(`Bölüm ${n}: Ölçüm Başlığı`) && t.includes("OfficeHere ölçüm derlemi"))).toBe(true);
  await close(app);
});

test("Belge: saydam görselde PDF/A-1b sorulur, onayla PDF/A-2b olarak kaydedilir", async () => {
  const file = fixture("test-derlem/poi/WithGIF.docx");
  const out = path.join(path.dirname(file), "gif.pdf");
  const { app, win } = await launch([file]);
  await win.locator(".oh-titlebar").filter({ hasText: "WithGIF" }).waitFor();
  await win.waitForSelector(".bl-body .ProseMirror img:not(.ProseMirror-separator)");
  await stub(app, out, 0);
  await exportPdf(win, { standard: /^PDF\/A-1b/ });
  await expect.poll(() => existsSync(out)).toBe(true);
  const k = await kayit(app);
  expect(k.messages).toHaveLength(1);
  expect(k.messages[0].message).toBe("Belge PDF/A-1b olarak kaydedilemiyor");
  expect(k.messages[0].detail).toContain("1. sayfa: saydam arka planlı ya da yarı saydam görsel");
  expect((await pdfBilgi(out)).pdfa).toBe("2B");
  await expect(win.locator(".oh-messagebar, [role=status]").filter({ hasText: "PDF/A-2b olarak kaydedildi" }).first()).toBeVisible();
  await close(app);
});

test("Belge: saydamlık sorusunda İptal dosya yazmaz", async () => {
  const file = fixture("test-derlem/poi/WithGIF.docx");
  const out = path.join(path.dirname(file), "gif-iptal.pdf");
  const { app, win } = await launch([file]);
  await win.locator(".oh-titlebar").filter({ hasText: "WithGIF" }).waitFor();
  await win.waitForSelector(".bl-body .ProseMirror img:not(.ProseMirror-separator)");
  await stub(app, out, 1);
  await exportPdf(win, { standard: /^PDF\/A-1b/ });
  await expect.poll(async () => (await kayit(app)).messages.length).toBe(1);
  expect(existsSync(out)).toBe(false);
  await close(app);
});

test("Tablo: kenarlıksız hücrelerin görünmez kenarlıkları PDF/A-1b'yi engellemez", async () => {
  const file = fixture("test-derlem/lo/lo-butce.xlsx");
  const out = path.join(path.dirname(file), "butce-1b.pdf");
  const { app, win } = await launch([file]);
  await win.locator(".oh-titlebar").filter({ hasText: "lo-butce" }).waitFor();
  await stub(app, out);
  await exportPdf(win, { standard: /^PDF\/A-1b/ });
  await expect.poll(() => existsSync(out)).toBe(true);
  expect((await kayit(app)).messages).toEqual([]);
  const pdf = await pdfBilgi(out);
  expect(pdf.pdfa).toBe("1B");
  expect(pdf.yaziTipleri.every((f) => f.gomulu && (f.cidSet || f.ad === "/Type3"))).toBe(true);
  expect(pdf.metin.join("\n")).toContain("2024 Satış Raporu — İĞÜŞÖÇ");
  await close(app);
});

test("Belge: geçersiz aralık kutuda bildirilir, iptal dosya yazmaz", async () => {
  const file = fixture("apps/belge/src/io/__fixtures__/libreoffice-resaved-report.docx");
  const out = path.join(path.dirname(file), "iptal.pdf");
  const { app, win } = await launch([file]);
  await win.waitForSelector(".ProseMirror");
  await stub(app, out);
  await win.locator(".oh-ribbon-file").click();
  await win.locator(".oh-bs-item", { hasText: "Dışarı Aktar" }).first().click();
  await win.locator(".oh-bs-action", { hasText: /PDF.*Oluştur/ }).click();
  const dlg = win.getByRole("dialog", { name: "PDF Olarak Dışa Aktar" });
  await dlg.getByRole("textbox", { name: "Sayfa aralığı" }).fill("40-45");
  await expect(dlg.getByRole("alert")).toContainText("belgenin dışında");
  await expect(dlg.getByRole("button", { name: "Dışa aktar" })).toBeDisabled();
  await dlg.getByRole("button", { name: "İptal" }).click();
  await expect(dlg).toHaveCount(0);
  expect((await kayit(app)).save).toEqual([]);
  expect(existsSync(out)).toBe(false);
  await close(app);
});

test("Tablo: PDF/A-2b, belge bilgileri ve klasörü göster", async () => {
  const file = fixture("apps/tablo/test/fixtures/lo-generated.xlsx");
  const out = path.join(path.dirname(file), "tablo.pdf");
  const { app, win } = await launch([file]);
  await win.waitForSelector(".oh-ribbon");
  await stub(app, out);
  await exportPdf(win, { standard: /^PDF\/A-2b/, after: "folder", title: "Bütçe çizelgesi" });
  await expect.poll(() => existsSync(out)).toBe(true);
  await expect.poll(async () => (await kayit(app)).shown).toEqual([out]);
  const pdf = await pdfBilgi(out);
  expect(pdf.pdfa).toBe("2B");
  expect(pdf.outputIntent && pdf.id).toBe(true);
  expect(pdf.baslik).toBe("Bütçe çizelgesi");
  expect(pdf.olusturan).toBe("OfficeHere Tablo");
  expect(pdf.yaziTipleri.filter((f) => !f.gomulu)).toEqual([]);
  expect(pdf.sayfa).toBeGreaterThan(0);
  await close(app);
});

test("Sunum: sayfa aralığı slaytları seçer; standart PDF'te PDF/A işareti yok", async () => {
  const file = fixture("apps/sunum/src/io/__fixtures__/lo-proje.pptx");
  const out = path.join(path.dirname(file), "sunu.pdf");
  const { app, win } = await launch([file]);
  await expect(win.locator(".thumb-pane .thumb-item")).toHaveCount(7);
  await stub(app, out);
  await exportPdf(win, { standard: /^PDF$/, range: "2-3" });
  await expect.poll(() => existsSync(out)).toBe(true);
  const pdf = await pdfBilgi(out);
  expect(pdf.sayfa).toBe(2);
  expect(pdf.pdfa).toBeNull();
  expect(pdf.olusturan).toBe("OfficeHere Sunum");
  expect(pdf.metin[0]).toContain("Gündem");
  expect(pdf.metin[1]).toContain("İş hedefleri");
  expect(pdf.metin.join("\n")).not.toContain("Müşteri Portalı Yenileme Projesi");
  await close(app);
});
