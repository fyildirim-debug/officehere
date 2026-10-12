import { _electron as electron, expect, test, type ElectronApplication, type Page } from "@playwright/test";
import type {} from "../packages/core/src/bridge";
import { execFileSync } from "node:child_process";
import { copyFileSync, mkdtempSync, readFileSync, statSync } from "node:fs";
import { tmpdir } from "node:os";
import path from "node:path";

const root = path.resolve(__dirname, "..");
const desktop = path.join(root, "apps/desktop");
const appsDir = path.join(root, "dist-web");

function electronPath(): string {
  // apps/desktop/node_modules/electron paketinin döndürdüğü ikili yol.
  return require(path.join(desktop, "node_modules/electron")) as unknown as string;
}

async function launch(args: string[] = []): Promise<{ app: ElectronApplication; win: Page; userData: string }> {
  const userData = mkdtempSync(path.join(tmpdir(), "oh-masaustu-"));
  const app = await electron.launch({
    executablePath: electronPath(),
    args: [desktop, "--no-sandbox", ...args],
    env: { ...process.env, OH_APPS_DIR: appsDir, OH_USER_DATA: userData, ELECTRON_DISABLE_SECURITY_WARNINGS: "1" },
  });
  const win = await app.firstWindow();
  await win.waitForLoadState("domcontentloaded");
  await win.waitForSelector(".oh-titlebar");
  return { app, win, userData };
}

function fixture(rel: string): string {
  // Testler sırasında asıl fixture bozulmasın diye geçici kopya.
  const dir = mkdtempSync(path.join(tmpdir(), "oh-dosya-"));
  const dest = path.join(dir, path.basename(rel));
  copyFileSync(path.join(root, rel), dest);
  return dest;
}

test("başlangıç ekranı açılır ve yerel köprü hazırdır", async () => {
  const { app, win } = await launch();
  await expect(win).toHaveTitle(/Başlangıç - OfficeHere/);
  const info = await win.evaluate(() => ({ native: typeof window.officehereNative, cls: document.documentElement.className }));
  expect(info.native).toBe("object");
  for (const name of ["Belge", "Tablo", "Sunum"]) await expect(win.locator(".ohl-app", { hasText: name })).toBeVisible();
  await app.close();
});

test("üçüncü taraf lisans bildirimi uygulama içinde açılır", async () => {
  const { app, win } = await launch();
  await win.locator(".oh-bs-item", { hasText: "Seçenekler" }).click();
  const [notices] = await Promise.all([app.waitForEvent("window"), win.getByRole("link", { name: "Üçüncü taraf yazılım bildirimleri" }).click()]);
  await notices.waitForURL(/^file:.*\/THIRD_PARTY_NOTICES\.txt$/);
  await expect(notices.locator("body")).toContainText("Third-Party Software Notices");
  await app.close();
});

test("başlangıçtan yeni belge penceresi açılır", async () => {
  const { app, win } = await launch();
  const [belge] = await Promise.all([app.waitForEvent("window"), win.locator(".ohl-app", { hasText: "Belge" }).click()]);
  await belge.waitForSelector(".oh-ribbon");
  await expect(belge).toHaveTitle(/OfficeHere Belge/);
  await app.close();
});

for (const [rel, appName] of [
  ["apps/belge/src/io/__fixtures__/libreoffice-resaved-report.docx", "OfficeHere Belge"],
  ["apps/tablo/test/fixtures/lo-generated.xlsx", "OfficeHere Tablo"],
  ["apps/sunum/src/io/__fixtures__/lo-proje.pptx", "OfficeHere Sunum"],
] as const) {
  test(`komut satırından açılan dosya doğru uygulamada açılır: ${path.basename(rel)}`, async () => {
    const file = fixture(rel);
    const { app, win } = await launch([file]);
    await win.waitForSelector(".oh-ribbon");
    await expect(win.locator(".oh-titlebar")).toContainText(appName);
    await expect(win.locator(".oh-titlebar")).toContainText(path.basename(file).replace(/\.[^.]+$/, ""));
    await app.close();
  });
}

test("belge düzenlenir, Ctrl+S ile aynı dosyaya yazılır ve son kullanılanlara eklenir", async () => {
  const file = fixture("apps/belge/src/io/__fixtures__/libreoffice-resaved-report.docx");
  const before = statSync(file).mtimeMs;
  const { app, win } = await launch([file]);
  await win.waitForSelector(".ProseMirror");
  await win.locator(".ProseMirror").first().click();
  await win.keyboard.press("Control+Home");
  await win.keyboard.type("MASAUSTU-TESTI ");
  await expect(win.locator(".oh-titlebar-dirty")).toBeVisible();
  await win.keyboard.press("Control+S");
  await expect(win.locator(".oh-titlebar-dirty")).toHaveCount(0);
  await expect.poll(() => statSync(file).mtimeMs).toBeGreaterThan(before);
  // Kaydedilen DOCX'in içinde yazılan metin bulunmalı.
  const xml = execFileSync("unzip", ["-p", file, "word/document.xml"]).toString("utf8");
  expect(xml).toContain("MASAUSTU-TESTI");
  const recent = await win.evaluate(() => window.officehereNative!.getRecent());
  expect(recent.some((r) => r.path === file && r.app === "belge")).toBe(true);
  await app.close();
});

test("kaydedilmemiş belge kapatılırken onay sorulur", async () => {
  const file = fixture("apps/belge/src/io/__fixtures__/libreoffice-html.docx");
  const original = readFileSync(file);
  const { app, win } = await launch([file]);
  await win.waitForSelector(".ProseMirror");
  await win.locator(".ProseMirror").first().click();
  await win.keyboard.type("kaydedilmeyecek");
  await expect(win.locator(".oh-titlebar-dirty")).toBeVisible();
  // Pencereyi kapatmayı dene: ana süreç kapanmayı engelleyip sayfaya sormalı.
  await app.evaluate(({ BrowserWindow }) => BrowserWindow.getAllWindows()[0]?.close());
  const dialog = win.getByRole("dialog", { name: "OfficeHere" });
  await expect(dialog).toBeVisible();
  await expect(dialog).toContainText("kaydetmek istiyor musunuz");
  // İptal: pencere açık kalır.
  await dialog.getByRole("button", { name: "İptal" }).click();
  await expect(dialog).toHaveCount(0);
  expect(await app.evaluate(({ BrowserWindow }) => BrowserWindow.getAllWindows().length)).toBe(1);
  // Kaydetme: pencere kapanır, dosya değişmez.
  const closed = app.waitForEvent("close");
  await app.evaluate(({ BrowserWindow }) => BrowserWindow.getAllWindows()[0]?.close());
  // Tıklama pencereyi (ve son pencere olduğu için uygulamayı) kapatır; kapanan sayfa hatası beklenen sonuçtur.
  await win.getByRole("dialog", { name: "OfficeHere" }).getByRole("button", { name: "Kaydetme" }).click().catch(() => {});
  await closed;
  expect(readFileSync(file).equals(original)).toBe(true);
});
