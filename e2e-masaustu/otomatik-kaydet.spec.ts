import { _electron as electron, expect, test, type ElectronApplication, type Page } from "@playwright/test";
import type {} from "../packages/core/src/bridge";
import { execFileSync } from "node:child_process";
import { copyFileSync, existsSync, mkdtempSync, readFileSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import path from "node:path";

// Otomatik kaydetme (masaüstü): diskteki .docx/.xlsx/.pptx düzenlendikçe Ctrl+S'siz
// aynı dosyaya yazılır, ilk yazmadan önce özgün dosya <ad>.bak olarak yedeklenir.

const root = path.resolve(__dirname, "..");
const desktop = path.join(root, "apps/desktop");
const appsDir = path.join(root, "dist-web");
const SAVED = /^Kaydedildi \d{2}:\d{2}$/;

function electronPath(): string {
  return require(path.join(desktop, "node_modules/electron")) as unknown as string;
}

async function launch(file: string): Promise<{ app: ElectronApplication; win: Page }> {
  const userData = mkdtempSync(path.join(tmpdir(), "oh-otokayit-"));
  const app = await electron.launch({
    executablePath: electronPath(),
    args: [desktop, "--no-sandbox", file],
    env: { ...process.env, OH_APPS_DIR: appsDir, OH_USER_DATA: userData, ELECTRON_DISABLE_SECURITY_WARNINGS: "1" },
  });
  const win = await app.firstWindow();
  await win.waitForLoadState("domcontentloaded");
  await win.waitForSelector(".oh-ribbon");
  // Dosya ?file= ile eşzamansız yüklenir: düzenlemeden önce yüklendiği görülsün.
  await expect(win.locator(".oh-titlebar-title")).toContainText(path.basename(file).replace(/\.[^.]+$/, ""));
  return { app, win };
}

/** Testler asıl fixture'ı bozmasın diye geçici kopya. */
function fixture(rel: string): string {
  const dir = mkdtempSync(path.join(tmpdir(), "oh-otokayit-dosya-"));
  const dest = path.join(dir, path.basename(rel));
  copyFileSync(path.join(root, rel), dest);
  return dest;
}

/** Zip (OOXML) içindeki tüm parçaların metni. */
const zipText = (file: string) => execFileSync("unzip", ["-p", file]).toString("utf8");
const badge = (win: Page) => win.locator(".oh-titlebar .oh-savebadge");

test("Belge: değiştirilen .docx Ctrl+S'siz diske yazılır, .bak alınır, taslak silinir", async () => {
  const file = fixture("apps/belge/src/io/__fixtures__/libreoffice-resaved-report.docx");
  const original = readFileSync(file);
  const { app, win } = await launch(file);
  await win.waitForSelector(".ProseMirror");
  await win.locator(".ProseMirror").first().click();
  await win.keyboard.press("Control+Home");
  await win.keyboard.type("OTOMATIK-KAYIT ");
  await expect(badge(win)).toHaveText("Kaydedilmedi");
  // Kurtarma taslağını hemen yazdır: otomatik kaydetmeden sonra silinmeli.
  await win.evaluate(() => window.dispatchEvent(new Event("officehere:autosave")));
  const drafts = () => win.evaluate(() => Object.keys(localStorage).filter((k) => k.startsWith("officehere.belge.draft.")).length);
  expect(await drafts()).toBe(1);

  await expect(badge(win)).toHaveText(SAVED);
  expect(existsSync(`${file}.bak`)).toBe(true);
  expect(readFileSync(`${file}.bak`).equals(original)).toBe(true);
  expect(execFileSync("unzip", ["-p", file, "word/document.xml"]).toString("utf8")).toContain("OTOMATIK-KAYIT");
  expect(await drafts()).toBe(0);
  await expect(win).not.toHaveTitle(/^●/);

  // İkinci değişiklik de yazılır; yedek aynı oturumda yeniden alınmaz.
  await win.keyboard.type("IKINCI ");
  await expect.poll(() => execFileSync("unzip", ["-p", file, "word/document.xml"]).toString("utf8"), { timeout: 15_000 }).toContain("IKINCI");
  await expect(badge(win)).toHaveText(SAVED);
  expect(readFileSync(`${file}.bak`).equals(original)).toBe(true);
  await app.close();
});

test("Belge: seçenek kapalıyken kendiliğinden yazılmaz; Ctrl+S rozeti günceller", async () => {
  const file = fixture("apps/belge/src/io/__fixtures__/libreoffice-html.docx");
  const original = readFileSync(file);
  const { app, win } = await launch(file);
  await win.waitForSelector(".ProseMirror");
  await win.locator(".oh-ribbon-file").click();
  await win.locator(".oh-bs-item", { hasText: "Seçenekler" }).first().click();
  const option = win.getByRole("checkbox", { name: "Otomatik kaydet" });
  await expect(option).toBeChecked();
  await expect(win.getByText("Bilgisayarınızdaki belgeler değiştikçe kendiliğinden kaydedilir.")).toBeVisible();
  await option.uncheck();
  expect(await win.evaluate(() => localStorage.getItem("officehere.autoSave"))).toBe("0");
  await win.getByRole("button", { name: "Geri", exact: true }).click();

  await win.locator(".ProseMirror").first().click();
  await win.keyboard.type("KAPALI ");
  await expect(badge(win)).toHaveText("Kaydedilmedi");
  await win.waitForTimeout(4000);
  expect(readFileSync(file).equals(original)).toBe(true);
  expect(existsSync(`${file}.bak`)).toBe(false);
  await expect(badge(win)).toHaveText("Kaydedilmedi");

  // Elle kaydetme rozeti günceller (yedek yalnız otomatik kaydetmeden önce alınır).
  await win.keyboard.press("Control+S");
  await expect(badge(win)).toHaveText(SAVED);
  expect(readFileSync(file).equals(original)).toBe(false);
  expect(existsSync(`${file}.bak`)).toBe(false);
  await app.close();
});

test("Tablo: .xlsx kendiliğinden kaydedilir", async () => {
  const file = fixture("apps/tablo/test/fixtures/lo-generated.xlsx");
  const original = readFileSync(file);
  const { app, win } = await launch(file);
  const box = win.getByRole("textbox", { name: "Ad Kutusu" });
  await box.fill("H40");
  await box.press("Enter");
  await win.keyboard.type("OTOKAYITTABLO");
  await win.keyboard.press("Enter");
  await expect(badge(win)).toHaveText(SAVED);
  expect(readFileSync(`${file}.bak`).equals(original)).toBe(true);
  expect(zipText(file)).toContain("OTOKAYITTABLO");
  await app.close();
});

test("Tablo: biçim kaybettiren .csv kendiliğinden kaydedilmez", async () => {
  const dir = mkdtempSync(path.join(tmpdir(), "oh-otokayit-csv-"));
  const file = path.join(dir, "veri.csv");
  writeFileSync(file, "Ad;Tutar\nElma;12\n");
  const original = readFileSync(file);
  const { app, win } = await launch(file);
  const box = win.getByRole("textbox", { name: "Ad Kutusu" });
  await box.fill("C3");
  await box.press("Enter");
  await win.keyboard.type("99");
  await win.keyboard.press("Enter");
  await expect(badge(win)).toHaveText("Kaydedilmedi");
  await win.waitForTimeout(4000);
  expect(readFileSync(file).equals(original)).toBe(true);
  expect(existsSync(`${file}.bak`)).toBe(false);
  await expect(badge(win)).toHaveText("Kaydedilmedi");
  // Kirli pencere kapatılırken sorar; soru beklenmeden pencere yok edilir.
  await app.evaluate(({ BrowserWindow }) => BrowserWindow.getAllWindows().forEach((w) => w.destroy()));
  await app.close();
});

test("Sunum: .pptx kendiliğinden kaydedilir", async () => {
  const file = fixture("apps/sunum/src/io/__fixtures__/lo-proje.pptx");
  const original = readFileSync(file);
  const { app, win } = await launch(file);
  await win.locator(".stage .sv [data-shape-id]").first().click();
  await win.keyboard.type("OTOKAYITSUNUM");
  await win.keyboard.press("Escape");
  await win.keyboard.press("Escape");
  await expect(badge(win)).toHaveText(SAVED);
  expect(readFileSync(`${file}.bak`).equals(original)).toBe(true);
  expect(zipText(file)).toContain("OTOKAYITSUNUM");
  await app.close();
});
