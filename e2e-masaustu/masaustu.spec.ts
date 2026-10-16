import { _electron as electron, expect, test, type ElectronApplication, type Page } from "@playwright/test";
import type {} from "../packages/core/src/bridge";
import { execFileSync } from "node:child_process";
import { copyFileSync, mkdtempSync, readFileSync, statSync } from "node:fs";
import { createServer } from "node:http";
import type { AddressInfo } from "node:net";
import { tmpdir } from "node:os";
import path from "node:path";

const root = path.resolve(__dirname, "..");
const desktop = path.join(root, "apps/desktop");
// Test için ayrı derleme dizini verilebilir (paralel derlemeler çakışmasın).
const appsDir = process.env.OH_E2E_APPS_DIR ?? path.join(root, "dist-web");

function electronPath(): string {
  // apps/desktop/node_modules/electron paketinin döndürdüğü ikili yol.
  return require(path.join(desktop, "node_modules/electron")) as unknown as string;
}

async function launch(args: string[] = [], env: Record<string, string> = {}): Promise<{ app: ElectronApplication; win: Page; userData: string }> {
  const userData = mkdtempSync(path.join(tmpdir(), "oh-masaustu-"));
  const app = await electron.launch({
    executablePath: electronPath(),
    args: [desktop, "--no-sandbox", ...args],
    env: { ...process.env, OH_APPS_DIR: appsDir, OH_USER_DATA: userData, ELECTRON_DISABLE_SECURITY_WARNINGS: "1", ...env },
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
  // Alt bilgideki vaat satırı (README › Ağ erişimi)
  await expect(win.locator(".ohl-vaat")).toHaveText("Hesap yok · Reklam yok · Abonelik yok · Telemetri yok · Yapay zekâ yok");
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

/** GitHub "latest release" yanıtını taklit eden yerel sunucu. */
async function releaseFeed(tag: string) {
  const server = createServer((_req, res) => {
    res.setHeader("Content-Type", "application/json");
    res.end(JSON.stringify({ tag_name: tag, html_url: `https://github.com/fyildirim-debug/officehere/releases/tag/${tag}`, draft: false, prerelease: false }));
  });
  await new Promise<void>((r) => server.listen(0, "127.0.0.1", r));
  return { url: `http://127.0.0.1:${(server.address() as AddressInfo).port}/latest`, close: () => server.close() };
}

test("yeni sürüm başlık çubuğunda ve Seçenekler'de bildirilir", async () => {
  const feed = await releaseFeed("v9.9.9");
  const { app, win } = await launch([], { OH_UPDATE_MODE: "manual", OH_UPDATE_FEED: feed.url, OH_UPDATE_DELAY_MS: "0" });
  const pill = win.locator(".oh-update-pill");
  await expect(pill).toHaveText("Yeni sürüm: 9.9.9");
  await win.locator(".oh-bs-item", { hasText: "Seçenekler" }).click();
  const section = win.locator(".oh-update-section");
  await expect(section).toContainText("Sürüm 9.9.9 yayımlandı.");
  await expect(section.getByRole("button", { name: "İndirme sayfasını aç" })).toBeVisible();
  await app.close();
  feed.close();
});

test("güncel sürümde elle denetim 'güncel' der", async () => {
  const feed = await releaseFeed("v0.0.1");
  const { app, win } = await launch([], { OH_UPDATE_MODE: "manual", OH_UPDATE_FEED: feed.url, OH_UPDATE_DELAY_MS: "600000" });
  await win.locator(".oh-bs-item", { hasText: "Seçenekler" }).click();
  const section = win.locator(".oh-update-section");
  await section.getByRole("button", { name: "Güncelleştirmeleri denetle" }).click();
  await expect(section).toContainText("OfficeHere güncel.");
  await expect(win.locator(".oh-update-pill")).toHaveCount(0);
  // Otomatik denetim seçeneği kalıcıdır.
  const auto = section.getByRole("checkbox", { name: "Güncelleştirmeleri otomatik denetle" });
  await expect(auto).toBeChecked();
  await auto.uncheck();
  await expect(auto).not.toBeChecked();
  await app.close();
  feed.close();
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

// Yeni biçimler: çift tık / komut satırı → doğru uygulama → içerik okunur.
type TabloHook = { __ohTablo?: { cell(ref: string): { value: unknown } | null } };
for (const [rel, appName, check] of [
  ["apps/belge/src/io/__fixtures__/bicimler/lo-ornek.doc", "OfficeHere Belge", (w: Page) => expect(w.locator(".bl-body")).toContainText("Giriş ve Amaç")],
  ["apps/belge/src/io/__fixtures__/bicimler/lo-ornek.odt", "OfficeHere Belge", (w: Page) => expect(w.locator(".bl-body")).toContainText("Giriş ve Amaç")],
  ["apps/belge/src/io/__fixtures__/bicimler/lo-ornek.rtf", "OfficeHere Belge", (w: Page) => expect(w.locator(".bl-body")).toContainText("Giriş ve Amaç")],
  ["apps/tablo/src/io/__fixtures__/ornek.xls", "OfficeHere Tablo", (w: Page) => expect.poll(() => w.evaluate(() => (window as unknown as TabloHook).__ohTablo?.cell("A1")?.value ?? null)).toBe("Ürün")],
  ["apps/tablo/src/io/__fixtures__/ornek.ods", "OfficeHere Tablo", (w: Page) => expect.poll(() => w.evaluate(() => (window as unknown as TabloHook).__ohTablo?.cell("A1")?.value ?? null)).toBe("Ürün")],
  ["apps/sunum/src/io/__fixtures__/lo-bicim.ppt", "OfficeHere Sunum", (w: Page) => expect(w.locator(".stage .sv")).toContainText("Çeyrek Sonuçları")],
  ["apps/sunum/src/io/__fixtures__/lo-proje.odp", "OfficeHere Sunum", (w: Page) => expect(w.locator(".thumb-pane .thumb-item")).toHaveCount(7)],
] as const) {
  test(`masaüstünde ${path.extname(rel)} dosyası ${appName} ile açılır ve içeriği okunur`, async () => {
    const { app, win } = await launch([fixture(rel)]);
    await win.waitForSelector(".oh-ribbon");
    await expect(win.locator(".oh-titlebar")).toContainText(appName);
    await check(win);
    await app.close();
  });
}

test("belge düzenlenir, Ctrl+S ile aynı dosyaya yazılır ve son kullanılanlara eklenir", async () => {
  const file = fixture("apps/belge/src/io/__fixtures__/libreoffice-resaved-report.docx");
  const before = statSync(file).mtimeMs;
  const { app, win } = await launch([file]);
  await win.waitForSelector(".ProseMirror");
  // Dosya ?file= ile eşzamansız (Worker'da) yüklenir: düzenlemeden önce yüklendiği görülsün.
  await expect(win.locator(".oh-titlebar-title")).toContainText("libreoffice-resaved-report");
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
  await expect(win.locator(".oh-titlebar-title")).toContainText("libreoffice-html");
  // Otomatik kaydetme kapalı: belge kirli kalsın ve dosya değişmesin.
  await win.evaluate(() => localStorage.setItem("officehere.autoSave", "0"));
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
