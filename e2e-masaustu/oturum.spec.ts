import { _electron as electron, expect, test, type ElectronApplication, type Page } from "@playwright/test";
import type {} from "../packages/core/src/bridge";
import { copyFileSync, existsSync, mkdtempSync, readFileSync, rmSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import path from "node:path";

// Son kullanılanlar yönetimi, pencere durumu ve oturum geri yükleme (masaüstü).
const root = path.resolve(__dirname, "..");
const desktop = path.join(root, "apps/desktop");
const appsDir = path.join(root, "dist-web");

function electronPath(): string {
  return require(path.join(desktop, "node_modules/electron")) as unknown as string;
}

const tempDir = (prefix: string) => mkdtempSync(path.join(tmpdir(), prefix));

function start(userData: string, args: string[] = []): Promise<ElectronApplication> {
  return electron.launch({
    executablePath: electronPath(),
    args: [desktop, "--no-sandbox", ...args],
    env: { ...process.env, OH_APPS_DIR: appsDir, OH_USER_DATA: userData, ELECTRON_DISABLE_SECURITY_WARNINGS: "1" },
  });
}

async function launch(userData: string, args: string[] = []): Promise<{ app: ElectronApplication; win: Page }> {
  const app = await start(userData, args);
  const win = await app.firstWindow();
  await win.waitForSelector(".oh-titlebar");
  return { app, win };
}

/** Boş (içeriği önemsiz) bir dosya oluşturur. */
function touch(dir: string, name: string): string {
  const p = path.join(dir, name);
  writeFileSync(p, "x");
  return p;
}

function fixture(rel: string): string {
  const dest = path.join(tempDir("oh-dosya-"), path.basename(rel));
  copyFileSync(path.join(root, rel), dest);
  return dest;
}

function seedRecent(userData: string, files: { path: string; app: "belge" | "tablo" | "sunum"; ago: number; pinned?: boolean }[]) {
  const now = Date.now();
  const list = files.map((f) => ({ path: f.path, name: path.basename(f.path), app: f.app, openedAt: now - f.ago, ...(f.pinned ? { pinned: true } : {}) }));
  writeFileSync(path.join(userData, "son-kullanilanlar.json"), JSON.stringify(list));
}

const rowNames = (win: Page) => win.locator(".oh-recent-row .oh-recent-filename").allTextContents();
const row = (win: Page, name: string) => win.locator(".oh-recent-row", { hasText: name });

test("son kullanılanlar: sabitleme sırayı değiştirir, listeden kaldırma yalnız kaydı siler", async () => {
  const userData = tempDir("oh-oturum-");
  const dir = tempDir("oh-dosya-");
  const alfa = touch(dir, "Alfa.docx");
  const beta = touch(dir, "Beta.xlsx");
  const gama = touch(dir, "Gama.pptx");
  seedRecent(userData, [
    { path: alfa, app: "belge", ago: 60_000 },
    { path: beta, app: "tablo", ago: 2 * 3_600_000 },
    { path: gama, app: "sunum", ago: 3 * 86_400_000 },
  ]);
  const { app, win } = await launch(userData);
  await expect.poll(() => rowNames(win)).toEqual(["Alfa.docx", "Beta.xlsx", "Gama.pptx"]);
  await expect(row(win, "Alfa.docx").locator(".oh-recent-time")).toHaveText("1 dakika önce");
  await expect(row(win, "Gama.pptx").locator(".oh-recent-time")).toHaveText("3 gün önce");
  await expect(win.locator(".oh-recent-group")).toHaveCount(0);

  // Sağ tık → Sabitle: dosya "Sabitlenmiş" grubunda başa geçer.
  await row(win, "Gama.pptx").click({ button: "right" });
  await win.getByRole("menuitem", { name: "Sabitle" }).click();
  await expect.poll(() => rowNames(win)).toEqual(["Gama.pptx", "Alfa.docx", "Beta.xlsx"]);
  await expect(win.locator(".oh-recent-group")).toHaveText(["Sabitlenmiş", "Son"]);

  // Satırdaki raptiye düğmesi; sabitlenenler kendi aralarında en son açılan önce.
  await row(win, "Beta.xlsx").hover();
  await win.getByRole("button", { name: "Beta.xlsx: sabitle" }).click();
  await expect.poll(() => rowNames(win)).toEqual(["Beta.xlsx", "Gama.pptx", "Alfa.docx"]);
  const recent = await win.evaluate(() => window.officehereNative!.getRecent());
  expect(recent.map((r) => [r.name, r.pinned === true])).toEqual([
    ["Beta.xlsx", true],
    ["Gama.pptx", true],
    ["Alfa.docx", false],
  ]);

  // Listeden kaldır: kayıt silinir, dosya yerinde kalır. Masaüstünde "Klasörde göster" de vardır.
  await row(win, "Alfa.docx").click({ button: "right" });
  await expect(win.getByRole("menuitem", { name: "Klasörde göster" })).toBeVisible();
  await win.getByRole("menuitem", { name: "Listeden kaldır" }).click();
  await expect.poll(() => rowNames(win)).toEqual(["Beta.xlsx", "Gama.pptx"]);
  expect((await win.evaluate(() => window.officehereNative!.getRecent())).map((r) => r.name)).toEqual(["Beta.xlsx", "Gama.pptx"]);
  expect(existsSync(alfa)).toBe(true);

  // Aç sayfasındaki liste de güncel durumu gösterir; sabitleme kaldırılabilir.
  await win.locator(".oh-bs-item", { hasText: "Aç" }).click();
  await expect.poll(() => rowNames(win)).toEqual(["Beta.xlsx", "Gama.pptx"]);
  await row(win, "Beta.xlsx").click({ button: "right" });
  await win.getByRole("menuitem", { name: "Sabitlemeyi kaldır" }).click();
  await expect.poll(() => rowNames(win)).toEqual(["Gama.pptx", "Beta.xlsx"]);
  await app.close();
});

test("taşınmış dosya 'Bulunamadı' gösterilir, tıklanınca listeden kaldırılması sorulur", async () => {
  const userData = tempDir("oh-oturum-");
  const dir = tempDir("oh-dosya-");
  const var1 = touch(dir, "Mevcut.docx");
  const sonra = touch(dir, "Sonra silinecek.xlsx");
  const kayip = path.join(dir, "Kayıp.docx");
  seedRecent(userData, [
    { path: var1, app: "belge", ago: 1_000 },
    { path: kayip, app: "belge", ago: 2_000 },
    { path: sonra, app: "tablo", ago: 3_000 },
  ]);
  const { app, win } = await launch(userData);
  // Ana süreç olmayan dosyayı artık sessizce düşürmez.
  expect((await win.evaluate(() => window.officehereNative!.getRecent())).map((r) => r.name)).toEqual(["Mevcut.docx", "Kayıp.docx", "Sonra silinecek.xlsx"]);
  await expect(row(win, "Kayıp.docx").locator(".oh-recent-badge")).toHaveText("Bulunamadı");
  await expect(row(win, "Kayıp.docx")).toHaveClass(/is-missing/);
  await expect(row(win, "Mevcut.docx")).not.toHaveClass(/is-missing/);
  await expect(win.locator(".oh-recent-badge")).toHaveCount(1);

  await row(win, "Kayıp.docx").locator(".oh-recent-open").click();
  const dialog = win.getByRole("dialog", { name: "OfficeHere" });
  await expect(dialog).toContainText("Dosya bulunamadı. Listeden kaldırılsın mı?");
  await dialog.getByRole("button", { name: "Kaldır" }).click();
  await expect(dialog).toHaveCount(0);
  await expect(row(win, "Kayıp.docx")).toHaveCount(0);
  expect((await win.evaluate(() => window.officehereNative!.getRecent())).map((r) => r.name)).toEqual(["Mevcut.docx", "Sonra silinecek.xlsx"]);

  // Liste açıkken silinen dosya tıklanınca denetlenir; İptal satırı bırakır ve rozetler.
  rmSync(sonra);
  await row(win, "Sonra silinecek.xlsx").locator(".oh-recent-open").focus();
  await win.keyboard.press("Enter");
  await expect(dialog).toBeVisible();
  await dialog.getByRole("button", { name: "İptal" }).click();
  await expect(row(win, "Sonra silinecek.xlsx").locator(".oh-recent-badge")).toHaveText("Bulunamadı");
  expect(app.windows()).toHaveLength(1);
  await app.close();
});

const bounds = (app: ElectronApplication) =>
  app.evaluate(({ BrowserWindow }) =>
    BrowserWindow.getAllWindows()
      .sort((a, b) => a.id - b.id)
      .map((w) => w.getBounds()),
  );

test("pencere boyutu ve konumu türüne göre saklanır, aynı türden ikinci pencere kaydırılır", async () => {
  const userData = tempDir("oh-oturum-");
  let { app } = await launch(userData);
  const target = { x: 120, y: 90, width: 980, height: 660 };
  await app.evaluate(({ BrowserWindow }, b) => BrowserWindow.getAllWindows()[0].setBounds(b), target);
  await expect.poll(() => bounds(app)).toEqual([target]);
  await app.close();
  expect(JSON.parse(readFileSync(path.join(userData, "pencere.json"), "utf8")).baslangic).toEqual({ ...target, maximized: false });

  // Belge pencereleri için ayrı kayıt.
  const saved = JSON.parse(readFileSync(path.join(userData, "pencere.json"), "utf8"));
  writeFileSync(path.join(userData, "pencere.json"), JSON.stringify({ ...saved, belge: { x: 60, y: 40, width: 900, height: 640, maximized: false } }));

  let win: Page;
  ({ app, win } = await launch(userData));
  expect(await bounds(app)).toEqual([target]);
  for (let i = 0; i < 2; i++) {
    await Promise.all([app.waitForEvent("window"), win.locator(".ohl-app", { hasText: "Belge" }).click()]);
  }
  await expect.poll(async () => (await bounds(app)).length).toBe(3);
  const [, first, second] = await bounds(app);
  expect(first).toEqual({ x: 60, y: 40, width: 900, height: 640 });
  expect(second).toEqual({ x: 84, y: 64, width: 900, height: 640 });
  await app.close();
});

test("oturum: seçenek açıksa son açık belgeler yeniden açılır, yoksa başlangıç ekranı", async () => {
  const userData = tempDir("oh-oturum-");
  // Ayar dosyasındaki güncelleştirme anahtarı korunmalı.
  writeFileSync(path.join(userData, "ayarlar.json"), JSON.stringify({ otomatikGuncellemeDenetimi: false }));
  const docx = fixture("apps/belge/src/io/__fixtures__/libreoffice-resaved-report.docx");
  const xlsx = fixture("apps/tablo/test/fixtures/lo-generated.xlsx");

  let { app, win } = await launch(userData);
  await win.locator(".oh-bs-item", { hasText: "Seçenekler" }).click();
  const box = win.getByRole("checkbox", { name: "Açılışta son oturumu geri yükle" });
  await expect(box).not.toBeChecked();
  await box.check();
  await expect.poll(() => JSON.parse(readFileSync(path.join(userData, "ayarlar.json"), "utf8"))).toEqual({ otomatikGuncellemeDenetimi: false, oturumuGeriYukle: true });

  for (const [kind, file] of [["belge", docx], ["tablo", xlsx]] as const) {
    const [doc] = await Promise.all([app.waitForEvent("window"), win.evaluate(([k, p]) => window.officehereNative!.newWindow(k, p), [kind, file] as const)]);
    await doc.waitForSelector(".oh-ribbon");
  }
  await app.close();
  const session = JSON.parse(readFileSync(path.join(userData, "oturum.json"), "utf8"));
  expect(session.belgeler).toEqual([
    { path: docx, app: "belge" },
    { path: xlsx, app: "tablo" },
  ]);

  // Dosyasız açılış: başlangıç ekranı yerine iki belge kendi uygulamasında açılır.
  app = await start(userData);
  await expect.poll(() => app.windows().length).toBe(2);
  const stem = (p: string) => path.basename(p).replace(/\.[^.]+$/, "");
  const opened = async () => {
    const titles = await Promise.all(app.windows().map((w) => w.locator(".oh-titlebar").innerText().catch(() => "")));
    return [
      titles.some((t) => t.includes("OfficeHere Belge") && t.includes(stem(docx))),
      titles.some((t) => t.includes("OfficeHere Tablo") && t.includes(stem(xlsx))),
      titles.some((t) => t.includes("Başlangıç")),
    ];
  };
  await expect.poll(opened).toEqual([true, true, false]);
  await app.close();

  // Belgeler artık yoksa başlangıç ekranı açılır.
  rmSync(docx);
  rmSync(xlsx);
  ({ app, win } = await launch(userData));
  await expect(win).toHaveTitle(/Başlangıç - OfficeHere/);
  expect(app.windows()).toHaveLength(1);
  await app.close();
});

test("oturum: belge yolu setDocumentState ile öğrenilir; son pencere kapatılınca oturumda kalır", async () => {
  const userData = tempDir("oh-oturum-");
  const docx = fixture("apps/belge/src/io/__fixtures__/libreoffice-html.docx");
  const saved = path.join(tempDir("oh-dosya-"), "Farklı kaydedilen.docx");
  const { app, win } = await launch(userData, [docx]);
  await win.waitForSelector(".oh-ribbon");
  // Boş belge penceresi "Farklı Kaydet" sonrası yolunu bildirir.
  const [blank] = await Promise.all([app.waitForEvent("window"), win.evaluate(() => window.officehereNative!.newWindow("belge"))]);
  await blank.waitForSelector(".oh-ribbon");
  writeFileSync(saved, "x");
  await blank.evaluate((p) => window.officehereNative!.setDocumentState({ title: "Farklı kaydedilen", dirty: false, path: p }), saved);
  const session = () => JSON.parse(readFileSync(path.join(userData, "oturum.json"), "utf8")).belgeler;
  await expect.poll(session).toEqual([
    { path: docx, app: "belge" },
    { path: saved, app: "belge" },
  ]);
  // İlk pencere kapanınca oturumdan düşer; son pencere kapatılarak çıkılınca o belge kalır.
  await app.evaluate(({ BrowserWindow }) => BrowserWindow.getAllWindows().sort((a, b) => a.id - b.id)[0]?.close());
  await expect.poll(session).toEqual([{ path: saved, app: "belge" }]);
  const closed = app.waitForEvent("close");
  await app.evaluate(({ BrowserWindow }) => BrowserWindow.getAllWindows()[0]?.close()).catch(() => {});
  await closed;
  expect(session()).toEqual([{ path: saved, app: "belge" }]);
});

test("oturum: seçenek kapalıyken ya da dosya verilince oturum geri yüklenmez", async () => {
  const userData = tempDir("oh-oturum-");
  const docx = fixture("apps/belge/src/io/__fixtures__/libreoffice-html.docx");
  const pptx = fixture("apps/sunum/src/io/__fixtures__/lo-proje.pptx");
  writeFileSync(path.join(userData, "oturum.json"), JSON.stringify({ belgeler: [{ path: docx, app: "belge" }] }));

  // Seçenek kapalı (varsayılan): başlangıç ekranı.
  let { app, win } = await launch(userData);
  await expect(win).toHaveTitle(/Başlangıç - OfficeHere/);
  await app.close();

  // Seçenek açık ama komut satırında dosya var: yalnız o dosya açılır.
  writeFileSync(path.join(userData, "ayarlar.json"), JSON.stringify({ oturumuGeriYukle: true }));
  writeFileSync(path.join(userData, "oturum.json"), JSON.stringify({ belgeler: [{ path: docx, app: "belge" }] }));
  ({ app, win } = await launch(userData, [pptx]));
  await win.waitForSelector(".oh-ribbon");
  await expect(win.locator(".oh-titlebar")).toContainText("OfficeHere Sunum");
  await new Promise((r) => setTimeout(r, 1000));
  expect(app.windows()).toHaveLength(1);
  await app.close();
});
