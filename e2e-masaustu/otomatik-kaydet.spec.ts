import { _electron as electron, expect, test, type ElectronApplication, type Page } from "@playwright/test";
import type {} from "../packages/core/src/bridge";
import { execFileSync } from "node:child_process";
import { copyFileSync, existsSync, mkdtempSync, readdirSync, readFileSync, rmSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import path from "node:path";

// Otomatik kaydetme (masaüstü): diskteki .docx/.xlsx/.pptx düzenlendikçe Ctrl+S'siz
// aynı dosyaya yazılır, ilk yazmadan önce özgün dosya <ad>.bak olarak yedeklenir;
// pencere düzgün kapanınca .bak silinir, özgün içerik önceki sürümlerde kalır.

const root = path.resolve(__dirname, "..");
const desktop = path.join(root, "apps/desktop");
// Test için ayrı derleme dizini verilebilir (paralel derlemeler çakışmasın).
const appsDir = process.env.OH_E2E_APPS_DIR ?? path.join(root, "dist-web");
const SAVED = /^Kaydedildi \d{2}:\d{2}$/;

function electronPath(): string {
  return require(path.join(desktop, "node_modules/electron")) as unknown as string;
}

async function launch(file: string, userData = mkdtempSync(path.join(tmpdir(), "oh-otokayit-"))): Promise<{ app: ElectronApplication; win: Page; userData: string }> {
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
  return { app, win, userData };
}

/** Kullanıcı veri dizinindeki önceki sürümlerin içerikleri. */
function versionContents(userData: string): Buffer[] {
  const root = path.join(userData, "yedekler");
  if (!existsSync(root)) return [];
  return readdirSync(root).flatMap((d) =>
    readdirSync(path.join(root, d))
      .filter((n) => n.startsWith("surum-"))
      .map((n) => readFileSync(path.join(root, d, n))),
  );
}

/** Pencereyi kapatır (X); soru gelirse verilen düğmeye basılır. Uygulama son pencereyle kapanır. */
async function closeWindow(app: ElectronApplication, win: Page, answer?: string) {
  const closed = new Promise<void>((r) => app.once("close", () => r()));
  await app.evaluate(({ BrowserWindow }) => BrowserWindow.getAllWindows()[0]?.close());
  if (answer) await win.getByRole("dialog").getByRole("button", { name: answer, exact: true }).click();
  await closed;
}

/** Testler asıl fixture'ı bozmasın diye geçici kopya. */
function fixture(rel: string): string {
  const dir = mkdtempSync(path.join(tmpdir(), "oh-otokayit-dosya-"));
  const dest = path.join(dir, path.basename(rel));
  copyFileSync(path.join(root, rel), dest);
  return dest;
}

/** Belge'nin IndexedDB'deki kurtarma taslağı sayısı (veritabanı yoksa oluşturmadan 0). */
function belgeDraftCount(win: Page): Promise<number> {
  return win.evaluate(async () => {
    const dbs = await indexedDB.databases();
    if (!dbs.some((d) => d.name === "officehere-belge")) return 0;
    return new Promise<number>((res) => {
      const r = indexedDB.open("officehere-belge");
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
  // Kurtarma taslağını hemen yazdır (IndexedDB): otomatik kaydetmeden sonra silinmeli.
  await win.evaluate(() => window.dispatchEvent(new Event("officehere:autosave")));
  const drafts = () => belgeDraftCount(win);
  await expect.poll(drafts).toBe(1);
  // Taslak yerel depoya (5 MB sınırlı) değil IndexedDB'ye yazılır.
  expect(await win.evaluate(() => Object.keys(localStorage).filter((k) => k.startsWith("officehere.belge.draft.")).length)).toBe(0);

  await expect(badge(win)).toHaveText(SAVED);
  expect(existsSync(`${file}.bak`)).toBe(true);
  expect(readFileSync(`${file}.bak`).equals(original)).toBe(true);
  expect(execFileSync("unzip", ["-p", file, "word/document.xml"]).toString("utf8")).toContain("OTOMATIK-KAYIT");
  await expect.poll(drafts).toBe(0);
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

test("Belge: iki oturum — temiz kapanışta .bak silinir, özgün içerik önceki sürümlerde kalır; sonraki oturum yeni yedek alır", async () => {
  const file = fixture("apps/belge/src/io/__fixtures__/libreoffice-resaved-report.docx");
  for (const word of ["BIRINCI-OTURUM ", "IKINCI-OTURUM "]) {
    const original = readFileSync(file);
    const { app, win, userData } = await launch(file);
    await win.waitForSelector(".ProseMirror");
    await win.locator(".ProseMirror").first().click();
    await win.keyboard.press("Control+Home");
    await win.keyboard.type(word);
    await expect.poll(() => execFileSync("unzip", ["-p", file, "word/document.xml"]).toString("utf8"), { timeout: 20_000 }).toContain(word.trim());
    await expect(badge(win)).toHaveText(SAVED);
    expect(readFileSync(`${file}.bak`).equals(original)).toBe(true);
    await closeWindow(app, win);
    expect(existsSync(`${file}.bak`)).toBe(false);
    expect(versionContents(userData).some((b) => b.equals(original))).toBe(true);
  }
});

test("Tablo: kaydedilmemiş değişiklikle (\"Kaydetme\") kapatılınca .bak kalır; uygulamadan çıkışta temizse silinir", async () => {
  const file = fixture("apps/tablo/test/fixtures/lo-generated.xlsx");
  const original = readFileSync(file);
  const box = (w: Page) => w.getByRole("textbox", { name: "Ad Kutusu" });
  // 1. oturum: otomatik kaydedilir, sonra yeni değişiklik kaydedilmeden "Kaydetme" ile kapatılır.
  let { app, win } = await launch(file);
  await box(win).fill("H40");
  await box(win).press("Enter");
  await win.keyboard.type("OTOKAYIT");
  await win.keyboard.press("Enter");
  await expect(badge(win)).toHaveText(SAVED);
  expect(readFileSync(`${file}.bak`).equals(original)).toBe(true);
  // Sonraki değişiklik kendiliğinden kaydedilmesin (seçenek bu pencerede kapatılır).
  await win.evaluate(() => {
    localStorage.setItem("officehere.autoSave", "0");
    window.dispatchEvent(new Event("officehere:autosave-pref"));
  });
  await box(win).fill("H41");
  await box(win).press("Enter");
  await win.keyboard.type("KAYDEDILMEDI");
  await win.keyboard.press("Enter");
  await expect(badge(win)).toHaveText("Kaydedilmedi");
  // Kapatırken sorulur, "Kaydetme": kaydedilmemiş değişiklikle kapanış temiz sayılmaz.
  await closeWindow(app, win, "Kaydetme");
  expect(zipText(file)).not.toContain("KAYDEDILMEDI");
  expect(readFileSync(`${file}.bak`).equals(original)).toBe(true);
  rmSync(`${file}.bak`);

  // 2. oturum: otomatik kaydedilir, uygulamadan çıkılır (pencere temiz): .bak silinir.
  const beforeSecond = readFileSync(file);
  ({ app, win } = await launch(file));
  await box(win).fill("H42");
  await box(win).press("Enter");
  await win.keyboard.type("IKINCI");
  await win.keyboard.press("Enter");
  await expect(badge(win)).toHaveText(SAVED);
  expect(readFileSync(`${file}.bak`).equals(beforeSecond)).toBe(true);
  await app.close();
  expect(existsSync(`${file}.bak`)).toBe(false);
});

test("Sunum: kapatırken \"Kaydet\" seçilirse pencere temiz kapanmış sayılır, .bak silinir", async () => {
  const file = fixture("apps/sunum/src/io/__fixtures__/lo-proje.pptx");
  const original = readFileSync(file);
  const { app, win, userData } = await launch(file);
  await win.locator(".stage .sv [data-shape-id]").first().click();
  await win.keyboard.type("OTOSUNUM");
  await win.keyboard.press("Escape");
  await win.keyboard.press("Escape");
  await expect(badge(win)).toHaveText(SAVED);
  expect(readFileSync(`${file}.bak`).equals(original)).toBe(true);
  // Sonraki değişiklik kendiliğinden kaydedilmez; kapatırken sorulur ve "Kaydet" seçilir.
  await win.evaluate(() => {
    localStorage.setItem("officehere.autoSave", "0");
    window.dispatchEvent(new Event("officehere:autosave-pref"));
  });
  await win.locator(".thumb-add").click();
  await expect(badge(win)).toHaveText("Kaydedilmedi");
  await closeWindow(app, win, "Kaydet");
  expect(zipText(file)).toContain("OTOSUNUM");
  expect(existsSync(`${file}.bak`)).toBe(false);
  expect(versionContents(userData).some((b) => b.equals(original))).toBe(true);
});

test("Belge: uyarıyla açılan dosya ilk elle kayda kadar otomatik kaydedilmez; rozet nedenini söyler", async () => {
  const file = fixture("e2e-masaustu/fixtures/uyarili-nesne.docx");
  const original = readFileSync(file);
  const { app, win } = await launch(file);
  await win.waitForSelector(".ProseMirror");
  await win.locator(".ProseMirror").first().click();
  await win.keyboard.press("Control+Home");
  await win.keyboard.type("UYARILI ");
  // Rozet: otomatik kaydetme ilk elle kaydı bekliyor; neden sakin bir dille söylenir.
  await expect(badge(win)).toHaveText("Otomatik kayıt bekliyor");
  await expect(badge(win)).toHaveAttribute("aria-label", /desteklenmeyen öğeler var; ilk kaydı siz onaylayın/);
  await win.waitForTimeout(4000);
  expect(readFileSync(file).equals(original)).toBe(true);
  expect(existsSync(`${file}.bak`)).toBe(false);
  // Elle Kaydet: kullanıcı kararı; sonrasında otomatik kaydetme sürer.
  await win.keyboard.press("Control+S");
  await expect(badge(win)).toHaveText(SAVED);
  await win.locator(".ProseMirror").first().click();
  await win.keyboard.type("SONRA ");
  await expect.poll(() => execFileSync("unzip", ["-p", file, "word/document.xml"]).toString("utf8"), { timeout: 20_000 }).toContain("SONRA");
  await expect(badge(win)).toHaveText(SAVED);
  await app.close();
});

test("Kaydet: yazma hatası Türkçe nedeniyle rozette görünür, belge kirli kalır", async () => {
  const file = fixture("apps/belge/src/io/__fixtures__/libreoffice-html.docx");
  const { app, win } = await launch(file);
  await win.waitForSelector(".ProseMirror");
  // Otomatik kaydetme kapalı: yalnız elle Kaydet denenir.
  await win.evaluate(() => localStorage.setItem("officehere.autoSave", "0"));
  await win.locator(".ProseMirror").first().click();
  await win.keyboard.type("HATA ");
  // Klasör silindi (USB çıkarıldı, ağ sürücüsü koptu…): ENOENT → "Klasör bulunamadı".
  rmSync(path.dirname(file), { recursive: true, force: true });
  await win.keyboard.press("Control+S");
  await expect(badge(win)).toHaveText("Kaydedilemedi: Klasör bulunamadı");
  await expect(win).toHaveTitle(/^●/);
  await app.evaluate(({ BrowserWindow }) => BrowserWindow.getAllWindows().forEach((w) => w.destroy()));
  await app.close();
});

test("Önceki sürümler: Dosya › Bilgi'de listelenir ve geri yüklenir", async () => {
  const file = fixture("apps/belge/src/io/__fixtures__/libreoffice-resaved-report.docx");
  const original = readFileSync(file);
  const { app, win } = await launch(file);
  await win.waitForSelector(".ProseMirror");
  await win.locator(".ProseMirror").first().click();
  await win.keyboard.press("Control+Home");
  await win.keyboard.type("SURUM-DEGISIKLIGI ");
  await expect(badge(win)).toHaveText(SAVED);
  expect(execFileSync("unzip", ["-p", file, "word/document.xml"]).toString("utf8")).toContain("SURUM-DEGISIKLIGI");
  await win.locator(".oh-ribbon-file").click();
  await win.locator(".oh-bs-item", { hasText: "Bilgi" }).first().click();
  const versions = win.getByRole("region", { name: "Önceki sürümler" });
  await expect(versions.locator("li")).toHaveCount(1);
  await versions.getByRole("button", { name: "Geri yükle" }).click();
  const dlg = win.getByRole("dialog", { name: "Önceki sürümü geri yükle" });
  await dlg.getByRole("button", { name: "Geri yükle" }).click();
  // Dosya özgün içeriğe döner ve pencere belgeyi diskten yeniden açar.
  await expect.poll(() => readFileSync(file).equals(original), { timeout: 15_000 }).toBe(true);
  await expect(win.locator(".ProseMirror").first()).not.toContainText("SURUM-DEGISIKLIGI", { timeout: 15_000 });
  await expect(win.locator(".oh-titlebar-title")).toContainText("libreoffice-resaved-report");
  await app.close();
});

