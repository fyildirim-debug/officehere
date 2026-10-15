import { _electron as electron, expect, test, type ElectronApplication, type Page } from "@playwright/test";
import type {} from "../packages/core/src/bridge";
import { execFileSync } from "node:child_process";
import { copyFileSync, existsSync, mkdirSync, mkdtempSync, readFileSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import path from "node:path";

// Görüntüleyici kipi (masaüstü).
// - Seçenekler › "İndirilenler ve e-posta eklerinden gelen dosyaları salt okunur aç" açıkken internetten
//   gelen dosya sarı şeritle salt okunur açılır: düzenleme ve otomatik kayıt kapalı; "Düzenlemeyi
//   etkinleştir" sonrası her şey normal. Linux'ta kaynak İndirilenler dizinidir: test sahte bir dizini
//   XDG_CONFIG_HOME/user-dirs.dirs ile gösterir (Windows'ta Zone.Identifier, macOS'ta karantina özniteliği).

const root = path.resolve(__dirname, "..");
const desktop = path.join(root, "apps/desktop");
const appsDir = process.env.OH_E2E_APPS_DIR ?? path.join(root, "dist-web");
const SAVED = /^Kaydedildi \d{2}:\d{2}$/;
const BANNER_TEXT = "Bu dosya internetten geldi; güvenliğiniz için salt okunur açıldı.";

function electronPath(): string {
  return require(path.join(desktop, "node_modules/electron")) as unknown as string;
}

interface Ortam {
  /** Sahte İndirilenler dizini. */
  downloads: string;
  userData: string;
  env: Record<string, string>;
}

/** Yalıtılmış kullanıcı verisi, sahte İndirilenler ve (verilirse) görüntüleyici kipi ayarı. */
function ortam(saltOkunur?: boolean): Ortam {
  const base = mkdtempSync(path.join(tmpdir(), "oh-goruntuleyici-"));
  const downloads = path.join(base, "İndirilenler");
  const config = path.join(base, "config");
  const userData = path.join(base, "kullanici");
  for (const d of [downloads, config, userData]) mkdirSync(d);
  writeFileSync(path.join(config, "user-dirs.dirs"), `# xdg-user-dirs\nXDG_DOWNLOAD_DIR="${downloads}"\n`);
  if (saltOkunur !== undefined) writeFileSync(path.join(userData, "ayarlar.json"), JSON.stringify({ indirilenleriSaltOkunurAc: saltOkunur }));
  return { downloads, userData, env: { XDG_CONFIG_HOME: config } };
}

/** Fixture'ın verilen dizine (ya da yeni geçici dizine) kopyası; İndirilenler'e kopyalanınca internetten gelmiş sayılır. */
function kopya(rel: string, dir?: string): string {
  const target = dir ?? mkdtempSync(path.join(tmpdir(), "oh-goruntuleyici-dosya-"));
  const dest = path.join(target, path.basename(rel));
  copyFileSync(path.join(root, rel), dest);
  // Windows ve macOS'ta kaynak dizinden değil dosyanın işaretinden anlaşılır.
  if (dir && process.platform === "win32") writeFileSync(`${dest}:Zone.Identifier`, "[ZoneTransfer]\r\nZoneId=3\r\n");
  if (dir && process.platform === "darwin") execFileSync("xattr", ["-w", "com.apple.quarantine", "0083;00000000;Safari;", dest]);
  return dest;
}

async function launch(args: string[], o: Ortam): Promise<ElectronApplication> {
  return electron.launch({
    executablePath: electronPath(),
    args: [desktop, "--no-sandbox", ...args],
    env: { ...process.env, OH_APPS_DIR: appsDir, OH_USER_DATA: o.userData, ELECTRON_DISABLE_SECURITY_WARNINGS: "1", ...o.env },
  });
}

async function openDoc(file: string, o: Ortam): Promise<{ app: ElectronApplication; win: Page }> {
  const app = await launch([file], o);
  const win = await app.firstWindow();
  await win.waitForLoadState("domcontentloaded");
  await win.waitForSelector(".oh-ribbon");
  await expect(win.locator(".oh-titlebar-title")).toContainText(path.basename(file).replace(/\.[^.]+$/, ""));
  return { app, win };
}

const zipPart = (file: string, part: string) => execFileSync("unzip", ["-p", file, part]).toString("utf8");
const zipText = (file: string) => execFileSync("unzip", ["-p", file]).toString("utf8");
const badge = (win: Page) => win.locator(".oh-titlebar .oh-savebadge");

/** Belge gövdesi: parça parça yükleme bitince düzenleyici açılır (salt okunur kip içeriği değil girdiyi süzer). */
async function belgeBody(win: Page) {
  const doc = win.locator(".ProseMirror").first();
  await expect(doc).toHaveAttribute("contenteditable", "true");
  return doc;
}
const banner = (win: Page) => win.locator(".oh-readonly-banner");

/**
 * Salt okunur kipte hiçbir şey yazılmadı: pencere kirli değil (otomatik kaydetme yalnız kirli belgeyi
 * yazar; salt okunur bayrağı ayrıca autoSave birim testinde), dosya özgün, yedek yok, rozet yok.
 */
async function expectUntouched(win: Page, file: string, original: Buffer) {
  await expect(win).not.toHaveTitle(/^●/);
  await expect(badge(win)).toHaveCount(0);
  expect(readFileSync(file).equals(original)).toBe(true);
  expect(existsSync(`${file}.bak`)).toBe(false);
}

async function enableEditing(win: Page) {
  await banner(win).getByRole("button", { name: "Düzenlemeyi etkinleştir" }).click();
  await expect(banner(win)).toHaveCount(0);
}

test.describe("görüntüleyici kipi", () => {
  test("Belge: İndirilenler'deki .docx salt okunur açılır; düzenleme ve otomatik kayıt kapalı, etkinleştirince açılır", async () => {
    const o = ortam(true);
    const file = kopya("apps/belge/src/io/__fixtures__/libreoffice-resaved-report.docx", o.downloads);
    const original = readFileSync(file);
    const { app, win } = await openDoc(file, o);
    await expect(banner(win)).toContainText(BANNER_TEXT);
    const doc = await belgeBody(win);
    await doc.click();
    await win.keyboard.press("Control+Home");
    await win.keyboard.type("SALT-OKUNUR ");
    await win.keyboard.press("Enter");
    await expect(doc).not.toContainText("SALT-OKUNUR");
    await expectUntouched(win, file, original);

    await enableEditing(win);
    await doc.click();
    await win.keyboard.press("Control+Home");
    await win.keyboard.type("ETKIN ");
    await expect.poll(() => zipPart(file, "word/document.xml"), { timeout: 20_000 }).toContain("ETKIN");
    await expect(badge(win)).toHaveText(SAVED);
    expect(zipPart(file, "word/document.xml")).not.toContain("SALT-OKUNUR");
    expect(readFileSync(`${file}.bak`).equals(original)).toBe(true);
    await app.close();
  });

  test("Tablo: İndirilenler'deki .xlsx'te hücre düzenlenemez, etkinleştirince otomatik kaydedilir", async () => {
    const o = ortam(true);
    const file = kopya("apps/tablo/test/fixtures/lo-generated.xlsx", o.downloads);
    const original = readFileSync(file);
    const { app, win } = await openDoc(file, o);
    await expect(banner(win)).toContainText(BANNER_TEXT);
    const box = win.getByRole("textbox", { name: "Ad Kutusu" });
    const fx = win.getByRole("textbox", { name: "Formül Çubuğu" });
    await box.fill("H40");
    await box.press("Enter");
    await win.keyboard.type("SALTTABLO");
    await win.keyboard.press("Enter");
    await box.fill("H40");
    await box.press("Enter");
    await expect(fx).toHaveValue("");
    await expectUntouched(win, file, original);

    await enableEditing(win);
    await box.fill("H40");
    await box.press("Enter");
    await win.keyboard.type("ETKINTABLO");
    await win.keyboard.press("Enter");
    await expect(badge(win)).toHaveText(SAVED, { timeout: 20_000 });
    expect(zipText(file)).toContain("ETKINTABLO");
    expect(zipText(file)).not.toContain("SALTTABLO");
    expect(readFileSync(`${file}.bak`).equals(original)).toBe(true);
    await app.close();
  });

  test("Sunum: İndirilenler'deki .pptx'te metin değişmez, etkinleştirince otomatik kaydedilir", async () => {
    const o = ortam(true);
    const file = kopya("apps/sunum/src/io/__fixtures__/lo-proje.pptx", o.downloads);
    const original = readFileSync(file);
    const { app, win } = await openDoc(file, o);
    await expect(banner(win)).toContainText(BANNER_TEXT);
    const shape = win.locator(".stage .sv [data-shape-id]").first();
    await shape.click();
    await win.keyboard.type("SALTSUNUM");
    await win.keyboard.press("Escape");
    await win.keyboard.press("Escape");
    await expect(win.locator(".stage")).not.toContainText("SALTSUNUM");
    await expectUntouched(win, file, original);

    await enableEditing(win);
    await shape.click();
    await win.keyboard.type("ETKINSUNUM");
    await win.keyboard.press("Escape");
    await win.keyboard.press("Escape");
    await expect(badge(win)).toHaveText(SAVED, { timeout: 20_000 });
    expect(zipText(file)).toContain("ETKINSUNUM");
    expect(zipText(file)).not.toContain("SALTSUNUM");
    expect(readFileSync(`${file}.bak`).equals(original)).toBe(true);
    await app.close();
  });

  test("ayar kapalıyken (varsayılan) İndirilenler'deki dosyada şerit yok, otomatik kayıt sürer", async () => {
    const o = ortam();
    const file = kopya("apps/belge/src/io/__fixtures__/libreoffice-html.docx", o.downloads);
    const { app, win } = await openDoc(file, o);
    const doc = await belgeBody(win);
    await doc.click();
    await win.keyboard.press("Control+Home");
    await win.keyboard.type("AYAR-KAPALI ");
    await expect.poll(() => zipPart(file, "word/document.xml"), { timeout: 20_000 }).toContain("AYAR-KAPALI");
    await expect(banner(win)).toHaveCount(0);
    await app.close();
  });

  test("ayar açıkken İndirilenler dışındaki dosya düzenlenebilir açılır", async () => {
    const o = ortam(true);
    const file = kopya("apps/belge/src/io/__fixtures__/libreoffice-html.docx");
    const { app, win } = await openDoc(file, o);
    const doc = await belgeBody(win);
    await doc.click();
    await win.keyboard.press("Control+Home");
    await win.keyboard.type("YEREL-DOSYA ");
    await expect.poll(() => zipPart(file, "word/document.xml"), { timeout: 20_000 }).toContain("YEREL-DOSYA");
    await expect(banner(win)).toHaveCount(0);
    await app.close();
  });

  test("Başlangıç › Seçenekler: seçenek varsayılan kapalı, açılınca ayarlara yazılır", async () => {
    const o = ortam();
    const app = await launch([], o);
    const win = await app.firstWindow();
    await win.waitForLoadState("domcontentloaded");
    await win.locator(".oh-bs-item", { hasText: "Seçenekler" }).first().click();
    const option = win.getByRole("checkbox", { name: "İndirilenler ve e-posta eklerinden gelen dosyaları salt okunur aç" });
    await expect(option).not.toBeChecked();
    await option.check();
    await expect(option).toBeChecked();
    await expect.poll(() => JSON.parse(readFileSync(path.join(o.userData, "ayarlar.json"), "utf8")).indirilenleriSaltOkunurAc).toBe(true);
    await app.close();
  });
});
