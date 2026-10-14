// "Ağ isteği yok" vaadi: onay verilmemiş profilde uygulama dışarıya hiçbir istek göndermez.
// Chromium ağ günlüğü (--log-net-log) kapanışta okunur; URL isteği, DNS çözümleme ve bağlantı
// havuzu olaylarındaki her yerel olmayan adres dış istek sayılır. Yazım sözlüğü (Google) indirmesi
// yalnız "İndir" onayından sonra başlar; testte indirme adresi yerel bir sunucuya çevrilir.
import { _electron as electron, expect, test, type ElectronApplication, type Page } from "@playwright/test";
import type {} from "../packages/core/src/bridge";
import { existsSync, mkdirSync, mkdtempSync, readFileSync, readdirSync, writeFileSync } from "node:fs";
import { createServer } from "node:http";
import type { AddressInfo } from "node:net";
import { tmpdir } from "node:os";
import path from "node:path";

const root = path.resolve(__dirname, "..");
const desktop = path.join(root, "apps/desktop");
const appsDir = process.env.OH_E2E_APPS_DIR ?? path.join(root, "dist-web");
const isMac = process.platform === "darwin";

function electronPath(): string {
  return require(path.join(desktop, "node_modules/electron")) as unknown as string;
}

async function launch(userData: string, env: Record<string, string> = {}): Promise<{ app: ElectronApplication; win: Page; netLog: string }> {
  const netLog = path.join(mkdtempSync(path.join(tmpdir(), "oh-ag-")), "net.json");
  const app = await electron.launch({
    executablePath: electronPath(),
    args: [desktop, "--no-sandbox", `--log-net-log=${netLog}`, "--net-log-capture-mode=Default"],
    env: { ...process.env, OH_APPS_DIR: appsDir, OH_USER_DATA: userData, ELECTRON_DISABLE_SECURITY_WARNINGS: "1", ...env },
  });
  const win = await app.firstWindow();
  await win.waitForSelector(".oh-titlebar");
  return { app, win, netLog };
}

/**
 * Kaydedilmemiş belge sorusu çıkmadan kapatır; ağ günlüğü süreç çıkarken tamamlanır. macOS'ta son
 * pencere kapanınca uygulama çıkmadığı için ayrıca app.quit() çağrılır (yoksa süreç hiç bitmez).
 */
async function close(app: ElectronApplication) {
  const exited = new Promise((r) => app.process().once("exit", r));
  await app.evaluate(({ app: electronApp, BrowserWindow }) => {
    for (const w of BrowserWindow.getAllWindows()) w.destroy();
    electronApp.quit();
  });
  await exited;
}

const LOCAL = /^(127\.\d+\.\d+\.\d+|localhost|\[?::1\]?)$/i;

/** Ağ günlüğündeki dış adresler (URL isteği, DNS çözümleme, bağlantı havuzu). */
function externalHosts(netLog: string): string[] {
  const d = JSON.parse(readFileSync(netLog, "utf8")) as { constants: { logEventTypes: Record<string, number> }; events: { type: number; params?: Record<string, unknown> }[] };
  const t = d.constants.logEventTypes;
  const found = new Set<string>();
  const add = (u: unknown) => {
    if (typeof u !== "string") return;
    const m = /^(?:pm\/)?(https?|wss?|ftp):\/\/([^/:?#]+|\[[^\]]+\])/i.exec(u);
    if (m && !LOCAL.test(m[2])) found.add(`${m[1]}://${m[2]}`);
  };
  for (const e of d.events) {
    if (e.type === t.URL_REQUEST_START_JOB) add(e.params?.url);
    else if (e.type === t.HOST_RESOLVER_MANAGER_REQUEST) add(e.params?.host);
    else if (e.type === t.SOCKET_POOL_CONNECT_JOB_CREATED) add(e.params?.group_id);
  }
  return [...found];
}

/** İndirme isteklerini kaydeden yerel sözlük sunucusu (dosya vermez: 404). */
async function dictionaryServer() {
  const requests: string[] = [];
  const server = createServer((req, res) => {
    requests.push(path.basename(new URL(req.url ?? "/", "http://yerel").pathname));
    res.statusCode = 404;
    res.end();
  });
  await new Promise<void>((r) => server.listen(0, "127.0.0.1", r));
  return { url: `http://127.0.0.1:${(server.address() as AddressInfo).port}/sozluk/`, requests, close: () => server.close() };
}

const spellState = (app: ElectronApplication) =>
  app.evaluate(({ session }) => ({ langs: session.defaultSession.getSpellCheckerLanguages(), on: session.defaultSession.isSpellCheckerEnabled() }));

/** Belge'de yazı yazar: yazım denetimi istenen bir düzenleyici etkin olsun. */
async function typeInBelge(app: ElectronApplication, win: Page) {
  const [belge] = await Promise.all([app.waitForEvent("window"), win.locator(".ohl-app", { hasText: "Belge" }).click()]);
  await belge.locator(".bl-editor").first().click();
  await belge.keyboard.type("merhabaa dunya yanlis yazim ");
  // Belge cümle başını büyük harfe çevirir.
  await expect(belge.locator(".bl-editor").first()).toContainText("erhabaa dunya");
  return belge;
}

test("onay verilmemiş temiz profilde dış ağ isteği yok", async () => {
  const userData = mkdtempSync(path.join(tmpdir(), "oh-ag-profil-"));
  const { app, win, netLog } = await launch(userData);
  const prompt = win.locator(".oh-spell-prompt");
  if (isMac) await expect(prompt).toHaveCount(0);
  else {
    await expect(prompt).toContainText("Google sunucusundan indirilir");
    await expect.poll(() => spellState(app)).toEqual({ langs: [], on: false });
  }
  await typeInBelge(app, win);
  await close(app);
  expect(externalHosts(netLog)).toEqual([]);
  expect(existsSync(path.join(userData, "Dictionaries")) ? readdirSync(path.join(userData, "Dictionaries")).filter((f) => f.endsWith(".bdic")) : []).toEqual([]);
});

test.describe("yazım sözlüğü onayı", () => {
  test.skip(isMac, "macOS sistemin yazım denetleyicisini kullanır; sözlük indirilmez");

  test("İndir: indirme yalnız onaydan sonra ve yalnız iki sözlük için başlar", async () => {
    const dict = await dictionaryServer();
    const userData = mkdtempSync(path.join(tmpdir(), "oh-ag-profil-"));
    const { app, win, netLog } = await launch(userData, { OH_SPELLCHECK_DICT_URL: dict.url });
    await typeInBelge(app, win);
    expect(dict.requests).toEqual([]);

    await win.locator(".oh-spell-prompt").getByRole("button", { name: "İndir" }).click();
    await expect.poll(() => [...new Set(dict.requests)].sort()).toEqual(["en-us-10-1.bdic", "tr-tr-4-0.bdic"]);
    await expect(win.locator(".oh-spell-prompt")).toHaveCount(0);
    expect(JSON.parse(readFileSync(path.join(userData, "ayarlar.json"), "utf8")).yazimSozlugu).toBe("izinli");

    // Sunucu dosya vermedi: Seçenekler'de durum ve yeniden deneme.
    await win.locator(".oh-bs-item", { hasText: "Seçenekler" }).click();
    const section = win.locator(".oh-spell-section");
    await expect(section).toContainText("Sözlükler indirilemedi");
    const before = dict.requests.length;
    await section.getByRole("button", { name: "İndir" }).click();
    await expect.poll(() => dict.requests.length).toBeGreaterThan(before);
    await close(app);
    dict.close();
    expect(externalHosts(netLog)).toEqual([]);
  });

  test("Şimdi değil: soru bir daha sorulmaz, istek gitmez", async () => {
    const dict = await dictionaryServer();
    const userData = mkdtempSync(path.join(tmpdir(), "oh-ag-profil-"));
    const first = await launch(userData, { OH_SPELLCHECK_DICT_URL: dict.url });
    await first.win.locator(".oh-spell-prompt").getByRole("button", { name: "Şimdi değil" }).click();
    await expect(first.win.locator(".oh-spell-prompt")).toHaveCount(0);
    await close(first.app);

    const again = await launch(userData, { OH_SPELLCHECK_DICT_URL: dict.url });
    await expect(again.win.locator(".ohl-app", { hasText: "Belge" })).toBeVisible();
    await expect(again.win.locator(".oh-spell-prompt")).toHaveCount(0);
    await expect.poll(() => spellState(again.app)).toEqual({ langs: [], on: false });
    await typeInBelge(again.app, again.win);
    await again.win.locator(".oh-bs-item", { hasText: "Seçenekler" }).click();
    await expect(again.win.locator(".oh-spell-section")).toContainText("Sözlükler indirilmedi");
    await close(again.app);
    dict.close();
    expect(dict.requests).toEqual([]);
    expect(externalHosts(first.netLog)).toEqual([]);
    expect(externalHosts(again.netLog)).toEqual([]);
  });

  test("sözlükler zaten yereldeyse soru sorulmaz ve indirme denenmez", async () => {
    const dict = await dictionaryServer();
    const userData = mkdtempSync(path.join(tmpdir(), "oh-ag-profil-"));
    // Önceki sürümlerin indirdiği dosyalar (adları Chromium'un beklediği adlar).
    mkdirSync(path.join(userData, "Dictionaries"));
    for (const f of ["tr-TR-4-0.bdic", "en-US-10-1.bdic"]) writeFileSync(path.join(userData, "Dictionaries", f), "");
    const { app, win, netLog } = await launch(userData, { OH_SPELLCHECK_DICT_URL: dict.url });
    await expect(win.locator(".ohl-app", { hasText: "Belge" })).toBeVisible();
    await expect(win.locator(".oh-spell-prompt")).toHaveCount(0);
    await expect.poll(() => spellState(app)).toEqual({ langs: ["tr", "en-US"], on: true });
    await typeInBelge(app, win);
    await close(app);
    dict.close();
    // Onay yokken indirme adresi ağa çıkmaz: dosya geçersiz bulunsa bile (test dosyaları boş) istek gitmez.
    expect(dict.requests).toEqual([]);
    expect(externalHosts(netLog)).toEqual([]);
  });
});
