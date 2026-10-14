import { _electron as electron, expect, test, type ElectronApplication, type Page } from "@playwright/test";
import type {} from "../packages/core/src/bridge";
import { execFileSync } from "node:child_process";
import { copyFileSync, existsSync, mkdtempSync, readdirSync, readFileSync } from "node:fs";
import { tmpdir } from "node:os";
import path from "node:path";

// Çökme ve donma yönetimi (ölçüm: /tmp/olcum/rapor/electron.json): kirli belgede sayfa
// işlemi çöktüğünde pencere kapatılamıyor ve uygulamadan çıkılamıyordu; donan sayfada
// kapatma isteği karşılanmıyordu. Ayrıca: aynı dosya iki kez açılınca tek pencere,
// "Hata bildir" penceresi ve yerel günlük.

const root = path.resolve(__dirname, "..");
const desktop = path.join(root, "apps/desktop");
const appsDir = process.env.OH_E2E_APPS_DIR ?? path.join(root, "dist-web");

interface Launched {
  app: ElectronApplication;
  first: Page;
  userData: string;
}

async function launch(args: string[] = [], env: Record<string, string> = {}): Promise<Launched> {
  const userData = mkdtempSync(path.join(tmpdir(), "oh-cokme-"));
  const app = await electron.launch({
    executablePath: require(path.join(desktop, "node_modules/electron")) as unknown as string,
    args: [desktop, "--no-sandbox", ...args],
    env: { ...process.env, OH_APPS_DIR: appsDir, OH_USER_DATA: userData, ELECTRON_DISABLE_SECURITY_WARNINGS: "1", ...env },
  });
  const first = await app.firstWindow();
  await first.waitForSelector(".oh-titlebar");
  await stubDialogs(app);
  return { app, first, userData };
}

/**
 * Yerel ileti kutuları otomasyonla tıklanamaz: ana süreçte dialog.showMessageBox
 * kaydeden ve verilen yanıtı dönen bir sahteyle değiştirilir.
 */
async function stubDialogs(app: ElectronApplication) {
  await app.evaluate(({ dialog }) => {
    const g = globalThis as unknown as { __dialogs: { message: string; buttons: string[]; detail?: string }[]; __dialogResponse: number };
    g.__dialogs = [];
    g.__dialogResponse = 0;
    (dialog as unknown as { showMessageBox: unknown }).showMessageBox = async (...args: unknown[]) => {
      const opts = (args.length > 1 ? args[1] : args[0]) as { message: string; buttons: string[]; detail?: string };
      g.__dialogs.push({ message: opts.message, buttons: opts.buttons, detail: opts.detail });
      return { response: g.__dialogResponse, checkboxChecked: false };
    };
  });
}

const dialogs = (app: ElectronApplication) => app.evaluate(() => (globalThis as unknown as { __dialogs: { message: string; buttons: string[] }[] }).__dialogs);
const setResponse = (app: ElectronApplication, n: number) =>
  app.evaluate((_e, v) => {
    (globalThis as unknown as { __dialogResponse: number }).__dialogResponse = v;
  }, n);
const windowCount = (app: ElectronApplication) => app.evaluate(({ BrowserWindow }) => BrowserWindow.getAllWindows().filter((w) => !w.isDestroyed()).length);

/** Başlangıç ekranından yeni bir Belge penceresi açar, metin yazar ve kurtarma taslağını hemen yazdırır. */
async function dirtyBelge(app: ElectronApplication, launcher: Page, text: string): Promise<Page> {
  const [win] = await Promise.all([app.waitForEvent("window"), launcher.locator(".ohl-app", { hasText: "Belge" }).click()]);
  await win.waitForSelector(".ProseMirror");
  await win.locator(".ProseMirror").first().click();
  await win.keyboard.type(text);
  await expect(win.locator(".oh-titlebar-dirty")).toBeVisible();
  await win.evaluate(() => window.dispatchEvent(new Event("officehere:autosave")));
  // Taslak IndexedDB'ye yazılana kadar beklenir.
  await expect.poll(() => belgeDraftCount(win)).toBeGreaterThan(0);
  return win;
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

const belgeTitle = (app: ElectronApplication) => app.evaluate(({ BrowserWindow }) => BrowserWindow.getAllWindows().find((w) => w.getTitle().includes("Belge"))?.id ?? -1);

test("sayfa çökünce kurtarma önerilir: 'Yeniden yükle ve taslaktan kurtar' taslağı geri getirir", async () => {
  const { app, first, userData } = await launch();
  const win = await dirtyBelge(app, first, "ÇÖKMEDEN ÖNCE YAZILAN");
  const id = await belgeTitle(app);
  await setResponse(app, 0);
  await app.evaluate(({ BrowserWindow }, wid) => BrowserWindow.fromId(wid)?.webContents.forcefullyCrashRenderer(), id);
  await expect.poll(async () => (await dialogs(app)).map((d) => d.message)).toContain("Pencere beklenmedik biçimde kapandı");
  const crash = (await dialogs(app)).find((d) => d.message === "Pencere beklenmedik biçimde kapandı")!;
  expect(crash.buttons).toEqual(["Yeniden yükle ve taslaktan kurtar", "Pencereyi kapat"]);
  // Pencere aynı pencerede yeniden yüklenir; taslağın sahibi kapanmış sayıldığında kurtarma
  // çubuğu çıkar. (Playwright çöken sayfanın nesnesini kullanılamaz bırakır: sayfaya ana süreçten erişilir.)
  const js = <T,>(code: string) =>
    app.evaluate(
      ({ BrowserWindow }, [wid, c]) => {
        const wc = BrowserWindow.fromId(wid as number)!.webContents;
        // Yükleme sürerken yanıt gelmeyebilir: kısa zaman aşımıyla yeniden denenir.
        if (!wc.getURL().includes("/belge/") || wc.isLoading()) return null;
        return Promise.race([wc.executeJavaScript(c as string), new Promise((r) => setTimeout(() => r(null), 2000))]);
      },
      [id, code] as const,
    ) as Promise<T>;
  await expect.poll(() => js<string>("document.querySelector('.oh-messagebar')?.textContent ?? ''"), { timeout: 30_000 }).toContain("Kurtarılan belge");
  expect(await windowCount(app)).toBe(2);
  await js("[...document.querySelectorAll('.oh-messagebar button')].find((b) => b.textContent.trim() === 'Geri Yükle').click()");
  await expect.poll(() => js<string>("document.querySelector('.ProseMirror')?.textContent ?? ''")).toContain("ÇÖKMEDEN ÖNCE YAZILAN");
  void win;
  // Çökme yerel günlüğe yazıldı.
  expect(readFileSync(path.join(userData, "gunluk", "officehere.log"), "utf8")).toMatch(/HATA \[belge\] Pencere işlemi sonlandı/);
  // Kurtarılan belge kirli: kapatırken sorulmasın diye pencereler doğrudan kapatılır.
  await app.evaluate(({ BrowserWindow }) => BrowserWindow.getAllWindows().forEach((w) => w.destroy()));
  await app.close();
});

test("çöken kirli pencere kapatılabilir ve uygulamadan çıkılabilir", async () => {
  const { app, first } = await launch();
  await dirtyBelge(app, first, "KAYDEDİLMEMİŞ");
  const id = await belgeTitle(app);
  // Soru yanıtlanmadan önce bile kapatma engellenmez.
  await app.evaluate(({ dialog }) => {
    (dialog as unknown as { showMessageBox: unknown }).showMessageBox = () => new Promise(() => undefined);
  });
  await app.evaluate(({ BrowserWindow }, wid) => BrowserWindow.fromId(wid)?.webContents.forcefullyCrashRenderer(), id);
  await expect.poll(() => app.evaluate(({ BrowserWindow }, wid) => BrowserWindow.fromId(wid)?.webContents.isCrashed() ?? true, id)).toBe(true);
  await app.evaluate(({ BrowserWindow }, wid) => BrowserWindow.fromId(wid)?.close(), id);
  await expect.poll(() => app.evaluate(({ BrowserWindow }, wid) => !BrowserWindow.fromId(wid) || BrowserWindow.fromId(wid)!.isDestroyed(), id)).toBe(true);
  // Çıkış takılmaz.
  const exited = new Promise<void>((r) => app.process().once("exit", () => r()));
  await app.evaluate(({ app: a }) => a.quit());
  await Promise.race([exited, new Promise((_, rej) => setTimeout(() => rej(new Error("uygulama kapanmadı")), 10_000))]);
});

test("donan sayfa: kapatma isteğine yanıt gelmezse Bekle / Yeniden yükle / Kapat sorulur", async () => {
  const { app, first } = await launch([], { OH_CLOSE_ACK_MS: "1500" });
  const win = await dirtyBelge(app, first, "DONMADAN ÖNCE");
  const id = await belgeTitle(app);
  await setResponse(app, 2); // "Kapat"
  // Sayfa 20 sn yanıt vermez.
  void win.evaluate(() => {
    const t = Date.now();
    while (Date.now() - t < 20_000) {
      /* meşgul */
    }
  }).catch(() => undefined);
  await new Promise((r) => setTimeout(r, 300));
  await app.evaluate(({ BrowserWindow }, wid) => BrowserWindow.fromId(wid)?.close(), id);
  await expect.poll(async () => (await dialogs(app)).map((d) => d.message), { timeout: 10_000 }).toContain("Pencere yanıt vermiyor");
  const hang = (await dialogs(app)).find((d) => d.message === "Pencere yanıt vermiyor")!;
  expect(hang.buttons).toEqual(["Bekle", "Yeniden yükle", "Kapat"]);
  await expect.poll(() => app.evaluate(({ BrowserWindow }, wid) => !BrowserWindow.fromId(wid) || BrowserWindow.fromId(wid)!.isDestroyed(), id), { timeout: 10_000 }).toBe(true);
  expect(await windowCount(app)).toBe(1);
  await app.close();
});

test("aynı dosya ikinci kez açılınca yeni pencere açılmaz, var olan öne gelir", async () => {
  const dir = mkdtempSync(path.join(tmpdir(), "oh-ayni-"));
  const file = path.join(dir, "rapor.docx");
  copyFileSync(path.join(root, "apps/belge/src/io/__fixtures__/libreoffice-resaved-report.docx"), file);
  const { app, first } = await launch([file]);
  await expect(first.locator(".oh-titlebar-title")).toContainText("rapor");
  expect(await windowCount(app)).toBe(1);
  // İşletim sistemi dosyayı yeniden açtı (çift tık / ikinci örnek).
  await app.evaluate(({ app: a }, f) => a.emit("second-instance", {}, [process.execPath, ".", f], process.cwd()), file);
  // Son kullanılanlar / başlangıç ekranı: aynı dosya için yeni pencere istendi.
  // Pencere ana süreçte eşzamanlı oluşturulur: istek döndüğünde yeni pencere olsaydı sayılırdı.
  await first.evaluate((f) => window.officehereNative!.newWindow("belge", f), file);
  expect(await windowCount(app)).toBe(1);
  // Boş bir Belge penceresi aynı dosyayı belge olarak açmak isterse okunmaz; var olan pencere öne gelir.
  const [second] = await Promise.all([app.waitForEvent("window"), first.evaluate(() => window.officehereNative!.newWindow("belge"))]);
  await second.waitForSelector(".ProseMirror");
  expect(await second.evaluate((f) => window.officehereNative!.readFile(f).then((r) => r === null), file)).toBe(true);
  expect(await windowCount(app)).toBe(2);
  await app.close();
});

test("Hata bildir: sürüm, işletim sistemi ve son günlük satırları gösterilir; GitHub adresi önceden doldurulur", async () => {
  const { app, first, userData } = await launch();
  await app.evaluate(({ shell }) => {
    const g = globalThis as unknown as { __opened: string[] };
    g.__opened = [];
    (shell as unknown as { openExternal: unknown }).openExternal = async (u: string) => void g.__opened.push(u);
  });
  const [win] = await Promise.all([app.waitForEvent("window"), first.locator(".ohl-app", { hasText: "Belge" }).click()]);
  await win.waitForSelector(".ProseMirror");
  // Sayfada yakalanmamış hata: günlüğe yazılır (belge içeriği yazılmaz).
  await win.evaluate(() => {
    setTimeout(() => {
      throw new Error("deneme hatası 42");
    });
  });
  await expect.poll(() => readFileSync(path.join(userData, "gunluk", "officehere.log"), "utf8")).toContain("deneme hatası 42");
  await win.getByRole("tab", { name: "Yardım" }).click();
  await win.getByRole("button", { name: "Hata Bildir" }).click();
  const dlg = win.getByRole("dialog", { name: "Hata bildir" });
  await expect(dlg).toBeVisible();
  const body = dlg.getByRole("textbox", { name: "Rapor" });
  await expect(body).toHaveValue(/OfficeHere Belge \d+\.\d+\.\d+ \(masaüstü\)/);
  await expect(body).toHaveValue(/İşletim sistemi: linux/);
  await expect(body).toHaveValue(/deneme hatası 42/);
  // Kullanıcı düzenleyebilir.
  await body.fill("Kullanıcının kendi açıklaması");
  await dlg.getByRole("button", { name: "GitHub'da bildir" }).click();
  await expect.poll(() => app.evaluate(() => (globalThis as unknown as { __opened: string[] }).__opened.length)).toBe(1);
  const url = (await app.evaluate(() => (globalThis as unknown as { __opened: string[] }).__opened[0])) as string;
  expect(url.startsWith("https://github.com/fyildirim-debug/officehere/issues/new?title=")).toBe(true);
  expect(decodeURIComponent(url)).toContain("Kullanıcının kendi açıklaması");
  // Günlük kullanıcı veri dizininde, döner dosyada.
  expect(readdirSync(path.join(userData, "gunluk"))).toContain("officehere.log");
  expect(existsSync(path.join(userData, "gunluk", "officehere.log"))).toBe(true);
  await app.close();
});

test("aynı dosya Aç iletişim kutusuyla ikinci pencerede açılırsa uyarılır ve orada otomatik kaydedilmez", async () => {
  const dir = mkdtempSync(path.join(tmpdir(), "oh-catisma-"));
  const file = path.join(dir, "ortak.docx");
  copyFileSync(path.join(root, "apps/belge/src/io/__fixtures__/libreoffice-resaved-report.docx"), file);
  const before = readFileSync(file);
  const { app, first, userData } = await launch([file]);
  await expect(first.locator(".oh-titlebar-title")).toContainText("ortak");
  const [second] = await Promise.all([app.waitForEvent("window"), first.evaluate(() => window.officehereNative!.newWindow("belge"))]);
  await second.waitForSelector(".ProseMirror");
  // Aç iletişim kutusunda aynı dosya seçilir.
  await app.evaluate(({ dialog }, f) => {
    (dialog as unknown as { showOpenDialog: unknown }).showOpenDialog = async () => ({ canceled: false, filePaths: [f] });
  }, file);
  await second.locator(".ProseMirror").first().click();
  await second.keyboard.press("Control+O");
  await expect(second.locator(".oh-titlebar-title")).toContainText("ortak");
  const alert = second.getByRole("dialog", { name: "Dosya başka bir pencerede açık" });
  await expect(alert).toBeVisible();
  await alert.getByRole("button", { name: "Tamam" }).click();
  await second.locator(".ProseMirror").first().click();
  await second.keyboard.type("IKINCI PENCERE ");
  await expect(second.locator(".oh-titlebar .oh-savebadge")).toHaveText("Otomatik kayıt kapalı");
  // Saat olarak ilk pencerenin otomatik kaydı kullanılır: ikinci pencere daha önce yazdığı için onun
  // otomatik kaydetme turu (yazacak olsaydı) ilk pencerenin yazmasından önce çalışmış olur.
  await first.locator(".ProseMirror").first().click();
  await first.keyboard.press("Control+Home");
  await first.keyboard.type("BIRINCI ");
  await expect.poll(() => execFileSync("unzip", ["-p", file, "word/document.xml"]).toString("utf8"), { timeout: 20_000 }).toContain("BIRINCI");
  // Belgeyi açmış ilk pencerenin kaydı doğrulandı (kendi dosyasını okuması engellenmez).
  await expect(first.locator(".oh-titlebar .oh-savebadge")).toHaveText(/^Kaydedildi \d{2}:\d{2}$/);
  expect(execFileSync("unzip", ["-p", file, "word/document.xml"]).toString("utf8")).not.toContain("IKINCI");
  // İkinci pencere yazmış olsaydı, ilk pencerenin yazmasından önce alınan "önceki sürüm" onun içeriğini taşırdı.
  const versions = readdirSync(path.join(userData, "yedekler"), { recursive: true, withFileTypes: true }).filter((d) => d.isFile() && d.name.startsWith("surum-"));
  expect(versions.length).toBeGreaterThan(0);
  for (const v of versions) expect(execFileSync("unzip", ["-p", path.join(v.parentPath, v.name), "word/document.xml"]).toString("utf8")).not.toContain("IKINCI");
  expect(before.length).toBeGreaterThan(0);
  await app.evaluate(({ BrowserWindow }) => BrowserWindow.getAllWindows().forEach((w) => w.destroy()));
  await app.close();
});

test("ana süreçteki beklenmeyen hata günlüğe yazılır, kullanıcı uyarılır, uygulama çalışmayı sürdürür", async () => {
  const { app, first, userData } = await launch();
  await app.evaluate(() => {
    setTimeout(() => {
      throw new Error("ana süreç deneme hatası");
    });
  });
  await expect.poll(async () => (await dialogs(app)).map((d) => d.message)).toContain("Beklenmeyen bir hata oluştu");
  expect(readFileSync(path.join(userData, "gunluk", "officehere.log"), "utf8")).toContain("ana süreç deneme hatası");
  // Uygulama yanıt vermeye devam eder.
  const [win] = await Promise.all([app.waitForEvent("window"), first.locator(".ohl-app", { hasText: "Belge" }).click()]);
  await win.waitForSelector(".ProseMirror");
  await app.close();
});
