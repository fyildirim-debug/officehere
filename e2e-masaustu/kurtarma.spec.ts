import { _electron as electron, expect, test, type ElectronApplication, type Page } from "@playwright/test";
import { mkdtempSync } from "node:fs";
import { tmpdir } from "node:os";
import path from "node:path";

// Kurtarma taslakları: kullanıcı "Kaydetme" diyerek kapattığı belgenin taslağı
// bir sonraki açılışta önerilmez; yalnız çökme sonrası kalan taslak önerilir.
// Kurtarma çubuğu Kapat (×) ile kapatılırsa taslak silinir, yeniden önerilmez.

const root = path.resolve(__dirname, "..");
const desktop = path.join(root, "apps/desktop");

async function launch(userData = mkdtempSync(path.join(tmpdir(), "oh-kurtarma-"))): Promise<{ app: ElectronApplication; launcher: Page }> {
  const app = await electron.launch({
    executablePath: require(path.join(desktop, "node_modules/electron")) as unknown as string,
    args: [desktop, "--no-sandbox"],
    env: { ...process.env, OH_APPS_DIR: process.env.OH_E2E_APPS_DIR ?? path.join(root, "dist-web"), OH_USER_DATA: userData, ELECTRON_DISABLE_SECURITY_WARNINGS: "1" },
  });
  const launcher = await app.firstWindow();
  await launcher.waitForSelector(".oh-titlebar");
  return { app, launcher };
}

/** Başlangıç ekranından yeni bir Sunum penceresi açar. */
async function openSunum(app: ElectronApplication, launcher: Page): Promise<Page> {
  const [win] = await Promise.all([app.waitForEvent("window"), launcher.locator(".ohl-app", { hasText: "Sunum" }).click()]);
  await win.waitForSelector(".thumb-pane");
  return win;
}

/** Sunuyu kirletir ve taslağın yazılmasını bekler (boşta kalma süresi + yazma). */
async function dirtyAndWaitDraft(win: Page) {
  await win.locator(".thumb-add").click();
  await expect(win.locator(".oh-titlebar-dirty")).toBeVisible();
  await win.waitForTimeout(4000);
}

const recoveryBar = (win: Page) => win.locator(".oh-messagebar", { hasText: "Kurtarılan sunu" });

test("Kaydetme ile kapatılan sununun taslağı yeni pencerede önerilmez", async () => {
  const { app, launcher } = await launch();
  const first = await openSunum(app, launcher);
  await dirtyAndWaitDraft(first);
  // Pencereyi kapat: sorulur, "Kaydetme" seçilir.
  await app.evaluate(({ BrowserWindow }) => BrowserWindow.getAllWindows().find((w) => w.getTitle().includes("Sunum"))?.close());
  const dialog = first.getByRole("dialog", { name: "OfficeHere" });
  await expect(dialog).toBeVisible();
  await dialog.getByRole("button", { name: "Kaydetme" }).click().catch(() => {});
  await expect.poll(() => app.evaluate(({ BrowserWindow }) => BrowserWindow.getAllWindows().length)).toBe(1);
  // Yeni pencerede kurtarma önerisi çıkmaz.
  const second = await openSunum(app, launcher);
  await second.waitForTimeout(1500);
  await expect(recoveryBar(second)).toHaveCount(0);
  await app.close();
});

test("çöken sununun taslağı yeni pencerede okunur düğmelerle önerilir", async () => {
  const { app, launcher } = await launch();
  const first = await openSunum(app, launcher);
  await dirtyAndWaitDraft(first);
  // Çökme: süreç sorulmadan sonlanır.
  await app.evaluate(({ BrowserWindow }) => {
    const w = BrowserWindow.getAllWindows().find((x) => x.getTitle().includes("Sunum"));
    w?.webContents.forcefullyCrashRenderer();
    w?.destroy();
  });
  await expect.poll(() => app.evaluate(({ BrowserWindow }) => BrowserWindow.getAllWindows().length)).toBe(1);
  const second = await openSunum(app, launcher);
  await expect(recoveryBar(second)).toBeVisible();
  const sil = recoveryBar(second).getByRole("button", { name: "Sil", exact: true });
  await expect(sil).toBeVisible();
  // Koyu temada da ikincil düğmenin yazısı zemininden ayrışır.
  await second.evaluate(() => document.documentElement.setAttribute("data-oh-theme", "dark"));
  const { fg, bg } = await sil.evaluate((el) => ({ fg: getComputedStyle(el).color, bg: getComputedStyle(el).backgroundColor }));
  expect(fg).not.toBe(bg);
  expect(fg).toBe("rgb(240, 240, 240)");
  await app.close();
});

/** Üç uygulama: yeni pencere, belgeyi kirletme ve kurtarma çubuğunun metni. */
const APPS = [
  {
    name: "Belge",
    ready: ".ProseMirror",
    bar: "Kurtarılan belge",
    dirty: async (w: Page) => {
      await w.locator(".ProseMirror").first().click();
      await w.keyboard.type("Kurtarma denemesi");
    },
  },
  {
    name: "Tablo",
    ready: ".oh-ribbon",
    bar: "Kurtarılan çalışma kitabı",
    dirty: async (w: Page) => {
      const box = w.getByRole("textbox", { name: "Ad Kutusu" });
      await box.fill("B2");
      await box.press("Enter");
      await w.keyboard.type("Kurtarma");
      await w.keyboard.press("Enter");
    },
  },
  { name: "Sunum", ready: ".thumb-pane", bar: "Kurtarılan sunu", dirty: (w: Page) => w.locator(".thumb-add").click() },
] as const;

async function openApp(app: ElectronApplication, launcher: Page, a: (typeof APPS)[number]): Promise<Page> {
  const [win] = await Promise.all([app.waitForEvent("window"), launcher.locator(".ohl-app", { hasText: a.name }).click()]);
  await win.waitForSelector(a.ready);
  return win;
}

/** Uygulamanın IndexedDB'deki kurtarma taslağı sayısı (veritabanı yoksa oluşturmadan 0). */
function draftCount(win: Page, a: (typeof APPS)[number]): Promise<number> {
  return win.evaluate(async (name) => {
    const dbs = await indexedDB.databases();
    if (!dbs.some((d) => d.name === name)) return 0;
    return new Promise<number>((res) => {
      const r = indexedDB.open(name);
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
  }, `officehere-${a.name.toLowerCase()}`); // yerel-ayar: ascii
}

/** Belgeyi kirletir ve taslağın yazılmasını bekler. */
async function dirtyWithDraft(win: Page, a: (typeof APPS)[number]) {
  await a.dirty(win);
  await expect(win.locator(".oh-titlebar-dirty")).toBeVisible();
  await win.evaluate(() => window.dispatchEvent(new Event("officehere:autosave")));
  await expect.poll(() => draftCount(win, a), { timeout: 20_000 }).toBeGreaterThan(0);
}

/** Yeni pencerede kurtarma önerisi yok: depoda taslak kalmamış ve çubuk görünmüyor. */
async function expectNoOffer(win: Page, a: (typeof APPS)[number]) {
  await expect.poll(() => draftCount(win, a)).toBe(0);
  await expect(win.locator(".oh-messagebar", { hasText: a.bar })).toHaveCount(0);
}

const windows = (app: ElectronApplication) => app.evaluate(({ BrowserWindow }) => BrowserWindow.getAllWindows().length);

for (const a of APPS) {
  test(`${a.name}: X ile kapatılıp "Kaydetme" seçilen belgenin taslağı uygulama yeniden açılınca önerilmez`, async () => {
    const userData = mkdtempSync(path.join(tmpdir(), "oh-kurtarma-"));
    const first = await launch(userData);
    const win = await openApp(first.app, first.launcher, a);
    await dirtyWithDraft(win, a);
    await first.app.evaluate(({ BrowserWindow }, n) => BrowserWindow.getAllWindows().find((w) => w.getTitle().includes(n))?.close(), a.name);
    const dialog = win.getByRole("dialog", { name: "OfficeHere" });
    await expect(dialog).toBeVisible();
    await dialog.getByRole("button", { name: "Kaydetme" }).click().catch(() => {});
    await expect.poll(() => windows(first.app)).toBe(1);
    await first.app.close();
    // Uygulama yeniden açılır: kurtarma önerisi çıkmaz.
    const again = await launch(userData);
    await expectNoOffer(await openApp(again.app, again.launcher, a), a);
    await again.app.close();
  });

  test(`${a.name}: kurtarma çubuğu Kapat (×) ile kapatılınca taslak silinir, yeniden önerilmez`, async () => {
    const userData = mkdtempSync(path.join(tmpdir(), "oh-kurtarma-"));
    const first = await launch(userData);
    const win = await openApp(first.app, first.launcher, a);
    await dirtyWithDraft(win, a);
    // Çökme: pencere sorulmadan yok edilir, taslak kalır.
    await first.app.evaluate(({ BrowserWindow }, n) => {
      const w = BrowserWindow.getAllWindows().find((x) => x.getTitle().includes(n));
      w?.webContents.forcefullyCrashRenderer();
      w?.destroy();
    }, a.name);
    await expect.poll(() => windows(first.app)).toBe(1);
    // Çöken pencerenin canlılık nabzı eskiyene kadar taslağı "sahipli" sayılır: öneri çıkana dek yeniden yüklenir.
    const second = await openApp(first.app, first.launcher, a);
    const bar = second.locator(".oh-messagebar", { hasText: a.bar });
    await expect
      .poll(
        async () => {
          if (await bar.isVisible()) return true;
          await second.reload();
          await second.waitForSelector(a.ready);
          return bar.waitFor({ timeout: 2000 }).then(
            () => true,
            () => false,
          );
        },
        { timeout: 30_000 },
      )
      .toBe(true);
    await bar.getByRole("button", { name: "Taslağı sil ve kapat" }).click();
    await expect(bar).toHaveCount(0);
    await expect.poll(() => draftCount(second, a)).toBe(0);
    await first.app.close();
    // Uygulama yeniden açılır: kapatılan taslak yeniden önerilmez.
    const again = await launch(userData);
    await expectNoOffer(await openApp(again.app, again.launcher, a), a);
    await again.app.close();
  });
}
