import { _electron as electron, expect, test, type ElectronApplication, type Page } from "@playwright/test";
import { mkdtempSync } from "node:fs";
import { tmpdir } from "node:os";
import path from "node:path";

// Kurtarma taslakları: kullanıcı "Kaydetme" diyerek kapattığı belgenin taslağı
// bir sonraki açılışta önerilmez; yalnız çökme sonrası kalan taslak önerilir.

const root = path.resolve(__dirname, "..");
const desktop = path.join(root, "apps/desktop");

async function launch(): Promise<{ app: ElectronApplication; launcher: Page }> {
  const userData = mkdtempSync(path.join(tmpdir(), "oh-kurtarma-"));
  const app = await electron.launch({
    executablePath: require(path.join(desktop, "node_modules/electron")) as unknown as string,
    args: [desktop, "--no-sandbox"],
    env: { ...process.env, OH_APPS_DIR: path.join(root, "dist-web"), OH_USER_DATA: userData, ELECTRON_DISABLE_SECURITY_WARNINGS: "1" },
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
  const sil = recoveryBar(second).getByRole("button", { name: "Sil" });
  await expect(sil).toBeVisible();
  // Koyu temada da ikincil düğmenin yazısı zemininden ayrışır.
  await second.evaluate(() => document.documentElement.setAttribute("data-oh-theme", "dark"));
  const { fg, bg } = await sil.evaluate((el) => ({ fg: getComputedStyle(el).color, bg: getComputedStyle(el).backgroundColor }));
  expect(fg).not.toBe(bg);
  expect(fg).toBe("rgb(240, 240, 240)");
  await app.close();
});
