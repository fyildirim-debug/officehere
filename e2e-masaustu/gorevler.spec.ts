// Atlama listesi / Dock / masaüstü eylemi görevleri: "--yeni=<uygulama>" ile açılış ve çalışan örneğe
// iletilen istek (second-instance).
import { _electron as electron, expect, test, type ElectronApplication, type Page } from "@playwright/test";
import { spawn } from "node:child_process";
import { mkdtempSync } from "node:fs";
import { tmpdir } from "node:os";
import path from "node:path";

const root = path.resolve(__dirname, "..");
const desktop = path.join(root, "apps/desktop");
const appsDir = process.env.OH_E2E_APPS_DIR ?? path.join(root, "dist-web");

function electronPath(): string {
  return require(path.join(desktop, "node_modules/electron")) as unknown as string;
}

async function launch(args: string[], userData = mkdtempSync(path.join(tmpdir(), "oh-gorev-"))): Promise<{ app: ElectronApplication; win: Page }> {
  const app = await electron.launch({
    executablePath: electronPath(),
    args: [desktop, "--no-sandbox", ...args],
    env: { ...process.env, OH_APPS_DIR: appsDir, OH_USER_DATA: userData, ELECTRON_DISABLE_SECURITY_WARNINGS: "1" },
  });
  const win = await app.firstWindow();
  await win.waitForLoadState("domcontentloaded");
  return { app, win };
}

test("--yeni=tablo ile açılışta yalnız boş Tablo penceresi açılır (başlangıç ekranı yok)", async () => {
  const { app, win } = await launch(["--yeni=tablo"]);
  await expect(win).toHaveURL(/\/tablo\/index\.html/);
  await win.waitForSelector(".oh-titlebar");
  expect(app.windows()).toHaveLength(1);
  await app.close();
});

test("çalışan örneğe iletilen görev o uygulamanın penceresini, görevsiz istek başlangıç ekranını açar", async () => {
  const { app, win } = await launch([]);
  await expect(win).toHaveURL(/\/index\.html/);
  const yeni = app.waitForEvent("window");
  await app.evaluate(({ app: a }) => a.emit("second-instance", {}, [process.execPath, ".", "--yeni=sunum"], process.cwd()));
  const sunum = await yeni;
  await expect(sunum).toHaveURL(/\/sunum\/index\.html/);
  const baska = app.waitForEvent("window");
  await app.evaluate(({ app: a }) => a.emit("second-instance", {}, [process.execPath, "."], process.cwd()));
  const baslangic = await baska;
  await baslangic.waitForLoadState("domcontentloaded");
  expect(baslangic.url()).not.toMatch(/\/(belge|tablo|sunum)\/index\.html/);
  await app.close();
});

test("ikinci süreç olarak başlatılan görev çalışan örnekte açılır ve kendisi kapanır", async () => {
  const userData = mkdtempSync(path.join(tmpdir(), "oh-gorev-"));
  const { app } = await launch([], userData);
  const yeni = app.waitForEvent("window");
  // Atlama listesi görevinin yaptığı: aynı programı bayrakla başlatmak.
  const ikinci = spawn(electronPath(), [desktop, "--no-sandbox", "--yeni=belge"], {
    env: { ...process.env, OH_APPS_DIR: appsDir, OH_USER_DATA: userData, ELECTRON_DISABLE_SECURITY_WARNINGS: "1" },
    stdio: "ignore",
  });
  const cikis = new Promise<number | null>((r) => ikinci.on("exit", r));
  const belge = await yeni;
  await expect(belge).toHaveURL(/\/belge\/index\.html/);
  expect(await cikis).toBe(0);
  await app.close();
});
