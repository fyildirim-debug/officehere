// Electron ölçümü: ana süreç başlangıç süresi (başlangıç ekranı / belgeyle) ve "renderer çökünce ya da donunca
// kirli pencere kapanabiliyor mu?" denemesi.
// Ön koşul: pnpm web && pnpm --filter @officehere/desktop build:electron (Electron ikilisi kurulu).
// Kullanım: xvfb-run -a node scripts/derlem/olcum/electron_olcum.mjs → .derlem/rapor/electron.json
import { _electron as electron } from "playwright";
import { copyFileSync, existsSync, mkdirSync, mkdtempSync, writeFileSync } from "node:fs";
import { createRequire } from "node:module";
import { tmpdir } from "node:os";
import path from "node:path";
import { DERLEM, KOK, RAPOR } from "../ortak.mjs";

const desktop = path.join(KOK, "apps/desktop");
const exe = createRequire(import.meta.url)(path.join(desktop, "node_modules/electron"));
const UYG = existsSync(path.join(KOK, "dist-web")) ? { OH_APPS_DIR: path.join(KOK, "dist-web") } : {};
const RAPOR_YOLU = path.join(RAPOR, "electron.json");
const ORNEK = path.join(KOK, "apps/belge/src/io/__fixtures__/bicimler/lo-ornek.docx");
const out = {};
const withTimeout = (p, ms, label) => Promise.race([p, new Promise((r) => setTimeout(() => r(`ZAMAN AŞIMI (${label})`), ms))]);
const dump = () => {
  mkdirSync(RAPOR, { recursive: true });
  writeFileSync(RAPOR_YOLU, JSON.stringify(out, null, 1));
};

async function launch(args = [], env = {}) {
  const userData = mkdtempSync(path.join(tmpdir(), "oh-olcum-"));
  const t0 = Date.now();
  const app = await electron.launch({ executablePath: exe, args: [desktop, "--no-sandbox", ...args], env: { ...process.env, ...UYG, OH_USER_DATA: userData, ELECTRON_DISABLE_SECURITY_WARNINGS: "1", ...env } });
  const tLaunch = Date.now() - t0;
  const win = await app.firstWindow();
  const tWin = Date.now() - t0;
  await win.waitForSelector(".oh-titlebar", { timeout: 60000 });
  const tReady = Date.now() - t0;
  return { app, win, t: { süreçMs: tLaunch, ilkPencereMs: tWin, arayüzHazırMs: tReady } };
}

// 1) Soğuk başlangıç (başlangıç ekranı) x3 ve belgeyle açılış x3
const starts = [];
for (let i = 0; i < 3; i++) {
  const { app, t } = await launch();
  starts.push(t);
  await app.close();
}
out.başlangıçEkranı = starts;
console.log("başlangıç", JSON.stringify(starts)); dump();
const docStarts = [];
for (let i = 0; i < 3; i++) {
  const dir = mkdtempSync(path.join(tmpdir(), "oh-dosya-"));
  const f = path.join(dir, "rapor.docx");
  copyFileSync(path.join(DERLEM, "lo/lo-rapor.docx"), f);
  const t0 = Date.now();
  const { app, win, t } = await launch([f]);
  await win.waitForFunction(() => [...document.querySelectorAll(".ProseMirror")].some((e) => e.textContent?.includes("Bölüm 8")), null, { timeout: 60000 });
  docStarts.push({ ...t, belgeYüklendiMs: Date.now() - t0 });
  await app.close();
}
out.belgeyleAçılış = docStarts;
console.log("belgeyle", JSON.stringify(docStarts)); dump();

// 2) Kirli belge + renderer çökmesi → pencere kapatılabiliyor mu?
{
  const dir = mkdtempSync(path.join(tmpdir(), "oh-dosya-"));
  const f = path.join(dir, "cokme.docx");
  copyFileSync(ORNEK, f);
  const { app, win } = await launch([f]);
  await win.waitForSelector(".ProseMirror");
  await win.click(".ProseMirror");
  await win.keyboard.type("kaydedilmemiş değişiklik");
  await win.waitForTimeout(800);
  const title = await app.evaluate(({ BrowserWindow }) => BrowserWindow.getAllWindows()[0].getTitle());
  await app.evaluate(({ BrowserWindow }) => BrowserWindow.getAllWindows()[0].webContents.forcefullyCrashRenderer());
  await new Promise((r) => setTimeout(r, 1500));
  await withTimeout(app.evaluate(({ BrowserWindow }) => { BrowserWindow.getAllWindows()[0].close(); }), 5000, "close");
  await new Promise((r) => setTimeout(r, 3000));
  const alive = await withTimeout(app.evaluate(({ BrowserWindow }) => BrowserWindow.getAllWindows().map((w) => ({ destroyed: w.isDestroyed(), visible: w.isVisible(), crashed: w.webContents.isCrashed(), title: w.getTitle() }))), 5000, "durum");
  console.log("çökme sonrası pencereler", JSON.stringify(alive));
  // uygulamadan çıkış denemesi
  await withTimeout(app.evaluate(({ app }) => { app.quit(); }).catch(() => {}), 5000, "quit");
  await new Promise((r) => setTimeout(r, 3000));
  let stillRunning = app.process().exitCode === null && !app.process().killed;
  out.rendererÇökmesi = { kirliBaşlık: title, kapatmaSonrasıPencereler: alive, quitSonrasıHâlâÇalışıyor: stillRunning };
  console.log("çökme", JSON.stringify(out.rendererÇökmesi)); dump();
  try { app.process().kill("SIGKILL"); } catch {}
}

// 3) Kirli belge + renderer yanıt vermiyor (sonsuz döngü) → kapatma
{
  const dir = mkdtempSync(path.join(tmpdir(), "oh-dosya-"));
  const f = path.join(dir, "donma.docx");
  copyFileSync(ORNEK, f);
  const { app, win } = await launch([f]);
  await win.waitForSelector(".ProseMirror");
  await win.click(".ProseMirror");
  await win.keyboard.type("değişiklik");
  await win.waitForTimeout(800);
  console.log("donma testi başlıyor");
  // 20 sn sürecek senkron döngü
  win.evaluate(() => { const t = Date.now(); while (Date.now() - t < 20000) {} }).catch(() => {});
  await new Promise((r) => setTimeout(r, 500));
  const t0 = Date.now();
  await withTimeout(app.evaluate(({ BrowserWindow }) => { BrowserWindow.getAllWindows()[0].close(); }), 5000, "close");
  await new Promise((r) => setTimeout(r, 3000));
  const n = await withTimeout(app.evaluate(({ BrowserWindow }) => BrowserWindow.getAllWindows().filter((w) => !w.isDestroyed()).length), 5000, "durum");
  out.rendererDonması = { kapatmaİsteğinden3snSonraAçıkPencere: n, not: "unresponsive olayı için işleyici yok; kapatma isteği donmuş sayfaya iletilir" };
  try { app.process().kill("SIGKILL"); } catch {}
}
dump();
console.log(JSON.stringify(out, null, 1));
process.exit(0);
