// Worker RPC duman testi: Vite ile derlenen modül Worker'ı, paketlenmiş uygulamadaki gibi file://
// sayfasından (OH_APPS_DIR/belge/index.html), uygulamanın pencere ayarları ve CSP'siyle açılır.
import { _electron as electron, expect, test, type ElectronApplication, type Page } from "@playwright/test";
import { execFileSync } from "node:child_process";
import { copyFileSync, mkdtempSync, readFileSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import path from "node:path";
import type { api } from "./worker-sayfa/main";

const root = path.resolve(__dirname, "..");
const desktop = path.join(root, "apps/desktop");
let appsDir = "";

type Api = typeof api;
/** Sayfadaki test arayüzünü çağırır (işlev metin olarak sayfaya gider; dönen söz beklenir). */
const call = <T>(win: Page, fn: (api: Api) => T | Promise<T>) => win.evaluate(`(${fn.toString()})(window.ohWorkerTest)`) as Promise<Awaited<T>>;

function electronPath(): string {
  return require(path.join(desktop, "node_modules/electron")) as unknown as string;
}

test.beforeAll(() => {
  // Sayfa uygulamanın yerine konur: OH_APPS_DIR/belge/index.html.
  appsDir = mkdtempSync(path.join(tmpdir(), "oh-worker-sayfa-"));
  const out = path.join(appsDir, "belge");
  execFileSync(process.execPath, [path.join(root, "node_modules/vite/bin/vite.js"), "build", "--config", "e2e-masaustu/worker-sayfa/vite.config.mjs", "--outDir", out], { cwd: root, stdio: "pipe" });
  const csp = /http-equiv="Content-Security-Policy" content="([^"]+)"/.exec(readFileSync(path.join(root, "apps/belge/index.html"), "utf8"))?.[1];
  expect(csp, "Belge CSP'si").toContain("worker-src");
  const html = path.join(out, "index.html");
  writeFileSync(html, readFileSync(html, "utf8").replace("__UYGULAMA_CSP__", csp!));
});

async function launch(): Promise<{ app: ElectronApplication; win: Page; errors: string[] }> {
  const userData = mkdtempSync(path.join(tmpdir(), "oh-masaustu-"));
  // Belge penceresi açılsın diye var olan bir .docx verilir; sayfa onu okumaz.
  const docx = path.join(mkdtempSync(path.join(tmpdir(), "oh-dosya-")), "bos.docx");
  copyFileSync(path.join(root, "apps/belge/src/io/__fixtures__/libreoffice-resaved-report.docx"), docx);
  const app = await electron.launch({
    executablePath: electronPath(),
    args: [desktop, "--no-sandbox", docx],
    env: { ...process.env, OH_APPS_DIR: appsDir, OH_USER_DATA: userData, ELECTRON_DISABLE_SECURITY_WARNINGS: "1" },
  });
  const win = await app.firstWindow();
  const errors: string[] = [];
  win.on("pageerror", (e) => errors.push(`pageerror: ${e.message}`));
  win.on("console", (m) => m.type() === "error" && errors.push(`console: ${m.text()}`));
  await win.waitForFunction(() => "ohWorkerTest" in window);
  return { app, win, errors };
}

test("file:// modül Worker açılır, ArrayBuffer aktarılır, hata sınıfı korunur, İptal 1 sn içinde Worker'ı durdurur", async () => {
  const { app, win, errors } = await launch();
  expect(win.url()).toMatch(/^file:.*\/belge\/index\.html/);

  const env = await call(win, (t) => t.ortam());
  expect(env).toMatchObject({ protocol: "file:", domParser: "undefined", compression: "function" });

  const moved = await call(win, (t) => t.aktar(32));
  expect(moved.detached).toBe(true);
  expect(moved.back).toBe(32 << 20);
  expect(moved.sum).toBe(moved.expected);

  expect(await call(win, (t) => t.hata())).toEqual({
    format: true,
    name: "FormatError",
    message: "Dosya bozuk ya da beklenmeyen bir yapıda; okunamadı.",
    detail: "zip: end of central directory not found",
    save: true,
    code: "ENOSPC",
    reason: "Disk dolu",
  });

  // Eşzamanlı döngüde kilitli Worker; durum çubuğundaki İptal klavyeyle etkinleştirilir.
  await call(win, (t) => t.baslat(20_000));
  const bar = win.getByRole("progressbar", { name: "Açılıyor…" });
  await expect(bar).toHaveAttribute("aria-valuenow", /^[1-9]\d*$/);
  await expect(win.locator(".oh-progress")).toHaveText(/^Açılıyor…%\d+İptal$/);
  await win.keyboard.press("Tab");
  const cancel = win.getByRole("button", { name: "İptal" });
  await expect(cancel).toBeFocused();
  await win.keyboard.press("Enter");
  await expect.poll(() => call(win, (t) => t.sonuc()?.error)).toBe("AbortError");
  const run = (await call(win, (t) => t.sonuc()))!;
  expect(run.cancelMs).toBeLessThan(1000);
  await expect(bar).toHaveCount(0);

  // Yeni Worker kurulup hemen yanıt verir (eski Worker sürseydi 20 sn beklenirdi).
  const ping = await call(win, (t) => t.ping());
  expect(ping.id).not.toBe(env.instance);
  expect(ping.ms).toBeLessThan(1000);
  expect(await call(win, (t) => t.kurulan())).toBe(2);
  expect(errors).toEqual([]);

  const missing = await call(win, (t) => t.yuklenemez());
  expect(missing).toMatchObject({ cls: true, code: "load" });

  console.log(`Worker ölçümleri: 32 MB aktarım ${moved.ms} ms · iptal ${run.cancelMs} ms · yeni Worker ilk yanıt ${ping.ms} ms · ilerleme iletisi ${run.progress}`);
  await app.close();
});

test("ilerleme göstergesi açık ve koyu temalarda okunur", async ({}, testInfo) => {
  const { app, win, errors } = await launch();
  await call(win, (t) => t.goster(0.4));
  await expect(win.locator(".oh-progress")).toHaveText("Açılıyor…%40İptal");
  await expect(win.getByRole("progressbar", { name: "Açılıyor…" })).toHaveAttribute("aria-valuetext", "%40");
  const fills = new Set<string>();
  for (const app of ["belge", "tablo", "sunum"]) {
    for (const theme of ["colorful", "dark", "black"]) {
      const k = await win.evaluate(([a, t]) => (window as unknown as { ohWorkerTest: Api }).ohWorkerTest.karsitlik(a, t), [app, theme]);
      const label = `${app}/${theme}: ${JSON.stringify(k)}`;
      // WCAG: arayüz bileşeni 3:1, metin 4,5:1.
      expect(k.dolguIz, label).toBeGreaterThanOrEqual(3);
      expect(k.dolguZemin, label).toBeGreaterThanOrEqual(3);
      expect(k.metin, label).toBeGreaterThanOrEqual(4.5);
      expect(k.dugme, label).toBeGreaterThanOrEqual(4.5);
      fills.add(k.fill);
      if (app === "belge" && theme !== "black") await testInfo.attach(`ilerleme-${theme}.png`, { body: await win.locator(".oh-statusbar").screenshot({ animations: "disabled" }), contentType: "image/png" });
    }
  }
  // Koyu temada dolgu rengi açılır (her uygulamada açık ve koyu iki ayrı renk).
  expect(fills.size).toBe(6);
  expect(errors).toEqual([]);
  await app.close();
});
