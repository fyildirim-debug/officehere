import { _electron as electron, expect, test, type ElectronApplication, type Page } from "@playwright/test";
import type {} from "../packages/core/src/bridge";
import { execFileSync } from "node:child_process";
import { copyFileSync, existsSync, mkdirSync, mkdtempSync, readFileSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import path from "node:path";

// Görüntüleyici kipi ve PDF görüntüleme (masaüstü).
// - Seçenekler › "İndirilenler ve e-posta eklerinden gelen dosyaları salt okunur aç" açıkken internetten
//   gelen dosya sarı şeritle salt okunur açılır: düzenleme ve otomatik kayıt kapalı; "Düzenlemeyi
//   etkinleştir" sonrası her şey normal. Linux'ta kaynak İndirilenler dizinidir: test sahte bir dizini
//   XDG_CONFIG_HOME/user-dirs.dirs ile gösterir (Windows'ta Zone.Identifier, macOS'ta karantina özniteliği).
// - PDF ayrı pencerede Chromium'un yerleşik görüntüleyicisiyle açılır; .pdf dosya ilişkilendirmesi yoktur.

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

/** PDF görüntüleyicisinin ana süreçteki penceresi: güvenlik ayarları ve yüklenen adres. */
function pdfWindowInfo(app: ElectronApplication) {
  return app.evaluate(({ BrowserWindow }) =>
    BrowserWindow.getAllWindows()
      .filter((w) => /\.pdf$/i.test(w.webContents.getURL()))
      .map((w) => {
        // Belgelenmemiş ama Electron'da bulunan okuyucu: pencerenin gerçekte aldığı ayarlar.
        const wc = w.webContents as unknown as { getLastWebPreferences(): Electron.WebPreferences | null };
        const p = wc.getLastWebPreferences() ?? {};
        return { url: w.webContents.getURL(), title: w.getTitle(), node: p.nodeIntegration, sandbox: p.sandbox, isolation: p.contextIsolation, preload: p.preload ?? null };
      }),
  );
}

/** Görüntüleyici yüklendi: PDF eklentisinin iç çerçevesi (yerleşik uzantı) var. */
async function waitForViewer(win: Page) {
  await expect.poll(() => win.frames().some((f) => f.url().startsWith("chrome-extension://")), { timeout: 20_000 }).toBe(true);
}

/** PDF penceresine yerel tuş girdisi (gerçek klavyeyle aynı yol; ana süreç kısayolları bunu görür). */
function pdfKey(app: ElectronApplication, keyCode: string, modifiers: ("control" | "shift")[] = []) {
  return app.evaluate(
    ({ BrowserWindow }, [k, m]) => {
      const wc = BrowserWindow.getAllWindows().find((w) => /\.pdf$/i.test(w.webContents.getURL()))!.webContents;
      wc.sendInputEvent({ type: "keyDown", keyCode: k as string, modifiers: m as ("control" | "shift")[] });
      wc.sendInputEvent({ type: "keyUp", keyCode: k as string, modifiers: m as ("control" | "shift")[] });
    },
    [keyCode, modifiers] as const,
  );
}

/** PDF penceresinin arama çubuğu (ayrı görünüm, data: adresli sayfa). */
async function findBar(app: ElectronApplication): Promise<Page> {
  const pick = () => app.windows().find((p) => p.url().startsWith("data:"));
  await expect.poll(() => !!pick(), { timeout: 10_000 }).toBe(true);
  return pick()!;
}

function samplePdf(link?: string): string {
  const file = path.join(mkdtempSync(path.join(tmpdir(), "oh-pdf-")), "örnek belge.pdf");
  // Üç sayfa, her birinde aranacak sözcük (elle yazılmış, sıkıştırılmamış en küçük PDF).
  const pages = [1, 2, 3];
  const objects: string[] = [];
  objects.push("<< /Type /Catalog /Pages 2 0 R >>");
  objects.push(`<< /Type /Pages /Kids [${pages.map((_, i) => `${3 + i * 2} 0 R`).join(" ")}] /Count ${pages.length} >>`);
  for (const n of pages) {
    const stream = `BT /F1 18 Tf 50 760 Td (Sayfa ${n} ARANACAK kelime) Tj ET`;
    const annots = link ? ` /Annots [<< /Type /Annot /Subtype /Link /Rect [50 700 300 720] /A << /S /URI /URI (${link}) >> >>]` : "";
    objects.push(`<< /Type /Page /Parent 2 0 R /MediaBox [0 0 595 842] /Resources << /Font << /F1 ${3 + pages.length * 2} 0 R >> >> /Contents ${4 + (n - 1) * 2} 0 R${annots} >>`);
    objects.push(`<< /Length ${stream.length} >>\nstream\n${stream}\nendstream`);
  }
  objects.push("<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>");
  let body = "%PDF-1.4\n";
  const offsets: number[] = [];
  objects.forEach((o, i) => {
    offsets.push(Buffer.byteLength(body, "latin1"));
    body += `${i + 1} 0 obj\n${o}\nendobj\n`;
  });
  const xref = Buffer.byteLength(body, "latin1");
  body += `xref\n0 ${objects.length + 1}\n0000000000 65535 f \n${offsets.map((x) => `${String(x).padStart(10, "0")} 00000 n \n`).join("")}`;
  body += `trailer\n<< /Size ${objects.length + 1} /Root 1 0 R >>\nstartxref\n${xref}\n%%EOF\n`;
  writeFileSync(file, body, "latin1");
  return file;
}

test.describe("PDF görüntüleme", () => {
  test("komut satırından (Birlikte aç) .pdf ayrı, güvenli pencerede açılır; arama ve yazdırma çalışır", async () => {
    const o = ortam();
    const pdf = samplePdf();
    const app = await launch([pdf], o);
    const win = await app.firstWindow();
    await waitForViewer(win);
    const [info] = await pdfWindowInfo(app);
    expect(info.url).toMatch(/^file:\/\/.*\.pdf$/);
    expect(info.title).toContain("örnek belge.pdf");
    // Node yok, korumalı alan ve bağlam yalıtımı açık, ön yükleme betiği yok.
    expect(info).toMatchObject({ node: false, sandbox: true, isolation: true, preload: null });
    // Ayrı oturum: yazım denetimi (sözlük indirmesi) kapalı, ana oturumdaki belge deposundan yalıtılmış.
    expect(
      await app.evaluate(({ BrowserWindow, session }) => {
        const ses = BrowserWindow.getAllWindows()[0].webContents.session;
        return { ayri: ses !== session.defaultSession, kalici: ses.isPersistent(), yazim: ses.isSpellCheckerEnabled(), diller: ses.getSpellCheckerLanguages() };
      }),
    ).toEqual({ ayri: true, kalici: false, yazim: false, diller: [] });

    // Yeni pencere reddedilir; gezinme isteği görülür ve engellenir, pencere PDF'te kalır.
    expect(await win.evaluate(() => window.open("file:///etc/hostname") === null)).toBe(true);
    await app.evaluate(({ BrowserWindow }) => {
      const g = globalThis as { __gezinme?: string[] };
      g.__gezinme = [];
      BrowserWindow.getAllWindows()[0].webContents.on("will-navigate", (_e, url) => g.__gezinme!.push(url));
    });
    await win.evaluate(() => {
      location.href = "file:///etc/hostname";
    });
    await expect.poll(() => app.evaluate(() => (globalThis as { __gezinme?: string[] }).__gezinme)).toEqual(["file:///etc/hostname"]);
    expect(await win.evaluate(() => location.href)).toBe(info.url);
    expect(await app.evaluate(({ BrowserWindow }) => BrowserWindow.getAllWindows().length)).toBe(1);
    expect((await pdfWindowInfo(app))[0]?.url).toBe(info.url);

    // Ctrl+F: arama çubuğu; üç sayfadaki eşleşmeler sayılır.
    await pdfKey(app, "F", ["control"]);
    const find = await findBar(app);
    const input = find.getByRole("searchbox", { name: "PDF'te bul" });
    await expect(input).toBeFocused();
    await input.fill("ARANACAK");
    await expect(find.locator("output")).toHaveText("1/3");
    await input.press("Enter");
    await expect(find.locator("output")).toHaveText("2/3");
    await input.press("Shift+Enter");
    await expect(find.locator("output")).toHaveText("1/3");
    await input.fill("YOK-BOYLE-SOZ");
    await expect(find.locator("output")).toHaveText("Bulunamadı");
    await input.press("Escape");
    await expect.poll(() => app.evaluate(({ BrowserWindow }) => BrowserWindow.getAllWindows()[0].contentView.children.some((v) => v.getVisible()))).toBe(false);
    // Yeniden Ctrl+F: aynı çubuk açılır.
    await pdfKey(app, "F", ["control"]);
    await expect.poll(() => app.evaluate(({ BrowserWindow }) => BrowserWindow.getAllWindows()[0].contentView.children.map((v) => v.getVisible()))).toEqual([true]);

    // Ctrl+P yazdırmayı başlatır (sistem iletişim kutusu yerine sayaç).
    await app.evaluate(({ BrowserWindow }) => {
      const wc = BrowserWindow.getAllWindows()[0].webContents;
      (globalThis as { __ohPrint?: number }).__ohPrint = 0;
      wc.print = ((_o: unknown, cb?: (ok: boolean, reason: string) => void) => {
        (globalThis as { __ohPrint?: number }).__ohPrint! += 1;
        cb?.(true, "");
      }) as typeof wc.print;
    });
    await pdfKey(app, "P", ["control"]);
    await expect.poll(() => app.evaluate(() => (globalThis as { __ohPrint?: number }).__ohPrint)).toBe(1);

    // Yazdırma hattı PDF'i sayfalarıyla işler (yazıcısız ortamda PDF'e yazdırma).
    const pages = await app.evaluate(async ({ BrowserWindow }) => {
      const out = await BrowserWindow.getAllWindows()[0].webContents.printToPDF({});
      return (out.toString("latin1").match(/\/Type\s*\/Page[^s]/g) ?? []).length;
    });
    expect(pages).toBe(3);
    await app.close();
  });

  test("Başlangıç › Aç › PDF aç: seçilen PDF görüntüleyicide açılır", async () => {
    const o = ortam();
    const pdf = samplePdf();
    const app = await launch([], o);
    const win = await app.firstWindow();
    await win.waitForLoadState("domcontentloaded");
    // Sistem dosya seçicisi yerine seçilmiş dosya.
    await app.evaluate(({ dialog }, file) => {
      dialog.showOpenDialog = (async () => ({ canceled: false, filePaths: [file] })) as unknown as typeof dialog.showOpenDialog;
    }, pdf);
    await win.locator(".oh-bs-item", { hasText: "Aç" }).first().click();
    await win.getByRole("button", { name: /PDF aç/ }).click();
    const viewer = () => app.windows().find((p) => /\.pdf$/i.test(p.url()));
    await expect.poll(() => !!viewer(), { timeout: 15_000 }).toBe(true);
    expect(await pdfWindowInfo(app)).toHaveLength(1);
    await waitForViewer(viewer()!);
    await app.close();
  });

  test("PDF penceresi dışarıya istek göndermez (PDF'teki web bağlantısı dahil)", async () => {
    const o = ortam();
    const pdf = samplePdf("https://ornek.invalid/izleme.png");
    const netLog = path.join(mkdtempSync(path.join(tmpdir(), "oh-pdf-ag-")), "net.json");
    const app = await electron.launch({
      executablePath: electronPath(),
      args: [desktop, "--no-sandbox", `--log-net-log=${netLog}`, "--net-log-capture-mode=Default", pdf],
      env: { ...process.env, OH_APPS_DIR: appsDir, OH_USER_DATA: o.userData, ELECTRON_DISABLE_SECURITY_WARNINGS: "1", ...o.env },
    });
    const win = await app.firstWindow();
    await waitForViewer(win);
    await pdfKey(app, "F", ["control"]);
    const find = await findBar(app);
    await find.getByRole("searchbox", { name: "PDF'te bul" }).fill("ARANACAK");
    await expect(find.locator("output")).toHaveText("1/3");
    // Ağ günlüğü süreç çıkarken tamamlanır.
    const exited = new Promise((r) => app.process().once("exit", r));
    await app.evaluate(({ app: electronApp, BrowserWindow }) => {
      for (const w of BrowserWindow.getAllWindows()) w.destroy();
      electronApp.quit();
    });
    await exited;
    const log = JSON.parse(readFileSync(netLog, "utf8")) as { constants: { logEventTypes: Record<string, number> }; events: { type: number; params?: Record<string, unknown> }[] };
    const t = log.constants.logEventTypes;
    const hosts = log.events
      .filter((e) => e.type === t.URL_REQUEST_START_JOB || e.type === t.HOST_RESOLVER_MANAGER_REQUEST || e.type === t.SOCKET_POOL_CONNECT_JOB_CREATED)
      .map((e) => String(e.params?.url ?? e.params?.host ?? e.params?.group_id ?? ""))
      .filter((u) => /^(?:pm\/)?(https?|wss?|ftp):\/\//i.test(u) && !/:\/\/(127\.|localhost|\[?::1)/i.test(u));
    expect(hosts).toEqual([]);
  });

  test(".pdf için dosya ilişkilendirmesi yok (electron-builder.yml)", () => {
    const yml = readFileSync(path.join(desktop, "electron-builder.yml"), "utf8");
    const exts = [...yml.matchAll(/^\s*-\s*ext:\s*(\S+)/gm)].map((m) => m[1].toLowerCase());
    expect(exts.length).toBeGreaterThan(5);
    expect(exts).not.toContain("pdf");
    expect(yml).not.toMatch(/application\/pdf/i);
  });
});
