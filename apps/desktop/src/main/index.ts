import { app, BrowserWindow, dialog, ipcMain, Menu, shell, type FileFilter, type IpcMainInvokeEvent } from "electron";
import fs from "node:fs/promises";
import { existsSync } from "node:fs";
import path from "node:path";
import { APP_COLORS, APP_TITLES, DEV_PORTS, appForPath, isSupportedFile, type AppKind, type WindowKind } from "./apps";
import { addRecent, pinRecent, readRecent, recentStatus, removeRecent, type RecentFile } from "./recent";
import { restorableSession, writeSession, type SessionDoc } from "./session";
import { desktopSettings, setDesktopSetting } from "./settings";
import { initUpdates } from "./updates";
import { planFor, rememberBounds } from "./windowState";

const isDev = process.env.OH_DEV === "1";
const isMac = process.platform === "darwin";

interface WindowState {
  kind: WindowKind;
  dirty: boolean;
  title: string;
  /** Sayfa kapatmadan önce sorma işleyicisi kaydettiyse. */
  hasCloseHandler: boolean;
  /** Sayfa onay verdi; bir sonraki kapatma engellenmez. */
  allowClose: boolean;
  /** Belgenin diskteki tam yolu (oturum için). */
  path?: string;
}

// Testler için yalıtılmış kullanıcı verisi dizini (son kullanılanlar vb.).
if (process.env.OH_USER_DATA) app.setPath("userData", process.env.OH_USER_DATA);

const states = new Map<number, WindowState>();
/** Uygulama hazır olmadan gelen dosya açma istekleri (macOS open-file). */
const pendingPaths: string[] = [];
/** Uygulama kapanıyor: pencereler tek tek kapanırken oturum kaydı değişmesin. */
let quitting = false;

function resourceDir(...parts: string[]) {
  // Paketlenmiş uygulamada web uygulamaları resources/apps altına kopyalanır.
  return app.isPackaged ? path.join(process.resourcesPath, ...parts) : path.join(__dirname, "..", "..", "..", ...parts);
}

function pageLocation(kind: WindowKind): { url?: string; file?: string } {
  if (isDev) return { url: `http://localhost:${DEV_PORTS[kind]}/` };
  // Testler için: web sürümü dizini (pnpm web → dist-web: kökte başlangıç, alt dizinlerde uygulamalar).
  if (process.env.OH_APPS_DIR) return { file: path.join(process.env.OH_APPS_DIR, ...(kind === "baslangic" ? [] : [kind]), "index.html") };
  if (app.isPackaged) return { file: path.join(process.resourcesPath, "apps", kind, "index.html") };
  // Paketlenmemiş üretim derlemesi (pnpm start): depo içindeki dist dizinleri.
  const file = kind === "baslangic" ? path.join(__dirname, "..", "launcher", "index.html") : resourceDir(kind, "dist", "index.html");
  return { file };
}

function iconPath() {
  const p = app.isPackaged ? path.join(process.resourcesPath, "icon.png") : path.join(__dirname, "..", "..", "build", "icon.png");
  return existsSync(p) ? p : undefined;
}

/** Açık belge pencerelerinin yolları. */
function openDocuments(): SessionDoc[] {
  const docs: SessionDoc[] = [];
  for (const [id, st] of states) {
    if (st.kind === "baslangic" || !st.path || docs.some((d) => d.path === st.path)) continue;
    const w = BrowserWindow.fromId(id);
    if (w && !w.isDestroyed()) docs.push({ path: st.path, app: st.kind });
  }
  return docs;
}

/** Açık belgeleri oturum dosyasına yazar (kapanış sırasında yazılmaz; before-quit bir kez yazar). */
function snapshotSession() {
  if (!quitting) writeSession(openDocuments());
}

/** Aynı türden en son açılan pencere (yeni pencere ondan kaydırılır). */
function lastWindowOf(kind: WindowKind): BrowserWindow | undefined {
  let last: BrowserWindow | undefined;
  for (const [id, st] of states) {
    if (st.kind !== kind) continue;
    const w = BrowserWindow.fromId(id);
    if (w && !w.isDestroyed()) last = w;
  }
  return last;
}

function createWindow(kind: WindowKind, filePath?: string): BrowserWindow {
  const launcher = kind === "baslangic";
  const plan = planFor(kind, { width: launcher ? 1100 : 1360, height: launcher ? 720 : 860 }, lastWindowOf(kind));
  const win = new BrowserWindow({
    width: plan.width,
    height: plan.height,
    ...(plan.x !== undefined && plan.y !== undefined ? { x: plan.x, y: plan.y } : {}),
    minWidth: 640,
    minHeight: 420,
    show: false,
    title: APP_TITLES[kind],
    icon: iconPath(),
    backgroundColor: "#f3f3f3",
    titleBarStyle: "hidden",
    ...(isMac ? { trafficLightPosition: { x: 10, y: 9 } } : { titleBarOverlay: { color: APP_COLORS[kind], symbolColor: "#ffffff", height: 32 } }),
    webPreferences: {
      preload: path.join(__dirname, "..", "preload", "index.js"),
      contextIsolation: true,
      nodeIntegration: false,
      sandbox: true,
      spellcheck: true,
    },
  });
  states.set(win.id, { kind, dirty: false, title: APP_TITLES[kind], hasCloseHandler: false, allowClose: false, path: filePath });
  rememberBounds(kind, win);
  if (filePath) snapshotSession();
  win.webContents.session.setSpellCheckerLanguages(["tr-TR", "en-US"]);

  const loc = pageLocation(kind);
  const query = filePath ? { file: filePath } : undefined;
  if (loc.url) {
    const url = new URL(loc.url);
    if (filePath) url.searchParams.set("file", filePath);
    void win.loadURL(url.toString());
  } else if (loc.file) {
    void win.loadFile(loc.file, query ? { query } : undefined);
  }

  win.once("ready-to-show", () => {
    if (plan.maximized) win.maximize();
    win.show();
  });

  win.on("close", (e) => {
    const st = states.get(win.id);
    if (!st || st.allowClose || !st.hasCloseHandler || !st.dirty) return;
    e.preventDefault();
    // Engellenen kapatma uygulamadan çıkışı da iptal eder.
    quitting = false;
    win.webContents.send("oh:before-close");
  });
  win.on("closed", () => {
    states.delete(win.id);
    // Son pencere kapanınca uygulama da kapanır; oturumda o an açık olan belgeler kalsın.
    if (states.size > 0) snapshotSession();
  });

  // Bağlantılar sistem tarayıcısında açılır; uygulama içinde gezinme engellenir.
  win.webContents.setWindowOpenHandler(({ url }) => {
    // Paketle gelen üçüncü taraf lisans bildirimi uygulama içinde, salt metin olarak açılır.
    if (url.startsWith("file:") && url.endsWith("/THIRD_PARTY_NOTICES.txt")) {
      return {
        action: "allow",
        overrideBrowserWindowOptions: {
          width: 820,
          height: 720,
          autoHideMenuBar: true,
          title: "Üçüncü taraf yazılım bildirimleri",
          webPreferences: { contextIsolation: true, nodeIntegration: false, sandbox: true },
        },
      };
    }
    if (/^https?:|^mailto:/.test(url)) void shell.openExternal(url);
    return { action: "deny" };
  });
  win.webContents.on("will-navigate", (e, url) => {
    const current = win.webContents.getURL();
    if (url.split("?")[0] !== current.split("?")[0]) {
      e.preventDefault();
      if (/^https?:|^mailto:/.test(url)) void shell.openExternal(url);
    }
  });
  return win;
}

/** Bir dosyayı uygun uygulamanın yeni penceresinde açar. */
function openPath(filePath: string) {
  const kind = appForPath(filePath);
  if (!kind) return;
  createWindow(kind, filePath);
}

/** Ayar açıksa son oturumdaki, hâlâ var olan belgeleri açar; açılan yoksa false döner. */
function restoreSession(): boolean {
  if (!desktopSettings().oturumuGeriYukle) return false;
  const docs = restorableSession();
  for (const d of docs) createWindow(d.app, d.path);
  return docs.length > 0;
}

function filesFromArgv(argv: string[]): string[] {
  return argv.slice(app.isPackaged ? 1 : 2).filter((a) => !a.startsWith("-") && isSupportedFile(a) && existsSync(a)).map((a) => path.resolve(a));
}

function senderState(e: IpcMainInvokeEvent | Electron.IpcMainEvent) {
  const win = BrowserWindow.fromWebContents(e.sender);
  return { win, state: win ? states.get(win.id) : undefined };
}

function toArrayBuffer(buf: Buffer): ArrayBuffer {
  return buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength) as ArrayBuffer;
}

async function readAsResult(filePath: string) {
  const data = await fs.readFile(filePath);
  return { path: filePath, name: path.basename(filePath), data: toArrayBuffer(data) };
}

function registerIpc() {
  ipcMain.handle("oh:open-file", async (e, filters: FileFilter[]) => {
    const { win } = senderState(e);
    const opts = { properties: ["openFile" as const], filters: [...filters, { name: "Tüm Dosyalar", extensions: ["*"] }] };
    const r = win ? await dialog.showOpenDialog(win, opts) : await dialog.showOpenDialog(opts);
    if (r.canceled || !r.filePaths[0]) return null;
    return readAsResult(r.filePaths[0]);
  });

  ipcMain.handle("oh:read-file", async (_e, filePath: string) => {
    try {
      return await readAsResult(filePath);
    } catch {
      return null;
    }
  });

  ipcMain.handle("oh:save-file-as", async (e, suggestedName: string, filters: FileFilter[], data: ArrayBuffer) => {
    const { win } = senderState(e);
    const opts = { defaultPath: path.join(app.getPath("documents"), suggestedName), filters };
    const r = win ? await dialog.showSaveDialog(win, opts) : await dialog.showSaveDialog(opts);
    if (r.canceled || !r.filePath) return null;
    await fs.writeFile(r.filePath, Buffer.from(data));
    return { path: r.filePath, name: path.basename(r.filePath) };
  });

  ipcMain.handle("oh:write-file", async (_e, filePath: string, data: ArrayBuffer) => {
    try {
      // Önce geçici dosyaya yaz, sonra yerine taşı: yarım kalan kayıt belgeyi bozmasın.
      const tmp = `${filePath}.oh-${process.pid}.tmp`;
      await fs.writeFile(tmp, Buffer.from(data));
      await fs.rename(tmp, filePath);
      return true;
    } catch {
      return false;
    }
  });

  ipcMain.handle("oh:get-recent", () => readRecent());
  ipcMain.handle("oh:add-recent", (_e, file: RecentFile) => addRecent(file));
  ipcMain.handle("oh:pin-recent", (_e, filePath: string, pinned: boolean) => pinRecent(filePath, pinned));
  ipcMain.handle("oh:remove-recent", (_e, filePath: string) => removeRecent(filePath));
  ipcMain.handle("oh:recent-status", (_e, paths: unknown) => recentStatus(paths));
  ipcMain.handle("oh:show-in-folder", (_e, filePath: string) => {
    if (typeof filePath === "string" && path.isAbsolute(filePath)) shell.showItemInFolder(filePath);
  });

  ipcMain.handle("oh:settings-get", () => desktopSettings());
  ipcMain.handle("oh:settings-set", (_e, key: unknown, value: unknown) => setDesktopSetting(key, value));

  ipcMain.handle("oh:new-window", (_e, kind: AppKind, filePath?: string) => {
    createWindow(kind, filePath);
  });

  ipcMain.on("oh:set-document-state", (e, s: { title: string; dirty: boolean; path?: string }) => {
    const { win, state } = senderState(e);
    if (!win || !state) return;
    state.dirty = s.dirty;
    state.title = s.title;
    if (typeof s.path === "string" && path.isAbsolute(s.path) && s.path !== state.path) {
      state.path = s.path;
      snapshotSession();
    }
    win.setTitle(`${s.dirty ? "● " : ""}${s.title} - ${APP_TITLES[state.kind]}`);
    if (isMac) win.setDocumentEdited(s.dirty);
  });

  ipcMain.on("oh:register-close-handler", (e, on: boolean) => {
    const { state } = senderState(e);
    if (state) state.hasCloseHandler = on;
  });

  ipcMain.on("oh:close-response", (e, ok: boolean) => {
    const { win, state } = senderState(e);
    if (!win || !state || !ok) return;
    state.allowClose = true;
    win.close();
  });

  ipcMain.on("oh:set-titlebar-overlay", (e, color: string, symbolColor: string) => {
    const { win } = senderState(e);
    if (!win || isMac) return;
    try {
      win.setTitleBarOverlay({ color, symbolColor, height: 32 });
    } catch {
      // Bazı Linux pencere yöneticileri katmanı desteklemez.
    }
  });

  ipcMain.handle("oh:export-pdf", async (e, suggestedName: string, options?: { landscape?: boolean; pageSize?: Electron.PrintToPDFOptions["pageSize"] }) => {
    const { win } = senderState(e);
    if (!win) return null;
    const r = await dialog.showSaveDialog(win, {
      defaultPath: path.join(app.getPath("documents"), suggestedName),
      filters: [{ name: "PDF", extensions: ["pdf"] }],
    });
    if (r.canceled || !r.filePath) return null;
    const pdf = await win.webContents.printToPDF({
      landscape: options?.landscape ?? false,
      pageSize: options?.pageSize ?? "A4",
      printBackground: true,
      preferCSSPageSize: true,
      margins: { top: 0, bottom: 0, left: 0, right: 0 },
    });
    await fs.writeFile(r.filePath, pdf);
    return { path: r.filePath, name: path.basename(r.filePath) };
  });
}

function buildMenu() {
  if (!isMac) {
    Menu.setApplicationMenu(null);
    return;
  }
  // macOS'ta kopyala/yapıştır kısayolları uygulama menüsü üzerinden çalışır.
  Menu.setApplicationMenu(
    Menu.buildFromTemplate([
      { role: "appMenu" },
      {
        label: "Dosya",
        submenu: [
          { label: "Yeni Belge", click: () => createWindow("belge") },
          { label: "Yeni Çalışma Kitabı", click: () => createWindow("tablo") },
          { label: "Yeni Sunu", click: () => createWindow("sunum") },
          { type: "separator" },
          { label: "Başlangıç Ekranı", click: () => createWindow("baslangic") },
          { type: "separator" },
          { role: "close", label: "Pencereyi Kapat" },
        ],
      },
      { role: "editMenu", label: "Düzen" },
      { role: "windowMenu", label: "Pencere" },
    ]),
  );
}

if (!app.requestSingleInstanceLock()) {
  app.quit();
} else {
  app.on("second-instance", (_e, argv) => {
    const files = filesFromArgv(argv);
    if (files.length) files.forEach(openPath);
    else createWindow("baslangic");
  });

  app.on("open-file", (e, filePath) => {
    e.preventDefault();
    if (app.isReady()) openPath(filePath);
    else pendingPaths.push(filePath);
  });

  app.whenReady().then(() => {
    if (process.platform === "win32") app.setAppUserModelId("com.officehere.suite");
    buildMenu();
    registerIpc();
    initUpdates();
    const files = [...filesFromArgv(process.argv), ...pendingPaths];
    if (files.length) files.forEach(openPath);
    else if (!restoreSession()) createWindow("baslangic");
    app.on("activate", () => {
      if (BrowserWindow.getAllWindows().length === 0) createWindow("baslangic");
    });
  });

  app.on("before-quit", () => {
    // Pencereler açıkken çıkılıyorsa oturum şimdi yazılır; hepsi kapandıktan sonraysa son kayıt geçerlidir.
    if (states.size > 0) writeSession(openDocuments());
    quitting = true;
  });

  app.on("window-all-closed", () => {
    if (!isMac) app.quit();
  });
}
