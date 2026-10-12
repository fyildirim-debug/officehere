import { app, BrowserWindow, dialog, ipcMain, Menu, shell, type FileFilter, type IpcMainInvokeEvent } from "electron";
import fs from "node:fs/promises";
import { existsSync } from "node:fs";
import path from "node:path";
import { APP_COLORS, APP_TITLES, DEV_PORTS, appForPath, isSupportedFile, type AppKind, type WindowKind } from "./apps";
import { addRecent, readRecent, type RecentFile } from "./recent";

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
}

// Testler için yalıtılmış kullanıcı verisi dizini (son kullanılanlar vb.).
if (process.env.OH_USER_DATA) app.setPath("userData", process.env.OH_USER_DATA);

const states = new Map<number, WindowState>();
/** Uygulama hazır olmadan gelen dosya açma istekleri (macOS open-file). */
const pendingPaths: string[] = [];

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

function createWindow(kind: WindowKind, filePath?: string): BrowserWindow {
  const launcher = kind === "baslangic";
  const win = new BrowserWindow({
    width: launcher ? 1100 : 1360,
    height: launcher ? 720 : 860,
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
  states.set(win.id, { kind, dirty: false, title: APP_TITLES[kind], hasCloseHandler: false, allowClose: false });
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

  win.once("ready-to-show", () => win.show());

  win.on("close", (e) => {
    const st = states.get(win.id);
    if (!st || st.allowClose || !st.hasCloseHandler || !st.dirty) return;
    e.preventDefault();
    win.webContents.send("oh:before-close");
  });
  win.on("closed", () => states.delete(win.id));

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

  ipcMain.handle("oh:new-window", (_e, kind: AppKind, filePath?: string) => {
    createWindow(kind, filePath);
  });

  ipcMain.on("oh:set-document-state", (e, s: { title: string; dirty: boolean }) => {
    const { win, state } = senderState(e);
    if (!win || !state) return;
    state.dirty = s.dirty;
    state.title = s.title;
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
    const files = [...filesFromArgv(process.argv), ...pendingPaths];
    if (files.length) files.forEach(openPath);
    else createWindow("baslangic");
    app.on("activate", () => {
      if (BrowserWindow.getAllWindows().length === 0) createWindow("baslangic");
    });
  });

  app.on("window-all-closed", () => {
    if (!isMac) app.quit();
  });
}
