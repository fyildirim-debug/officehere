import { app, BrowserWindow, dialog, ipcMain, Menu, shell, type FileFilter, type IpcMainInvokeEvent, type MessageBoxOptions } from "electron";
import fs from "node:fs/promises";
import { existsSync } from "node:fs";
import os from "node:os";
import path from "node:path";
import { APP_COLORS, APP_TITLES, DEV_PORTS, appForPath, isSupportedFile, type AppKind, type WindowKind } from "./apps";
import { BackupStore } from "./backups";
import { saveErrorMessage, saveErrorReason } from "./errors";
import { describePath, LocalLog, type LogLevel } from "./log";
import { addRecent, pinRecent, readRecent, recentStatus, removeRecent, type RecentFile } from "./recent";
import { errorCode, pendingWrites, queueKey, hasPendingWrites } from "./safeWrite";
import { SaveService } from "./saveService";
import { restorableSession, writeSession, type SessionDoc } from "./session";
import { desktopSettings, setDesktopSetting } from "./settings";
import { initSpellcheck } from "./spellcheck";
import { initTasks, startWindow } from "./tasks";
import { initUpdates } from "./updates";
import { planFor, rememberBounds } from "./windowState";

const isDev = process.env.OH_DEV === "1";
const isMac = process.platform === "darwin";

/** Kapatma isteğine sayfanın "aldım" yanıtı için beklenen süre; gelmezse sayfa donmuş sayılır. */
const CLOSE_ACK_MS = Number(process.env.OH_CLOSE_ACK_MS ?? 4000);
/**
 * Çöken pencere yeniden yüklenmeden önce beklenen süre: uygulamalar taslağın sahibini
 * "kapanmış" saymak için pencerenin nabzının (≤6 sn) eskimesini bekler.
 */
const RECOVERY_WAIT_MS = Number(process.env.OH_RECOVERY_WAIT_MS ?? 6500);

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
  /** Sayfa yüklemesi sayacı: "önceki sürüm" her yüklemede dosyaya ilk yazmadan önce alınır. */
  loads: number;
  /** Sayfa işlemi çöktü (ya da sonlandırıldı): kapatma engellenmez. */
  crashed?: boolean;
  crashedAt?: number;
  /** Donan pencerede "Yeniden yükle" seçildi: çökme sonrası soru sorulmadan yüklenir. */
  reloadAfterCrash?: boolean;
  /** Kapatma isteği sayacı ve sayfanın son "aldım" yanıtı. */
  closeSeq: number;
  closeAcked: number;
  ackTimer?: NodeJS.Timeout;
  /** Açık "yanıt vermiyor" ya da "çöktü" iletişim kutusu (kapatmak için). */
  prompt?: AbortController;
  /** Hata günlüğü sınırı (dakikada). */
  logWindowStart: number;
  logCount: number;
}

// Testler için yalıtılmış kullanıcı verisi dizini (son kullanılanlar vb.).
if (process.env.OH_USER_DATA) app.setPath("userData", process.env.OH_USER_DATA);

const log = new LocalLog({ dir: path.join(app.getPath("userData"), "gunluk") });
const backups = new BackupStore(path.join(app.getPath("userData"), "yedekler"));
const saver = new SaveService(backups, { warn: (message, err) => log.write("uyari", "kaydet", message, { code: errorCode(err) }) });

const states = new Map<number, WindowState>();
/** Uygulama hazır olmadan gelen dosya açma istekleri (macOS open-file). */
const pendingPaths: string[] = [];
/** Uygulama kapanıyor: pencereler tek tek kapanırken oturum kaydı değişmesin. */
let quitting = false;

/* ------------------------------------------------------------------ */
/* Ana süreç hataları                                                   */
/* ------------------------------------------------------------------ */

let lastMainErrorDialog = 0;

function reportMainError(kind: string, err: unknown) {
  const e = err instanceof Error ? err : new Error(String(err));
  log.write("hata", "ana", `${kind}: ${e.message}`, { stack: e.stack, code: (e as { code?: string }).code });
  // Uygulama tutarlı kaldığı sürece çalışmaya devam eder; kullanıcı en çok dakikada bir uyarılır.
  if (!app.isReady() || Date.now() - lastMainErrorDialog < 60_000) return;
  lastMainErrorDialog = Date.now();
  void dialog
    .showMessageBox({
      type: "warning",
      title: "OfficeHere",
      message: "Beklenmeyen bir hata oluştu",
      detail: "OfficeHere çalışmaya devam ediyor; yine de açık belgelerinizi kaydetmeniz önerilir. Ayrıntılar bilgisayarınızdaki hata günlüğüne yazıldı.",
      buttons: ["Tamam", "Hata bildir…"],
      defaultId: 0,
      cancelId: 0,
      noLink: true,
    })
    .then(({ response }) => {
      if (response === 1) openBugReport();
    })
    .catch(() => undefined);
}

process.on("uncaughtException", (err) => reportMainError("Yakalanmamış hata", err));
process.on("unhandledRejection", (reason) => reportMainError("İşlenmemiş söz reddi", reason));

/* ------------------------------------------------------------------ */
/* Pencereler                                                           */
/* ------------------------------------------------------------------ */

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

/** Dosyayı belge olarak açmış (başka) pencere. */
function windowForPath(filePath: string, exceptId?: number): BrowserWindow | undefined {
  const key = queueKey(filePath);
  for (const [id, st] of states) {
    if (id === exceptId || st.kind === "baslangic" || !st.path || queueKey(st.path) !== key) continue;
    const w = BrowserWindow.fromId(id);
    if (w && !w.isDestroyed()) return w;
  }
  return undefined;
}

function focusWindow(w: BrowserWindow) {
  if (w.isMinimized()) w.restore();
  w.show();
  w.focus();
}

/** Pencereye uygulama sayfasını (isteğe bağlı dosya ve sorgu parametreleriyle) yükler. */
function loadPage(win: BrowserWindow, kind: WindowKind, filePath?: string, extra: Record<string, string> = {}) {
  const st = states.get(win.id);
  if (st) {
    // Yeni sayfa kendi kapatma işleyicisini ve kirli durumunu bildirir.
    st.hasCloseHandler = false;
    st.dirty = false;
    st.allowClose = false;
    st.crashed = false;
    st.loads += 1;
    clearTimeout(st.ackTimer);
  }
  const loc = pageLocation(kind);
  const query: Record<string, string> = { ...extra, ...(filePath ? { file: filePath } : {}) };
  if (loc.url) {
    const url = new URL(loc.url);
    for (const [k, v] of Object.entries(query)) url.searchParams.set(k, v);
    void win.loadURL(url.toString());
  } else if (loc.file) {
    void win.loadFile(loc.file, Object.keys(query).length ? { query } : undefined);
  }
}

function createWindow(kind: WindowKind, filePath?: string, extra?: Record<string, string>): BrowserWindow {
  // Aynı dosya ikinci kez açılmak istenirse var olan pencere öne getirilir.
  if (filePath && kind !== "baslangic") {
    const existing = windowForPath(filePath);
    if (existing) {
      focusWindow(existing);
      return existing;
    }
  }
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
      // Diller ve sözlükler oturum düzeyinde (spellcheck.ts): onay yoksa dil listesi boştur, sözlük indirilmez.
      spellcheck: true,
    },
  });
  states.set(win.id, { kind, dirty: false, title: APP_TITLES[kind], hasCloseHandler: false, allowClose: false, path: filePath, loads: 0, closeSeq: 0, closeAcked: 0, logWindowStart: 0, logCount: 0 });
  rememberBounds(kind, win);
  if (filePath) snapshotSession();
  loadPage(win, kind, filePath, extra);

  win.once("ready-to-show", () => {
    if (plan.maximized) win.maximize();
    win.show();
  });

  win.on("close", (e) => {
    const st = states.get(win.id);
    st?.prompt?.abort();
    if (!st || st.allowClose || st.crashed || !st.hasCloseHandler || !st.dirty) return;
    e.preventDefault();
    // Engellenen kapatma uygulamadan çıkışı da iptal eder.
    quitting = false;
    requestClose(win, st);
  });
  const wcId = win.webContents.id;
  win.on("closed", () => {
    const st = states.get(win.id);
    if (st) clearTimeout(st.ackTimer);
    states.delete(win.id);
    saver.forgetSession(wcId);
    // Son pencere kapanınca uygulama da kapanır; oturumda o an açık olan belgeler kalsın.
    if (states.size > 0) snapshotSession();
  });

  watchRenderer(win);

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

/* ------------------------------------------------------------------ */
/* Kapatma, çökme ve donma                                              */
/* ------------------------------------------------------------------ */

/** Sayfaya "kapatılmak isteniyor" der; sayfa hemen "aldım" demezse donmuş sayılır. */
function requestClose(win: BrowserWindow, st: WindowState) {
  const seq = ++st.closeSeq;
  win.webContents.send("oh:before-close");
  clearTimeout(st.ackTimer);
  st.ackTimer = setTimeout(() => {
    if (!win.isDestroyed() && st.closeAcked < seq && !st.crashed) showHangPrompt(win, st, "close");
  }, CLOSE_ACK_MS);
}

/** "Pencere yanıt vermiyor": Bekle / Yeniden yükle / Kapat. */
function showHangPrompt(win: BrowserWindow, st: WindowState, why: "close" | "unresponsive") {
  if (st.prompt || win.isDestroyed()) return;
  const ac = new AbortController();
  st.prompt = ac;
  log.write("uyari", st.kind, why === "close" ? "Kapatma isteğine yanıt gelmedi (pencere donmuş)" : "Pencere yanıt vermiyor");
  const opts: MessageBoxOptions = {
    type: "warning",
    title: APP_TITLES[st.kind],
    message: "Pencere yanıt vermiyor",
    detail: `${st.title} penceresi şu anda yanıt vermiyor. Bekleyebilir, pencereyi yeniden yükleyip kaydedilmemiş değişiklikleri son kurtarma taslağından geri alabilir ya da pencereyi kapatabilirsiniz.`,
    buttons: ["Bekle", "Yeniden yükle", "Kapat"],
    defaultId: 0,
    cancelId: 0,
    noLink: true,
    signal: ac.signal,
  };
  void dialog
    .showMessageBox(win, opts)
    .then(({ response }) => {
      if (st.prompt === ac) st.prompt = undefined;
      if (ac.signal.aborted || win.isDestroyed()) return;
      if (response === 1) {
        log.write("bilgi", st.kind, "Donan pencere yeniden yükleniyor");
        st.reloadAfterCrash = true;
        win.webContents.forcefullyCrashRenderer();
      } else if (response === 2) {
        log.write("bilgi", st.kind, "Donan pencere kapatıldı");
        st.allowClose = true;
        win.destroy();
      } else if (why === "close" && st.closeAcked < st.closeSeq) {
        // Bekle: sayfa yanıt verince kapatma sorusu kendiliğinden gelir; yeniden denetlenir.
        clearTimeout(st.ackTimer);
        const seq = st.closeSeq;
        st.ackTimer = setTimeout(() => {
          if (!win.isDestroyed() && st.closeAcked < seq && !st.crashed) showHangPrompt(win, st, "close");
        }, CLOSE_ACK_MS * 2);
      }
    })
    .catch(() => {
      if (st.prompt === ac) st.prompt = undefined;
    });
}

/** Çöken sayfayı yeniden yükler; taslağı varsa uygulama açılışta kurtarmayı önerir. */
function reloadAfterCrash(win: BrowserWindow, st: WindowState) {
  if (win.isDestroyed()) return;
  // Kirli belge: dosya yeniden açılmaz, boş pencere taslağı önerir (taslak dosya yolunu taşır).
  const file = st.dirty ? undefined : st.path;
  if (!file) st.path = undefined;
  const wait = Math.max(0, (st.crashedAt ?? 0) + RECOVERY_WAIT_MS - Date.now());
  if (wait > 0) {
    const html = `<!doctype html><meta charset="utf-8"><title>OfficeHere</title><body style="margin:0;height:100vh;display:grid;place-items:center;font:15px system-ui,sans-serif;background:#f3f3f3;color:#333"><div>Pencere yeniden yükleniyor…</div></body>`;
    void win.loadURL(`data:text/html;charset=utf-8,${encodeURIComponent(html)}`).catch(() => undefined);
  }
  setTimeout(() => {
    if (!win.isDestroyed()) loadPage(win, st.kind, file);
  }, wait);
}

function watchRenderer(win: BrowserWindow) {
  win.webContents.on("render-process-gone", (_e, details) => {
    const st = states.get(win.id);
    if (details.reason === "clean-exit" || !st) return;
    log.write("hata", st.kind, `Pencere işlemi sonlandı: ${details.reason} (çıkış kodu ${details.exitCode})`);
    st.crashed = true;
    st.crashedAt = Date.now();
    st.hasCloseHandler = false;
    clearTimeout(st.ackTimer);
    st.prompt?.abort();
    st.prompt = undefined;
    if (win.isDestroyed()) return;
    if (st.reloadAfterCrash) {
      st.reloadAfterCrash = false;
      reloadAfterCrash(win, st);
      return;
    }
    if (quitting) {
      win.destroy();
      return;
    }
    const ac = new AbortController();
    st.prompt = ac;
    void dialog
      .showMessageBox(win, {
        type: "error",
        title: APP_TITLES[st.kind],
        message: "Pencere beklenmedik biçimde kapandı",
        detail: st.dirty
          ? `“${st.title}” kaydedilmemiş değişiklikler içeriyordu. Pencereyi yeniden yükleyip son otomatik kurtarma taslağından devam edebilirsiniz. Sorunu bildirmek için Yardım › Hata bildir.`
          : "Pencereyi yeniden yükleyebilir ya da kapatabilirsiniz. Sorunu bildirmek için Yardım › Hata bildir.",
        buttons: ["Yeniden yükle ve taslaktan kurtar", "Pencereyi kapat"],
        defaultId: 0,
        cancelId: 1,
        noLink: true,
        signal: ac.signal,
      })
      .then(({ response }) => {
        if (st.prompt === ac) st.prompt = undefined;
        if (ac.signal.aborted || win.isDestroyed()) return;
        if (response === 0) reloadAfterCrash(win, st);
        else win.close();
      })
      .catch(() => {
        if (st.prompt === ac) st.prompt = undefined;
      });
  });
  win.on("unresponsive", () => {
    const st = states.get(win.id);
    if (st && !st.crashed) showHangPrompt(win, st, "unresponsive");
  });
  win.on("responsive", () => {
    const st = states.get(win.id);
    st?.prompt?.abort();
  });
  // Sayfa kendisi yeniden yüklendiyse (ör. önceki sürüm geri yüklendi) yeni bir kaydetme oturumu başlar.
  win.webContents.on("did-navigate", () => {
    const st = states.get(win.id);
    if (st) st.loads += 1;
  });
  win.webContents.on("did-fail-load", (_e, code, desc, url, isMainFrame) => {
    if (isMainFrame && code !== -3) log.write("hata", states.get(win.id)?.kind ?? "pencere", `Sayfa yüklenemedi: ${desc} (${code}) ${url.startsWith("data:") ? "" : path.basename(url.split("?")[0])}`);
  });
}

/** Başlangıç ekranında "Hata bildir" penceresini açar. */
function openBugReport() {
  for (const [id, st] of states) {
    if (st.kind !== "baslangic") continue;
    const w = BrowserWindow.fromId(id);
    if (w && !w.isDestroyed()) {
      loadPage(w, "baslangic", undefined, { "hata-bildir": "1" });
      focusWindow(w);
      return;
    }
  }
  createWindow("baslangic", undefined, { "hata-bildir": "1" });
}

/** Bir dosyayı uygun uygulamanın yeni penceresinde açar (zaten açıksa o pencere öne gelir). */
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

function toBytes(data: unknown): Uint8Array | null {
  if (data instanceof ArrayBuffer) return new Uint8Array(data);
  if (ArrayBuffer.isView(data)) return new Uint8Array(data.buffer, data.byteOffset, data.byteLength);
  return null;
}

async function readAsResult(filePath: string) {
  const data = await fs.readFile(filePath);
  return { path: filePath, name: path.basename(filePath), data: toArrayBuffer(data) };
}

/* ------------------------------------------------------------------ */
/* Kaydetme                                                             */
/* ------------------------------------------------------------------ */

type WriteFailure = { ok: false; code: string; reason: string; message: string };
type WriteOutcome = { ok: true } | WriteFailure;

function failure(err: unknown, filePath: string, what: string): WriteFailure {
  const code = errorCode(err);
  log.write("hata", "kaydet", `${what} başarısız: ${describePath(filePath)}`, { code, stack: code === "EUNKNOWN" ? (err as Error)?.stack : undefined });
  return { ok: false, code, reason: saveErrorReason(code), message: saveErrorMessage(code) };
}

/** Kaydetme oturumu: pencere + sayfa yüklemesi (yeniden yüklemede yeni "önceki sürüm" alınır). */
function writeSessionOf(e: IpcMainInvokeEvent) {
  const { state } = senderState(e);
  return `${e.sender.id}:${state?.loads ?? 0}`;
}

function validPath(p: unknown): p is string {
  return typeof p === "string" && p.length > 0 && path.isAbsolute(p);
}

/** Konum soran kaydetme (Farklı Kaydet, PDF) başarısız: kullanıcıya nedeni söylenir. */
function showSaveError(win: BrowserWindow | null, f: WriteFailure, filePath: string) {
  const opts: MessageBoxOptions = { type: "error", title: "OfficeHere", message: `“${path.basename(filePath)}” kaydedilemedi`, detail: f.message, buttons: ["Tamam"], noLink: true };
  void (win ? dialog.showMessageBox(win, opts) : dialog.showMessageBox(opts)).catch(() => undefined);
}

function registerIpc() {
  ipcMain.handle("oh:open-file", async (e, filters: FileFilter[]) => {
    const { win } = senderState(e);
    const opts = { properties: ["openFile" as const], filters: [...filters, { name: "Tüm Dosyalar", extensions: ["*"] }] };
    const r = win ? await dialog.showOpenDialog(win, opts) : await dialog.showOpenDialog(opts);
    if (r.canceled || !r.filePaths[0]) return null;
    return readAsResult(r.filePaths[0]);
  });

  ipcMain.handle("oh:read-file", async (e, filePath: string) => {
    // Pencere kendi türünden bir belgeyi açmak istiyor ama dosya başka pencerede açık: o pencere öne gelir.
    // Pencerenin kendi belgesini okuması (otomatik kaydetme doğrulaması) hiçbir zaman engellenmez.
    const { win, state } = senderState(e);
    const own = !!state?.path && typeof filePath === "string" && queueKey(state.path) === queueKey(filePath);
    if (win && state && !own && typeof filePath === "string" && appForPath(filePath) === state.kind) {
      const other = windowForPath(filePath, win.id);
      if (other) {
        focusWindow(other);
        return null;
      }
    }
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
    const bytes = toBytes(data);
    try {
      if (!bytes) throw Object.assign(new Error("geçersiz veri"), { code: "EINVAL" });
      // Var olan dosyanın üzerine yazma da atomiktir; üzerine yazılan içerik önceki sürüm olarak saklanır.
      await saver.write(r.filePath, bytes, { session: writeSessionOf(e) });
      return { path: r.filePath, name: path.basename(r.filePath) };
    } catch (err) {
      const f = failure(err, r.filePath, "Farklı Kaydet");
      showSaveError(win, f, r.filePath);
      return { error: { code: f.code, reason: f.reason, message: f.message } };
    }
  });

  ipcMain.handle("oh:write-file", async (e, filePath: unknown, data: unknown, options?: { sidecar?: boolean; keepPrevious?: boolean }): Promise<WriteOutcome> => {
    if (!validPath(filePath)) return { ok: false, code: "EINVAL", reason: saveErrorReason("EINVAL"), message: saveErrorMessage("EINVAL") };
    const bytes = toBytes(data);
    if (!bytes) return failure(Object.assign(new Error("geçersiz veri"), { code: "EINVAL" }), filePath, "Kaydet");
    try {
      await saver.write(filePath, bytes, { session: writeSessionOf(e), sidecar: options?.sidecar === true, keepPrevious: options?.keepPrevious === true });
      return { ok: true };
    } catch (err) {
      return failure(err, filePath, options?.sidecar ? "Otomatik kaydetme" : "Kaydet");
    }
  });

  ipcMain.handle("oh:revert-file", async (_e, filePath: unknown): Promise<WriteOutcome> => {
    if (!validPath(filePath)) return { ok: false, code: "EINVAL", reason: saveErrorReason("EINVAL"), message: saveErrorMessage("EINVAL") };
    try {
      await saver.revertPrevious(filePath);
      log.write("uyari", "kaydet", `Doğrulanamayan kayıt geri alındı: ${describePath(filePath)}`);
      return { ok: true };
    } catch (err) {
      return failure(err, filePath, "Önceki içeriği geri yükleme");
    }
  });

  ipcMain.handle("oh:list-versions", async (_e, filePath: unknown) => (validPath(filePath) ? backups.list(filePath) : []));

  ipcMain.handle("oh:restore-version", async (_e, filePath: unknown, id: unknown): Promise<WriteOutcome> => {
    if (!validPath(filePath) || typeof id !== "string") return { ok: false, code: "EINVAL", reason: saveErrorReason("EINVAL"), message: saveErrorMessage("EINVAL") };
    try {
      await saver.restoreVersion(filePath, id);
      log.write("bilgi", "kaydet", `Önceki sürüm geri yüklendi: ${describePath(filePath)}`);
      return { ok: true };
    } catch (err) {
      return failure(err, filePath, "Sürüm geri yükleme");
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
      // Dosya bu pencerede (ör. Aç iletişim kutusuyla) açıldı ama başka pencerede de açık: sayfa uyarılır.
      const other = windowForPath(s.path, win.id);
      state.path = s.path;
      snapshotSession();
      if (other) {
        log.write("uyari", state.kind, "Aynı dosya iki pencerede açıldı");
        win.webContents.send("oh:path-conflict", s.path);
      }
    }
    win.setTitle(`${s.dirty ? "● " : ""}${s.title} - ${APP_TITLES[state.kind]}`);
    if (isMac) win.setDocumentEdited(s.dirty);
  });

  ipcMain.on("oh:register-close-handler", (e, on: boolean) => {
    const { state } = senderState(e);
    if (state) state.hasCloseHandler = on;
  });

  ipcMain.on("oh:close-ack", (e) => {
    const { state } = senderState(e);
    if (!state) return;
    state.closeAcked = state.closeSeq;
    clearTimeout(state.ackTimer);
    state.prompt?.abort();
  });

  ipcMain.on("oh:close-response", (e, ok: boolean | "error") => {
    const { win, state } = senderState(e);
    if (!win || !state) return;
    if (ok === "error") {
      // Sayfanın kapatma işleyicisi hata verdi: pencere kapatılamaz hâle gelmesin, kullanıcıya sorulur.
      log.write("hata", state.kind, "Kapatma işleyicisi hata verdi");
      void dialog
        .showMessageBox(win, {
          type: "warning",
          title: APP_TITLES[state.kind],
          message: "Pencere kapatılırken bir sorun oluştu",
          detail: "Kaydedilmemiş değişiklikler kaybolabilir. Son kurtarma taslağı bir sonraki açılışta önerilir. Yine de kapatılsın mı?",
          buttons: ["İptal", "Kapat"],
          defaultId: 0,
          cancelId: 0,
          noLink: true,
        })
        .then(({ response }) => {
          if (response === 1 && !win.isDestroyed()) {
            state.allowClose = true;
            win.close();
          }
        })
        .catch(() => undefined);
      return;
    }
    if (!ok) return;
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
    try {
      await saver.write(r.filePath, pdf, { session: writeSessionOf(e), snapshot: "never" });
    } catch (err) {
      // Sayfalar null'u "iptal" sayar: neden burada kullanıcıya gösterilir.
      showSaveError(win, failure(err, r.filePath, "PDF dışa aktarma"), r.filePath);
      return null;
    }
    return { path: r.filePath, name: path.basename(r.filePath) };
  });

  /* Hata günlüğü ve "Hata bildir" */
  ipcMain.on("oh:log", (e, entry: { level?: unknown; message?: unknown; stack?: unknown; kind?: unknown }) => {
    const { state } = senderState(e);
    const now = Date.now();
    // Hata fırtınasında günlük şişmesin: pencere başına dakikada en çok 30 kayıt.
    if (state) {
      if (now - state.logWindowStart > 60_000) {
        state.logWindowStart = now;
        state.logCount = 0;
      }
      if (++state.logCount > 30) return;
    }
    const level: LogLevel = entry?.level === "bilgi" || entry?.level === "uyari" ? entry.level : "hata";
    log.write(level, state?.kind ?? "sayfa", typeof entry?.message === "string" ? entry.message : String(entry?.message ?? ""), {
      stack: typeof entry?.stack === "string" ? entry.stack : undefined,
      code: typeof entry?.kind === "string" ? entry.kind : undefined,
    });
  });
  ipcMain.handle("oh:log-tail", (_e, lines?: unknown) => log.tail(typeof lines === "number" ? Math.max(1, Math.min(200, Math.floor(lines))) : 40));
  ipcMain.handle("oh:app-info", () => ({
    version: app.getVersion(),
    platform: process.platform,
    arch: process.arch,
    osVersion: typeof os.version === "function" ? `${os.version()} (${os.release()})` : os.release(),
    electron: process.versions.electron,
    chrome: process.versions.chrome,
  }));
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
      { label: "Yardım", submenu: [{ label: "Hata bildir…", click: () => openBugReport() }] },
    ]),
  );
}

if (!app.requestSingleInstanceLock()) {
  app.quit();
} else {
  app.on("second-instance", (_e, argv) => {
    const files = filesFromArgv(argv);
    if (files.length) files.forEach(openPath);
    else createWindow(startWindow(argv));
  });

  app.on("open-file", (e, filePath) => {
    e.preventDefault();
    if (app.isReady()) openPath(filePath);
    else pendingPaths.push(filePath);
  });

  app.on("child-process-gone", (_e, details) => {
    if (details.reason !== "clean-exit") log.write("uyari", "ana", `Yardımcı işlem sonlandı: ${details.type} ${details.reason}`);
  });

  app.whenReady().then(() => {
    // Oturuma ilk erişim burada olmalı: yazım sözlüğü onayı sorulmadan hiçbir indirme başlamaz.
    initSpellcheck();
    log.write("bilgi", "ana", `OfficeHere ${app.getVersion()} başlatıldı · ${process.platform} ${os.release()} ${process.arch} · Electron ${process.versions.electron}`);
    if (process.platform === "win32") app.setAppUserModelId("com.officehere.suite");
    buildMenu();
    initTasks((kind) => createWindow(kind));
    registerIpc();
    initUpdates();
    void backups.prune().catch(() => 0);
    const files = [...filesFromArgv(process.argv), ...pendingPaths];
    if (files.length) files.forEach(openPath);
    else if (startWindow(process.argv) !== "baslangic" || !restoreSession()) createWindow(startWindow(process.argv));
    app.on("activate", () => {
      if (BrowserWindow.getAllWindows().length === 0) createWindow("baslangic");
    });
  });

  app.on("before-quit", () => {
    // Pencereler açıkken çıkılıyorsa oturum şimdi yazılır; hepsi kapandıktan sonraysa son kayıt geçerlidir.
    if (states.size > 0) writeSession(openDocuments());
    quitting = true;
  });

  // Süren bir yazma varsa (otomatik kaydetme) bitmesi beklenir; en çok 5 sn.
  let writesFlushed = false;
  app.on("will-quit", (e) => {
    if (writesFlushed || !hasPendingWrites()) return;
    e.preventDefault();
    void Promise.race([pendingWrites(), new Promise((r) => setTimeout(r, 5000))]).finally(() => {
      writesFlushed = true;
      app.quit();
    });
  });

  app.on("window-all-closed", () => {
    if (!isMac) app.quit();
  });
}
