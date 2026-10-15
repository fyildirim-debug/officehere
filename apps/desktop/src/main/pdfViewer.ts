// PDF görüntüleme: Chromium'un yerleşik PDF görüntüleyicisiyle ayrı pencere. Düzenleme yok; görüntüleyicinin
// araç çubuğu yakınlaştırma, döndürme, küçük resimler ve yazdırmayı sunar. Electron'da tarayıcının "Bul"
// çubuğu olmadığından Ctrl+F küçük bir arama çubuğu (ayrı görünüm) açar; Ctrl+P yazdırır.
// Güvenlik: PDF sayfasında ön yükleme betiği ve Node yok, korumalı alan açık; ayrı ve kalıcı olmayan oturum
// (uygulama pencerelerinin deposuna erişmez); yalnız yerel (file://) PDF yüklenir; gezinme ve yeni pencere
// engellenir (http/https/mailto bağlantıları sistem tarayıcısında açılır); oturumdan ağ isteği çıkmaz.
// .pdf için dosya ilişkilendirmesi yoktur: "Birlikte aç" ve başlangıç ekranındaki "PDF aç" buraya gelir.
import { BrowserWindow, dialog, ipcMain, session, shell, WebContentsView, type Input, type Session, type WebContents } from "electron";
import fs from "node:fs/promises";
import path from "node:path";
import { pathToFileURL } from "node:url";
import { queueKey } from "./safeWrite";
import { NO_DOWNLOAD_URL } from "./spellcheckCore";

const PARTITION = "oh-pdf";
const FIND_W = 380;
const FIND_H = 44;
/** Arama çubuğu görüntüleyicinin araç çubuğunun (56 px) altında durur. */
const FIND_TOP = 64;

/** Uzantıya göre PDF mi. */
export function isPdfPath(p: string): boolean {
  return /\.pdf$/i.test(p);
}

/** PDF imzası ("%PDF-") belirtimin izin verdiği gibi ilk 1024 baytta mı. */
export function hasPdfHeader(head: Uint8Array): boolean {
  const sig = [0x25, 0x50, 0x44, 0x46, 0x2d];
  const end = Math.min(head.length, 1024) - sig.length;
  for (let i = 0; i <= end; i++) if (sig.every((b, j) => head[i + j] === b)) return true;
  return false;
}

async function readHead(file: string): Promise<Uint8Array | null> {
  let fh: fs.FileHandle | undefined;
  try {
    fh = await fs.open(file, "r");
    const buf = new Uint8Array(1024);
    const { bytesRead } = await fh.read(buf, 0, buf.length, 0);
    return buf.subarray(0, bytesRead);
  } catch {
    return null;
  } finally {
    await fh?.close().catch(() => undefined);
  }
}

export interface PdfViewerOptions {
  icon?: string;
  /** Günlüğe uyarı. */
  warn?: (message: string) => void;
}

let options: PdfViewerOptions = {};
let configured: Session | null = null;
const windows = new Map<string, BrowserWindow>();

/** Arama çubuğu görünümü (webContents kimliğiyle) → PDF sayfası. */
const finders = new Map<number, { pdf: WebContents; hide: () => void; query: string }>();

const FIND_HTML = `<!doctype html>
<html lang="tr"><head><meta charset="utf-8">
<meta http-equiv="Content-Security-Policy" content="default-src 'none'; style-src 'unsafe-inline'">
<title>PDF'te bul</title>
<style>
:root{color-scheme:light dark;font:13px system-ui,"Segoe UI",sans-serif}
html,body{margin:0;height:100%;overflow:hidden;background:transparent}
form{display:flex;align-items:center;gap:4px;height:100%;box-sizing:border-box;padding:6px 6px 6px 8px;background:Canvas;color:CanvasText;border:1px solid color-mix(in srgb,CanvasText 22%,transparent);border-radius:6px}
input{flex:1;min-width:0;height:28px;box-sizing:border-box;padding:0 8px;border:1px solid color-mix(in srgb,CanvasText 30%,transparent);border-radius:4px;background:Field;color:FieldText;font:inherit}
input:focus{outline:2px solid Highlight;outline-offset:-1px}
output{min-width:72px;text-align:right;color:GrayText;font-variant-numeric:tabular-nums;white-space:nowrap}
button{width:28px;height:28px;border:0;border-radius:4px;background:transparent;color:inherit;font:15px/1 system-ui,sans-serif;cursor:pointer}
button:hover{background:color-mix(in srgb,CanvasText 12%,transparent)}
button:focus-visible{outline:2px solid Highlight}
</style></head>
<body><form role="search" aria-label="PDF'te bul">
<input type="search" aria-label="PDF'te bul" placeholder="PDF'te bul" autocomplete="off" spellcheck="false">
<output aria-live="polite"></output>
<button type="button" id="onceki" aria-label="Önceki (Shift+Enter)" title="Önceki (Shift+Enter)">&#8593;</button>
<button type="button" id="sonraki" aria-label="Sonraki (Enter)" title="Sonraki (Enter)">&#8595;</button>
<button type="button" id="kapat" aria-label="Kapat (Esc)" title="Kapat (Esc)">&#10005;</button>
</form></body></html>`;

/**
 * PDF oturumu: izin istekleri reddedilir (tam ekran sunum dışında), ağ istekleri kesilir; yazım denetimi
 * kapalıdır (Chromium bu oturum için de sözlük indirmeye kalkmasın; onay yalnız ana oturumda sorulur).
 */
function pdfSession(): Session {
  if (configured) return configured;
  const ses = session.fromPartition(PARTITION);
  ses.setSpellCheckerEnabled(false);
  if (process.platform !== "darwin") {
    ses.setSpellCheckerLanguages([]);
    ses.setSpellCheckerDictionaryDownloadURL(NO_DOWNLOAD_URL);
  }
  ses.setPermissionRequestHandler((_wc, permission, cb) => cb(permission === "fullscreen"));
  ses.webRequest.onBeforeRequest({ urls: ["http://*/*", "https://*/*", "ws://*/*", "wss://*/*", "ftp://*/*"] }, (_d, cb) => cb({ cancel: true }));
  configured = ses;
  return ses;
}

function openIfWeb(url: string) {
  if (/^(https?:|mailto:)/i.test(url)) void shell.openExternal(url);
}

/** Gezinme ve yeni pencere: yalnız izinli adres; web bağlantıları sistem tarayıcısına. */
function lockNavigation(wc: WebContents, allowed: string) {
  wc.setWindowOpenHandler(({ url }) => {
    openIfWeb(url);
    return { action: "deny" };
  });
  wc.on("will-navigate", (e, url) => {
    if (url === allowed) return;
    e.preventDefault();
    openIfWeb(url);
  });
  wc.on("will-redirect", (e, url) => {
    if (url !== allowed) e.preventDefault();
  });
}

/** Arama çubuğu (ilk Ctrl+F'de oluşturulur). */
function createFinder(win: BrowserWindow, pdf: WebContents) {
  const view = new WebContentsView({
    webPreferences: { preload: path.join(__dirname, "..", "preload", "pdfBul.js"), sandbox: true, contextIsolation: true, nodeIntegration: false, partition: PARTITION, spellcheck: false },
  });
  const url = `data:text/html;charset=utf-8,${encodeURIComponent(FIND_HTML)}`;
  lockNavigation(view.webContents, url);
  view.setBackgroundColor("#00000000");
  view.setVisible(false);
  win.contentView.addChildView(view);
  void view.webContents.loadURL(url);
  const layout = () => {
    const [w] = win.getContentSize();
    const width = Math.min(FIND_W, Math.max(160, w - 16));
    view.setBounds({ x: Math.max(8, w - width - 16), y: FIND_TOP, width, height: FIND_H });
  };
  layout();
  win.on("resize", layout);
  const id = view.webContents.id;
  const entry = {
    pdf,
    query: "",
    hide: () => {
      view.setVisible(false);
      if (!pdf.isDestroyed()) {
        pdf.stopFindInPage("keepSelection");
        pdf.focus();
      }
    },
  };
  finders.set(id, entry);
  const onResult = (_e: unknown, r: Electron.Result) => {
    if (!view.webContents.isDestroyed()) view.webContents.send("oh:pdf-bul-sonuc", { active: r.activeMatchOrdinal, matches: r.matches });
  };
  pdf.on("found-in-page", onResult);
  win.once("closed", () => {
    finders.delete(id);
    if (!view.webContents.isDestroyed()) view.webContents.close();
  });
  const show = () => {
    view.setVisible(true);
    view.webContents.focus();
    const focusInput = () => view.webContents.send("oh:pdf-bul-odak");
    if (view.webContents.isLoading()) view.webContents.once("did-finish-load", focusInput);
    else focusInput();
  };
  const again = (forward: boolean) => {
    if (!entry.query) return show();
    pdf.findInPage(entry.query, { forward, findNext: false });
  };
  return { show, again, hide: entry.hide, visible: () => view.getVisible(), keys: view.webContents };
}

/** PDF penceresinin kısayolları: Ctrl+F ara, F3/Shift+F3 sonraki/önceki, Ctrl+P yazdır, Ctrl+W kapat. */
function handleKeys(win: BrowserWindow, pdf: WebContents) {
  let finder: ReturnType<typeof createFinder> | null = null;
  const getFinder = () => {
    if (!finder) {
      finder = createFinder(win, pdf);
      finder.keys.on("before-input-event", (e, input) => onKey(e, input));
    }
    return finder;
  };
  const onKey = (e: Electron.Event, input: Input) => {
    if (input.type !== "keyDown") return;
    const mod = (input.control || input.meta) && !input.alt;
    if (mod && input.code === "KeyF") getFinder().show();
    else if (mod && input.code === "KeyP") printPdf(pdf);
    else if (mod && input.code === "KeyW") win.close();
    else if (input.key === "F3") getFinder().again(!input.shift);
    else if (input.key === "Escape" && finder?.visible()) finder.hide();
    else return;
    e.preventDefault();
  };
  pdf.on("before-input-event", onKey);
}

function printPdf(pdf: WebContents) {
  pdf.print({ printBackground: true }, (ok, reason) => {
    if (!ok && reason && reason !== "cancelled") options.warn?.(`PDF yazdırılamadı: ${reason}`);
  });
}

/** PDF'i ayrı görüntüleyici penceresinde açar (zaten açıksa öne getirir); açılamazsa nedenini söyler. */
export async function openPdfWindow(filePath: string): Promise<BrowserWindow | null> {
  const file = path.resolve(filePath);
  if (!isPdfPath(file)) return null;
  const key = queueKey(file);
  const open = windows.get(key);
  if (open && !open.isDestroyed()) {
    if (open.isMinimized()) open.restore();
    open.show();
    open.focus();
    return open;
  }
  const head = await readHead(file);
  if (!head || !hasPdfHeader(head)) {
    void dialog
      .showMessageBox({ type: "error", title: "OfficeHere", message: `“${path.basename(file)}” açılamadı`, detail: head ? "Dosya bir PDF belgesi değil ya da bozuk." : "Dosya bulunamadı ya da okunamadı.", buttons: ["Tamam"], noLink: true })
      .catch(() => undefined);
    return null;
  }
  const win = new BrowserWindow({
    width: 1000,
    height: 820,
    minWidth: 480,
    minHeight: 360,
    show: false,
    title: `${path.basename(file)} - OfficeHere PDF Görüntüleyici`,
    icon: options.icon,
    backgroundColor: "#525659",
    autoHideMenuBar: true,
    webPreferences: {
      session: pdfSession(),
      plugins: true,
      sandbox: true,
      contextIsolation: true,
      nodeIntegration: false,
      webviewTag: false,
      navigateOnDragDrop: false,
      spellcheck: false,
      safeDialogs: true,
    },
  });
  windows.set(key, win);
  win.on("closed", () => {
    if (windows.get(key) === win) windows.delete(key);
  });
  // Başlık dosya adıyla kalır (PDF'in kendi başlık bilgisi yanıltıcı olabilir).
  win.on("page-title-updated", (e) => e.preventDefault());
  win.once("ready-to-show", () => win.show());
  const url = pathToFileURL(file).href;
  lockNavigation(win.webContents, url);
  handleKeys(win, win.webContents);
  void win.loadURL(url).catch((err: Error) => options.warn?.(`PDF yüklenemedi: ${err.message}`));
  return win;
}

/** "PDF aç": dosya sorulur, seçilen PDF görüntüleyicide açılır. */
export async function pickAndOpenPdf(parent: BrowserWindow | null): Promise<boolean> {
  const opts: Electron.OpenDialogOptions = { properties: ["openFile"], filters: [{ name: "PDF belgeleri", extensions: ["pdf"] }] };
  const r = parent ? await dialog.showOpenDialog(parent, opts) : await dialog.showOpenDialog(opts);
  if (r.canceled || !r.filePaths[0]) return false;
  return !!(await openPdfWindow(r.filePaths[0]));
}

/** Arama çubuğunun iletileri (bir kez kaydedilir). */
export function initPdfViewer(o: PdfViewerOptions) {
  options = o;
  ipcMain.on("oh:pdf-bul", (e, q: { text?: unknown; forward?: unknown; findNext?: unknown }) => {
    const f = finders.get(e.sender.id);
    if (!f || f.pdf.isDestroyed() || typeof q?.text !== "string") return;
    f.query = q.text.slice(0, 500);
    if (!f.query) {
      f.pdf.stopFindInPage("clearSelection");
      e.sender.send("oh:pdf-bul-sonuc", null);
      return;
    }
    f.pdf.findInPage(f.query, { forward: q.forward !== false, findNext: q.findNext === true });
  });
  ipcMain.on("oh:pdf-bul-kapat", (e) => finders.get(e.sender.id)?.hide());
}
