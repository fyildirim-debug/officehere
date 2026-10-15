import { contextBridge, ipcRenderer } from "electron";

// Sayfaya açılan köprü; tür tanımı packages/core/src/bridge.ts içindeki OfficehereNative'dir.

let closeHandler: (() => boolean | Promise<boolean>) | null = null;

ipcRenderer.on("oh:before-close", async () => {
  // Hemen "aldım": ana süreç yanıt alamazsa sayfayı donmuş sayar ve kullanıcıya sorar.
  ipcRenderer.send("oh:close-ack");
  let ok: boolean | "error" = true;
  try {
    ok = closeHandler ? await closeHandler() : true;
  } catch (err) {
    // Kapatma işleyicisi hata verdi (ör. kaydetme başarısız): pencere sessizce kapanmaz, ana süreç sorar.
    ok = "error";
    ipcRenderer.send("oh:log", { level: "hata", message: `Kapatma işleyicisi: ${(err as Error)?.message ?? String(err)}`, stack: (err as Error)?.stack, kind: "before-close" });
  }
  ipcRenderer.send("oh:close-response", ok);
});

const pathListeners = new Set<(p: string) => void>();
ipcRenderer.on("oh:open-path", (_e, p: string) => pathListeners.forEach((cb) => cb(p)));

const conflictListeners = new Set<(p: string) => void>();
ipcRenderer.on("oh:path-conflict", (_e, p: string) => conflictListeners.forEach((cb) => cb(p)));

const updateListeners = new Set<(s: unknown) => void>();
ipcRenderer.on("oh:update-state", (_e, s: unknown) => updateListeners.forEach((cb) => cb(s)));

const spellcheckListeners = new Set<(s: unknown) => void>();
ipcRenderer.on("oh:spellcheck-state", (_e, s: unknown) => spellcheckListeners.forEach((cb) => cb(s)));

type WriteOutcome = { ok: boolean; code?: string; reason?: string; message?: string };

contextBridge.exposeInMainWorld("officehereNative", {
  platform: process.platform,
  openFile: (filters: unknown) => ipcRenderer.invoke("oh:open-file", filters),
  readFile: (path: string) => ipcRenderer.invoke("oh:read-file", path),
  saveFileAs: (suggestedName: string, filters: unknown, data: ArrayBuffer) => ipcRenderer.invoke("oh:save-file-as", suggestedName, filters, data),
  writeFile: async (path: string, data: ArrayBuffer) => ((await ipcRenderer.invoke("oh:write-file", path, data)) as WriteOutcome)?.ok === true,
  writeFileSafe: (path: string, data: ArrayBuffer, options?: { sidecar?: boolean; keepPrevious?: boolean }) =>
    ipcRenderer.invoke("oh:write-file", path, data, { sidecar: options?.sidecar === true, keepPrevious: options?.keepPrevious === true }),
  revertFile: (path: string) => ipcRenderer.invoke("oh:revert-file", path),
  listVersions: (path: string) => ipcRenderer.invoke("oh:list-versions", path),
  restoreVersion: (path: string, id: string) => ipcRenderer.invoke("oh:restore-version", path, id),
  log: (entry: { level?: string; message?: string; stack?: string; kind?: string }) =>
    ipcRenderer.send("oh:log", { level: entry?.level, message: String(entry?.message ?? "").slice(0, 2000), stack: typeof entry?.stack === "string" ? entry.stack.slice(0, 4000) : undefined, kind: entry?.kind }),
  getLogTail: (lines?: number) => ipcRenderer.invoke("oh:log-tail", lines),
  getAppInfo: () => ipcRenderer.invoke("oh:app-info"),
  getRecent: () => ipcRenderer.invoke("oh:get-recent"),
  addRecent: (file: unknown) => ipcRenderer.invoke("oh:add-recent", file),
  pinRecent: (path: string, pinned: boolean) => ipcRenderer.invoke("oh:pin-recent", path, pinned),
  removeRecent: (path: string) => ipcRenderer.invoke("oh:remove-recent", path),
  showInFolder: (path: string) => ipcRenderer.invoke("oh:show-in-folder", path),
  recentStatus: (paths: string[]) => ipcRenderer.invoke("oh:recent-status", paths),
  getSettings: () => ipcRenderer.invoke("oh:settings-get"),
  setSetting: (key: string, value: unknown) => ipcRenderer.invoke("oh:settings-set", key, value),
  newWindow: (app: string, path?: string) => ipcRenderer.invoke("oh:new-window", app, path),
  setDocumentState: (state: { title: string; dirty: boolean; path?: string }) =>
    ipcRenderer.send("oh:set-document-state", { title: state.title, dirty: state.dirty, ...(typeof state.path === "string" ? { path: state.path } : {}) }),
  onOpenPath: (cb: (p: string) => void) => {
    pathListeners.add(cb);
    return () => pathListeners.delete(cb);
  },
  onPathConflict: (cb: (p: string) => void) => {
    conflictListeners.add(cb);
    return () => conflictListeners.delete(cb);
  },
  onBeforeClose: (cb: () => boolean | Promise<boolean>) => {
    closeHandler = cb;
    ipcRenderer.send("oh:register-close-handler", true);
    return () => {
      if (closeHandler === cb) {
        closeHandler = null;
        ipcRenderer.send("oh:register-close-handler", false);
      }
    };
  },
  setTitleBarOverlay: (color: string, symbolColor: string) => ipcRenderer.send("oh:set-titlebar-overlay", color, symbolColor),
  exportPdf: (suggestedName: string, options?: unknown) => ipcRenderer.invoke("oh:export-pdf", suggestedName, options),
  updates: {
    getState: () => ipcRenderer.invoke("oh:update-get"),
    check: () => ipcRenderer.invoke("oh:update-check"),
    install: () => ipcRenderer.invoke("oh:update-install"),
    setAutoCheck: (on: boolean) => ipcRenderer.invoke("oh:update-auto", on),
    setChannel: (channel: "kararli" | "beta") => ipcRenderer.invoke("oh:update-channel", channel === "beta" ? "beta" : "kararli"),
    onChange: (cb: (s: unknown) => void) => {
      updateListeners.add(cb);
      return () => updateListeners.delete(cb);
    },
  },
  viewer: {
    check: (path: string) => ipcRenderer.invoke("oh:viewer-check", path),
    allowEditing: (path: string) => ipcRenderer.invoke("oh:viewer-allow", path),
  },
  openPdf: (path?: string) => ipcRenderer.invoke("oh:open-pdf", typeof path === "string" ? path : undefined),
  spellcheck: {
    getState: () => ipcRenderer.invoke("oh:spellcheck-get"),
    setDownload: (on: boolean) => ipcRenderer.invoke("oh:spellcheck-set", on === true),
    onChange: (cb: (s: unknown) => void) => {
      spellcheckListeners.add(cb);
      return () => spellcheckListeners.delete(cb);
    },
  },
});

// Yerel pencere düğmeleri için başlık çubuğunda yer aç.
window.addEventListener("DOMContentLoaded", () => {
  const root = document.documentElement;
  root.classList.add(`oh-platform-${process.platform}`);
  if (process.platform !== "darwin") root.classList.add("oh-native-frame");
});
