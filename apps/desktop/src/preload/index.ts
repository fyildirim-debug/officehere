import { contextBridge, ipcRenderer } from "electron";

// Sayfaya açılan köprü; tür tanımı packages/core/src/bridge.ts içindeki OfficehereNative'dir.

let closeHandler: (() => boolean | Promise<boolean>) | null = null;

ipcRenderer.on("oh:before-close", async () => {
  let ok = true;
  try {
    ok = closeHandler ? await closeHandler() : true;
  } catch {
    ok = true;
  }
  ipcRenderer.send("oh:close-response", ok);
});

const pathListeners = new Set<(p: string) => void>();
ipcRenderer.on("oh:open-path", (_e, p: string) => pathListeners.forEach((cb) => cb(p)));

const updateListeners = new Set<(s: unknown) => void>();
ipcRenderer.on("oh:update-state", (_e, s: unknown) => updateListeners.forEach((cb) => cb(s)));

contextBridge.exposeInMainWorld("officehereNative", {
  platform: process.platform,
  openFile: (filters: unknown) => ipcRenderer.invoke("oh:open-file", filters),
  readFile: (path: string) => ipcRenderer.invoke("oh:read-file", path),
  saveFileAs: (suggestedName: string, filters: unknown, data: ArrayBuffer) => ipcRenderer.invoke("oh:save-file-as", suggestedName, filters, data),
  writeFile: (path: string, data: ArrayBuffer) => ipcRenderer.invoke("oh:write-file", path, data),
  getRecent: () => ipcRenderer.invoke("oh:get-recent"),
  addRecent: (file: unknown) => ipcRenderer.invoke("oh:add-recent", file),
  newWindow: (app: string, path?: string) => ipcRenderer.invoke("oh:new-window", app, path),
  setDocumentState: (state: { title: string; dirty: boolean }) => ipcRenderer.send("oh:set-document-state", state),
  onOpenPath: (cb: (p: string) => void) => {
    pathListeners.add(cb);
    return () => pathListeners.delete(cb);
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
    onChange: (cb: (s: unknown) => void) => {
      updateListeners.add(cb);
      return () => updateListeners.delete(cb);
    },
  },
});

// Yerel pencere düğmeleri için başlık çubuğunda yer aç.
window.addEventListener("DOMContentLoaded", () => {
  const root = document.documentElement;
  root.classList.add(`oh-platform-${process.platform}`);
  if (process.platform !== "darwin") root.classList.add("oh-native-frame");
});
