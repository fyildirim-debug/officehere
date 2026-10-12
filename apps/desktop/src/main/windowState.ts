// Pencere boyutu, konumu ve büyütülmüş durumu uygulama türüne göre userData/pencere.json'da saklanır.
import { app, screen, type BrowserWindow } from "electron";
import fs from "node:fs";
import path from "node:path";
import type { WindowKind } from "./apps";
import { planWindow, sanitizeSaved, type SavedWindow, type Size, type WindowPlan } from "./windowBounds";

type Store = Partial<Record<WindowKind, SavedWindow>>;

function storePath() {
  return path.join(app.getPath("userData"), "pencere.json");
}

function readStore(): Store {
  try {
    const data = JSON.parse(fs.readFileSync(storePath(), "utf8")) as unknown;
    return data && typeof data === "object" ? (data as Store) : {};
  } catch {
    return {};
  }
}

function writeStore(store: Store) {
  try {
    fs.mkdirSync(path.dirname(storePath()), { recursive: true });
    fs.writeFileSync(storePath(), JSON.stringify(store, null, 2));
  } catch {
    // Yazılamazsa pencere bir sonraki açılışta varsayılan boyutta açılır.
  }
}

/** Yeni pencerenin yerleşimi: aynı türden açık pencere varsa ondan kaydırılır, yoksa saklanan durum geri yüklenir. */
export function planFor(kind: WindowKind, defaults: Size, previous?: BrowserWindow): WindowPlan {
  const areas = screen.getAllDisplays().map((d) => d.workArea);
  const primary = screen.getPrimaryDisplay().workArea;
  const prev = previous && !previous.isDestroyed() ? { ...previous.getNormalBounds(), maximized: previous.isMaximized() } : undefined;
  return planWindow({ defaults, saved: sanitizeSaved(readStore()[kind]), previous: prev, areas, primary });
}

/** Pencere kapanırken boyutunu, konumunu ve büyütülmüş durumunu türü için saklar. */
export function rememberBounds(kind: WindowKind, win: BrowserWindow) {
  win.on("close", () => {
    if (win.isDestroyed()) return;
    const store = readStore();
    store[kind] = { ...win.getNormalBounds(), maximized: win.isMaximized() };
    writeStore(store);
  });
}
