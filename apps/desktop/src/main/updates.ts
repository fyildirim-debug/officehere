// GitHub Sürümleri üzerinden güncelleştirme denetimi ve kurulumu.
// Kendini güncelleyebilen paketlerde (Windows kurulumu, AppImage) electron-updater
// güncelleştirmeyi arka planda indirir ve yeniden başlatınca kurar; diğerlerinde
// GitHub API ile yeni sürüm bildirilir ve indirme sayfası açılır.
import { app, BrowserWindow, ipcMain, net, shell } from "electron";
import fs from "node:fs";
import path from "node:path";
import type { AppUpdater } from "electron-updater";
import { compareVersions, releaseFromGitHub, updateMode } from "./updateCore";

const REPO = "fyildirim-debug/officehere";
const LATEST_API = process.env.OH_UPDATE_FEED ?? `https://api.github.com/repos/${REPO}/releases/latest`;
const FIRST_CHECK_MS = Number(process.env.OH_UPDATE_DELAY_MS ?? 10_000);
const INTERVAL_MS = 6 * 60 * 60 * 1000;

type Phase = "idle" | "checking" | "available" | "downloading" | "ready" | "latest" | "error";

interface UpdateState {
  phase: Phase;
  current: string;
  mode: "auto" | "manual" | "off";
  version?: string;
  percent?: number;
  url?: string;
  error?: string;
  checkedAt?: number;
  autoCheck: boolean;
}

let state: UpdateState;
let updater: AppUpdater | null = null;
let timer: NodeJS.Timeout | null = null;

const releasePage = (version: string) => `https://github.com/${REPO}/releases/tag/v${version}`;

function settingsPath() {
  return path.join(app.getPath("userData"), "ayarlar.json");
}

function readSettings(): Record<string, unknown> {
  try {
    return JSON.parse(fs.readFileSync(settingsPath(), "utf8")) as Record<string, unknown>;
  } catch {
    return {};
  }
}

function writeSettings(patch: Record<string, unknown>) {
  try {
    fs.mkdirSync(path.dirname(settingsPath()), { recursive: true });
    fs.writeFileSync(settingsPath(), JSON.stringify({ ...readSettings(), ...patch }, null, 2));
  } catch {
    // Ayar yazılamazsa yalnız bu oturum için geçerli olur.
  }
}

function set(patch: Partial<UpdateState>) {
  state = { ...state, ...patch };
  for (const w of BrowserWindow.getAllWindows()) if (!w.isDestroyed()) w.webContents.send("oh:update-state", state);
}

function message(e: unknown) {
  return e instanceof Error ? e.message : String(e);
}

/** GitHub API ile son sürümü denetler (her pakette çalışır). */
async function checkLatestRelease() {
  const res = await net.fetch(LATEST_API, {
    headers: { Accept: "application/vnd.github+json", "User-Agent": `OfficeHere/${app.getVersion()}` },
  });
  if (!res.ok) throw new Error(`Sürüm bilgisi alınamadı (HTTP ${res.status}).`);
  const rel = releaseFromGitHub(await res.json());
  if (!rel) throw new Error("Sürüm bilgisi okunamadı.");
  if (compareVersions(rel.version, app.getVersion()) > 0) set({ phase: "available", version: rel.version, url: rel.url || releasePage(rel.version), checkedAt: Date.now() });
  else set({ phase: "latest", version: undefined, url: undefined, checkedAt: Date.now() });
}

async function getUpdater(): Promise<AppUpdater> {
  if (updater) return updater;
  const { autoUpdater } = await import("electron-updater");
  autoUpdater.logger = null;
  autoUpdater.autoDownload = true;
  autoUpdater.autoInstallOnAppQuit = true;
  autoUpdater.on("update-available", (i) => set({ phase: "downloading", version: i.version, percent: 0, url: releasePage(i.version) }));
  autoUpdater.on("download-progress", (p) => set({ phase: "downloading", percent: Math.round(p.percent) }));
  autoUpdater.on("update-downloaded", (i) => set({ phase: "ready", version: i.version, percent: 100 }));
  autoUpdater.on("update-not-available", () => set({ phase: "latest", checkedAt: Date.now() }));
  autoUpdater.on("error", (e) => {
    // İndirme yarıda kalırsa yeni sürüm yine bildirilir; kullanıcı elle indirebilir.
    if (state.phase === "downloading" && state.version) set({ phase: "available", percent: undefined, error: message(e) });
  });
  updater = autoUpdater;
  return autoUpdater;
}

export async function checkForUpdates(): Promise<UpdateState> {
  if (state.mode === "off" || state.phase === "checking" || state.phase === "downloading" || state.phase === "ready") return state;
  set({ phase: "checking", error: undefined });
  try {
    if (state.mode === "auto") {
      try {
        await (await getUpdater()).checkForUpdates();
        if ((state.phase as Phase) !== "checking") return state;
      } catch {
        // Sürümde güncelleştirme bilgisi (latest.yml) yoksa yalnız bildirime düşülür.
      }
    }
    await checkLatestRelease();
  } catch (e) {
    set({ phase: "error", error: message(e) });
  }
  return state;
}

async function install() {
  if (state.phase === "ready" && updater) {
    // Sessiz kurulum ve ardından uygulamayı yeniden aç. Kaydedilmemiş belge varsa pencere önce sorar.
    setImmediate(() => updater?.quitAndInstall(true, true));
    return;
  }
  if (state.url) await shell.openExternal(state.url);
}

function schedule() {
  if (timer) clearInterval(timer);
  timer = null;
  if (state.mode === "off" || !state.autoCheck) return;
  setTimeout(() => void checkForUpdates(), FIRST_CHECK_MS);
  timer = setInterval(() => void checkForUpdates(), INTERVAL_MS);
}

export function initUpdates() {
  const mode = updateMode({ platform: process.platform, packaged: app.isPackaged, env: process.env });
  state = { phase: "idle", current: app.getVersion(), mode, autoCheck: readSettings().otomatikGuncellemeDenetimi !== false };

  ipcMain.handle("oh:update-get", () => state);
  ipcMain.handle("oh:update-check", () => checkForUpdates());
  ipcMain.handle("oh:update-install", () => install());
  ipcMain.handle("oh:update-auto", (_e, on: boolean) => {
    writeSettings({ otomatikGuncellemeDenetimi: on });
    set({ autoCheck: on });
    schedule();
    return state;
  });

  schedule();
}
