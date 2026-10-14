// GitHub Sürümleri üzerinden güncelleştirme denetimi ve kurulumu.
// Kendini güncelleyebilen paketlerde (Windows kurulumu, AppImage) electron-updater
// güncelleştirmeyi arka planda indirir ve yeniden başlatınca kurar; diğerlerinde
// GitHub API ile yeni sürüm bildirilir ve indirme sayfası açılır.
// Kanal (kararlı/beta) ayarlar.json'daki guncellemeKanali'dır; kurallar updateCore.ts'te.
import { app, BrowserWindow, ipcMain, net, shell } from "electron";
import type { AppUpdater } from "electron-updater";
import { readSettings, writeSettings } from "./settings";
import { acceptsVersion, bindChannel, channelFrom, pickRelease, updateMode, type UpdateChannel } from "./updateCore";

const OWNER = "fyildirim-debug";
const NAME = "officehere";
const REPO = `${OWNER}/${NAME}`;
const LATEST_API = process.env.OH_UPDATE_FEED ?? `https://api.github.com/repos/${REPO}/releases/latest`;
// Beta kanalı ön sürümleri de görmek için listeyi okur ("latest" ön sürümü hiç döndürmez).
const RELEASES_API = process.env.OH_UPDATE_RELEASES ?? `https://api.github.com/repos/${REPO}/releases?per_page=20`;
// Testler için: electron-updater'ın GitHub sağlayıcısı yerel bir sunucuya yöneltilir (ör. http://127.0.0.1:4000).
const GITHUB_HOST = process.env.OH_UPDATE_GITHUB;
const FIRST_CHECK_MS = Number(process.env.OH_UPDATE_DELAY_MS ?? 10_000);
const INTERVAL_MS = 6 * 60 * 60 * 1000;

type Phase = "idle" | "checking" | "available" | "downloading" | "ready" | "latest" | "error";

interface UpdateState {
  phase: Phase;
  current: string;
  mode: "auto" | "manual" | "off";
  channel: UpdateChannel;
  version?: string;
  percent?: number;
  url?: string;
  error?: string;
  checkedAt?: number;
  autoCheck: boolean;
}

let state: UpdateState;
let updater: AppUpdater | null = null;
let applyChannel: (() => void) | null = null;
let timer: NodeJS.Timeout | null = null;

const releasePage = (version: string) => `https://github.com/${REPO}/releases/tag/v${version}`;

function set(patch: Partial<UpdateState>) {
  state = { ...state, ...patch };
  for (const w of BrowserWindow.getAllWindows()) if (!w.isDestroyed()) w.webContents.send("oh:update-state", state);
}

function message(e: unknown) {
  return e instanceof Error ? e.message : String(e);
}

/** Bulunan sürümü bu kanal alabilir mi (ör. kanal bu arada kararlıya döndüyse beta artık alınmaz). */
function wanted(version: string) {
  return acceptsVersion(version, app.getVersion(), state.channel);
}

/** GitHub API ile kanalın en yeni sürümünü denetler (her pakette çalışır). */
async function checkLatestRelease() {
  const res = await net.fetch(state.channel === "beta" ? RELEASES_API : LATEST_API, {
    headers: { Accept: "application/vnd.github+json", "User-Agent": `OfficeHere/${app.getVersion()}` },
  });
  if (!res.ok) throw new Error(`Sürüm bilgisi alınamadı (HTTP ${res.status}).`);
  const rel = pickRelease(await res.json(), state.channel);
  if (!rel) throw new Error("Sürüm bilgisi okunamadı.");
  if (wanted(rel.version)) set({ phase: "available", version: rel.version, url: rel.url || releasePage(rel.version), checkedAt: Date.now() });
  else set({ phase: "latest", version: undefined, url: undefined, checkedAt: Date.now() });
}

async function getUpdater(): Promise<AppUpdater> {
  if (updater) return updater;
  const { autoUpdater } = await import("electron-updater");
  autoUpdater.logger = null;
  autoUpdater.autoDownload = true;
  autoUpdater.autoInstallOnAppQuit = true;
  if (GITHUB_HOST) {
    const u = new URL(GITHUB_HOST);
    autoUpdater.forceDevUpdateConfig = true;
    autoUpdater.setFeedURL({ provider: "github", owner: OWNER, repo: NAME, host: u.host, protocol: u.protocol === "http:" ? "http" : "https" });
  }
  applyChannel = bindChannel(autoUpdater, app.getVersion(), () => state.channel);
  autoUpdater.on("update-available", (i) => set({ phase: "downloading", version: i.version, percent: 0, url: releasePage(i.version) }));
  autoUpdater.on("download-progress", (p) => {
    if (state.phase === "downloading") set({ percent: Math.round(p.percent) });
  });
  autoUpdater.on("update-downloaded", (i) => {
    // İndirme sürerken kanal değiştiyse bu sürüm kurulmaz.
    autoUpdater.autoInstallOnAppQuit = wanted(i.version);
    if (autoUpdater.autoInstallOnAppQuit) set({ phase: "ready", version: i.version, percent: 100 });
  });
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

/** Kanal değişimi: bekleyen sürümü yeni kanal almıyorsa (betadan kararlıya dönüş) unutulur, kurulmaz. */
function setChannel(value: unknown) {
  const channel = channelFrom(value);
  writeSettings({ guncellemeKanali: channel });
  const keep = !state.version || acceptsVersion(state.version, app.getVersion(), channel);
  set({ channel, ...(keep ? {} : { phase: "idle" as const, version: undefined, url: undefined, percent: undefined, error: undefined }) });
  applyChannel?.();
  if (updater) updater.autoInstallOnAppQuit = keep;
  // Otomatik denetim kapalıysa kullanıcı "denetle" demeden ağa çıkılmaz.
  if (state.mode !== "off" && state.autoCheck) void checkForUpdates();
  return state;
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
  const settings = readSettings();
  state = { phase: "idle", current: app.getVersion(), mode, channel: channelFrom(settings.guncellemeKanali), autoCheck: settings.otomatikGuncellemeDenetimi !== false };

  ipcMain.handle("oh:update-get", () => state);
  ipcMain.handle("oh:update-check", () => checkForUpdates());
  ipcMain.handle("oh:update-install", () => install());
  ipcMain.handle("oh:update-auto", (_e, on: boolean) => {
    writeSettings({ otomatikGuncellemeDenetimi: on });
    set({ autoCheck: on });
    schedule();
    return state;
  });
  ipcMain.handle("oh:update-channel", (_e, channel: unknown) => setChannel(channel));

  schedule();
}
