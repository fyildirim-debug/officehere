// Güncelleştirme kararlarının yan etkisiz çekirdeği (birim testli).

export type UpdateMode = "auto" | "manual" | "off";

/** Güncelleme kanalı: kararlı yalnız kararlı sürümleri, beta ayrıca "-beta.N" ön sürümlerini alır. */
export type UpdateChannel = "kararli" | "beta";

export interface ReleaseInfo {
  version: string;
  url: string;
}

/** Ayarlardaki kanal değeri; tanınmayan değer kararlı sayılır. */
export function channelFrom(value: unknown): UpdateChannel {
  return value === "beta" ? "beta" : "kararli";
}

/** "1.2.3", "v1.2.3" ya da "1.2.3-beta.1" biçimindeki iki sürümü karşılaştırır (a > b ise pozitif). */
export function compareVersions(a: string, b: string): number {
  const parse = (v: string) => {
    const [main, pre = ""] = v.trim().replace(/^v/i, "").split("-", 2);
    const nums = main.split(".").map((x) => Number.parseInt(x, 10) || 0);
    while (nums.length < 3) nums.push(0);
    return { nums, pre };
  };
  const x = parse(a);
  const y = parse(b);
  for (let i = 0; i < Math.max(x.nums.length, y.nums.length); i++) {
    const d = (x.nums[i] ?? 0) - (y.nums[i] ?? 0);
    if (d) return d;
  }
  // Ön sürüm, aynı numaralı kararlı sürümden eskidir (1.0.0-beta < 1.0.0).
  if (x.pre === y.pre) return 0;
  if (!x.pre) return 1;
  if (!y.pre) return -1;
  return x.pre.localeCompare(y.pre, "en", { numeric: true });
}

/** Sürümün ön sürüm kimliği ("0.2.0-beta.1" → "beta"); kararlı sürümde boş. */
function prereleaseId(version: string): string {
  const i = version.indexOf("-");
  return i < 0 ? "" : version.slice(i + 1).split(/[.+]/)[0].toLowerCase(); // yerel-ayar: ascii
}

/**
 * Kanal bu sürümü alabilir mi: kararlı kanal yalnız kararlı sürümleri, beta kanalı kararlı ve
 * "-beta.N" sürümlerini alır (rc/alpha gibi diğer ön sürümler uygulamaya bildirilmez). Geri sürüme
 * inilmez: yalnız çalışan sürümden yenisi kabul edilir.
 */
export function acceptsVersion(version: string, current: string, channel: UpdateChannel): boolean {
  const v = version.trim().replace(/^v/i, "");
  if (!/^\d+\.\d+\.\d+/.test(v)) return false;
  const id = prereleaseId(v);
  if (id && !(channel === "beta" && id === "beta")) return false;
  return compareVersions(v, current) > 0;
}

/**
 * Mağaza (Microsoft Store, MSIX) paketi mi: OH_STORE=1 ile derlenen pakette package.json'a
 * `"ohDagitim": "store"` yazılır (scripts/dist.mjs); AppX kabında Electron ayrıca
 * process.windowsStore'u true yapar.
 */
export function isStoreBuild(o: { windowsStore?: boolean; dagitim?: unknown }): boolean {
  return o.windowsStore === true || o.dagitim === "store";
}

/**
 * Paket türüne göre güncelleştirme kipi: Windows kurulumu (NSIS) ve AppImage kendini
 * günceller; taşınabilir .exe, .deb ve imzasız macOS paketi yalnız yeni sürümü bildirir.
 * Mağaza paketini mağaza günceller: denetim hiçbir koşulda yapılmaz (OH_UPDATE_MODE da açamaz).
 * Paketlenmemiş geliştirme ortamında denetim kapalıdır (OH_UPDATE_MODE ile zorlanabilir).
 */
export function updateMode(o: { platform: string; packaged: boolean; env: Record<string, string | undefined>; store?: boolean }): UpdateMode {
  if (o.store) return "off";
  const forced = o.env.OH_UPDATE_MODE;
  if (forced === "auto" || forced === "manual" || forced === "off") return forced;
  if (!o.packaged) return "off";
  if (o.platform === "win32") return o.env.PORTABLE_EXECUTABLE_DIR ? "manual" : "auto";
  if (o.platform === "linux") return o.env.APPIMAGE ? "auto" : "manual";
  return "manual";
}

/** GitHub sürüm yanıtındaki tek sürüm; taslak ya da bozuk kayıtta null. */
function parseRelease(json: unknown): (ReleaseInfo & { prerelease: boolean }) | null {
  if (!json || typeof json !== "object") return null;
  const r = json as { tag_name?: unknown; html_url?: unknown; draft?: unknown; prerelease?: unknown };
  if (r.draft || typeof r.tag_name !== "string") return null;
  const version = r.tag_name.replace(/^v/i, "");
  if (!/^\d+\.\d+\.\d+/.test(version)) return null;
  const url = typeof r.html_url === "string" && /^https:\/\//.test(r.html_url) ? r.html_url : "";
  return { version, url, prerelease: r.prerelease === true };
}

/**
 * GitHub "releases" listesinden (ya da "releases/latest" yanıtından) kanalın alabileceği en yeni sürüm.
 * Kararlı kanal ön sürüm işaretli ya da "-" içeren sürümü hiçbir zaman seçmez. Uygun sürüm yoksa null.
 */
export function pickRelease(json: unknown, channel: UpdateChannel): ReleaseInfo | null {
  let best: ReleaseInfo | null = null;
  for (const item of Array.isArray(json) ? json : [json]) {
    const r = parseRelease(item);
    if (!r) continue;
    const id = prereleaseId(r.version);
    const ok = channel === "beta" ? !id || id === "beta" : !id && !r.prerelease;
    if (ok && (!best || compareVersions(r.version, best.version) > 0)) best = { version: r.version, url: r.url };
  }
  return best;
}

/** electron-updater'ın kanala göre ayarlanan alanları (AppUpdater). */
export interface ChannelUpdater {
  allowPrerelease: boolean;
  allowDowngrade: boolean;
  channel: string | null;
  isUpdateSupported(info: { version: string }): boolean | Promise<boolean>;
}

/**
 * electron-updater'ı (6.8, GitHub sağlayıcısı) kanala bağlar ve kanal değişince yeniden uygulanacak
 * işlevi döndürür. Kanal her denetimde channelOf()'tan okunur.
 * - Ön sürümler yalnız beta kanalında görünür (allowPrerelease). Varsayılan değer çalışan sürüme
 *   göredir (beta kuruluysa true); bu yüzden her iki kanalda da açıkça atanır.
 * - Kanal adı açıkça verilir: kararlıda latest*.yml, betada beta*.yml (yoksa latest*.yml) okunur.
 * - `channel` ataması allowDowngrade'i true yapar (AppUpdater.js); geri sürüme inilmesin diye
 *   ardından false yapılır.
 * - Son kapı: sağlayıcı ne bulursa bulsun acceptsVersion'ın reddettiği sürüm "güncelleme yok" sayılır.
 */
export function bindChannel(u: ChannelUpdater, current: string, channelOf: () => UpdateChannel): () => void {
  const base = u.isUpdateSupported.bind(u);
  u.isUpdateSupported = (info) => acceptsVersion(info.version, current, channelOf()) && base(info);
  const apply = () => {
    const beta = channelOf() === "beta";
    u.allowPrerelease = beta;
    u.channel = beta ? "beta" : "latest";
    u.allowDowngrade = false;
  };
  apply();
  return apply;
}
