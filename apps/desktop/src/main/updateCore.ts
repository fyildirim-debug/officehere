// Güncelleştirme kararlarının yan etkisiz çekirdeği (birim testli).

export type UpdateMode = "auto" | "manual" | "off";

export interface ReleaseInfo {
  version: string;
  url: string;
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

/**
 * Paket türüne göre güncelleştirme kipi: Windows kurulumu (NSIS) ve AppImage kendini
 * günceller; taşınabilir .exe, .deb ve imzasız macOS paketi yalnız yeni sürümü bildirir.
 * Paketlenmemiş geliştirme ortamında denetim kapalıdır (OH_UPDATE_MODE ile zorlanabilir).
 */
export function updateMode(o: { platform: string; packaged: boolean; env: Record<string, string | undefined> }): UpdateMode {
  const forced = o.env.OH_UPDATE_MODE;
  if (forced === "auto" || forced === "manual" || forced === "off") return forced;
  if (!o.packaged) return "off";
  if (o.platform === "win32") return o.env.PORTABLE_EXECUTABLE_DIR ? "manual" : "auto";
  if (o.platform === "linux") return o.env.APPIMAGE ? "auto" : "manual";
  return "manual";
}

/** GitHub "latest release" yanıtından sürüm bilgisi; taslak, ön sürüm ya da bozuk yanıtta null. */
export function releaseFromGitHub(json: unknown): ReleaseInfo | null {
  if (!json || typeof json !== "object") return null;
  const r = json as { tag_name?: unknown; html_url?: unknown; draft?: unknown; prerelease?: unknown };
  if (r.draft || r.prerelease || typeof r.tag_name !== "string") return null;
  const version = r.tag_name.replace(/^v/i, "");
  if (!/^\d+\.\d+\.\d+/.test(version)) return null;
  const url = typeof r.html_url === "string" && /^https:\/\//.test(r.html_url) ? r.html_url : "";
  return { version, url };
}
