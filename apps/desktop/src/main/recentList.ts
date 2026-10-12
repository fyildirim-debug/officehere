// Son kullanılanlar listesi üzerindeki saf işlemler (dosya sistemi ve Electron'dan bağımsız; birim testli).

export interface RecentEntry {
  path: string;
  name: string;
  app: "belge" | "tablo" | "sunum";
  openedAt: number;
  pinned?: boolean;
}

/** Diskten okunan listeyi doğrular: bozuk kayıtlar ve yinelenen yollar atılır. */
export function normalizeRecent(raw: unknown): RecentEntry[] {
  if (!Array.isArray(raw)) return [];
  const seen = new Set<string>();
  const out: RecentEntry[] = [];
  for (const f of raw as Partial<RecentEntry>[]) {
    if (!f || typeof f.path !== "string" || typeof f.name !== "string" || seen.has(f.path)) continue;
    if (f.app !== "belge" && f.app !== "tablo" && f.app !== "sunum") continue;
    seen.add(f.path);
    const entry: RecentEntry = { path: f.path, name: f.name, app: f.app, openedAt: typeof f.openedAt === "number" ? f.openedAt : 0 };
    if (f.pinned === true) entry.pinned = true;
    out.push(entry);
  }
  return orderRecent(out);
}

/** Sabitlenenler başta, iki grup da en son açılan önce. */
export function orderRecent(list: readonly RecentEntry[]): RecentEntry[] {
  const byTime = (a: RecentEntry, b: RecentEntry) => b.openedAt - a.openedAt;
  return [...list.filter((f) => f.pinned).sort(byTime), ...list.filter((f) => !f.pinned).sort(byTime)];
}

/** Açılan dosyayı başa alır; sabitlenmişse sabit kalır. Sınır yalnız sabitlenmemişlere uygulanır. */
export function addRecentEntry(list: readonly RecentEntry[], file: RecentEntry, max: number): RecentEntry[] {
  const old = list.find((f) => f.path === file.path);
  const entry: RecentEntry = { path: file.path, name: file.name, app: file.app, openedAt: file.openedAt };
  if (file.pinned ?? old?.pinned) entry.pinned = true;
  let unpinned = 0;
  return orderRecent([entry, ...list.filter((f) => f.path !== file.path)]).filter((f) => f.pinned || ++unpinned <= max);
}

export function pinRecentEntry(list: readonly RecentEntry[], filePath: string, pinned: boolean): RecentEntry[] {
  return orderRecent(
    list.map((f) => {
      if (f.path !== filePath) return f;
      const next: RecentEntry = { path: f.path, name: f.name, app: f.app, openedAt: f.openedAt };
      if (pinned) next.pinned = true;
      return next;
    }),
  );
}

export function removeRecentEntry(list: readonly RecentEntry[], filePath: string): RecentEntry[] {
  return list.filter((f) => f.path !== filePath);
}
