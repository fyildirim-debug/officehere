import type { RecentFile } from "./bridge";

// Son kullanılan dosyalar listesi için saf yardımcılar (sıralama, sabitleme, arama, görünüm metinleri).

/** Listede bir dosyayı ayırt eden anahtar (tarayıcıda yol yalnız dosya adıdır, uygulama da katılır). */
export function recentKey(f: Pick<RecentFile, "app" | "path">): string {
  return `${f.app}:${f.path}`;
}

/** Sabitlenenler başta, ardından diğerleri; iki grup da en son açılan önce (eşitlikte sıra korunur). */
export function orderRecent<T extends { pinned?: boolean; openedAt?: number }>(list: readonly T[]): T[] {
  const byTime = (a: T, b: T) => (b.openedAt ?? 0) - (a.openedAt ?? 0);
  return [...list.filter((f) => f.pinned).sort(byTime), ...list.filter((f) => !f.pinned).sort(byTime)];
}

/** Sınır yalnız sabitlenmemiş dosyalara uygulanır. */
function limitRecent(list: RecentFile[], max: number): RecentFile[] {
  let unpinned = 0;
  return list.filter((f) => f.pinned || ++unpinned <= max);
}

/** Açılan dosyayı başa alır; dosya önceden sabitlenmişse sabit kalır. */
export function withRecentAdded(list: readonly RecentFile[], file: RecentFile, max: number): RecentFile[] {
  const key = recentKey(file);
  const old = list.find((f) => recentKey(f) === key);
  const entry: RecentFile = { ...file };
  if (entry.pinned === undefined && old?.pinned) entry.pinned = true;
  return limitRecent(orderRecent([entry, ...list.filter((f) => recentKey(f) !== key)]), max);
}

export function withRecentPinned(list: readonly RecentFile[], target: Pick<RecentFile, "app" | "path">, pinned: boolean): RecentFile[] {
  const key = recentKey(target);
  return orderRecent(
    list.map((f) => {
      if (recentKey(f) !== key) return f;
      const next = { ...f };
      if (pinned) next.pinned = true;
      else delete next.pinned;
      return next;
    }),
  );
}

export function withRecentRemoved(list: readonly RecentFile[], target: Pick<RecentFile, "app" | "path">): RecentFile[] {
  const key = recentKey(target);
  return list.filter((f) => recentKey(f) !== key);
}

/**
 * Ad ve yolda büyük/küçük harf duyarsız arama. Türkçe kurallarla (İ→i, I→ı) ve genel
 * kurallarla (I→i) ayrı ayrı denenir: "istanbul" İSTANBUL'u, "invoice" Invoice'u bulur.
 * Boşlukla ayrılmış her sözcük eşleşmelidir.
 */
export function matchesRecentQuery(f: Pick<RecentFile, "name" | "path">, query: string): boolean {
  const terms = query.trim().split(/\s+/).filter(Boolean);
  if (terms.length === 0) return true;
  const hay = `${f.name}\n${f.path}`;
  const hayTr = hay.toLocaleLowerCase("tr-TR");
  const hayRoot = hay.toLowerCase();
  return terms.every((t) => hayTr.includes(t.toLocaleLowerCase("tr-TR")) || hayRoot.includes(t.toLowerCase()));
}

export function filterRecent<T extends Pick<RecentFile, "name" | "path">>(list: readonly T[], query: string): T[] {
  return list.filter((f) => matchesRecentQuery(f, query));
}

const MINUTE = 60_000;
const HOUR = 60 * MINUTE;
const DAY = 24 * HOUR;

function startOfDay(ts: number): number {
  const d = new Date(ts);
  d.setHours(0, 0, 0, 0);
  return d.getTime();
}

/** gg.aa.yyyy */
export function formatShortDate(ts: number): string {
  const d = new Date(ts);
  const pad = (n: number) => String(n).padStart(2, "0");
  return `${pad(d.getDate())}.${pad(d.getMonth() + 1)}.${d.getFullYear()}`;
}

/** "az önce", "5 dakika önce", "2 saat önce", "dün", "3 gün önce"; bir haftadan eskiyse gg.aa.yyyy. */
export function formatRelativeTime(ts: number, now: number = Date.now()): string {
  const diff = now - ts;
  if (diff < MINUTE) return "az önce";
  if (diff < HOUR) return `${Math.floor(diff / MINUTE)} dakika önce`;
  // Takvim günü farkı (yaz saati geçişinde gün 23/25 saat sürebilir).
  const days = Math.round((startOfDay(now) - startOfDay(ts)) / DAY);
  if (days <= 0) return `${Math.floor(diff / HOUR)} saat önce`;
  if (days === 1) return "dün";
  if (days < 7) return `${days} gün önce`;
  return formatShortDate(ts);
}

/** Yolun klasör kısmı: "C:\\Belgeler\\a.docx" → "C:\\Belgeler", "/a.docx" → "/". Klasör yoksa "". */
export function folderOf(filePath: string): string {
  const i = Math.max(filePath.lastIndexOf("/"), filePath.lastIndexOf("\\"));
  if (i < 0) return "";
  if (i === 0) return filePath[0];
  const dir = filePath.slice(0, i);
  // "C:\\a.docx" → "C:\\"
  return /^[A-Za-z]:$/.test(dir) ? dir + filePath[i] : dir;
}

/** Metni ortasından "…" ile kısaltır. */
export function truncateMiddle(text: string, max: number): string {
  if (text.length <= max) return text;
  if (max <= 1) return "…";
  const keep = max - 1;
  const head = Math.ceil(keep / 2);
  return `${text.slice(0, head)}…${text.slice(text.length - (keep - head))}`;
}

/**
 * Uzun yolu ortadan kısaltır: kök (ör. "C:", "/home") ve sığdığı kadar son klasör korunur,
 * aradakiler "…" olur: "C:\\Users\\…\\Projeler\\2026". Klasör adları sığmazsa karakter düzeyinde kısaltır.
 */
export function shortenPath(filePath: string, max = 48): string {
  if (filePath.length <= max) return filePath;
  const back = (filePath.match(/\\/g) ?? []).length;
  const fwd = (filePath.match(/\//g) ?? []).length;
  const sep = back > fwd ? "\\" : "/";
  const parts = filePath.split(sep);
  const first = parts.findIndex((p) => p !== "");
  const headEnd = first + 1;
  if (first < 0 || parts.length - headEnd < 2) return truncateMiddle(filePath, max);
  const head = parts.slice(0, headEnd).join(sep);
  const tail: string[] = [];
  for (let i = parts.length - 1; i >= headEnd; i--) {
    const candidate = [head, "…", parts[i], ...tail].join(sep);
    if (candidate.length > max) break;
    tail.unshift(parts[i]);
  }
  if (tail.length === 0) return truncateMiddle(filePath, max);
  return [head, "…", ...tail].join(sep);
}
