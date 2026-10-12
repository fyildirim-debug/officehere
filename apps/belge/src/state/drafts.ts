// Otomatik kurtarma taslakları: her pencerenin kendi taslağı vardır (anahtar pencere kimliğini taşır).
// Açılışta yalnız sahibi kapanmış (canlı olmayan) pencerelerin taslakları önerilir. Canlılık
// BroadcastChannel üzerinden sorulur; kanal yoksa ya da pencere meşgulse yerel depodaki nabız
// kaydına bakılır. Taslak yalnız açıkça silinince, kaydedilince ya da başka pencereye
// devredilince (yeni sahibi kendi anahtarına yazdıktan sonra) kaldırılır.
import type { BelgeDocument, DocSettings } from "../model/document";

export interface Draft {
  content: BelgeDocument["content"];
  settings: DocSettings;
  fileName: string;
  savedAt: number;
}

/** Depodaki bir taslak: anahtarı ve sahibi pencere. */
export interface StoredDraft extends Draft {
  key: string;
  owner: string;
}

export const DRAFT_PREFIX = "officehere.belge.draft.";
/** Eski sürümlerin tek (pencereden bağımsız) taslak anahtarı. */
export const LEGACY_DRAFT_KEY = "officehere.belge.draft";
const ALIVE_PREFIX = "officehere.belge.alive.";
const CHANNEL = "officehere.belge.windows";
const HEARTBEAT_MS = 2000;
/** Nabız bu süreden yeniyse pencere canlı sayılır (kanal yanıtı gecikse bile). */
const ALIVE_MS = 5000;

function makeId(): string {
  try {
    return crypto.randomUUID().replace(/-/g, "").slice(0, 16);
  } catch {
    return `${Date.now().toString(36)}${Math.random().toString(36).slice(2, 10)}`;
  }
}

/** Bu pencerenin (sayfa yüklemesinin) kimliği. */
export const windowId: string = makeId();

export const draftKey = (id: string = windowId) => DRAFT_PREFIX + id;

function storage(): Storage | null {
  try {
    return typeof localStorage === "undefined" ? null : localStorage;
  } catch {
    return null;
  }
}

function parse(key: string, raw: string | null): StoredDraft | null {
  if (!raw) return null;
  try {
    const d = JSON.parse(raw) as Draft & { owner?: string };
    if (!d?.content) return null;
    const owner = key === LEGACY_DRAFT_KEY ? "" : (d.owner ?? key.slice(DRAFT_PREFIX.length));
    return { ...d, key, owner };
  } catch {
    return null;
  }
}

/** Pencere kullanıcı isteğiyle kapanıyor: artık taslak yazılmaz. */
let closing = false;

/** Bu pencerenin taslağını yazar. */
export function writeDraft(d: Draft): boolean {
  if (closing) return false;
  const s = storage();
  if (!s) return false;
  try {
    s.setItem(draftKey(), JSON.stringify({ ...d, owner: windowId }));
    return true;
  } catch {
    return false;
  }
}

export function readDraftByKey(key: string): StoredDraft | null {
  const s = storage();
  if (!s) return null;
  try {
    return parse(key, s.getItem(key));
  } catch {
    return null;
  }
}

/** Bu pencerenin taslağı (kök hata sınırından kurtarma). */
export function readOwnDraft(): StoredDraft | null {
  return readDraftByKey(draftKey());
}

export function removeDraft(key: string) {
  try {
    storage()?.removeItem(key);
  } catch {
    // yok say
  }
}

export function clearOwnDraft() {
  removeDraft(draftKey());
}

/**
 * Pencere kullanıcı isteğiyle kapanıyor (kaydedildi ya da "Kaydetme" seçildi):
 * bu pencerenin taslağı silinir ve kapanırken yeniden yazılmaz. Yalnız çökme
 * sonrası kalan taslaklar bir sonraki açılışta önerilir.
 */
export function forgetOwnDraft() {
  closing = true;
  clearOwnDraft();
}

/** Depodaki tüm taslaklar (en yeni önce). */
export function listDrafts(): StoredDraft[] {
  const s = storage();
  if (!s) return [];
  const out: StoredDraft[] = [];
  try {
    for (let i = 0; i < s.length; i++) {
      const key = s.key(i);
      if (!key || (key !== LEGACY_DRAFT_KEY && !key.startsWith(DRAFT_PREFIX))) continue;
      const d = parse(key, s.getItem(key));
      if (d) out.push(d);
    }
  } catch {
    // depolama okunamadı
  }
  return out.sort((a, b) => b.savedAt - a.savedAt);
}

/**
 * Başka bir pencerenin taslağını bu pencereye devralır: önce kendi anahtarına yazar, sonra eskisini
 * siler (arada çökme olsa bile taslak kaybolmaz).
 */
export function adoptDraft(d: StoredDraft): void {
  if (d.key === draftKey()) return;
  const ok = writeDraft({ content: d.content, settings: d.settings, fileName: d.fileName, savedAt: d.savedAt });
  if (ok) removeDraft(d.key);
}

/** Taslağın ilk sözcükleri (listede belgeleri ayırt etmek için). */
export function draftPreview(d: Pick<Draft, "content">, max = 48): string {
  let out = "";
  const walk = (n: { text?: string; content?: unknown[] } | undefined) => {
    if (!n || out.length > max) return;
    if (typeof n.text === "string") out += n.text;
    for (const c of n.content ?? []) {
      walk(c as { text?: string; content?: unknown[] });
      if ((c as { type?: string }).type === "paragraph" || (c as { type?: string }).type === "heading") out += " ";
    }
  };
  walk(d.content as { content?: unknown[] });
  const t = out.replace(/\s+/g, " ").trim();
  return t.length > max ? `${t.slice(0, max).trimEnd()}…` : t;
}

/* ------------------------------------------------------------------ */
/* Pencere canlılığı                                                   */
/* ------------------------------------------------------------------ */

type Msg = { t: "ping"; from: string; nonce: string } | { t: "pong"; from: string; nonce: string };

let channel: BroadcastChannel | null = null;
let beat = 0;

function heartbeat() {
  try {
    storage()?.setItem(ALIVE_PREFIX + windowId, String(Date.now()));
  } catch {
    // yok say
  }
}

function goodbye() {
  try {
    storage()?.removeItem(ALIVE_PREFIX + windowId);
  } catch {
    // yok say
  }
}

/** Pencerenin canlılık yanıtlayıcısını ve nabzını başlatır (bir kez). */
export function startPresence(): void {
  if (typeof window === "undefined" || beat) return;
  heartbeat();
  beat = window.setInterval(heartbeat, HEARTBEAT_MS);
  try {
    channel = new BroadcastChannel(CHANNEL);
    channel.onmessage = (e: MessageEvent<Msg>) => {
      const m = e.data;
      if (m?.t === "ping" && m.from !== windowId) channel?.postMessage({ t: "pong", from: windowId, nonce: m.nonce } satisfies Msg);
    };
  } catch {
    channel = null;
  }
  // Pencere kapanınca (ya da yenilenince) taslağı hemen "sahipsiz" sayılsın.
  window.addEventListener("pagehide", goodbye);
  window.addEventListener("pageshow", (e) => {
    if ((e as PageTransitionEvent).persisted) heartbeat();
  });
}

function heartbeatAge(id: string, now: number): number {
  try {
    const v = Number(storage()?.getItem(ALIVE_PREFIX + id));
    return Number.isFinite(v) && v > 0 ? now - v : Infinity;
  } catch {
    return Infinity;
  }
}

/** Yanıt veren (canlı) pencerelerin kimlikleri. */
export function pingWindows(timeout = 400): Promise<Set<string>> {
  const alive = new Set<string>();
  if (typeof BroadcastChannel === "undefined") return Promise.resolve(alive);
  return new Promise((resolve) => {
    let ch: BroadcastChannel;
    try {
      ch = new BroadcastChannel(CHANNEL);
    } catch {
      resolve(alive);
      return;
    }
    const nonce = Math.random().toString(36).slice(2);
    ch.onmessage = (e: MessageEvent<Msg>) => {
      if (e.data?.t === "pong" && e.data.nonce === nonce) alive.add(e.data.from);
    };
    ch.postMessage({ t: "ping", from: windowId, nonce } satisfies Msg);
    setTimeout(() => {
      ch.close();
      resolve(alive);
    }, timeout);
  });
}

/** Sahibi canlı olmayan taslaklar (bu pencereninki hariç), en yeni önce. */
export async function orphanDrafts(timeout = 400): Promise<StoredDraft[]> {
  const others = listDrafts().filter((d) => d.owner !== windowId);
  if (!others.length) return [];
  const alive = await pingWindows(timeout);
  const now = Date.now();
  // Yanıt beklenirken başka pencere taslağı devralmış olabilir: liste yeniden okunur.
  return listDrafts().filter((d) => d.owner !== windowId && (!d.owner || (!alive.has(d.owner) && heartbeatAge(d.owner, now) > ALIVE_MS)));
}

/** Eski nabız kayıtlarını temizler (sahibi çoktan kapanmış pencereler). */
export function pruneHeartbeats(maxAge = 24 * 3600_000) {
  const s = storage();
  if (!s) return;
  const now = Date.now();
  const stale: string[] = [];
  try {
    for (let i = 0; i < s.length; i++) {
      const key = s.key(i);
      if (key?.startsWith(ALIVE_PREFIX) && heartbeatAge(key.slice(ALIVE_PREFIX.length), now) > maxAge) stale.push(key);
    }
    stale.forEach((k) => s.removeItem(k));
  } catch {
    // yok say
  }
}
