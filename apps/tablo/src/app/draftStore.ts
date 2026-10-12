/**
 * Otomatik Kurtarma deposu (IndexedDB). Her pencere (sayfa yüklemesi) kendi
 * oturum kimliğiyle tek taslak tutar; anahtar pencereye özgüdür, pencereler
 * birbirinin taslağını ezmez. Açılışta yalnız sahibi kapanmış pencerelerin
 * (sahipsiz) taslakları önerilir. Canlılık BroadcastChannel ile sorulur;
 * yanıt gecikirse yerel depodaki nabız kaydına bakılır. Pencere kapanınca
 * (pagehide) nabız silinir, taslağı hemen "sahipsiz" sayılır.
 */
import type { WorkbookDraft } from "../io/draft";

export interface DraftRecord {
  session: string;
  fileName: string;
  path?: string;
  savedAt: number;
  sheetCount: number;
  cellCount: number;
  draft: WorkbookDraft;
}

export type DraftMeta = Omit<DraftRecord, "draft">;

const DB_NAME = "officehere-tablo";
const STORE = "drafts";
/** En çok bu kadar taslak saklanır (eskiler silinir). */
export const MAX_DRAFTS = 8;
const ALIVE_PREFIX = "officehere.tablo.alive.";
const HEARTBEAT_MS = 2000;
/** Nabız bu süreden yeniyse pencere canlı sayılır (kanal yanıtı gecikse bile). */
export const ALIVE_MS = 6000;

export function newSessionId(): string {
  try {
    return crypto.randomUUID().replace(/-/g, "").slice(0, 16);
  } catch {
    return `${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 10)}`;
  }
}

/** Bu pencerenin oturum kimliği (taslak anahtarı). */
export const SESSION_ID = newSessionId();

/** Sahibi canlı olmayan taslaklar (bu pencereninki hariç), en yeni önce. */
export function pickOrphans<T extends { session: string; savedAt: number }>(all: T[], live: Set<string>, own: string, heartbeatAge: (id: string) => number = () => Infinity): T[] {
  return all.filter((d) => d.session !== own && !live.has(d.session) && heartbeatAge(d.session) > ALIVE_MS).sort((a, b) => b.savedAt - a.savedAt);
}

/** Geriye uyumluluk: en yeni sahipsiz taslak. */
export function pickDraft<T extends { session: string; savedAt: number }>(all: T[], live: Set<string>, own: string): T | null {
  return pickOrphans(all, live, own)[0] ?? null;
}

export function draftsToPrune<T extends { session: string; savedAt: number }>(all: T[], keep = MAX_DRAFTS): string[] {
  return [...all].sort((a, b) => b.savedAt - a.savedAt).slice(keep).map((d) => d.session);
}

let dbp: Promise<IDBDatabase | null> | null = null;
function db(): Promise<IDBDatabase | null> {
  dbp ??= new Promise((resolve) => {
    try {
      if (typeof indexedDB === "undefined") return resolve(null);
      const req = indexedDB.open(DB_NAME, 1);
      req.onupgradeneeded = () => {
        if (!req.result.objectStoreNames.contains(STORE)) req.result.createObjectStore(STORE, { keyPath: "session" });
      };
      req.onsuccess = () => resolve(req.result);
      req.onerror = () => resolve(null);
      req.onblocked = () => resolve(null);
    } catch {
      resolve(null);
    }
  });
  return dbp;
}

function run<T>(d: IDBDatabase, mode: IDBTransactionMode, fn: (s: IDBObjectStore) => IDBRequest<T>): Promise<T | null> {
  return new Promise((resolve) => {
    try {
      const t = d.transaction(STORE, mode);
      const req = fn(t.objectStore(STORE));
      req.onsuccess = () => resolve(req.result);
      req.onerror = () => resolve(null);
    } catch {
      resolve(null);
    }
  });
}

export async function writeDraft(rec: DraftRecord): Promise<boolean> {
  const d = await db();
  if (!d) return false;
  const ok = (await run(d, "readwrite", (s) => s.put(rec))) !== null;
  if (ok) {
    const metas = await listDraftMeta();
    for (const id of draftsToPrune(metas)) if (id !== rec.session) await deleteDraft(id);
  }
  return ok;
}

/** Taslakların özeti (büyük veri okunmadan). */
export async function listDraftMeta(): Promise<DraftMeta[]> {
  const d = await db();
  if (!d) return [];
  const all = (await run(d, "readonly", (s) => s.getAll() as IDBRequest<DraftRecord[]>)) ?? [];
  return all.map((r) => ({ session: r.session, savedAt: r.savedAt, fileName: r.fileName, path: r.path, sheetCount: r.sheetCount, cellCount: r.cellCount }));
}

export async function readDraft(session: string): Promise<DraftRecord | null> {
  const d = await db();
  if (!d) return null;
  return (await run(d, "readonly", (s) => s.get(session) as IDBRequest<DraftRecord>)) ?? null;
}

export async function deleteDraft(session: string): Promise<void> {
  const d = await db();
  if (d) await run(d, "readwrite", (s) => s.delete(session));
}

/**
 * Başka pencerenin taslağını bu pencereye devralır: önce bu pencerenin
 * anahtarına yazılır, sonra eskisi silinir (arada kapanma olsa bile kaybolmaz).
 */
export async function adoptDraft(rec: DraftRecord): Promise<void> {
  if (rec.session === SESSION_ID) return;
  const ok = await writeDraft({ ...rec, session: SESSION_ID });
  if (ok) await deleteDraft(rec.session);
}

/* ------------------------------------------------------------------ */
/* Pencere canlılığı                                                   */
/* ------------------------------------------------------------------ */

const CHANNEL = "officehere-tablo-drafts";

function storage(): Storage | null {
  try {
    return typeof localStorage === "undefined" ? null : localStorage;
  } catch {
    return null;
  }
}

function heartbeat(): void {
  try {
    storage()?.setItem(ALIVE_PREFIX + SESSION_ID, String(Date.now()));
  } catch {
    /* yok say */
  }
}

/** Nabız kaydının yaşı (yoksa sonsuz). */
export function heartbeatAge(id: string, now = Date.now()): number {
  try {
    const v = Number(storage()?.getItem(ALIVE_PREFIX + id));
    return Number.isFinite(v) && v > 0 ? now - v : Infinity;
  } catch {
    return Infinity;
  }
}

/** Bu pencerenin canlı olduğunu soranlara bildirir ve nabız tutar. */
export function announceLive(session: string = SESSION_ID): () => void {
  heartbeat();
  const beat = window.setInterval(heartbeat, HEARTBEAT_MS);
  const goodbye = () => {
    try {
      storage()?.removeItem(ALIVE_PREFIX + session);
    } catch {
      /* yok say */
    }
  };
  const back = (e: PageTransitionEvent) => {
    if (e.persisted) heartbeat();
  };
  window.addEventListener("pagehide", goodbye);
  window.addEventListener("pageshow", back);
  let ch: BroadcastChannel | null = null;
  try {
    ch = typeof BroadcastChannel === "undefined" ? null : new BroadcastChannel(CHANNEL);
    if (ch)
      ch.onmessage = (e: MessageEvent<{ type?: string; nonce?: string }>) => {
        if (e.data?.type === "who") ch?.postMessage({ type: "alive", session, nonce: e.data.nonce });
      };
  } catch {
    ch = null;
  }
  return () => {
    window.clearInterval(beat);
    window.removeEventListener("pagehide", goodbye);
    window.removeEventListener("pageshow", back);
    ch?.close();
  };
}

export function liveSessions(waitMs = 400): Promise<Set<string>> {
  return new Promise((resolve) => {
    const set = new Set<string>();
    if (typeof BroadcastChannel === "undefined") return resolve(set);
    let ch: BroadcastChannel;
    try {
      ch = new BroadcastChannel(CHANNEL);
    } catch {
      return resolve(set);
    }
    const nonce = Math.random().toString(36).slice(2);
    ch.onmessage = (e: MessageEvent<{ type?: string; session?: string; nonce?: string }>) => {
      if (e.data?.type === "alive" && e.data.session && (!e.data.nonce || e.data.nonce === nonce)) set.add(e.data.session);
    };
    ch.postMessage({ type: "who", nonce });
    setTimeout(() => {
      ch.close();
      resolve(set);
    }, waitMs);
  });
}

/** Sahibi kapanmış pencerelerin taslakları (en yeni önce). */
export async function orphanDrafts(): Promise<DraftMeta[]> {
  const metas = await listDraftMeta();
  if (!metas.some((m) => m.session !== SESSION_ID)) return [];
  const live = await liveSessions();
  // Yanıt beklenirken başka pencere taslağı devralmış olabilir: liste yeniden okunur
  return pickOrphans(await listDraftMeta(), live, SESSION_ID, (id) => heartbeatAge(id));
}

/** Eski nabız kayıtlarını temizler (sahibi çoktan kapanmış pencereler). */
export function pruneHeartbeats(maxAge = 24 * 3600_000): void {
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
    /* yok say */
  }
}
