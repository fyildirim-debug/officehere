/**
 * Otomatik Kurtarma deposu. Kaydedilmemiş sunu taslakları IndexedDB'de
 * (resim/medya büyük olabileceği için) oturum kimliğiyle tutulur; IndexedDB
 * yoksa küçük taslaklar localStorage'a yazılır. Açılışta başka bir canlı
 * pencereye ait olmayan en yeni taslak önerilir.
 */
import type { Presentation } from "../model/types";

export interface Draft {
  /** Taslağı yazan pencere oturumu. */
  session: string;
  fileName: string;
  /** Masaüstünde dosya yolu (varsa). */
  path?: string;
  slideIdx: number;
  savedAt: number;
  pres: Presentation;
}

export interface DraftMeta {
  session: string;
  savedAt: number;
}

const DB_NAME = "officehere-sunum";
const STORE = "drafts";
const LS_PREFIX = "officehere.sunum.draft.";
/** En çok bu kadar taslak tutulur (eskiler silinir). */
export const MAX_DRAFTS = 5;
/** localStorage yedeğinde kabul edilen en büyük JSON boyutu (karakter). */
const LS_LIMIT = 3_500_000;

/** Başka canlı pencerelerin taslaklarını ayıklayıp en yenisini seçer. */
export function pickDraft<T extends DraftMeta>(drafts: T[], liveSessions: Set<string>, ownSession: string): T | null {
  const cand = drafts.filter((d) => d.session !== ownSession && !liveSessions.has(d.session)).sort((a, b) => b.savedAt - a.savedAt);
  return cand[0] ?? null;
}

/** Fazla taslakların (en eskiler) oturum kimlikleri. */
export function draftsToPrune<T extends DraftMeta>(drafts: T[], keep = MAX_DRAFTS): string[] {
  return drafts
    .slice()
    .sort((a, b) => b.savedAt - a.savedAt)
    .slice(keep)
    .map((d) => d.session);
}

export function newSessionId(): string {
  return `${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 8)}`;
}

let dbPromise: Promise<IDBDatabase | null> | null = null;

function openDb(): Promise<IDBDatabase | null> {
  if (dbPromise) return dbPromise;
  dbPromise = new Promise((resolve) => {
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
  return dbPromise;
}

function tx<T>(db: IDBDatabase, mode: IDBTransactionMode, fn: (s: IDBObjectStore) => IDBRequest<T>): Promise<T | null> {
  return new Promise((resolve) => {
    try {
      const t = db.transaction(STORE, mode);
      const req = fn(t.objectStore(STORE));
      req.onsuccess = () => resolve(req.result);
      req.onerror = () => resolve(null);
    } catch {
      resolve(null);
    }
  });
}

function lsDrafts(): Draft[] {
  const out: Draft[] = [];
  try {
    for (let i = 0; i < localStorage.length; i++) {
      const k = localStorage.key(i);
      if (!k?.startsWith(LS_PREFIX)) continue;
      const d = JSON.parse(localStorage.getItem(k) ?? "null") as Draft | null;
      if (d?.pres?.slides) out.push(d);
    }
  } catch {
    /* depo erişilemez */
  }
  return out;
}

/** Pencere kullanıcı isteğiyle kapanıyor: artık taslak yazılmaz. */
let closing = false;

export async function writeDraft(d: Draft): Promise<boolean> {
  if (closing) return false;
  const db = await openDb();
  if (closing) return false;
  if (db) {
    const ok = await tx(db, "readwrite", (s) => s.put(d));
    if (ok !== null) {
      const all = await listDrafts();
      for (const id of draftsToPrune(all)) await deleteDraft(id);
      return true;
    }
  }
  try {
    const json = JSON.stringify(d);
    if (json.length > LS_LIMIT) return false;
    localStorage.setItem(LS_PREFIX + d.session, json);
    return true;
  } catch {
    return false;
  }
}

export async function listDrafts(): Promise<Draft[]> {
  const db = await openDb();
  const fromDb = db ? ((await tx(db, "readonly", (s) => s.getAll() as IDBRequest<Draft[]>)) ?? []) : [];
  return [...fromDb, ...lsDrafts()];
}

/**
 * Pencere kullanıcı isteğiyle kapanıyor (kaydedildi ya da "Kaydetme" seçildi):
 * taslak silinir ve kapanırken yeniden yazılmaz. Yalnız çökme sonrası kalan
 * taslaklar bir sonraki açılışta önerilir.
 */
export async function forgetDraft(session: string): Promise<void> {
  closing = true;
  await deleteDraft(session);
}

export async function deleteDraft(session: string): Promise<void> {
  try {
    localStorage.removeItem(LS_PREFIX + session);
  } catch {
    /* yok say */
  }
  const db = await openDb();
  if (db) await tx(db, "readwrite", (s) => s.delete(session));
}

/* ---------------- Canlı pencereler (BroadcastChannel) ---------------- */

const CHANNEL = "officehere-sunum-drafts";

/** Bu pencerenin canlı olduğunu soran diğer pencerelere bildirir. */
export function announceLive(session: string): () => void {
  if (typeof BroadcastChannel === "undefined") return () => undefined;
  const ch = new BroadcastChannel(CHANNEL);
  ch.onmessage = (e: MessageEvent<{ type: string }>) => {
    if (e.data?.type === "who") ch.postMessage({ type: "alive", session });
  };
  return () => ch.close();
}

/** Diğer canlı pencerelerin oturumlarını kısa bir süre bekleyerek toplar. */
export function liveSessions(waitMs = 200): Promise<Set<string>> {
  return new Promise((resolve) => {
    const set = new Set<string>();
    if (typeof BroadcastChannel === "undefined") return resolve(set);
    const ch = new BroadcastChannel(CHANNEL);
    ch.onmessage = (e: MessageEvent<{ type: string; session?: string }>) => {
      if (e.data?.type === "alive" && e.data.session) set.add(e.data.session);
    };
    ch.postMessage({ type: "who" });
    window.setTimeout(() => {
      ch.close();
      resolve(set);
    }, waitMs);
  });
}
