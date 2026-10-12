/**
 * Otomatik Kurtarma deposu (IndexedDB). Her pencere kendi oturum kimliğiyle
 * tek taslak tutar; açılışta başka canlı pencereye ait olmayan en yeni
 * taslak önerilir. Canlı pencereler BroadcastChannel ile sorgulanır.
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

const DB_NAME = "officehere-tablo";
const STORE = "drafts";
/** En çok bu kadar taslak saklanır (eskiler silinir). */
export const MAX_DRAFTS = 5;

export function newSessionId(): string {
  return `${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 8)}`;
}

/** Başka canlı pencerelere ait olmayan en yeni taslak. */
export function pickDraft<T extends { session: string; savedAt: number }>(all: T[], live: Set<string>, own: string): T | null {
  return all.filter((d) => d.session !== own && !live.has(d.session)).sort((a, b) => b.savedAt - a.savedAt)[0] ?? null;
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
    for (const id of draftsToPrune(metas)) await deleteDraft(id);
  }
  return ok;
}

/** Taslakların özeti (büyük veri okunmadan). */
export async function listDraftMeta(): Promise<{ session: string; savedAt: number; fileName: string; sheetCount: number; cellCount: number }[]> {
  const d = await db();
  if (!d) return [];
  const all = (await run(d, "readonly", (s) => s.getAll() as IDBRequest<DraftRecord[]>)) ?? [];
  return all.map((r) => ({ session: r.session, savedAt: r.savedAt, fileName: r.fileName, sheetCount: r.sheetCount, cellCount: r.cellCount }));
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

const CHANNEL = "officehere-tablo-drafts";

/** Bu pencerenin canlı olduğunu soranlara bildirir. */
export function announceLive(session: string): () => void {
  if (typeof BroadcastChannel === "undefined") return () => undefined;
  const ch = new BroadcastChannel(CHANNEL);
  ch.onmessage = (e: MessageEvent<{ type?: string }>) => {
    if (e.data?.type === "who") ch.postMessage({ type: "alive", session });
  };
  return () => ch.close();
}

export function liveSessions(waitMs = 200): Promise<Set<string>> {
  return new Promise((resolve) => {
    const set = new Set<string>();
    if (typeof BroadcastChannel === "undefined") return resolve(set);
    const ch = new BroadcastChannel(CHANNEL);
    ch.onmessage = (e: MessageEvent<{ type?: string; session?: string }>) => {
      if (e.data?.type === "alive" && e.data.session) set.add(e.data.session);
    };
    ch.postMessage({ type: "who" });
    setTimeout(() => {
      ch.close();
      resolve(set);
    }, waitMs);
  });
}
