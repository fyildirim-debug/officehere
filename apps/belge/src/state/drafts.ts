// Otomatik kurtarma taslakları: her pencerenin kendi taslağı vardır (anahtar pencere kimliğini taşır).
// Taslaklar IndexedDB'de tutulur (resimli belgeler yerel deponun ~5 MB sınırını aşar); IndexedDB
// kullanılamazsa yerel depoya (localStorage) düşülür. Eski sürümlerin yerel depodaki taslakları
// ilk kullanımda bir kez IndexedDB'ye taşınır. Yazılamayan taslak (kota) sessizce yutulmaz:
// writeDraft "quota" döner, uygulama kullanıcıyı uyarır.
//
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
  /** Masaüstünde belgenin diskteki yolu (varsa): kurtarılan belge aynı dosyaya kaydedilebilir. */
  path?: string;
}

/** Depodaki bir taslak: anahtarı ve sahibi pencere. */
export interface StoredDraft extends Draft {
  key: string;
  owner: string;
}

/** Taslak yazma sonucu. */
export type DraftWriteResult = "ok" | "quota" | "error" | "closing";

export const DRAFT_PREFIX = "officehere.belge.draft.";
/** Eski sürümlerin tek (pencereden bağımsız) taslak anahtarı. */
export const LEGACY_DRAFT_KEY = "officehere.belge.draft";
const ALIVE_PREFIX = "officehere.belge.alive.";
const CHANNEL = "officehere.belge.windows";
const HEARTBEAT_MS = 2000;
/** Nabız bu süreden yeniyse pencere canlı sayılır (kanal yanıtı gecikse bile). */
const ALIVE_MS = 5000;
const DB_NAME = "officehere-belge";
const STORE = "drafts";

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

/* ------------------------------------------------------------------ */
/* Depo                                                                 */
/* ------------------------------------------------------------------ */

type PutResult = "ok" | "quota" | "error";

/** Taslak deposu (IndexedDB; yoksa yerel depo; testlerde bellek). */
export interface DraftBackend {
  readonly kind: "indexeddb" | "localStorage" | "memory";
  put(rec: StoredDraft): Promise<PutResult>;
  get(key: string): Promise<StoredDraft | null>;
  all(): Promise<StoredDraft[]>;
  delete(key: string): Promise<void>;
}

function isQuota(err: unknown): boolean {
  const name = (err as { name?: string } | null)?.name ?? "";
  return name === "QuotaExceededError" || name === "NS_ERROR_DOM_QUOTA_REACHED" || /quota/i.test(String((err as Error | null)?.message ?? ""));
}

function valid(rec: unknown, key?: string): StoredDraft | null {
  const d = rec as (Draft & { key?: string; owner?: string }) | null;
  if (!d?.content || typeof d !== "object") return null;
  const k = key ?? d.key ?? "";
  const owner = k === LEGACY_DRAFT_KEY ? "" : (d.owner ?? k.slice(DRAFT_PREFIX.length));
  return { content: d.content, settings: d.settings, fileName: d.fileName ?? "Belge", savedAt: Number(d.savedAt) || 0, ...(d.path ? { path: d.path } : {}), key: k, owner };
}

function openDb(): Promise<IDBDatabase | null> {
  return new Promise((resolve) => {
    try {
      if (typeof indexedDB === "undefined") return resolve(null);
      const req = indexedDB.open(DB_NAME, 1);
      req.onupgradeneeded = () => {
        if (!req.result.objectStoreNames.contains(STORE)) req.result.createObjectStore(STORE, { keyPath: "key" });
      };
      req.onsuccess = () => resolve(req.result);
      req.onerror = () => resolve(null);
      req.onblocked = () => resolve(null);
    } catch {
      resolve(null);
    }
  });
}

/** İşlem tamamlanınca (veri diske yazılınca) çözülür; kota aşımı işlemi iptal ettirir. */
function txDone<T>(db: IDBDatabase, mode: IDBTransactionMode, fn: (s: IDBObjectStore) => IDBRequest<T> | void): Promise<{ ok: boolean; quota: boolean; value?: T }> {
  return new Promise((resolve) => {
    try {
      const tx = db.transaction(STORE, mode);
      let value: T | undefined;
      const req = fn(tx.objectStore(STORE));
      if (req) req.onsuccess = () => (value = req.result);
      tx.oncomplete = () => resolve({ ok: true, quota: false, value });
      tx.onabort = () => resolve({ ok: false, quota: isQuota(tx.error) });
      tx.onerror = () => undefined;
    } catch (err) {
      resolve({ ok: false, quota: isQuota(err) });
    }
  });
}

function idbBackend(db: IDBDatabase): DraftBackend {
  return {
    kind: "indexeddb",
    async put(rec) {
      const r = await txDone(db, "readwrite", (s) => void s.put(rec));
      return r.ok ? "ok" : r.quota ? "quota" : "error";
    },
    async get(key) {
      const r = await txDone(db, "readonly", (s) => s.get(key) as IDBRequest<unknown>);
      return r.ok ? valid(r.value, key) : null;
    },
    async all() {
      const r = await txDone(db, "readonly", (s) => s.getAll() as IDBRequest<unknown[]>);
      return r.ok ? (r.value ?? []).map((v) => valid(v)).filter((d): d is StoredDraft => !!d) : [];
    },
    async delete(key) {
      await txDone(db, "readwrite", (s) => void s.delete(key));
    },
  };
}

/** IndexedDB kullanılamazsa: yerel depo (yalnız küçük belgeler sığar). */
function localBackend(): DraftBackend {
  return {
    kind: "localStorage",
    async put(rec) {
      const s = storage();
      if (!s) return "error";
      try {
        s.setItem(rec.key, JSON.stringify(rec));
        return "ok";
      } catch (err) {
        return isQuota(err) ? "quota" : "error";
      }
    },
    async get(key) {
      try {
        const raw = storage()?.getItem(key);
        return raw ? valid(JSON.parse(raw), key) : null;
      } catch {
        return null;
      }
    },
    async all() {
      return readLocalDrafts();
    },
    async delete(key) {
      try {
        storage()?.removeItem(key);
      } catch {
        // yok say
      }
    },
  };
}

/** Yerel depodaki (eski sürüm ya da yedek yol) taslaklar. */
function readLocalDrafts(): StoredDraft[] {
  const s = storage();
  if (!s) return [];
  const out: StoredDraft[] = [];
  try {
    for (let i = 0; i < s.length; i++) {
      const key = s.key(i);
      if (!key || (key !== LEGACY_DRAFT_KEY && !key.startsWith(DRAFT_PREFIX))) continue;
      try {
        const d = valid(JSON.parse(s.getItem(key) ?? "null"), key);
        if (d) out.push(d);
      } catch {
        // bozuk kayıt
      }
    }
  } catch {
    // depolama okunamadı
  }
  return out;
}

let backendPromise: Promise<DraftBackend> | null = null;

/** Depo: IndexedDB açılabiliyorsa o; eski yerel depo taslakları bir kez taşınır. */
function backend(): Promise<DraftBackend> {
  backendPromise ??= (async () => {
    const db = await openDb();
    if (!db) return localBackend();
    const b = idbBackend(db);
    await migrateLocal(b);
    return b;
  })();
  return backendPromise;
}

/** Eski yerel depo taslaklarını IndexedDB'ye taşır: önce yazılır, yazılabildiyse yerel kopya silinir. */
async function migrateLocal(b: DraftBackend): Promise<number> {
  let moved = 0;
  for (const d of readLocalDrafts()) {
    const key = d.key === LEGACY_DRAFT_KEY ? `${DRAFT_PREFIX}eski-${d.savedAt || Date.now()}` : d.key;
    if ((await b.put({ ...d, key, owner: d.key === LEGACY_DRAFT_KEY ? "" : d.owner })) !== "ok") continue;
    try {
      storage()?.removeItem(d.key);
    } catch {
      // yok say
    }
    moved++;
  }
  return moved;
}

/** Yalnız testler için: depoyu değiştirir (bellek deposu) ya da sıfırlar. */
export function setDraftBackendForTest(b: DraftBackend | null, migrate = true): Promise<number> {
  closing = false;
  queue = Promise.resolve();
  if (!b) {
    backendPromise = null;
    return Promise.resolve(0);
  }
  const p = migrate ? migrateLocal(b) : Promise.resolve(0);
  backendPromise = p.then(() => b);
  return p;
}

/** Bellekte depo (testler; isteğe bağlı kota sınırı). */
export function memoryDraftBackend(opts: { quotaBytes?: number } = {}): DraftBackend & { records: Map<string, StoredDraft> } {
  const records = new Map<string, StoredDraft>();
  const size = () => [...records.values()].reduce((n, r) => n + JSON.stringify(r).length, 0);
  return {
    kind: "memory",
    records,
    async put(rec) {
      const prev = records.get(rec.key);
      records.set(rec.key, structuredClone(rec));
      if (opts.quotaBytes !== undefined && size() > opts.quotaBytes) {
        if (prev) records.set(rec.key, prev);
        else records.delete(rec.key);
        return "quota";
      }
      return "ok";
    },
    async get(key) {
      const r = records.get(key);
      return r ? structuredClone(r) : null;
    },
    async all() {
      return [...records.values()].map((r) => structuredClone(r));
    },
    async delete(key) {
      records.delete(key);
    },
  };
}

/** Tüm taslak işlemleri sırayla çalışır: "yaz" sonra "sil" asla yer değiştirmez. */
let queue: Promise<unknown> = Promise.resolve();
function serial<T>(fn: (b: DraftBackend) => Promise<T>): Promise<T> {
  const run = queue.then(
    async () => fn(await backend()),
    async () => fn(await backend()),
  );
  queue = run.catch(() => undefined);
  return run;
}

/* ------------------------------------------------------------------ */
/* Taslak işlemleri                                                     */
/* ------------------------------------------------------------------ */

/** Pencere kullanıcı isteğiyle kapanıyor: artık taslak yazılmaz. */
let closing = false;

/** Bu pencerenin taslağını yazar. */
export function writeDraft(d: Draft): Promise<DraftWriteResult> {
  if (closing) return Promise.resolve("closing");
  const rec: StoredDraft = { content: d.content, settings: d.settings, fileName: d.fileName, savedAt: d.savedAt, ...(d.path ? { path: d.path } : {}), key: draftKey(), owner: windowId };
  return serial(async (b) => (closing ? "closing" : b.put(rec))).catch((): DraftWriteResult => "error");
}

export function readDraftByKey(key: string): Promise<StoredDraft | null> {
  return serial((b) => b.get(key)).catch(() => null);
}

/** Bu pencerenin taslağı (kök hata sınırından kurtarma). */
export function readOwnDraft(): Promise<StoredDraft | null> {
  return readDraftByKey(draftKey());
}

export function removeDraft(key: string): Promise<void> {
  return serial((b) => b.delete(key)).catch(() => undefined);
}

export function clearOwnDraft(): Promise<void> {
  return removeDraft(draftKey());
}

/**
 * Pencere kullanıcı isteğiyle kapanıyor (kaydedildi ya da "Kaydetme" seçildi):
 * bu pencerenin taslağı silinir ve kapanırken yeniden yazılmaz. Yalnız çökme
 * sonrası kalan taslaklar bir sonraki açılışta önerilir.
 */
export function forgetOwnDraft(): Promise<void> {
  closing = true;
  return clearOwnDraft();
}

/** Depodaki tüm taslaklar (en yeni önce). */
export async function listDrafts(): Promise<StoredDraft[]> {
  const all = await serial((b) => b.all()).catch(() => [] as StoredDraft[]);
  return all.sort((a, b) => b.savedAt - a.savedAt);
}

/**
 * Başka bir pencerenin taslağını bu pencereye devralır: önce kendi anahtarına yazar, sonra eskisini
 * siler (arada çökme olsa bile taslak kaybolmaz).
 */
export async function adoptDraft(d: StoredDraft): Promise<void> {
  if (d.key === draftKey()) return;
  const r = await writeDraft({ content: d.content, settings: d.settings, fileName: d.fileName, savedAt: d.savedAt, path: d.path });
  if (r === "ok") await removeDraft(d.key);
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
  // Depo erkenden açılır (eski taslaklar taşınır).
  void backend();
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
  const others = (await listDrafts()).filter((d) => d.owner !== windowId);
  if (!others.length) return [];
  const alive = await pingWindows(timeout);
  const now = Date.now();
  // Yanıt beklenirken başka pencere taslağı devralmış olabilir: liste yeniden okunur.
  return (await listDrafts()).filter((d) => d.owner !== windowId && (!d.owner || (!alive.has(d.owner) && heartbeatAge(d.owner, now) > ALIVE_MS)));
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
