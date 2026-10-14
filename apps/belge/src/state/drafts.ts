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
//
// Kaynak deposu ("kaynaklar"): açılan .docx'in özgün paketi (geçiş katmanı; settings.kaynak = SHA-256),
// onu anan ilk taslakla aynı işlemde bir kez yazılır. Kurtarılan taslak açılınca paket buradan yeniden
// kaydedilir (ensureSource); bulunamazsa belge taslak içeriğiyle açılır ve uygulama uyarır. Hiçbir taslağın
// anmadığı paketler silinir (taslak silinince ve depo açılırken).
import { getSource, openSource, sha256Hex } from "@officehere/core/opc";
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
  /** Andığı kaynak paket (settings.kaynak; kaynak deposunun dizini). */
  kaynak?: string;
}

/** Kaynak deposundaki özgün paket (anahtar: özgün baytların SHA-256'sı = settings.kaynak). */
export interface StoredSource {
  key: string;
  bytes: Uint8Array;
  savedAt: number;
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
/** 2: kaynak deposu ve taslaklarda "kaynak" dizini. */
const DB_VERSION = 2;
const STORE = "drafts";
const SOURCES = "kaynaklar";

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
  /** Taslağı yazar; `source` verilir ve depoda yoksa aynı işlemde (bölünmez) kaynak paketi de yazılır. */
  put(rec: StoredDraft, source?: StoredSource | null): Promise<PutResult>;
  get(key: string): Promise<StoredDraft | null>;
  all(): Promise<StoredDraft[]>;
  delete(key: string): Promise<void>;
  /** Kaynak deposundaki paket (yoksa null). */
  getSource(key: string): Promise<StoredSource | null>;
  /** Hiçbir taslağın anmadığı kaynak paketlerini siler; silinen sayısı. */
  pruneSources(): Promise<number>;
}

function isQuota(err: unknown): boolean {
  const name = (err as { name?: string } | null)?.name ?? "";
  return name === "QuotaExceededError" || name === "NS_ERROR_DOM_QUOTA_REACHED" || /quota/i.test(String((err as Error | null)?.message ?? ""));
}

function valid(rec: unknown, key?: string): StoredDraft | null {
  const d = rec as (Draft & { key?: string; owner?: string; kaynak?: string }) | null;
  if (!d?.content || typeof d !== "object") return null;
  const k = key ?? d.key ?? "";
  const owner = k === LEGACY_DRAFT_KEY ? "" : (d.owner ?? k.slice(DRAFT_PREFIX.length));
  const kaynak = d.kaynak ?? d.settings?.kaynak ?? undefined;
  return { content: d.content, settings: d.settings, fileName: d.fileName ?? "Belge", savedAt: Number(d.savedAt) || 0, ...(d.path ? { path: d.path } : {}), key: k, owner, ...(kaynak ? { kaynak } : {}) };
}

function openDb(): Promise<IDBDatabase | null> {
  return new Promise((resolve) => {
    try {
      if (typeof indexedDB === "undefined") return resolve(null);
      const req = indexedDB.open(DB_NAME, DB_VERSION);
      req.onupgradeneeded = () => {
        const db = req.result;
        const drafts = db.objectStoreNames.contains(STORE) ? req.transaction!.objectStore(STORE) : db.createObjectStore(STORE, { keyPath: "key" });
        if (!drafts.indexNames.contains("kaynak")) drafts.createIndex("kaynak", "kaynak");
        if (!db.objectStoreNames.contains(SOURCES)) db.createObjectStore(SOURCES, { keyPath: "key" });
      };
      req.onsuccess = () => {
        const db = req.result;
        // Daha yeni bir sürüm yükseltmek isterse bağlantı bırakılır; sonraki işlem depoyu yeniden açar.
        db.onversionchange = () => {
          db.close();
          backendPromise = null;
        };
        resolve(db);
      };
      req.onerror = () => resolve(null);
      req.onblocked = () => resolve(null);
    } catch {
      resolve(null);
    }
  });
}

/** İşlem tamamlanınca (veri diske yazılınca) çözülür; kota aşımı işlemi iptal ettirir. */
function txDone<T>(db: IDBDatabase, stores: string | string[], mode: IDBTransactionMode, fn: (tx: IDBTransaction) => IDBRequest<T> | void): Promise<{ ok: boolean; quota: boolean; value?: T }> {
  return new Promise((resolve) => {
    try {
      const tx = db.transaction(stores, mode);
      let value: T | undefined;
      const req = fn(tx);
      if (req) req.onsuccess = () => (value = req.result);
      tx.oncomplete = () => resolve({ ok: true, quota: false, value });
      tx.onabort = () => resolve({ ok: false, quota: isQuota(tx.error) });
      tx.onerror = () => undefined;
    } catch (err) {
      resolve({ ok: false, quota: isQuota(err) });
    }
  });
}

function validSource(rec: unknown): StoredSource | null {
  const s = rec as StoredSource | null;
  return s && typeof s.key === "string" && s.bytes instanceof Uint8Array ? s : null;
}

function idbBackend(db: IDBDatabase): DraftBackend {
  return {
    kind: "indexeddb",
    async put(rec, source) {
      const r = await txDone(db, source ? [STORE, SOURCES] : STORE, "readwrite", (tx) => {
        tx.objectStore(STORE).put(rec);
        if (!source) return;
        const s = tx.objectStore(SOURCES);
        const q = s.count(source.key);
        q.onsuccess = () => {
          if (!q.result) s.put(source);
        };
      });
      return r.ok ? "ok" : r.quota ? "quota" : "error";
    },
    async get(key) {
      const r = await txDone(db, STORE, "readonly", (tx) => tx.objectStore(STORE).get(key) as IDBRequest<unknown>);
      return r.ok ? valid(r.value, key) : null;
    },
    async all() {
      const r = await txDone(db, STORE, "readonly", (tx) => tx.objectStore(STORE).getAll() as IDBRequest<unknown[]>);
      return r.ok ? (r.value ?? []).map((v) => valid(v)).filter((d): d is StoredDraft => !!d) : [];
    },
    async delete(key) {
      await txDone(db, STORE, "readwrite", (tx) => void tx.objectStore(STORE).delete(key));
    },
    async getSource(key) {
      const r = await txDone(db, SOURCES, "readonly", (tx) => tx.objectStore(SOURCES).get(key) as IDBRequest<unknown>);
      return r.ok ? validSource(r.value) : null;
    },
    async pruneSources() {
      let removed = 0;
      // Taslaklar ve paketler tek işlemde: arada yazılan taslağın paketi silinmez.
      await txDone(db, [STORE, SOURCES], "readwrite", (tx) => {
        const sources = tx.objectStore(SOURCES);
        const index = tx.objectStore(STORE).index("kaynak");
        const keys = sources.getAllKeys();
        keys.onsuccess = () => {
          for (const key of keys.result) {
            const c = index.count(key);
            c.onsuccess = () => {
              if (c.result) return;
              sources.delete(key);
              removed++;
            };
          }
        };
      });
      return removed;
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
    // Yerel depoya paket sığmaz: kurtarılan taslak kaynaksız açılır (uyarıyla).
    async getSource() {
      return null;
    },
    async pruneSources() {
      return 0;
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
    await b.pruneSources();
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
export function memoryDraftBackend(opts: { quotaBytes?: number } = {}): DraftBackend & { records: Map<string, StoredDraft>; sources: Map<string, StoredSource> } {
  const records = new Map<string, StoredDraft>();
  const sources = new Map<string, StoredSource>();
  const size = () => [...records.values()].reduce((n, r) => n + JSON.stringify(r).length, 0) + [...sources.values()].reduce((n, s) => n + s.bytes.byteLength, 0);
  return {
    kind: "memory",
    records,
    sources,
    async put(rec, source) {
      const prev = records.get(rec.key);
      const newSource = !!source && !sources.has(source.key);
      records.set(rec.key, structuredClone(rec));
      if (newSource) sources.set(source.key, structuredClone(source));
      if (opts.quotaBytes !== undefined && size() > opts.quotaBytes) {
        if (prev) records.set(rec.key, prev);
        else records.delete(rec.key);
        if (newSource) sources.delete(source.key);
        return "quota";
      }
      return "ok";
    },
    async getSource(key) {
      const s = sources.get(key);
      return s ? structuredClone(s) : null;
    },
    async pruneSources() {
      const used = new Set([...records.values()].map((r) => r.kaynak).filter(Boolean));
      let removed = 0;
      for (const key of [...sources.keys()]) {
        if (used.has(key)) continue;
        sources.delete(key);
        removed++;
      }
      return removed;
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

/** Kotaya sığmadığı için bu oturumda yeniden denenmeyen kaynak paketleri. */
const tooBig = new Set<string>();

/** Bellekteki kaynak paketin depo kaydı (yoksa null). */
function sourceRecord(id: string | undefined): StoredSource | null {
  const pkg = id && !tooBig.has(id) ? getSource(id) : undefined;
  if (!pkg) return null;
  const b = pkg.bytes;
  // Görünüm daha büyük bir arabelleğin parçasıysa yalnız kendisi saklanır.
  const bytes = b.byteOffset === 0 && b.byteLength === b.buffer.byteLength ? b : b.slice();
  return { key: id!, bytes, savedAt: Date.now() };
}

/**
 * Bu pencerenin taslağını yazar. Belge bir kaynak paketi anıyorsa (settings.kaynak) ve paket depoda yoksa
 * aynı işlemde o da yazılır; paket kotaya sığmazsa taslak paketsiz yazılır (kurtarmada uyarılır).
 */
export function writeDraft(d: Draft): Promise<DraftWriteResult> {
  if (closing) return Promise.resolve("closing");
  const kaynak = d.settings?.kaynak ?? undefined;
  const rec: StoredDraft = { content: d.content, settings: d.settings, fileName: d.fileName, savedAt: d.savedAt, ...(d.path ? { path: d.path } : {}), key: draftKey(), owner: windowId, ...(kaynak ? { kaynak } : {}) };
  return serial(async (b): Promise<DraftWriteResult> => {
    if (closing) return "closing";
    const source = sourceRecord(kaynak);
    const r = await b.put(rec, source);
    if (r !== "quota" || !source) return r;
    tooBig.add(source.key);
    return b.put(rec);
  }).catch((): DraftWriteResult => "error");
}

/** Kaynak paketin durumu: bellekte, kurtarma deposundan yeniden açıldı ya da bulunamadı. */
export type SourceState = "bellek" | "depo" | "yok";

const restoring = new Map<string, Promise<SourceState>>();

/**
 * Belgenin andığı kaynak paketi kaydetmeye hazırlar: bellekte yoksa (kurtarılan taslak, yeni oturum) kaynak
 * deposundan açılıp aynı kimlikle kaydedilir. Kimlik yoksa null (taşınacak bir şey yok). "yok" dönerse
 * belge geçiş katmanı olmadan kaydedilir; uygulama kullanıcıyı uyarır.
 */
export function ensureSource(id: string | null | undefined): Promise<SourceState | null> {
  if (!id) return Promise.resolve(null);
  if (getSource(id)) return Promise.resolve("bellek");
  let p = restoring.get(id);
  if (!p) {
    p = (async (): Promise<SourceState> => {
      const rec = await serial((b) => b.getSource(id)).catch(() => null);
      // Bozulmuş kayıt kullanılmaz: içerik kimliğiyle (SHA-256) eşleşmeli.
      if (!rec || (await sha256Hex(rec.bytes)) !== id) return "yok";
      try {
        await openSource(rec.bytes, { id });
        return "depo";
      } catch {
        return "yok";
      }
    })();
    restoring.set(id, p);
    void p.finally(() => restoring.delete(id));
  }
  return p;
}

export function readDraftByKey(key: string): Promise<StoredDraft | null> {
  return serial((b) => b.get(key)).catch(() => null);
}

/** Bu pencerenin taslağı (kök hata sınırından kurtarma). */
export function readOwnDraft(): Promise<StoredDraft | null> {
  return readDraftByKey(draftKey());
}

/** Taslağı siler; artık hiçbir taslağın anmadığı kaynak paketi de silinir. */
export function removeDraft(key: string): Promise<void> {
  return serial(async (b) => {
    await b.delete(key);
    await b.pruneSources();
  }).catch(() => undefined);
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
