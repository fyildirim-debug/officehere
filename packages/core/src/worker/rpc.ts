// Dosya G/Ç'si için Worker RPC çekirdeği. Ana iş parçacığı `createIoWorker` ile çağırır,
// Worker tarafı `exposeIo` ile işleyicilerini yayımlar. Ayrıştırıcılar eşzamanlı döngüler
// olduğundan ileti dinleyemez: iptal, Worker'ı sonlandırıp sonraki çağrıda yeniden kurmaktır.
// Hatalar düz nesne olarak taşınır ve ana tarafta ada göre aynı sınıfa geri kurulur.

import { SaveError } from "../saveErrors";

/* ------------------------------------------------------------------ */
/* Türler                                                               */
/* ------------------------------------------------------------------ */

/** İlerleme bildirimi: `ratio` 0..1, `step` isteğe bağlı adım adı ("Sayfalar okunuyor"). */
export interface IoProgress {
  ratio: number;
  step?: string;
}

/** İşleyiciye verilen bağlam. */
export interface IoContext {
  /** İlerleme bildirir; aynı yüzde ve adım yinelenirse ileti gönderilmez. */
  progress(ratio: number, step?: string): void;
  /**
   * Yalnız süreç içi taşıyıcıda iptalde tetiklenir (eşzamansız işleyiciler erken bırakabilir).
   * Gerçek Worker'da iptal, Worker'ı sonlandırarak olur.
   */
  readonly signal: AbortSignal;
}

// eslint-disable-next-line @typescript-eslint/no-explicit-any
export type IoHandler = (args: any, ctx: IoContext) => unknown;
export type IoHandlers = Record<string, IoHandler>;

/** Sonucun aktarılacak (kopyalanmayacak) tamponlarla döndürülmesi: `withTransfer`. */
export class Transferred<T> {
  constructor(
    readonly value: T,
    readonly transfer: Transferable[],
  ) {}
}

/** İşleyici sonucunu aktarılacak tamponlarla sarar: `return withTransfer(bytes, [bytes.buffer])`. */
export function withTransfer<T>(value: T, transfer: Transferable[]): Transferred<T> {
  return new Transferred(value, transfer);
}

/** `Transferred<T>` sarmalını açar (birleşimlerde her üyeye ayrı uygulanır). */
type Unwrap<R> = R extends Transferred<infer T> ? T : R;

export type IoArgs<H extends IoHandlers, M extends keyof H> = Parameters<H[M]>[0];
export type IoResult<H extends IoHandlers, M extends keyof H> = Unwrap<Awaited<ReturnType<H[M]>>>;

export interface IoCallOptions {
  /** Aktarılacak nesneler (ör. `[bytes.buffer]`): çağrıdan sonra kaynakta boşalırlar. */
  transfer?: Transferable[];
  /** İptal: Worker sonlandırılır, çağrı `signal.reason` (AbortError) ile reddedilir. */
  signal?: AbortSignal;
  onProgress?: (progress: IoProgress) => void;
}

/** Worker'ın RPC'nin kullandığı alt kümesi; süreç içi taşıyıcı da bunu uygular. */
export interface WorkerLike {
  postMessage(message: unknown, transfer: Transferable[]): void;
  terminate(): void;
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  addEventListener(type: "message" | "messageerror" | "error", listener: (event: any) => void): void;
}

/** Worker tarafındaki uç (Worker'da `self`). */
export interface IoEndpoint {
  postMessage(message: unknown, transfer: Transferable[]): void;
  addEventListener(type: "message", listener: (event: { data: unknown }) => void): void;
}

/**
 * Worker kaynağı. Vite'ın Worker'ı paketleyebilmesi için fabrika biçimi gerekir:
 * `() => new Worker(new URL("./io.worker.ts", import.meta.url), { type: "module" })`.
 * Adres verilirse modül Worker'ı olarak kurulur.
 */
export type WorkerSource = string | URL | (() => WorkerLike);

// eslint-disable-next-line @typescript-eslint/no-explicit-any
export type ErrorClass = new (...args: any[]) => Error;

export interface IoWorkerOptions {
  /** Ada (`error.name`) göre geri kurulacak hata sınıfları, ör. `{ FormatError }`. */
  errors?: Record<string, ErrorClass>;
}

export interface IoClient<H extends IoHandlers> {
  call<M extends keyof H & string>(method: M, args: IoArgs<H, M>, options?: IoCallOptions): Promise<IoResult<H, M>>;
  /** Worker'ı hemen sonlandırır; bekleyen çağrılar `IoWorkerError("terminated")` ile reddedilir. */
  terminate(): void;
  /** Yanıt bekleyen çağrı sayısı. */
  readonly pending: number;
}

/* ------------------------------------------------------------------ */
/* Altyapı hataları                                                     */
/* ------------------------------------------------------------------ */

export type IoWorkerErrorCode = "load" | "crashed" | "interrupted" | "terminated" | "clone";

const WORKER_MESSAGES: Record<IoWorkerErrorCode, string> = {
  load: "Dosya işleme bileşeni yüklenemedi. Uygulamayı yeniden başlatıp tekrar deneyin.",
  crashed: "Dosya işlenirken arka plan işlemi beklenmedik biçimde durdu.",
  interrupted: "İşlem, aynı anda süren başka bir işlem iptal edildiği için yarıda kesildi.",
  terminated: "Arka plan işlemi kapatıldı.",
  clone: "Arka plan işlemiyle veri alışverişi başarısız oldu.",
};

/** Worker'ın kendisinden kaynaklanan hata (yüklenemedi, çöktü, kesildi…); iletisi Türkçedir. */
export class IoWorkerError extends Error {
  readonly code: IoWorkerErrorCode;
  /** Teknik ayrıntı (günlük ve "Ayrıntılar" için). */
  detail?: string;
  constructor(code: IoWorkerErrorCode, detail?: string) {
    super(WORKER_MESSAGES[code]);
    this.name = "IoWorkerError";
    this.code = code;
    if (detail) this.detail = detail;
  }
}

/** Kullanıcı iptali mi (gösterilecek hata değil). */
export function isAbortError(e: unknown): boolean {
  return (e as { name?: unknown } | null)?.name === "AbortError";
}

/* ------------------------------------------------------------------ */
/* Hataların taşınması                                                  */
/* ------------------------------------------------------------------ */

/** Sınırdan geçen hata: düz, klonlanabilir nesne. */
export interface SerializedError {
  name: string;
  message: string;
  stack?: string;
  /** DOMException mı (AbortError, QuotaExceededError…). */
  dom?: boolean;
  /** Kendi alanları (`detail`, `code`, `kind`, `path`…); yalnız klonlanabilenler. */
  props?: Record<string, unknown>;
  cause?: SerializedError;
}

const SKIP_KEYS = new Set(["name", "message", "stack", "cause"]);

function cloneable(v: unknown): boolean {
  if (typeof v === "function" || typeof v === "symbol") return false;
  if (v === null || typeof v !== "object") return true;
  try {
    structuredClone(v);
    return true;
  } catch {
    return false;
  }
}

/** Herhangi bir fırlatılan değeri düz nesneye çevirir. */
export function serializeError(e: unknown, depth = 0): SerializedError {
  if (e === null || typeof e !== "object") return { name: "Error", message: String(e) };
  const err = e as Partial<Error> & { cause?: unknown };
  const out: SerializedError = {
    name: typeof err.name === "string" && err.name ? err.name : "Error",
    message: typeof err.message === "string" ? err.message : String(e),
  };
  if (typeof err.stack === "string") out.stack = err.stack;
  if (typeof DOMException !== "undefined" && e instanceof DOMException) out.dom = true;
  const props: Record<string, unknown> = {};
  for (const key of Object.keys(e)) {
    if (SKIP_KEYS.has(key)) continue;
    const v = (e as Record<string, unknown>)[key];
    if (cloneable(v)) props[key] = v;
  }
  if (Object.keys(props).length) out.props = props;
  if (err.cause !== undefined && depth < 3) out.cause = serializeError(err.cause, depth + 1);
  return out;
}

const BUILTIN_ERRORS: Record<string, ErrorClass> = {
  Error,
  TypeError,
  RangeError,
  SyntaxError,
  ReferenceError,
  EvalError,
  URIError,
  SaveError,
  IoWorkerError,
};

function define(target: object, key: string, value: unknown, enumerable = true) {
  try {
    Object.defineProperty(target, key, { value, writable: true, configurable: true, enumerable });
  } catch {
    // dondurulmuş alan: yok say
  }
}

/**
 * Düz nesneden hatayı geri kurar. Sınıf `classes` ya da yerleşik listede adıyla aranır; kurucusu
 * çağrılmaz (imzalar farklı), alanlar düz nesneden geri yazılır. Bilinmeyen ad `Error` olur
 * (`name` ve alanlar korunur).
 */
export function reviveError(s: SerializedError, classes: Record<string, ErrorClass> = {}): Error {
  const Cls = Object.hasOwn(classes, s.name) ? classes[s.name] : Object.hasOwn(BUILTIN_ERRORS, s.name) ? BUILTIN_ERRORS[s.name] : undefined;
  let err: Error;
  if (!Cls && s.dom && typeof DOMException !== "undefined") {
    err = new DOMException(s.message, s.name);
  } else {
    err = Reflect.construct(Error, [s.message], Cls ?? Error) as Error;
    if (err.name !== s.name) define(err, "name", s.name);
  }
  if (s.stack) define(err, "stack", s.stack, false);
  for (const [k, v] of Object.entries(s.props ?? {})) define(err, k, v);
  if (s.cause) define(err, "cause", reviveError(s.cause, classes), false);
  return err;
}

/* ------------------------------------------------------------------ */
/* İletiler                                                             */
/* ------------------------------------------------------------------ */

type CallMessage = { t: "call"; id: number; method: string; args: unknown };
type ReplyMessage =
  | { t: "ready" }
  | { t: "progress"; id: number; ratio: number; step?: string }
  | { t: "ok"; id: number; value: unknown }
  | { t: "err"; id: number; error: SerializedError };

/* ------------------------------------------------------------------ */
/* Worker tarafı                                                        */
/* ------------------------------------------------------------------ */

/**
 * İşleyicileri yayımlar. Worker dosyasında yalnız `exposeIo(handlers)` çağrılır; işleyiciler
 * ayrı bir modülde tutulursa testler ve Node betikleri onları süreç içi taşıyıcıyla kullanır.
 */
export function exposeIo(handlers: IoHandlers, endpoint: IoEndpoint = globalThis as unknown as IoEndpoint, signal?: AbortSignal): void {
  const never = signal ?? new AbortController().signal;
  const send = (msg: ReplyMessage, transfer: Transferable[] = []) => endpoint.postMessage(msg, transfer);

  async function run({ id, method, args }: CallMessage) {
    let lastPct = -1;
    let lastStep: string | undefined;
    const ctx: IoContext = {
      signal: never,
      progress(ratio, step) {
        const r = Number.isFinite(ratio) ? Math.min(1, Math.max(0, ratio)) : 0;
        const pct = Math.floor(r * 100);
        if (pct === lastPct && step === lastStep) return;
        lastPct = pct;
        lastStep = step;
        send({ t: "progress", id, ratio: r, ...(step !== undefined ? { step } : {}) });
      },
    };
    try {
      const fn = Object.hasOwn(handlers, method) ? handlers[method] : undefined;
      if (typeof fn !== "function") throw new Error(`Bilinmeyen işlem: ${method}`);
      const out = await fn(args, ctx);
      const [value, transfer] = out instanceof Transferred ? [out.value, out.transfer] : [out, []];
      try {
        send({ t: "ok", id, value }, transfer);
      } catch (e) {
        send({ t: "err", id, error: serializeError(new IoWorkerError("clone", String((e as Error)?.message ?? e))) });
      }
    } catch (e) {
      send({ t: "err", id, error: serializeError(e) });
    }
  }

  endpoint.addEventListener("message", (e) => {
    const msg = e.data as CallMessage | null;
    if (msg && typeof msg === "object" && msg.t === "call") void run(msg);
  });
  send({ t: "ready" });
}

/* ------------------------------------------------------------------ */
/* Ana iş parçacığı tarafı                                              */
/* ------------------------------------------------------------------ */

function abortReason(signal: AbortSignal): unknown {
  return signal.reason ?? new DOMException("İşlem iptal edildi.", "AbortError");
}

function eventDetail(e: { message?: unknown; filename?: unknown; lineno?: unknown } | null | undefined): string | undefined {
  if (!e || typeof e.message !== "string" || !e.message) return undefined;
  return e.filename ? `${e.message} (${String(e.filename)}:${String(e.lineno ?? "")})` : e.message;
}

interface Instance {
  worker: WorkerLike;
  ready: boolean;
  dead: boolean;
}

interface PendingCall {
  resolve: (value: never) => void;
  reject: (reason: unknown) => void;
  onProgress?: (p: IoProgress) => void;
  cleanup: () => void;
}

/**
 * Worker istemcisi. Worker ilk çağrıda kurulur; iptal, çökme ya da `terminate` sonrasında
 * sonraki çağrıda yeniden kurulur. İptal edilen çağrı Worker'ı sonlandırdığından aynı anda
 * bekleyen diğer çağrılar `IoWorkerError("interrupted")` ile reddedilir.
 */
export function createIoWorker<H extends IoHandlers>(source: WorkerSource, options: IoWorkerOptions = {}): IoClient<H> {
  const classes = options.errors ?? {};
  const make = typeof source === "function" ? source : (): WorkerLike => new Worker(source, { type: "module" });
  const pending = new Map<number, PendingCall>();
  let current: Instance | null = null;
  let seq = 0;

  function kill(reason: IoWorkerError) {
    const inst = current;
    current = null;
    if (inst) {
      inst.dead = true;
      try {
        inst.worker.terminate();
      } catch {
        // zaten kapanmış
      }
    }
    const calls = [...pending.values()];
    pending.clear();
    for (const p of calls) {
      p.cleanup();
      p.reject(reason);
    }
  }

  function onReply(inst: Instance, msg: ReplyMessage | null) {
    if (!msg || typeof msg !== "object") return;
    inst.ready = true;
    if (msg.t === "ready") return;
    const p = pending.get(msg.id);
    if (!p) return;
    if (msg.t === "progress") {
      p.onProgress?.({ ratio: msg.ratio, ...(msg.step !== undefined ? { step: msg.step } : {}) });
      return;
    }
    pending.delete(msg.id);
    p.cleanup();
    if (msg.t === "ok") p.resolve(msg.value as never);
    else p.reject(reviveError(msg.error, classes));
  }

  function spawn(): Instance {
    const inst: Instance = { worker: make(), ready: false, dead: false };
    inst.worker.addEventListener("message", (e: { data: ReplyMessage | null }) => {
      if (!inst.dead) onReply(inst, e.data);
    });
    inst.worker.addEventListener("error", (e: { preventDefault?: () => void } & Parameters<typeof eventDetail>[0]) => {
      if (inst.dead) return;
      // Hata çağrının reddiyle bildirilir; sayfanın genel hata işleyicisine ayrıca düşmesin.
      e?.preventDefault?.();
      kill(new IoWorkerError(inst.ready ? "crashed" : "load", eventDetail(e)));
    });
    inst.worker.addEventListener("messageerror", () => {
      if (!inst.dead) kill(new IoWorkerError("clone", "messageerror"));
    });
    current = inst;
    return inst;
  }

  return {
    get pending() {
      return pending.size;
    },
    terminate() {
      kill(new IoWorkerError("terminated"));
    },
    call(method, args, { transfer = [], signal, onProgress } = {}) {
      if (signal?.aborted) return Promise.reject(abortReason(signal));
      return new Promise((resolve, reject) => {
        const id = ++seq;
        const onAbort = () => {
          if (!pending.has(id)) return;
          pending.delete(id);
          cleanup();
          kill(new IoWorkerError("interrupted"));
          reject(abortReason(signal!));
        };
        const cleanup = () => signal?.removeEventListener("abort", onAbort);
        let inst: Instance;
        try {
          inst = current ?? spawn();
        } catch (e) {
          reject(new IoWorkerError("load", String((e as Error)?.message ?? e)));
          return;
        }
        pending.set(id, { resolve: resolve as (v: never) => void, reject, onProgress, cleanup });
        signal?.addEventListener("abort", onAbort, { once: true });
        try {
          inst.worker.postMessage({ t: "call", id, method, args } satisfies CallMessage, transfer);
        } catch (e) {
          // Klonlanamayan bağımsız değişken (DataCloneError): ileti hiç gitmedi, Worker sağlam.
          pending.delete(id);
          cleanup();
          reject(e);
        }
      });
    },
  };
}
