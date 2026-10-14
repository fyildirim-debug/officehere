// Süreç içi taşıyıcı: Worker olmayan ortamlarda (vitest/jsdom, Node derlem betikleri) aynı RPC'yi
// aynı iş parçacığında çalıştırır. İletiler varsayılan olarak structuredClone ile kopyalanır;
// aktarılan ArrayBuffer'lar gerçek Worker'daki gibi kaynakta boşalır, klonlanamayan değer aynı
// DataCloneError'u verir. İptalde işleyicinin `ctx.signal`'i tetiklenir.

import { exposeIo, type IoEndpoint, type IoHandlers, type WorkerLike } from "./rpc";

export interface InProcessOptions {
  /** false: iletiler kopyalanmadan geçer (yalnız hız gereken betikler için). Varsayılan true. */
  clone?: boolean;
}

type Listener = (event: { data?: unknown; message?: string; error?: unknown }) => void;
type HandlerSource = IoHandlers | (() => IoHandlers | Promise<IoHandlers>);

class InProcessWorker implements WorkerLike {
  private readonly listeners = { message: new Set<Listener>(), messageerror: new Set<Listener>(), error: new Set<Listener>() };
  private readonly abort = new AbortController();
  private dead = false;
  /** Worker tarafının ileti dinleyicisi; işleyiciler yüklenene kadar iletiler kuyrukta bekler. */
  private inbox: ((event: { data: unknown }) => void) | null = null;
  private readonly queue: unknown[] = [];

  constructor(
    source: HandlerSource,
    private readonly clone: boolean,
  ) {
    const endpoint: IoEndpoint = {
      postMessage: (msg, transfer) => this.toMain(msg, transfer),
      addEventListener: (_type, fn) => {
        this.inbox = fn;
        for (const data of this.queue.splice(0)) this.dispatch(() => fn({ data }));
      },
    };
    const start = (handlers: IoHandlers) => {
      if (!this.dead) exposeIo(handlers, endpoint, this.abort.signal);
    };
    if (typeof source === "function") {
      Promise.resolve()
        .then(source)
        .then(start, (e: unknown) => this.emit("error", { message: String((e as Error)?.message ?? e), error: e }));
    } else {
      start(source);
    }
  }

  private copy(msg: unknown, transfer: Transferable[]): unknown {
    return this.clone ? structuredClone(msg, { transfer }) : msg;
  }

  private dispatch(fn: () => void) {
    queueMicrotask(() => {
      if (!this.dead) fn();
    });
  }

  private emit(type: keyof InProcessWorker["listeners"], event: Parameters<Listener>[0]) {
    this.dispatch(() => {
      for (const fn of this.listeners[type]) fn(event);
    });
  }

  /** Worker tarafından ana tarafa. */
  private toMain(msg: unknown, transfer: Transferable[]) {
    if (this.dead) return;
    this.emit("message", { data: this.copy(msg, transfer) });
  }

  /** Ana taraftan Worker tarafına. */
  postMessage(message: unknown, transfer: Transferable[] = []): void {
    if (this.dead) return;
    const data = this.copy(message, transfer);
    const inbox = this.inbox;
    if (inbox) this.dispatch(() => inbox({ data }));
    else this.queue.push(data);
  }

  terminate(): void {
    if (this.dead) return;
    this.dead = true;
    this.queue.length = 0;
    this.abort.abort(new DOMException("İşlem iptal edildi.", "AbortError"));
  }

  addEventListener(type: "message" | "messageerror" | "error", listener: Listener): void {
    this.listeners[type]?.add(listener);
  }
}

/**
 * Süreç içi Worker fabrikası: `createIoWorker(inProcessWorker(handlers))`. İşleyiciler
 * tembel de yüklenebilir: `inProcessWorker(() => import("./ioHandlers").then((m) => m.handlers))`.
 */
export function inProcessWorker(handlers: HandlerSource, options: InProcessOptions = {}): () => WorkerLike {
  return () => new InProcessWorker(handlers, options.clone !== false);
}

/**
 * Worker varsa (tarayıcı, Electron, Android WebView) gerçek Worker, yoksa (Node, jsdom) süreç içi
 * taşıyıcı. Uygulama istemcileri bunu kullanır; testler ve derlem betikleri değişmeden çalışır.
 */
export function workerOrInProcess(factory: () => WorkerLike, handlers: HandlerSource, options?: InProcessOptions): () => WorkerLike {
  return typeof Worker === "function" ? factory : inProcessWorker(handlers, options);
}
