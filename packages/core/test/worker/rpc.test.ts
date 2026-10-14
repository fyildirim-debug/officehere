// Worker RPC: iptal, ilerleme sırası, hata aktarımı ve ArrayBuffer aktarımı (süreç içi taşıyıcıyla
// ve yanıt vermeyen sahte Worker'la; gerçek Worker e2e-masaustu/worker.spec.ts'de).
import { describe, expect, it } from "vitest";
import {
  IoWorkerError,
  SaveError,
  createIoWorker,
  inProcessWorker,
  isAbortError,
  reviveError,
  serializeError,
  withTransfer,
  workerOrInProcess,
  type IoContext,
  type IoHandlers,
  type WorkerLike,
} from "../../src";

/** Uygulamalardaki FormatError'ın benzeri (kurucusu farklı imzalı). */
class FormatError extends Error {
  detail?: string;
  constructor(message: string, detail?: string) {
    super(message);
    this.name = "FormatError";
    if (detail) this.detail = detail;
  }
}

/** Kurucusu ilk bağımsız değişkeni ileti olarak almayan sınıf (Sunum PartError benzeri). */
class PartError extends Error {
  kind: "corrupt" | "tooLarge";
  path: string;
  constructor(kind: "corrupt" | "tooLarge", path: string, message: string) {
    super(message);
    this.name = "PartError";
    this.kind = kind;
    this.path = path;
  }
}

const sleep = (ms: number) => new Promise((r) => setTimeout(r, ms));

/** Sayaçlı fabrika: kaç Worker kurulduğunu sayar. */
function counted(factory: () => WorkerLike) {
  const made: WorkerLike[] = [];
  const make = () => {
    const w = factory();
    made.push(w);
    return w;
  };
  return { make, made };
}

const handlers = {
  echo: (a: { text: string }) => a.text.toLocaleUpperCase("tr-TR"),
  sum: (a: { bytes: ArrayBuffer }) => new Uint8Array(a.bytes).reduce((s, b) => s + b, 0),
  roundTrip: (a: { bytes: ArrayBuffer }) => withTransfer({ bytes: a.bytes, size: a.bytes.byteLength }, [a.bytes]),
  steps: (a: { n: number }, ctx: IoContext) => {
    for (let i = 1; i <= a.n; i++) ctx.progress(i / a.n, i <= a.n / 2 ? "Okunuyor" : "Kuruluyor");
    return "bitti";
  },
  noisy: (_: undefined, ctx: IoContext) => {
    // Aynı yüzde ve adım yinelenir: yalnız değişenler gider.
    for (let i = 0; i <= 1000; i++) ctx.progress(i / 1000, "Satırlar");
    return true;
  },
  slow: async (a: { ms: number }, ctx: IoContext) => {
    const end = Date.now() + a.ms;
    while (Date.now() < end) {
      if (ctx.signal.aborted) throw ctx.signal.reason;
      ctx.progress(1 - (end - Date.now()) / a.ms);
      await sleep(5);
    }
    return "yavaş bitti";
  },
  failFormat: () => {
    throw new FormatError("Dosya bozuk ya da beklenmeyen bir yapıda; okunamadı.", "zip: end of central directory not found");
  },
  failPart: () => {
    throw new PartError("tooLarge", "xl/worksheets/sheet1.xml", "Parça çok büyük.");
  },
  failRange: () => {
    throw new RangeError("Maximum call stack size exceeded");
  },
  failSave: () => {
    throw new SaveError("ENOSPC");
  },
  failQuota: () => {
    throw new DOMException("Kota doldu", "QuotaExceededError");
  },
  failString: () => {
    throw "düz metin hata";
  },
  failCause: () => {
    throw new Error("Dış hata", { cause: new TypeError("iç hata") });
  },
  badResult: () => ({ fn: () => 1 }),
} satisfies IoHandlers;

type H = typeof handlers;

describe("Worker RPC (süreç içi taşıyıcı)", () => {
  it("çağrı yapar, eşzamanlı çağrılar birbirine karışmaz", async () => {
    const io = createIoWorker<H>(inProcessWorker(handlers));
    const [a, b, c] = await Promise.all([io.call("echo", { text: "ılık" }), io.call("echo", { text: "iş" }), io.call("steps", { n: 4 })]);
    expect([a, b, c]).toEqual(["ILIK", "İŞ", "bitti"]);
    expect(io.pending).toBe(0);
  });

  it("ArrayBuffer aktarılınca kaynakta boşalır, aktarılmayınca kopyalanır", async () => {
    const io = createIoWorker<H>(inProcessWorker(handlers));
    const bytes = new Uint8Array(1 << 20).fill(3).buffer;
    const p = io.call("sum", { bytes }, { transfer: [bytes] });
    // postMessage eşzamanlıdır: çağrı döner dönmez kaynak boşalmıştır.
    expect(bytes.byteLength).toBe(0);
    expect(await p).toBe(3 << 20);

    const kept = new Uint8Array(1024).fill(1).buffer;
    expect(await io.call("sum", { bytes: kept })).toBe(1024);
    expect(kept.byteLength).toBe(1024);
  });

  it("sonuç withTransfer ile aktarılır, içerik bozulmaz", async () => {
    const io = createIoWorker<H>(inProcessWorker(handlers));
    const src = new Uint8Array(4096).map((_, i) => i & 255).buffer;
    const out = await io.call("roundTrip", { bytes: src }, { transfer: [src] });
    expect(src.byteLength).toBe(0);
    expect(out.size).toBe(4096);
    expect(out.bytes.byteLength).toBe(4096);
    expect(new Uint8Array(out.bytes)[255]).toBe(255);
  });

  it("ilerleme iletileri sırayla ve sonuçtan önce gelir", async () => {
    const io = createIoWorker<H>(inProcessWorker(handlers));
    const seen: string[] = [];
    const result = await io.call("steps", { n: 4 }, { onProgress: (p) => seen.push(`${p.step}:${p.ratio}`) }).then((r) => {
      seen.push(`sonuç:${r}`);
      return r;
    });
    expect(result).toBe("bitti");
    expect(seen).toEqual(["Okunuyor:0.25", "Okunuyor:0.5", "Kuruluyor:0.75", "Kuruluyor:1", "sonuç:bitti"]);
  });

  it("aynı yüzde ve adım yinelenmez (1001 bildirim → 101 ileti), oran artan sırada", async () => {
    const io = createIoWorker<H>(inProcessWorker(handlers));
    const ratios: number[] = [];
    await io.call("noisy", undefined, { onProgress: (p) => ratios.push(p.ratio) });
    expect(ratios).toHaveLength(101);
    expect(ratios[0]).toBe(0);
    expect(ratios.at(-1)).toBe(1);
    expect(ratios.every((r, i) => i === 0 || r > ratios[i - 1])).toBe(true);
  });

  it("hata sınıfı, iletisi ve alanları korunur", async () => {
    const io = createIoWorker<H>(inProcessWorker(handlers), { errors: { FormatError, PartError } });

    const f = await io.call("failFormat", undefined).catch((e: unknown) => e);
    expect(f).toBeInstanceOf(FormatError);
    expect(f).toBeInstanceOf(Error);
    expect((f as FormatError).name).toBe("FormatError");
    expect((f as FormatError).message).toBe("Dosya bozuk ya da beklenmeyen bir yapıda; okunamadı.");
    expect((f as FormatError).detail).toBe("zip: end of central directory not found");
    expect((f as Error).stack).toContain("failFormat");

    const p = await io.call("failPart", undefined).catch((e: unknown) => e);
    expect(p).toBeInstanceOf(PartError);
    expect(p).toMatchObject({ kind: "tooLarge", path: "xl/worksheets/sheet1.xml", message: "Parça çok büyük." });

    const r = await io.call("failRange", undefined).catch((e: unknown) => e);
    expect(r).toBeInstanceOf(RangeError);
    expect((r as Error).message).toBe("Maximum call stack size exceeded");

    const s = await io.call("failSave", undefined).catch((e: unknown) => e);
    expect(s).toBeInstanceOf(SaveError);
    expect(s).toMatchObject({ code: "ENOSPC", reason: "Disk dolu", message: "Disk dolu." });

    const q = await io.call("failQuota", undefined).catch((e: unknown) => e);
    expect(q).toBeInstanceOf(DOMException);
    expect((q as DOMException).name).toBe("QuotaExceededError");

    const t = await io.call("failString", undefined).catch((e: unknown) => e);
    expect(t).toBeInstanceOf(Error);
    expect((t as Error).message).toBe("düz metin hata");

    const c = await io.call("failCause", undefined).catch((e: unknown) => e);
    expect((c as Error).cause).toBeInstanceOf(TypeError);
    expect(((c as Error).cause as Error).message).toBe("iç hata");
  });

  it("kayıtlı olmayan sınıf adı ve alanlarıyla Error olarak gelir", async () => {
    const io = createIoWorker<H>(inProcessWorker(handlers));
    const e = (await io.call("failPart", undefined).catch((x: unknown) => x)) as Error & { kind?: string };
    expect(e).toBeInstanceOf(Error);
    expect(e.name).toBe("PartError");
    expect(e.kind).toBe("tooLarge");
  });

  it("bilinmeyen işlem ve klonlanamayan sonuç anlaşılır hatayla reddedilir", async () => {
    const io = createIoWorker<IoHandlers>(inProcessWorker(handlers));
    await expect(io.call("yok", undefined)).rejects.toThrow("Bilinmeyen işlem: yok");
    const e = await io.call("badResult", undefined).catch((x: unknown) => x);
    expect(e).toBeInstanceOf(IoWorkerError);
    expect((e as IoWorkerError).code).toBe("clone");
    // Worker sağlam kalır.
    expect(await io.call("echo", { text: "a" })).toBe("A");
  });

  it("klonlanamayan bağımsız değişken çağrıyı DataCloneError ile reddeder", async () => {
    const io = createIoWorker<IoHandlers>(inProcessWorker(handlers));
    const e = await io.call("echo", { text: "a", fn: () => 1 }).catch((x: unknown) => x);
    expect((e as DOMException).name).toBe("DataCloneError");
    expect(io.pending).toBe(0);
    expect(await io.call("echo", { text: "b" })).toBe("B");
  });

  it("iptal 1 sn içinde reddeder, Worker sonlandırılır ve sonraki çağrıda yeniden kurulur", async () => {
    const { make, made } = counted(inProcessWorker(handlers));
    const io = createIoWorker<H>(make);
    const ctrl = new AbortController();
    const progress: number[] = [];
    const p = io.call("slow", { ms: 10_000 }, { signal: ctrl.signal, onProgress: (x) => progress.push(x.ratio) });
    await expect.poll(() => progress.length).toBeGreaterThan(1);
    const t0 = performance.now();
    ctrl.abort();
    const err = await p.catch((e: unknown) => e);
    const elapsed = performance.now() - t0;
    expect(isAbortError(err)).toBe(true);
    expect(elapsed).toBeLessThan(1000);
    expect(made).toHaveLength(1);
    // İptalden sonra ilerleme gelmez.
    const count = progress.length;
    await sleep(30);
    expect(progress.length).toBe(count);
    expect(await io.call("echo", { text: "yeni" })).toBe("YENİ");
    expect(made).toHaveLength(2);
  });

  it("önceden iptal edilmiş sinyal Worker kurmadan reddeder; tamamlanmış çağrının iptali etkisizdir", async () => {
    const { make, made } = counted(inProcessWorker(handlers));
    const io = createIoWorker<H>(make);
    const done = new AbortController();
    done.abort();
    await expect(io.call("echo", { text: "a" }, { signal: done.signal })).rejects.toMatchObject({ name: "AbortError" });
    expect(made).toHaveLength(0);

    const late = new AbortController();
    expect(await io.call("echo", { text: "a" }, { signal: late.signal })).toBe("A");
    late.abort();
    expect(await io.call("echo", { text: "b" })).toBe("B");
    expect(made).toHaveLength(1);
  });

  it("terminate bekleyenleri reddeder, sonraki çağrı yeni Worker'la çalışır", async () => {
    const { make, made } = counted(inProcessWorker(handlers));
    const io = createIoWorker<H>(make);
    const p = io.call("slow", { ms: 10_000 });
    io.terminate();
    await expect(p).rejects.toMatchObject({ name: "IoWorkerError", code: "terminated" });
    expect(await io.call("echo", { text: "x" })).toBe("X");
    expect(made).toHaveLength(2);
  });

  it("tembel yüklenen işleyiciler; yüklenemezse 'load' hatası", async () => {
    const io = createIoWorker<H>(inProcessWorker(async () => handlers));
    expect(await io.call("echo", { text: "tembel" })).toBe("TEMBEL");
    const broken = createIoWorker<H>(
      inProcessWorker(() => {
        throw new Error("Failed to fetch dynamically imported module");
      }),
    );
    const e = await broken.call("echo", { text: "a" }).catch((x: unknown) => x);
    expect(e).toBeInstanceOf(IoWorkerError);
    expect((e as IoWorkerError).code).toBe("load");
    expect((e as IoWorkerError).detail).toContain("dynamically imported module");
  });

  it("Node'da workerOrInProcess süreç içi taşıyıcıyı seçer", async () => {
    const io = createIoWorker<H>(
      workerOrInProcess(() => {
        throw new Error("gerçek Worker kurulmamalı");
      }, handlers),
    );
    expect(await io.call("echo", { text: "node" })).toBe("NODE");
  });
});

/** Yanıt vermeyen (eşzamanlı döngüde kilitli) Worker taklidi. */
class StuckWorker implements WorkerLike {
  terminated = false;
  posted: unknown[] = [];
  private listeners: Record<string, ((e: unknown) => void)[]> = {};
  postMessage(message: unknown) {
    this.posted.push(message);
  }
  terminate() {
    this.terminated = true;
  }
  addEventListener(type: string, fn: (e: unknown) => void) {
    (this.listeners[type] ??= []).push(fn);
  }
  fire(type: string, event: unknown) {
    for (const fn of this.listeners[type] ?? []) fn(event);
  }
}

describe("Worker RPC (yanıt vermeyen Worker)", () => {
  it("iptal kilitli Worker'ı sonlandırır, aynı anda bekleyenleri 'interrupted' ile reddeder", async () => {
    const { make, made } = counted(() => new StuckWorker());
    const io = createIoWorker<IoHandlers>(make);
    const ctrl = new AbortController();
    const target = io.call("parse", { n: 1 }, { signal: ctrl.signal });
    const other = io.call("parse", { n: 2 });
    const t0 = performance.now();
    ctrl.abort();
    const [a, b] = await Promise.all([target.catch((e: unknown) => e), other.catch((e: unknown) => e)]);
    expect(performance.now() - t0).toBeLessThan(1000);
    expect(isAbortError(a)).toBe(true);
    expect(b).toBeInstanceOf(IoWorkerError);
    expect((b as IoWorkerError).code).toBe("interrupted");
    expect((made[0] as StuckWorker).terminated).toBe(true);
    expect(io.pending).toBe(0);
    void io.call("parse", { n: 3 });
    expect(made).toHaveLength(2);
    expect((made[1] as StuckWorker).posted).toEqual([{ t: "call", id: 3, method: "parse", args: { n: 3 } }]);
  });

  it("iptal nedeni (signal.reason) korunur", async () => {
    const io = createIoWorker<IoHandlers>(() => new StuckWorker());
    const ctrl = new AbortController();
    const p = io.call("parse", undefined, { signal: ctrl.signal });
    const reason = new DOMException("Kullanıcı vazgeçti", "AbortError");
    ctrl.abort(reason);
    await expect(p).rejects.toBe(reason);
  });

  it("hazır olmadan gelen hata 'load', sonra gelen 'crashed' sayılır", async () => {
    const { make, made } = counted(() => new StuckWorker());
    const io = createIoWorker<IoHandlers>(make);
    const p1 = io.call("parse", undefined);
    let prevented = false;
    (made[0] as StuckWorker).fire("error", { message: "", preventDefault: () => (prevented = true) });
    await expect(p1).rejects.toMatchObject({ name: "IoWorkerError", code: "load" });
    expect(prevented).toBe(true);
    expect((made[0] as StuckWorker).terminated).toBe(true);

    const p2 = io.call("parse", undefined);
    const w = made[1] as StuckWorker;
    w.fire("message", { data: { t: "ready" } });
    w.fire("error", { message: "Uncaught TypeError: x is not a function", filename: "io.worker.js", lineno: 12 });
    const e = (await p2.catch((x: unknown) => x)) as IoWorkerError;
    expect(e.code).toBe("crashed");
    expect(e.detail).toBe("Uncaught TypeError: x is not a function (io.worker.js:12)");
    expect(e.message).toMatch(/beklenmedik/);
  });

  it("sonlandırılmış Worker'dan gecikmiş ileti yok sayılır", async () => {
    const { make, made } = counted(() => new StuckWorker());
    const io = createIoWorker<IoHandlers>(make);
    const ctrl = new AbortController();
    const p = io.call("parse", undefined, { signal: ctrl.signal });
    ctrl.abort();
    await p.catch(() => undefined);
    const fresh = io.call("parse", undefined);
    // Eski Worker'ın (aynı kimlikli) geç yanıtı yeni çağrıyı etkilemez.
    (made[0] as StuckWorker).fire("message", { data: { t: "ok", id: 2, value: "eski" } });
    (made[1] as StuckWorker).fire("message", { data: { t: "ok", id: 2, value: "yeni" } });
    expect(await fresh).toBe("yeni");
  });
});

describe("hata dönüştürme", () => {
  it("işlev ve sembol alanları atlanır, düz nesne klonlanabilir", () => {
    const e = Object.assign(new FormatError("ileti", "ayrıntı"), { fn: () => 1, nested: { a: [1, 2] } });
    const s = serializeError(e);
    expect(s.props).toEqual({ detail: "ayrıntı", nested: { a: [1, 2] } });
    expect(() => structuredClone(s)).not.toThrow();
    const back = reviveError(structuredClone(s), { FormatError }) as FormatError;
    expect(back).toBeInstanceOf(FormatError);
    expect(back.detail).toBe("ayrıntı");
    expect(Object.prototype.toString.call(back)).toBe("[object Error]");
  });

  it("Belge'deki gibi farklı sınıf aynı adla eşlenebilir", () => {
    class BelgeOpenError extends Error {
      readonly detail: string;
      constructor(message: string, detail = "") {
        super(message);
        this.name = "FormatError";
        this.detail = detail;
      }
    }
    const back = reviveError(serializeError(new BelgeOpenError("Dosya bozuk.", "xml")), { FormatError: BelgeOpenError });
    expect(back).toBeInstanceOf(BelgeOpenError);
    expect(back.name).toBe("FormatError");
    expect((back as BelgeOpenError).detail).toBe("xml");
  });

  it("isAbortError yalnız AbortError'u tanır", () => {
    expect(isAbortError(new DOMException("x", "AbortError"))).toBe(true);
    expect(isAbortError(new Error("x"))).toBe(false);
    expect(isAbortError(null)).toBe(false);
  });
});
