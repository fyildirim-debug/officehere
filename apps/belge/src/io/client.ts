// Belge G/Ç istemcileri: ayrıştırma ve yazma Worker'da (yoksa süreç içinde) yapılır; ana iş parçacığı donmaz.
// Açma ve yazma aynı uzun ömürlü Worker'dadır (açılan belgenin kaynak paketi orada kayıtlıdır). Otomatik kayıt
// doğrulaması ayrı bir Worker'da: bir çağrının iptali aynı istemcideki diğer çağrıları da düşürür.
import { createIoWorker, inProcessWorker, IoWorkerError, workerOrInProcess, type IoCallOptions } from "@officehere/core";
import type { BelgeDocument } from "../model/document";
import { BelgeOpenError } from "./warnings";
import type { OpenReply } from "./ioHandlers";

type Handlers = typeof import("./ioHandlers").handlers;

const loadHandlers = () => import("./ioHandlers").then((m) => m.handlers);
/** Worker yüklenemediyse (eski WebView, kısıtlı ortam) işlemler süreç içinde sürer. */
let workerBroken = false;

function makeClient() {
  const inProcess = inProcessWorker(loadHandlers);
  return createIoWorker<Handlers>(
    workerOrInProcess(() => (workerBroken ? inProcess() : new Worker(new URL("./io.worker.ts", import.meta.url), { type: "module" })), loadHandlers),
    { errors: { FormatError: BelgeOpenError } },
  );
}

/** Açma, yazma ve kaynak paketler. */
export const io = makeClient();
/** Otomatik kaydedilen dosyanın yeniden açılabilirlik doğrulaması. */
export const verifier = makeClient();

/**
 * Worker yüklenemezse çağrı süreç içinde yinelenir. `interrupted`: başka bir çağrının iptali yüzünden yarıda
 * kalan çağrı da bir kez yinelenir (yeni Worker'da).
 */
async function retried<T>(run: () => Promise<T>, interrupted = true): Promise<T> {
  try {
    return await run();
  } catch (err) {
    if (!(err instanceof IoWorkerError)) throw err;
    if (err.code === "load" && !workerBroken) {
      workerBroken = true;
      if (typeof console !== "undefined") console.warn("[belge] dosya işleme Worker'ı yüklenemedi; ana iş parçacığında sürüyor:", err.detail ?? "");
      return run();
    }
    if (interrupted && err.code === "interrupted") return run();
    throw err;
  }
}

/** Dosyayı Worker'da okur (baytlar kopyalanır; çağıran kendi kopyasını korur). */
export function readInWorker(name: string, data: ArrayBuffer, opts: Pick<IoCallOptions, "signal" | "onProgress"> = {}): Promise<OpenReply> {
  return retried(() => io.call("open", { bytes: data, name }, opts), false);
}

/**
 * Belgeyi .docx olarak yazar. Worker'da olmayan kaynak paketler `provide` ile istenir (kimlik → baytlar; yoksa
 * null: geçiş katmanı olmadan yazılır ve bildirim döner).
 */
export async function writeDocxInWorker(doc: BelgeDocument, macro: boolean, provide: (ids: string[]) => Promise<Record<string, Uint8Array | null>>): Promise<{ bytes: Uint8Array; notes: string[] }> {
  return retried(async () => {
    let r = await io.call("write", { doc, macro });
    if (r.missing) r = await io.call("write", { doc, macro, sources: await provide(r.missing) });
    if (r.missing) throw new Error(`Kaynak paket bulunamadı: ${r.missing.join(", ")}`);
    return r;
  });
}

/** .odt, .rtf ve .md yazıcıları. */
export function encodeInWorker(format: "odt" | "rtf" | "md", doc: BelgeDocument): Promise<Uint8Array | string> {
  return retried(() => io.call("encode", { format, doc }));
}

/** Otomatik kaydedilen .docx yeniden açılabiliyor mu (ayrı Worker). */
export function verifyInWorker(bytes: ArrayBuffer): Promise<boolean> {
  return retried(() => verifier.call("verify", { bytes }));
}

/** Kapanan belgenin özgün paketi Worker'dan da bırakılır (Worker yoksa ya da kapanmışsa önemsiz). */
export function releaseInWorker(id: string): void {
  void io.call("release", { id }).catch(() => undefined);
}
