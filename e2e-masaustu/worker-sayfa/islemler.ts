// Duman testinin Worker işleyicileri. Çekirdeğin tüm dizini içe aktarılır: Worker'da yüklenemeyen
// (DOM'a üst düzeyde dokunan) bir modül çekirdeğe eklenirse bu test yakalar.
import { SaveError, withTransfer, type IoContext } from "../../packages/core/src/index";
import { FormatError } from "./ortak";

/** Her Worker örneğinin kimliği: iptalden sonra yeni Worker kurulduğu buradan anlaşılır. */
const INSTANCE = Math.random().toString(36).slice(2);

export const handlers = {
  ortam: () => ({
    instance: INSTANCE,
    protocol: location.protocol,
    domParser: typeof DOMParser,
    compression: typeof CompressionStream,
  }),
  ping: () => INSTANCE,
  topla: (a: { bytes: ArrayBuffer }) => {
    const v = new Uint8Array(a.bytes);
    let sum = 0;
    for (let i = 0; i < v.length; i++) sum = (sum + v[i]) >>> 0;
    return withTransfer({ bytes: a.bytes, sum }, [a.bytes]);
  },
  /** Ayrıştırıcı gibi eşzamanlı döngü: ileti dinleyemez, yalnız terminate durdurur. */
  mesgul: (a: { ms: number }, ctx: IoContext) => {
    const start = performance.now();
    for (let t = 0; t < a.ms; t = performance.now() - start) ctx.progress(t / a.ms, t < a.ms / 2 ? "Okunuyor" : "Kuruluyor");
    return "bitti";
  },
  bozuk: () => {
    throw new FormatError("Dosya bozuk ya da beklenmeyen bir yapıda; okunamadı.", "zip: end of central directory not found");
  },
  dolu: () => {
    throw new SaveError("ENOSPC");
  },
};
