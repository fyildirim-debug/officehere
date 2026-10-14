/**
 * Açma hatalarının sınıflandırılması: kitaplıklardan gelen ham teknik iletiler
 * ("Corrupted zip…", "Cannot read properties of undefined…", "[xmldom error]",
 * "3:15: undefined entity" …) kullanıcıya anlaşılır Türkçe iletiye çevrilir;
 * özgün ileti ayrıntı olarak (FormatError.detail) korunur.
 */
import { FormatError } from "./bookUtil";

export const OPEN_MESSAGES = {
  corruptZip: "Dosya bozuk ya da eksik indirilmiş: sıkıştırılmış paket okunamadı. Dosyayı yeniden indirip ya da kaynağından yeniden kaydedip deneyin.",
  corruptXml: "Dosya bozuk: içindeki veriler (XML) okunamadı. Dosya yarıda kesilmiş ya da değiştirilmiş olabilir.",
  corrupt: "Dosya bozuk ya da beklenmeyen bir yapıda; okunamadı.",
  tooLarge: "Dosya bu ortamda açılamayacak kadar büyük.",
  encrypted: "Dosya parola ile korunuyor. Parola korumalı dosyalar açılamaz; dosyayı parolasız kaydedip yeniden deneyin.",
  tooDeep: "Dosyada aşırı iç içe geçmiş bir yapı var; okunamadı.",
  unsupported: "Dosya desteklenmeyen bir sıkıştırma ya da kodlama kullanıyor.",
};

/** Ham hatanın sınıfı. */
export function classifyError(e: unknown): keyof typeof OPEN_MESSAGES {
  const err = e as Error | undefined;
  const m = String(err?.message ?? e ?? "");
  const name = String(err?.name ?? "");
  if (/encrypt|password|parola/i.test(m)) return "encrypted";
  if (name === "RangeError" && /call stack/i.test(m)) return "tooDeep";
  if (/invalid string length|array buffer allocation|out of memory|allocation failed|invalid array length|invalid typed array length/i.test(m)) return "tooLarge";
  if (/compression method|unsupported (compression|encoding)/i.test(m)) return "unsupported";
  if (
    /zip|central directory|end of data|signature|uncompressed data size|missing \d+ bytes|crc32|invalid (stored block|block type|distance|code|literal)|incorrect header|too many length|unknown compression|data error|corrupted/i.test(
      m,
    )
  )
    return "corruptZip";
  if (/xmldom|^\d+:\d+:|unclosed|unmatched|unexpected close|closing tag|root|entity|attribute|xml|character|text data outside|disallowed|unexpected end|malformed|parse/i.test(m)) return "corruptXml";
  return "corrupt";
}

/** Herhangi bir açma hatasını anlaşılır FormatError'a çevirir (zaten anlaşılırsa olduğu gibi bırakır). */
export function toOpenError(e: unknown): FormatError {
  if (e instanceof FormatError) return e;
  const err = e as Error | undefined;
  if (err?.name === "FormatError") return new FormatError(err.message, (err as FormatError).detail);
  const detail = `${err?.name ?? "Error"}: ${String(err?.message ?? e ?? "")}`.slice(0, 2000);
  return new FormatError(OPEN_MESSAGES[classifyError(e)], detail);
}
