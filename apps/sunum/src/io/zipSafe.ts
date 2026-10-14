/**
 * Sunu paketleri (PPTX/ODP ZIP) için korumalı okuma: açılmış boyut tavanları
 * (zip bombası), bozuk parça sınıflandırması ve Türkçe iletiler. Okuyucular
 * JSZip'e doğrudan değil bu katmana başvurur; böylece bozuk bir parça tüm
 * dosyayı ham bir teknik hatayla düşürmez, çağıran uyarıyla devam edebilir.
 */
import JSZip from "jszip";
import { describeZipError, PackageError } from "./errors";

export { describeZipError, PackageError };

const MB = 1024 * 1024;

/** Paket sınırları. Gerçek sunularda XML parçaları birkaç MB'ı nadiren aşar. */
export const ZIP_LIMITS = {
  /** Tek bir XML parçasının açılmış boyutu. */
  xml: 128 * MB,
  /** Tek bir ortam dosyası (resim/video/ses); veri adresine (base64) çevrilebilmesi için. */
  media: 300 * MB,
  /** Paketteki tüm parçaların bildirilen açılmış toplamı. */
  total: 2048 * MB,
  /** Paketteki girdi sayısı. */
  entries: 50_000,
};

/** Tek bir parça okunamadı: bozuk ya da sınırı aşıyor. Çağıran uyarıyla devam edebilir. */
export class PartError extends Error {
  kind: "corrupt" | "tooLarge";
  path: string;
  detail?: string;
  constructor(kind: "corrupt" | "tooLarge", path: string, message: string, detail?: string) {
    super(message);
    this.name = "PartError";
    this.kind = kind;
    this.path = path;
    this.detail = detail;
  }
}

export const formatMB = (n: number) => `${Math.max(1, Math.round(n / MB)).toLocaleString("tr-TR")} MB`;

/** Bildirilen açılmış boyut (merkezi dizinden; okunmadan önce). */
function declaredSize(f: JSZip.JSZipObject): number | undefined {
  const n = (f as unknown as { _data?: { uncompressedSize?: number } })._data?.uncompressedSize;
  return typeof n === "number" && n >= 0 ? n : undefined;
}

interface StreamHelper {
  on(ev: "data", fn: (chunk: Uint8Array) => void): StreamHelper;
  on(ev: "end", fn: () => void): StreamHelper;
  on(ev: "error", fn: (e: Error) => void): StreamHelper;
  pause(): StreamHelper;
  resume(): StreamHelper;
}

/**
 * Parçayı akışla açar; açılan bayt sayısı sınırı aşınca durur (merkezi dizinde
 * küçük boyut bildirip dev açılan sahte parçalara karşı da korur).
 */
function readLimited(f: JSZip.JSZipObject, limit: number): Promise<Uint8Array> {
  return new Promise((resolve, reject) => {
    const chunks: Uint8Array[] = [];
    let total = 0;
    let done = false;
    let h: StreamHelper;
    try {
      h = (f as unknown as { internalStream(t: string): StreamHelper }).internalStream("uint8array");
    } catch (e) {
      reject(e);
      return;
    }
    h.on("data", (chunk) => {
      if (done) return;
      total += chunk.length;
      if (total > limit) {
        done = true;
        h.pause();
        reject(new RangeError(`limit:${total}`));
        return;
      }
      chunks.push(chunk);
    })
      .on("error", (e) => {
        if (done) return;
        done = true;
        reject(e);
      })
      .on("end", () => {
        if (done) return;
        done = true;
        if (chunks.length === 1) return resolve(chunks[0]);
        const out = new Uint8Array(total);
        let o = 0;
        for (const c of chunks) {
          out.set(c, o);
          o += c.length;
        }
        resolve(out);
      })
      .resume();
  });
}

const decoder = () => new TextDecoder("utf-8");

export interface SafeZip {
  zip: JSZip;
  /** Paketteki dosya yolları (klasörler hariç). */
  paths: string[];
  has(path: string): boolean;
  /** XML/metin parçası; yoksa undefined. Bozuksa ya da sınırı aşıyorsa PartError. */
  text(path: string, limit?: number): Promise<string | undefined>;
  /** İkili parça; yoksa undefined. Bozuksa ya da sınırı aşıyorsa PartError. */
  bytes(path: string, limit?: number): Promise<Uint8Array | undefined>;
  /** Bildirilen açılmış boyut (bilinmiyorsa undefined). */
  size(path: string): number | undefined;
}

/**
 * ZIP paketini açar ve genel sınırları denetler. Okunamayan/zararlı paketlerde
 * Türkçe iletili PackageError atar (ayrıntı: özgün teknik ileti).
 */
export async function openZip(data: ArrayBuffer | Uint8Array, badZipMessage = "Dosya bozuk; sunu paketi okunamadı."): Promise<SafeZip> {
  let zip: JSZip;
  try {
    zip = await JSZip.loadAsync(data);
  } catch (e) {
    const m = describeZipError(e);
    throw new PackageError(m === "Dosya bozuk; sunu paketi okunamadı." ? badZipMessage : m, String((e as Error)?.message ?? e));
  }
  const paths = Object.keys(zip.files).filter((p) => !zip.files[p].dir);
  if (paths.length > ZIP_LIMITS.entries) throw new PackageError(`Dosya açılamadı: pakette olağandışı sayıda parça var (${paths.length.toLocaleString("tr-TR")}). Dosya bozuk ya da zararlı olabilir.`, `entries=${paths.length}`);
  let total = 0;
  for (const p of paths) total += declaredSize(zip.files[p]) ?? 0;
  if (total > ZIP_LIMITS.total) throw new PackageError(`Dosya açılamadı: paketin açılmış boyutu çok büyük (${formatMB(total)}). Dosya bozuk ya da zararlı olabilir.`, `uncompressed=${total}`);

  const read = async (path: string, limit: number): Promise<Uint8Array | undefined> => {
    const f = zip.file(path);
    if (!f) return undefined;
    const declared = declaredSize(f);
    if (declared !== undefined && declared > limit) throw new PartError("tooLarge", path, `“${path}” parçası çok büyük (${formatMB(declared)}; sınır ${formatMB(limit)}).`, `declared=${declared}`);
    try {
      return await readLimited(f, limit);
    } catch (e) {
      if (e instanceof RangeError && e.message.startsWith("limit:")) throw new PartError("tooLarge", path, `“${path}” parçası çok büyük (sınır ${formatMB(limit)}).`, e.message);
      throw new PartError("corrupt", path, `“${path}” parçası bozuk; okunamadı.`, String((e as Error)?.message ?? e));
    }
  };
  return {
    zip,
    paths,
    has: (p) => !!zip.file(p),
    size: (p) => {
      const f = zip.file(p);
      return f ? declaredSize(f) : undefined;
    },
    bytes: (p, limit = ZIP_LIMITS.media) => read(p, limit),
    text: async (p, limit = ZIP_LIMITS.xml) => {
      const b = await read(p, limit);
      return b === undefined ? undefined : decoder().decode(b);
    },
  };
}

/**
 * Ayrıştırılmış XML belgesinin durumu: kök beklenen öğe mi, ayrıştırma hatası
 * var mı. Tarayıcılar bozuk XML'de hata öğesi ekleyip okunabilen kısmı korur.
 */
export function xmlHealth(doc: Document | undefined, rootName: string | string[]): "ok" | "partial" | "bad" {
  const root = doc?.documentElement;
  if (!root) return "bad";
  const names = Array.isArray(rootName) ? rootName : [rootName];
  // Firefox/jsdom: kök hata öğesidir; Chromium/WebKit: hata öğesi kökün ilk çocuğu olarak eklenir (O(1) denetim).
  const first = root.firstElementChild;
  const hasError = root.localName === "parsererror" || first?.localName === "parsererror" || root.lastElementChild?.localName === "parsererror";
  if (!names.includes(root.localName)) return "bad";
  return hasError ? "partial" : "ok";
}
