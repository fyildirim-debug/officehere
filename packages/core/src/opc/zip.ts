// ZIP kapsayıcısı: merkez dizin okuyucu ve yazıcı.
// OPC paketlerinin taşınan parçaları yeniden sıkıştırılmadan (ham) kopyalanır; CRC ve boyutlar merkez
// dizinden alınır. Açma ve sıkıştırma tarayıcının (ve Node'un) yerel DecompressionStream /
// CompressionStream akışlarıyla yapılır; bağımlılık yoktur. Bozuk ya da kötü niyetli arşivde
// (taşan uzaklık, döngü, sıkıştırma bombası) çökmez: girdiyi atlar ya da OpcError fırlatır.

export type OpcErrorCode = "not-zip" | "corrupt" | "too-large" | "unsupported";

/** Paket katmanının tek hata türü; `code` ile ayırt edilir. */
export class OpcError extends Error {
  readonly code: OpcErrorCode;
  constructor(code: OpcErrorCode, message: string) {
    super(message);
    this.name = "OpcError";
    this.code = code;
  }
}

export interface ZipEntry {
  /** Arşivdeki ad (UTF-8 çözülmüş, "/" ile). */
  name: string;
  /** 0: sıkıştırmasız, 8: deflate. */
  method: number;
  flags: number;
  crc32: number;
  compressedSize: number;
  /** Açık boyut (merkez dizinin bildirdiği). */
  size: number;
  /** Sıkıştırılmış verinin arşivdeki başlangıcı (yerel başlıktan sonra). */
  dataOffset: number;
  /** DOS saat ve tarih alanları. */
  time: number;
  date: number;
}

export interface ZipArchive {
  bytes: Uint8Array;
  /** Dosya girdileri, arşiv sırasıyla. Dizinler, yinelenenler (büyük/küçük harf duyarsız) ve okunamayanlar yok. */
  entries: ZipEntry[];
  /** Atlanan girdi sayısı (yerel başlık bozuk, veri taşıyor, yinelenen ad). */
  skipped: number;
}

const SIG_LOCAL = 0x04034b50;
const SIG_CENTRAL = 0x02014b50;
const SIG_EOCD = 0x06054b50;
const SIG_EOCD64 = 0x06064b50;
const SIG_EOCD64_LOC = 0x07064b50;

const utf8 = new TextDecoder("utf-8");
const utf8Enc = new TextEncoder();

export function toBytes(data: Uint8Array | ArrayBuffer): Uint8Array {
  return data instanceof Uint8Array ? data : new Uint8Array(data);
}

/** İçerik ZIP yerel başlık imzasıyla mı başlıyor? */
export function isZip(data: Uint8Array | ArrayBuffer): boolean {
  const b = toBytes(data);
  return b.length >= 22 && b[0] === 0x50 && b[1] === 0x4b && b[2] === 0x03 && b[3] === 0x04;
}

/** ZIP merkez dizinini okur. ZIP değilse ya da merkez dizin bulunamazsa OpcError fırlatır. */
export function readZip(data: Uint8Array | ArrayBuffer): ZipArchive {
  const bytes = toBytes(data);
  if (bytes.length < 22) throw new OpcError("not-zip", "ZIP değil: dosya çok kısa");
  const dv = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  let eocd = -1;
  for (let i = bytes.length - 22; i >= Math.max(0, bytes.length - 65557); i--) {
    if (dv.getUint32(i, true) === SIG_EOCD) {
      eocd = i;
      break;
    }
  }
  if (eocd < 0) throw new OpcError("not-zip", "ZIP merkez dizini bulunamadı");
  let count = dv.getUint16(eocd + 10, true);
  let cdSize = dv.getUint32(eocd + 12, true);
  let cdOffset = dv.getUint32(eocd + 16, true);
  // ZIP64: sayılar taşmışsa konumlayıcıdan 64 bitlik bitiş kaydı okunur
  if ((count === 0xffff || cdOffset === 0xffffffff || cdSize === 0xffffffff) && eocd >= 20 && dv.getUint32(eocd - 20, true) === SIG_EOCD64_LOC) {
    const at = u64(dv, eocd - 12);
    if (at + 56 <= bytes.length && dv.getUint32(at, true) === SIG_EOCD64) {
      count = u64(dv, at + 32);
      cdSize = u64(dv, at + 40);
      cdOffset = u64(dv, at + 48);
    }
  }
  if (cdOffset > bytes.length) throw new OpcError("corrupt", "ZIP merkez dizini dosyanın dışında");
  // Girdi sayısı merkez dizin boyutunu aşamaz (her girdi en az 46 bayt): döngü sınırı
  const limit = Math.min(count, Math.floor((bytes.length - cdOffset) / 46));
  const entries: ZipEntry[] = [];
  const seen = new Set<string>();
  let skipped = 0;
  let p = cdOffset;
  for (let k = 0; k < limit && p + 46 <= bytes.length; k++) {
    if (dv.getUint32(p, true) !== SIG_CENTRAL) break;
    const flags = dv.getUint16(p + 8, true);
    const method = dv.getUint16(p + 10, true);
    const time = dv.getUint16(p + 12, true);
    const date = dv.getUint16(p + 14, true);
    const crc32 = dv.getUint32(p + 16, true);
    let compressedSize = dv.getUint32(p + 20, true);
    let size = dv.getUint32(p + 24, true);
    const nameLen = dv.getUint16(p + 28, true);
    const extraLen = dv.getUint16(p + 30, true);
    const commentLen = dv.getUint16(p + 32, true);
    let local = dv.getUint32(p + 42, true);
    const nameEnd = p + 46 + nameLen;
    if (nameEnd > bytes.length) break;
    const name = utf8.decode(bytes.subarray(p + 46, nameEnd));
    // ZIP64 ek alanı: yalnız taşan alanlar sırayla bulunur
    if (size === 0xffffffff || compressedSize === 0xffffffff || local === 0xffffffff) {
      let e = nameEnd;
      const end = Math.min(bytes.length, nameEnd + extraLen);
      while (e + 4 <= end) {
        const id = dv.getUint16(e, true);
        const len = dv.getUint16(e + 2, true);
        if (id === 0x0001) {
          let q = e + 4;
          const qEnd = Math.min(end, q + len);
          const next = () => {
            const v = q + 8 <= qEnd ? u64(dv, q) : 0xffffffff;
            q += 8;
            return v;
          };
          if (size === 0xffffffff) size = next();
          if (compressedSize === 0xffffffff) compressedSize = next();
          if (local === 0xffffffff) local = next();
          break;
        }
        e += 4 + len;
      }
    }
    p = nameEnd + extraLen + commentLen;
    if (name.endsWith("/")) continue;
    const key = name.toLowerCase();
    if (seen.has(key)) {
      skipped++;
      continue;
    }
    // Yerel başlık: imza ve veri sınırı denetlenir
    if (local + 30 > bytes.length || dv.getUint32(local, true) !== SIG_LOCAL) {
      skipped++;
      continue;
    }
    const dataOffset = local + 30 + dv.getUint16(local + 26, true) + dv.getUint16(local + 28, true);
    if (dataOffset + compressedSize > bytes.length) {
      skipped++;
      continue;
    }
    seen.add(key);
    entries.push({ name, method, flags, crc32, compressedSize, size, dataOffset, time, date });
  }
  return { bytes, entries, skipped };
}

function u64(dv: DataView, at: number): number {
  return dv.getUint32(at, true) + dv.getUint32(at + 4, true) * 0x100000000;
}

/** Girdinin sıkıştırılmış (ham) verisi; kopyalanmaz, arşivin görünümüdür. */
export function rawData(bytes: Uint8Array, entry: ZipEntry): Uint8Array {
  return bytes.subarray(entry.dataOffset, entry.dataOffset + entry.compressedSize);
}

type StreamCtor = new (format: string) => TransformStream<Uint8Array, Uint8Array>;
const streamCtor = (name: "DecompressionStream" | "CompressionStream") => (globalThis as unknown as Record<string, StreamCtor | undefined>)[name];

/** Girdiyi açar. `maxSize` aşılırsa (bildirilen ya da gerçek boyut) OpcError("too-large"). */
export async function inflateEntry(bytes: Uint8Array, entry: ZipEntry, maxSize: number): Promise<Uint8Array> {
  if (entry.flags & 1) throw new OpcError("unsupported", `Şifreli ZIP girdisi: ${entry.name}`);
  if (entry.size > maxSize) throw new OpcError("too-large", `Parça çok büyük: ${entry.name} (${entry.size} bayt)`);
  const raw = rawData(bytes, entry);
  if (entry.method === 0) {
    if (raw.length > maxSize) throw new OpcError("too-large", `Parça çok büyük: ${entry.name}`);
    return raw;
  }
  if (entry.method !== 8) throw new OpcError("unsupported", `Desteklenmeyen sıkıştırma yöntemi ${entry.method}: ${entry.name}`);
  const DS = streamCtor("DecompressionStream");
  if (!DS) throw new OpcError("unsupported", "Bu ortamda ZIP açıcı (DecompressionStream) yok");
  try {
    return await pump(new DS("deflate-raw"), raw, maxSize, entry.name);
  } catch (e) {
    if (e instanceof OpcError) throw e;
    throw new OpcError("corrupt", `Parça açılamadı: ${entry.name}`);
  }
}

/** Dönüştürme akışından tek girdiyi geçirir; çıktı `limit`i aşarsa akış kesilir. */
async function pump(ts: TransformStream<Uint8Array, Uint8Array>, input: Uint8Array, limit: number, name: string): Promise<Uint8Array> {
  const writer = ts.writable.getWriter();
  // Yazma hataları okuma tarafında görünür; burada yutulur (işlenmemiş ret kalmasın)
  writer.write(input).then(() => writer.close()).catch(() => undefined);
  const reader = ts.readable.getReader();
  const chunks: Uint8Array[] = [];
  let n = 0;
  for (;;) {
    const { value, done } = await reader.read();
    if (done) break;
    n += value.length;
    if (n > limit) {
      await reader.cancel().catch(() => undefined);
      throw new OpcError("too-large", `Parça açılırken sınır aşıldı: ${name}`);
    }
    chunks.push(value);
  }
  return concat(chunks, n);
}

function concat(chunks: Uint8Array[], n: number): Uint8Array {
  if (chunks.length === 1) return chunks[0];
  const out = new Uint8Array(n);
  let o = 0;
  for (const c of chunks) {
    out.set(c, o);
    o += c.length;
  }
  return out;
}

let crcTable: Uint32Array | null = null;

/** CRC-32 (ZIP, IEEE 802.3). */
export function crc32(data: Uint8Array): number {
  if (!crcTable) {
    crcTable = new Uint32Array(256);
    for (let i = 0; i < 256; i++) {
      let c = i;
      for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
      crcTable[i] = c >>> 0;
    }
  }
  const t = crcTable;
  let c = 0xffffffff;
  for (let i = 0; i < data.length; i++) c = t[(c ^ data[i]) & 0xff] ^ (c >>> 8);
  return (c ^ 0xffffffff) >>> 0;
}

/** Sabit DOS zamanı (1980-01-01 00:00): yeni girdilerde belirlenimci çıktı için. */
export const DOS_EPOCH = { time: 0, date: 0x21 };

interface PendingItem {
  name: string;
  /** Ham kopya: kaynak arşiv ve girdisi. */
  raw?: { bytes: Uint8Array; entry: ZipEntry };
  data?: Uint8Array;
  compress: boolean;
  time: number;
  date: number;
}

interface ReadyItem {
  name: Uint8Array;
  utf8Name: boolean;
  method: number;
  flags: number;
  crc: number;
  size: number;
  body: Uint8Array;
  time: number;
  date: number;
}

/**
 * ZIP yazıcı. Girdiler eklendiği sırayla yazılır; `addRaw` ile eklenen girdi yeniden sıkıştırılmaz.
 * Aynı ad (büyük/küçük harf duyarsız) ikinci kez eklenirse öncekinin yerine geçer, sırası korunur.
 */
export class ZipWriter {
  private items: PendingItem[] = [];
  private index = new Map<string, number>();
  private stamp: { time: number; date: number };

  constructor(opts: { time?: number; date?: number } = {}) {
    this.stamp = { time: opts.time ?? DOS_EPOCH.time, date: opts.date ?? DOS_EPOCH.date };
  }

  has(name: string): boolean {
    return this.index.has(name.toLowerCase());
  }

  get size(): number {
    return this.items.length;
  }

  /** Başka bir arşivin girdisini ham (sıkıştırılmış hâliyle) ekler. */
  addRaw(name: string, bytes: Uint8Array, entry: ZipEntry): void {
    this.put({ name, raw: { bytes, entry }, compress: false, time: entry.time, date: entry.date });
  }

  /** Yeni içerik ekler; varsayılan olarak deflate ile sıkıştırılır. */
  add(name: string, data: Uint8Array | string, opts: { compress?: boolean } = {}): void {
    const d = typeof data === "string" ? utf8Enc.encode(data) : data;
    this.put({ name, data: d, compress: opts.compress ?? true, time: this.stamp.time, date: this.stamp.date });
  }

  remove(name: string): void {
    const i = this.index.get(name.toLowerCase());
    if (i === undefined) return;
    this.items.splice(i, 1);
    this.reindex();
  }

  private put(item: PendingItem): void {
    const key = item.name.toLowerCase();
    const i = this.index.get(key);
    if (i !== undefined) this.items[i] = item;
    else {
      this.index.set(key, this.items.length);
      this.items.push(item);
    }
  }

  private reindex(): void {
    this.index.clear();
    this.items.forEach((it, i) => this.index.set(it.name.toLowerCase(), i));
  }

  /** Arşivi kurar. 65535'ten çok girdi ya da 4 GB'tan büyük çıktı OpcError("too-large"). */
  async finish(): Promise<Uint8Array> {
    if (this.items.length > 0xffff) throw new OpcError("too-large", "ZIP girdisi sayısı sınırı aşıldı");
    const CS = streamCtor("CompressionStream");
    const ready = await Promise.all(
      this.items.map(async (it): Promise<ReadyItem> => {
        const name = utf8Enc.encode(it.name);
        const utf8Name = name.some((b) => b > 0x7f);
        if (it.raw) {
          const e = it.raw.entry;
          // Bit 3 (veri tanımlayıcısı) kaldırılır: CRC ve boyutlar başlığa yazılır
          return { name, utf8Name, method: e.method, flags: e.flags & 0x0006, crc: e.crc32, size: e.size, body: rawData(it.raw.bytes, e), time: it.time, date: it.date };
        }
        const data = it.data!;
        const crc = crc32(data);
        if (it.compress && CS && data.length > 64) {
          try {
            const z = await pump(new CS("deflate-raw"), data, Number.MAX_SAFE_INTEGER, it.name);
            if (z.length < data.length) return { name, utf8Name, method: 8, flags: 0, crc, size: data.length, body: z, time: it.time, date: it.date };
          } catch {
            // sıkıştırılamazsa sıkıştırmasız yazılır
          }
        }
        return { name, utf8Name, method: 0, flags: 0, crc, size: data.length, body: data, time: it.time, date: it.date };
      }),
    );
    let total = 22;
    for (const r of ready) total += 30 + r.name.length + r.body.length + 46 + r.name.length;
    if (total > 0xffffffff) throw new OpcError("too-large", "ZIP çıktısı 4 GB sınırını aşıyor");
    const out = new Uint8Array(total);
    const dv = new DataView(out.buffer);
    const offsets: number[] = [];
    let p = 0;
    for (const r of ready) {
      offsets.push(p);
      dv.setUint32(p, SIG_LOCAL, true);
      dv.setUint16(p + 4, r.method === 8 ? 20 : 10, true);
      dv.setUint16(p + 6, r.flags | (r.utf8Name ? 0x0800 : 0), true);
      dv.setUint16(p + 8, r.method, true);
      dv.setUint16(p + 10, r.time, true);
      dv.setUint16(p + 12, r.date, true);
      dv.setUint32(p + 14, r.crc, true);
      dv.setUint32(p + 18, r.body.length, true);
      dv.setUint32(p + 22, r.size, true);
      dv.setUint16(p + 26, r.name.length, true);
      dv.setUint16(p + 28, 0, true);
      out.set(r.name, p + 30);
      out.set(r.body, p + 30 + r.name.length);
      p += 30 + r.name.length + r.body.length;
    }
    const cdStart = p;
    ready.forEach((r, i) => {
      dv.setUint32(p, SIG_CENTRAL, true);
      dv.setUint16(p + 4, 20, true);
      dv.setUint16(p + 6, r.method === 8 ? 20 : 10, true);
      dv.setUint16(p + 8, r.flags | (r.utf8Name ? 0x0800 : 0), true);
      dv.setUint16(p + 10, r.method, true);
      dv.setUint16(p + 12, r.time, true);
      dv.setUint16(p + 14, r.date, true);
      dv.setUint32(p + 16, r.crc, true);
      dv.setUint32(p + 20, r.body.length, true);
      dv.setUint32(p + 24, r.size, true);
      dv.setUint16(p + 28, r.name.length, true);
      // ek alan, yorum, disk, iç/dış öznitelikler: 0
      dv.setUint32(p + 42, offsets[i], true);
      out.set(r.name, p + 46);
      p += 46 + r.name.length;
    });
    dv.setUint32(p, SIG_EOCD, true);
    dv.setUint16(p + 8, ready.length, true);
    dv.setUint16(p + 10, ready.length, true);
    dv.setUint32(p + 12, p - cdStart, true);
    dv.setUint32(p + 16, cdStart, true);
    return out;
  }
}
