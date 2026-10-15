/**
 * Tek sıkıştırma: büyük parçalar (sayfa XML'i) üretilirken parça parça CompressionStream'e
 * ("deflate-raw"; yoksa "deflate" başlığı ve sağlaması kesilerek) akıtılır, CRC aynı geçişte
 * hesaplanır; sonuç ZIP'e yeniden sıkıştırılmadan (ham) eklenir. Küçük parçalar core ZipWriter'ın
 * kendi sıkıştırmasıyla yazılır. Ortamda sıkıştırıcı yoksa girdi sıkıştırmasız (stored) yazılır.
 */
import { ZipWriter, DOS_EPOCH, type ZipEntry } from "@officehere/core/opc";

/** Ham eklenecek (sıkıştırılmış ya da sıkıştırmasız) girdi. */
export interface PackedEntry {
  method: 0 | 8;
  crc: number;
  size: number;
  body: Uint8Array;
}

const CRC = (() => {
  const t = new Int32Array(256 * 8);
  for (let i = 0; i < 256; i++) {
    let c = i;
    for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
    t[i] = c;
  }
  for (let i = 0; i < 256; i++) for (let s = 1; s < 8; s++) t[s * 256 + i] = (t[(s - 1) * 256 + i] >>> 8) ^ t[t[(s - 1) * 256 + i] & 0xff];
  return t;
})();

/** CRC-32 (ZIP) sürdürme: crc32Update(crc32Update(0, a), b) = crc32(a+b). Sekizli dilimleme. */
export function crc32Update(crc: number, buf: Uint8Array): number {
  const t = CRC;
  let c = ~crc;
  let i = 0;
  const n = buf.length;
  for (; i + 8 <= n; i += 8) {
    const a = (buf[i] | (buf[i + 1] << 8) | (buf[i + 2] << 16) | (buf[i + 3] << 24)) ^ c;
    const b = buf[i + 4] | (buf[i + 5] << 8) | (buf[i + 6] << 16) | (buf[i + 7] << 24);
    c =
      t[1792 + (a & 0xff)] ^
      t[1536 + ((a >>> 8) & 0xff)] ^
      t[1280 + ((a >>> 16) & 0xff)] ^
      t[1024 + (a >>> 24)] ^
      t[768 + (b & 0xff)] ^
      t[512 + ((b >>> 8) & 0xff)] ^
      t[256 + ((b >>> 16) & 0xff)] ^
      t[b >>> 24];
  }
  for (; i < n; i++) c = t[(c ^ buf[i]) & 0xff] ^ (c >>> 8);
  return ~c >>> 0;
}

type StreamCtor = new (format: string) => TransformStream<Uint8Array, Uint8Array>;

/** Kullanılabilir sıkıştırıcı biçimi: "deflate-raw", "deflate" (zlib sarmalı) ya da yok. */
function openCompressor(): { ts: TransformStream<Uint8Array, Uint8Array>; zlib: boolean } | null {
  const CS = (globalThis as unknown as { CompressionStream?: StreamCtor }).CompressionStream;
  if (!CS) return null;
  try {
    return { ts: new CS("deflate-raw"), zlib: false };
  } catch {
    try {
      return { ts: new CS("deflate"), zlib: true };
    } catch {
      return null;
    }
  }
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

/** Sıkıştırıcıya verilip henüz işlenmemiş en çok parça sayısı (~64 KB'lık). */
const IN_FLIGHT = 8;
/** Node'un setImmediate'i (tarayıcıda yok: orada sıkıştırma akışın kendi adımında yürür). */
const immediate = (globalThis as { setImmediate?: (cb: () => void) => unknown }).setImmediate;

/** Bayt akışını sıkıştıran ve CRC'sini tutan alıcı. Yazmalar sırayla beklenmelidir. */
export class DeflateSink {
  private crc = 0;
  private size = 0;
  private stored: Uint8Array[] | null = null;
  private writer: WritableStreamDefaultWriter<Uint8Array> | null = null;
  private result: Promise<{ chunks: Uint8Array[]; n: number }> | null = null;
  private zlib = false;
  private failed: unknown = null;
  private inflight: Promise<void>[] = [];

  constructor() {
    const c = openCompressor();
    if (!c) {
      this.stored = [];
      return;
    }
    this.zlib = c.zlib;
    this.writer = c.ts.writable.getWriter();
    const reader = c.ts.readable.getReader();
    this.result = (async () => {
      const chunks: Uint8Array[] = [];
      let n = 0;
      for (;;) {
        const { value, done } = await reader.read();
        if (done) break;
        chunks.push(value);
        n += value.length;
      }
      return { chunks, n };
    })();
    // Okuma hatası finish'te görünür; burada işlenmemiş ret kalmasın
    this.result.catch((e) => (this.failed = e));
  }

  /**
   * Parçayı ekler. Sıkıştırıcıda en çok IN_FLIGHT parça bekler (bellek sınırlı); en eskisi bitince
   * devam edilir. Node'da sıkıştırma ayrı iş parçacığında yürür: her parçadan sonra olay döngüsüne
   * kısaca dönülür ki biten parçanın ardından sıradaki hemen başlasın (üretimle örtüşsün).
   */
  async write(chunk: Uint8Array): Promise<void> {
    if (!chunk.length) return;
    this.crc = crc32Update(this.crc, chunk);
    this.size += chunk.length;
    if (this.stored) {
      this.stored.push(chunk);
      return;
    }
    const p = this.writer!.write(chunk);
    p.catch((e) => (this.failed = e));
    this.inflight.push(p);
    if (this.inflight.length > IN_FLIGHT) await this.inflight.shift();
    if (immediate) await new Promise<void>((r) => immediate(r));
  }

  async finish(): Promise<PackedEntry> {
    if (this.stored) return { method: 0, crc: this.crc, size: this.size, body: concat(this.stored, this.size) };
    await this.writer!.close();
    const { chunks, n } = await this.result!;
    if (this.failed) throw this.failed;
    let body = concat(chunks, n);
    // zlib sarmalı: 2 baytlık başlık ve 4 baytlık Adler-32 atılınca ham deflate kalır
    if (this.zlib) body = body.subarray(2, body.length - 4);
    return { method: 8, crc: this.crc, size: this.size, body };
  }
}

const encoder = new TextEncoder();

/**
 * Metin parçalarını UTF-8'e çevirip alıcıya ~64 KB'lık dilimlerle akıtan yazıcı. Çağıran `put` ile
 * ekler, ara sıra `drain` ile bekler (eşzamanlı üretim sürerken sıkıştırma ilerlesin).
 */
export class ChunkedText {
  private buf = "";
  constructor(private readonly sink: DeflateSink) {}

  put(s: string): void {
    this.buf += s;
  }

  get pending(): number {
    return this.buf.length;
  }

  async drain(force = false): Promise<void> {
    if (!this.buf || (!force && this.buf.length < 32768)) return;
    const s = this.buf;
    this.buf = "";
    await this.sink.write(encoder.encode(s));
  }
}

/**
 * Bayt düzeyinde yazıcı (sayfa verisi): ASCII parçalar ve sayılar ara metin üretmeden doğrudan
 * tampona yazılır; tampon ~64 KB'ı geçince satır sonlarında alıcıya kopyası verilir.
 */
export class ByteOut {
  buf = new Uint8Array(1 << 17);
  pos = 0;
  constructor(private readonly sink: DeflateSink) {}

  private room(n: number): void {
    if (this.pos + n <= this.buf.length) return;
    const next = new Uint8Array(Math.max(this.buf.length * 2, this.pos + n));
    next.set(this.buf.subarray(0, this.pos));
    this.buf = next;
  }

  /** Yalnız ASCII karakterli metin. */
  ascii(s: string): void {
    const n = s.length;
    this.room(n);
    const b = this.buf;
    let p = this.pos;
    for (let i = 0; i < n; i++) b[p++] = s.charCodeAt(i);
    this.pos = p;
  }

  bytes(a: Uint8Array): void {
    const n = a.length;
    this.room(n);
    const b = this.buf;
    let p = this.pos;
    for (let i = 0; i < n; i++) b[p++] = a[i];
    this.pos = p;
  }

  byte(v: number): void {
    this.room(1);
    this.buf[this.pos++] = v;
  }

  /** Herhangi bir metin (UTF-8). */
  text(s: string): void {
    this.room(s.length * 3);
    this.pos += encoder.encodeInto(s, this.buf.subarray(this.pos)).written;
  }

  /** Negatif olmayan tam sayı (2^53'ten küçük). */
  int(n: number): void {
    if (n < 10) {
      this.room(1);
      this.buf[this.pos++] = 48 + n;
      return;
    }
    let d = 0;
    for (let x = n; x >= 1; x = Math.floor(x / 10)) d++;
    this.room(d);
    const b = this.buf;
    let p = this.pos + d;
    for (let x = n; p > this.pos; x = Math.floor(x / 10)) b[--p] = 48 + (x % 10);
    this.pos += d;
  }

  /** Sayı, String(v) ile aynı yazımla (tam sayılar ara metin üretmeden). */
  num(v: number): void {
    if (Number.isInteger(v) && v > -1e15 && v < 1e15) {
      if (v < 0) {
        this.byte(45);
        this.int(-v);
      } else this.int(v);
    } else this.ascii(String(v));
  }

  /** Tampon dolduysa (ya da zorla) alıcıya verir. */
  async flush(force = false): Promise<void> {
    if (!this.pos || (!force && this.pos < 65536)) return;
    const chunk = this.buf.slice(0, this.pos);
    this.pos = 0;
    await this.sink.write(chunk);
  }
}

/** ASCII metnin baytları (sabit parçalar). */
export function asciiBytes(s: string): Uint8Array {
  const out = new Uint8Array(s.length);
  for (let i = 0; i < s.length; i++) out[i] = s.charCodeAt(i);
  return out;
}

/** Tek seferlik metni sıkıştırır (küçük ve orta parçalar). */
export async function packText(text: string): Promise<PackedEntry> {
  const sink = new DeflateSink();
  const bytes = encoder.encode(text);
  for (let i = 0; i < bytes.length; i += 1 << 20) await sink.write(bytes.subarray(i, Math.min(bytes.length, i + (1 << 20))));
  return sink.finish();
}

/** ZIP'e hazır (ham) girdi olarak ekler. */
export function addPacked(zw: ZipWriter, name: string, e: PackedEntry): void {
  const entry: ZipEntry = { name, method: e.method, flags: 0, crc32: e.crc, compressedSize: e.body.length, size: e.size, dataOffset: 0, time: DOS_EPOCH.time, date: DOS_EPOCH.date };
  zw.addRaw(name, e.body, entry);
}
