/**
 * BIFF8/BIFF5 kayıt akışı (MS-XLS §2.1.4): kayıt başlığı (tür, uzunluk),
 * CONTINUE (0x003C) parçaları ve kayıt sınırında bölünen dizgiler.
 * Bozuk dosyalarda sınır denetimi yapılır; okuma dışına taşılmaz.
 */

export const R = {
  FORMULA: 0x0006,
  EOF: 0x000a,
  CALCMODE: 0x000d,
  DATEMODE: 0x0022,
  EXTERNNAME: 0x0023,
  FILEPASS: 0x002f,
  FONT: 0x0031,
  CONTINUE: 0x003c,
  WINDOW1: 0x003d,
  CODEPAGE: 0x0042,
  PANE: 0x0041,
  DEFCOLWIDTH: 0x0055,
  EXTERNSHEET: 0x0017,
  NAME: 0x0018,
  COLINFO: 0x007d,
  BOUNDSHEET: 0x0085,
  PALETTE: 0x0092,
  SCL: 0x00a0,
  MULRK: 0x00bd,
  MULBLANK: 0x00be,
  RSTRING: 0x00d6,
  XF: 0x00e0,
  MERGEDCELLS: 0x00e5,
  SST: 0x00fc,
  LABELSST: 0x00fd,
  SUPBOOK: 0x01ae,
  HLINK: 0x01b8,
  DIMENSIONS: 0x0200,
  BLANK: 0x0201,
  NUMBER: 0x0203,
  LABEL: 0x0204,
  BOOLERR: 0x0205,
  STRING: 0x0207,
  ROW: 0x0208,
  ARRAY: 0x0221,
  DEFAULTROWHEIGHT: 0x0225,
  WINDOW2: 0x023e,
  RK: 0x027e,
  FORMAT: 0x041e,
  SHRFMLA: 0x04bc,
  BOF: 0x0809,
  BOF_OLD: 0x0409,
} as const;

export interface BiffRecord {
  type: number;
  /** Akıştaki konum (BOUNDSHEET lbPlyPos ile eşleşir). */
  offset: number;
  data: Uint8Array;
  /** Ardından gelen CONTINUE kayıtlarının verileri. */
  cont: Uint8Array[];
}

/** Bu kayıtlardan sonra gelen CONTINUE'lar kayda eklenir. */
const CONTINUABLE = new Set<number>([R.SST, R.STRING, R.FORMULA, R.NAME, R.SHRFMLA, R.ARRAY, R.EXTERNNAME, R.EXTERNSHEET, R.LABEL]);

/** Kayıt verisi + CONTINUE parçaları tek dizide (dizgi bayrağı içermeyen kayıtlar için). */
export function joined(rec: BiffRecord): Uint8Array {
  if (!rec.cont.length) return rec.data;
  const out = new Uint8Array(rec.data.length + rec.cont.reduce((n, c) => n + c.length, 0));
  out.set(rec.data, 0);
  let p = rec.data.length;
  for (const c of rec.cont) {
    out.set(c, p);
    p += c.length;
  }
  return out;
}

/** Akışı kayıtlara böler. */
export function readRecords(stream: Uint8Array): BiffRecord[] {
  const out: BiffRecord[] = [];
  let p = 0;
  const n = stream.length;
  while (p + 4 <= n) {
    const type = stream[p] | (stream[p + 1] << 8);
    const len = stream[p + 2] | (stream[p + 3] << 8);
    const start = p + 4;
    const end = Math.min(n, start + len);
    const data = stream.subarray(start, end);
    if (type === R.CONTINUE && out.length && CONTINUABLE.has(out[out.length - 1].type)) out[out.length - 1].cont.push(data);
    else out.push({ type, offset: p, data, cont: [] });
    p = start + len;
    // Akışın sonundaki dolgu sıfırları
    if (type === 0 && len === 0) break;
  }
  return out;
}

/** Küçük uçlu okuma yardımcıları (sınır dışında 0 döner). */
export function u8(d: Uint8Array, o: number): number {
  return o < d.length ? d[o] : 0;
}
export function u16(d: Uint8Array, o: number): number {
  return o + 1 < d.length ? d[o] | (d[o + 1] << 8) : 0;
}
export function i16(d: Uint8Array, o: number): number {
  const v = u16(d, o);
  return v & 0x8000 ? v - 0x10000 : v;
}
export function u32(d: Uint8Array, o: number): number {
  return o + 3 < d.length ? (d[o] | (d[o + 1] << 8) | (d[o + 2] << 16)) + d[o + 3] * 0x1000000 : 0;
}
export function f64(d: Uint8Array, o: number): number {
  if (o + 8 > d.length) return 0;
  return new DataView(d.buffer, d.byteOffset + o, 8).getFloat64(0, true);
}

/** RK sayısı (MS-XLS §2.5.217): 30 bit tam sayı ya da IEEE kesik. */
export function rk(v: number): number {
  const div100 = v & 1;
  let n: number;
  if (v & 2) {
    n = (v | 0) >> 2;
  } else {
    const buf = new DataView(new ArrayBuffer(8));
    buf.setUint32(4, v & 0xfffffffc, true);
    buf.setUint32(0, 0, true);
    n = buf.getFloat64(0, true);
  }
  return div100 ? n / 100 : n;
}

/** Tek baytlık dizgiler için kod sayfası çözücüsü (BIFF5 ve 8 bit karakterler). */
export function byteDecoder(codepage: number): (b: Uint8Array) => string {
  const label = codepage === 1200 ? "utf-16le" : codepage >= 874 && codepage <= 1258 ? `windows-${codepage}` : codepage === 65001 ? "utf-8" : "windows-1252";
  let dec: TextDecoder;
  try {
    dec = new TextDecoder(label);
  } catch {
    dec = new TextDecoder("windows-1252");
  }
  return (b) => dec.decode(b);
}

/** BIFF8'de sıkıştırılmış (8 bit) karakterler: Unicode'un alt baytı (Latin-1). */
function latin1(d: Uint8Array, start: number, count: number): string {
  let s = "";
  const end = Math.min(d.length, start + count);
  for (let i = start; i < end; i += 0x2000) s += String.fromCharCode(...d.subarray(i, Math.min(end, i + 0x2000)));
  return s;
}

function utf16(d: Uint8Array, start: number, count: number): string {
  let s = "";
  const end = Math.min(d.length, start + count * 2);
  const codes: number[] = [];
  for (let i = start; i + 1 < end; i += 2) {
    codes.push(d[i] | (d[i + 1] << 8));
    if (codes.length >= 0x2000) {
      s += String.fromCharCode(...codes);
      codes.length = 0;
    }
  }
  return s + String.fromCharCode(...codes);
}

/**
 * Kayıt ve CONTINUE parçaları üzerinde sıralı okuyucu. Dizginin karakterleri
 * parça sınırına gelirse yeni parçanın ilk baytı yeniden seçenek baytıdır
 * (fHighByte); başlık, zengin metin ve uzantı verileri ise düz devam eder.
 */
export class ContinuedReader {
  private seg = 0;
  pos: number;
  private segs: Uint8Array[];

  constructor(rec: BiffRecord, start = 0) {
    this.segs = [rec.data, ...rec.cont];
    this.pos = start;
  }

  get eof(): boolean {
    return this.seg >= this.segs.length || (this.seg === this.segs.length - 1 && this.pos >= this.segs[this.seg].length);
  }

  private cur(): Uint8Array {
    while (this.seg < this.segs.length && this.pos >= this.segs[this.seg].length) {
      this.seg++;
      this.pos = 0;
    }
    return this.segs[this.seg] ?? new Uint8Array(0);
  }

  byte(): number {
    const d = this.cur();
    if (this.pos >= d.length) return 0;
    return d[this.pos++];
  }

  u16(): number {
    return this.byte() | (this.byte() << 8);
  }

  u32(): number {
    return (this.u16() + this.u16() * 0x10000) >>> 0;
  }

  skip(n: number): void {
    let left = n;
    while (left > 0 && !this.eof) {
      const d = this.cur();
      const take = Math.min(left, d.length - this.pos);
      this.pos += take;
      left -= take;
      if (take === 0) break;
    }
  }

  /** XLUnicodeRichExtendedString (SST öğesi) ya da XLUnicodeString. */
  string(cchBytes: 1 | 2 = 2): string {
    const cch = cchBytes === 2 ? this.u16() : this.byte();
    let flags = this.byte();
    const rich = flags & 0x08 ? this.u16() : 0;
    const ext = flags & 0x04 ? this.u32() : 0;
    let s = "";
    let left = cch;
    let guard = 0;
    while (left > 0 && guard++ < 100000) {
      const d = this.segs[this.seg];
      if (!d) break;
      if (this.pos >= d.length) {
        // Parça bitti: karakterler sonraki parçada, ilk bayt yeni seçenek baytı
        this.seg++;
        const next = this.segs[this.seg];
        if (!next?.length) break;
        flags = (flags & ~0x01) | (next[0] & 0x01);
        this.pos = 1;
        continue;
      }
      const high = flags & 0x01;
      const avail = high ? Math.floor((d.length - this.pos) / 2) : d.length - this.pos;
      const take = Math.min(left, avail);
      if (take <= 0) {
        // Tek bayt artığı (bozuk bölünme): atla
        this.pos = d.length;
        continue;
      }
      s += high ? utf16(d, this.pos, take) : latin1(d, this.pos, take);
      this.pos += high ? take * 2 : take;
      left -= take;
    }
    if (rich) this.skip(rich * 4);
    if (ext) this.skip(ext);
    return s;
  }
}

/** Tek kayıt içindeki XLUnicodeString (CONTINUE yok). */
export function readString(d: Uint8Array, o: number, cchBytes: 1 | 2): { text: string; end: number } {
  const cch = cchBytes === 2 ? u16(d, o) : u8(d, o);
  const flags = u8(d, o + cchBytes);
  let p = o + cchBytes + 1;
  const rich = flags & 0x08 ? u16(d, p) : 0;
  if (flags & 0x08) p += 2;
  const ext = flags & 0x04 ? u32(d, p) : 0;
  if (flags & 0x04) p += 4;
  const high = flags & 0x01;
  const text = high ? utf16(d, p, cch) : latin1(d, p, cch);
  p += high ? cch * 2 : cch;
  return { text, end: p + rich * 4 + ext };
}

/** BIFF5 bayt dizgisi (uzunluk + baytlar). */
export function readByteString(d: Uint8Array, o: number, cchBytes: 1 | 2, decode: (b: Uint8Array) => string): { text: string; end: number } {
  const cch = cchBytes === 2 ? u16(d, o) : u8(d, o);
  const start = o + cchBytes;
  return { text: decode(d.subarray(start, Math.min(d.length, start + cch))), end: start + cch };
}
