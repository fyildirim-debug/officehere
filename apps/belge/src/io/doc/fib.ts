// .doc dosya bilgi bloğu (FIB, MS-DOC 2.5) ve güvenli bayt okuyucu.

/** Sınır denetimli küçük-sonlu (little-endian) okuyucu; taşmada anlaşılır hata fırlatır. */
export class Bytes {
  readonly dv: DataView;
  constructor(readonly b: Uint8Array) {
    this.dv = new DataView(b.buffer, b.byteOffset, b.byteLength);
  }
  get length() {
    return this.b.length;
  }
  private check(o: number, n: number) {
    if (!(o >= 0 && o + n <= this.b.length)) throw new Error("Bozuk .doc dosyası (veri sınırı aşıldı).");
  }
  u8(o: number) {
    this.check(o, 1);
    return this.b[o];
  }
  u16(o: number) {
    this.check(o, 2);
    return this.dv.getUint16(o, true);
  }
  i16(o: number) {
    this.check(o, 2);
    return this.dv.getInt16(o, true);
  }
  u32(o: number) {
    this.check(o, 4);
    return this.dv.getUint32(o, true);
  }
  i32(o: number) {
    this.check(o, 4);
    return this.dv.getInt32(o, true);
  }
  slice(o: number, n: number): Uint8Array {
    this.check(o, n);
    return this.b.subarray(o, o + n);
  }
  /** UTF-16LE dizgi. */
  utf16(o: number, chars: number): string {
    this.check(o, chars * 2);
    let s = "";
    for (let i = 0; i < chars; i++) s += String.fromCharCode(this.dv.getUint16(o + i * 2, true));
    return s;
  }
}

/** FibRgFcLcb97 dizinleri (fc/lcb çiftleri). */
export const FC = {
  Stshf: 1,
  PlcffndRef: 2,
  PlcffndTxt: 3,
  PlcfSed: 6,
  PlcfHdd: 11,
  PlcfBteChpx: 12,
  PlcfBtePapx: 13,
  SttbfFfn: 15,
  Clx: 33,
  PlcfendRef: 46,
  PlcfendTxt: 47,
  DggInfo: 50,
  PlfLst: 73,
  PlfLfo: 74,
} as const;

export interface Fib {
  nFib: number;
  /** 1Table akışı mı (yoksa 0Table). */
  table1: boolean;
  ccpText: number;
  ccpFtn: number;
  ccpHdd: number;
  ccpMcr: number;
  ccpAtn: number;
  ccpEdn: number;
  ccpTxbx: number;
  ccpHdrTxbx: number;
  fcLcb(index: number): { fc: number; lcb: number };
}

/** WordDocument akışının başındaki FIB'i ayrıştırır; şifreli ve 97 öncesi dosyaları anlaşılır hatayla reddeder. */
export function parseFib(wd: Bytes): Fib {
  if (wd.length < 0x9a || wd.u16(0) !== 0xa5ec) throw new Error("Geçerli bir .doc dosyası değil (belge bilgi bloğu yok).");
  const nFib = wd.u16(2);
  const flags = wd.u16(0x0a);
  if (nFib < 0x00c0) throw new Error("Bu dosya 97 öncesi eski bir .doc biçiminde (6.0/95); bu sürüm açılamıyor.");
  if (flags & 0x0100 || flags & 0x8000) throw new Error("Bu .doc belgesi parolayla şifrelenmiş; şifreli belgeler açılamıyor.");
  const csw = wd.u16(0x20);
  const lwAt = 0x22 + csw * 2 + 2;
  const cslw = wd.u16(lwAt - 2);
  const fcAt = lwAt + cslw * 4 + 2;
  const cbRgFcLcb = wd.u16(fcAt - 2);
  const lw = (i: number) => (i < cslw ? wd.i32(lwAt + i * 4) : 0);
  const pos = (n: number) => Math.max(0, n);
  return {
    nFib,
    table1: (flags & 0x0200) !== 0,
    ccpText: pos(lw(3)),
    ccpFtn: pos(lw(4)),
    ccpHdd: pos(lw(5)),
    ccpMcr: pos(lw(6)),
    ccpAtn: pos(lw(7)),
    ccpEdn: pos(lw(8)),
    ccpTxbx: pos(lw(9)),
    ccpHdrTxbx: pos(lw(10)),
    fcLcb(index: number) {
      if (index >= cbRgFcLcb) return { fc: 0, lcb: 0 };
      return { fc: wd.u32(fcAt + index * 8), lcb: wd.u32(fcAt + index * 8 + 4) };
    },
  };
}

/** PLC: (n+1) CP ve n adet sabit boyutlu veri. */
export function plc(t: Bytes, fc: number, lcb: number, dataSize: number): { cps: number[]; data: number[] } {
  if (!lcb) return { cps: [], data: [] };
  const n = Math.floor((lcb - 4) / (4 + dataSize));
  if (n < 0 || n > 1_000_000) throw new Error("Bozuk .doc dosyası (geçersiz tablo).");
  const cps: number[] = [];
  for (let i = 0; i <= n; i++) cps.push(t.i32(fc + i * 4));
  const data: number[] = [];
  for (let i = 0; i < n; i++) data.push(fc + (n + 1) * 4 + i * dataSize);
  return { cps, data };
}
