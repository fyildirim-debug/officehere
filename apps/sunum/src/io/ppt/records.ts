/**
 * MS-PPT / MS-ODRAW kayıt okuyucusu. Her kayıt 8 baytlık başlıkla başlar:
 * recVer (4 bit) + recInstance (12 bit), recType (16 bit), recLen (32 bit).
 * recVer 0xF olan kayıtlar kapsayıcıdır. Sınırlar denetlenir; bozuk kayıtta
 * okuma o dalda durur.
 */

export class PptError extends Error {}

export interface Rec {
  type: number;
  ver: number;
  inst: number;
  /** Başlık dahil başlangıç (akış içinde). */
  offset: number;
  /** Veri başlangıcı ve uzunluğu. */
  start: number;
  len: number;
  /** Kapsayıcı ise alt kayıtlar (tembel okunur). */
  kids?: Rec[];
}

/** Kayıt türleri (yalnız kullanılanlar). */
export const RT = {
  Document: 0x03e8,
  DocumentAtom: 0x03e9,
  EndDocumentAtom: 0x03ea,
  Slide: 0x03ee,
  SlideAtom: 0x03ef,
  Notes: 0x03f0,
  NotesAtom: 0x03f1,
  Environment: 0x03f2,
  SlidePersistAtom: 0x03f3,
  MainMaster: 0x03f8,
  SlideShowSlideInfoAtom: 0x03f9,
  Drawing: 0x040c,
  DrawingGroup: 0x040b,
  ColorSchemeAtom: 0x07f0,
  FontCollection: 0x07d5,
  FontEntityAtom: 0x0fb7,
  PlaceholderAtom: 0x0bc3,
  OutlineTextRefAtom: 0x0f9e,
  TextHeaderAtom: 0x0f9f,
  TextCharsAtom: 0x0fa0,
  StyleTextPropAtom: 0x0fa1,
  TextMasterStyleAtom: 0x0fa3,
  TextBytesAtom: 0x0fa8,
  CString: 0x0fba,
  SlideListWithText: 0x0ff0,
  InteractiveInfo: 0x0ff2,
  UserEditAtom: 0x0ff5,
  CurrentUserAtom: 0x0ff6,
  PersistDirectoryAtom: 0x1772,
  CryptSession10Container: 0x2f14,
  // OfficeArt
  DggContainer: 0xf000,
  BStoreContainer: 0xf001,
  DgContainer: 0xf002,
  SpgrContainer: 0xf003,
  SpContainer: 0xf004,
  FBSE: 0xf007,
  FDG: 0xf008,
  FSPGR: 0xf009,
  FSP: 0xf00a,
  FOPT: 0xf00b,
  ClientTextbox: 0xf00d,
  ChildAnchor: 0xf00f,
  ClientAnchor: 0xf010,
  ClientData: 0xf011,
  TertiaryFOPT: 0xf122,
} as const;

const MAX_DEPTH = 32;
const MAX_KIDS = 100000;

export class Stream {
  readonly dv: DataView;
  constructor(readonly bytes: Uint8Array) {
    this.dv = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  }
  get length() {
    return this.bytes.length;
  }
  ok(off: number, n: number) {
    return off >= 0 && n >= 0 && off + n <= this.bytes.length;
  }
  u8(o: number) {
    return this.ok(o, 1) ? this.dv.getUint8(o) : 0;
  }
  u16(o: number) {
    return this.ok(o, 2) ? this.dv.getUint16(o, true) : 0;
  }
  i16(o: number) {
    return this.ok(o, 2) ? this.dv.getInt16(o, true) : 0;
  }
  u32(o: number) {
    return this.ok(o, 4) ? this.dv.getUint32(o, true) : 0;
  }
  i32(o: number) {
    return this.ok(o, 4) ? this.dv.getInt32(o, true) : 0;
  }

  /** Verilen konumdaki kayıt başlığı; sığmıyorsa undefined. */
  header(off: number, limit = this.bytes.length): Rec | undefined {
    if (off < 0 || off + 8 > limit) return undefined;
    const vi = this.u16(off);
    const len = this.u32(off + 4);
    if (off + 8 + len > limit) return undefined;
    return { ver: vi & 0xf, inst: vi >> 4, type: this.u16(off + 2), offset: off, start: off + 8, len };
  }

  /** [start, end) aralığındaki ardışık kayıtlar. */
  list(start: number, end: number, depth = 0): Rec[] {
    const out: Rec[] = [];
    if (depth > MAX_DEPTH) return out;
    let off = start;
    while (off + 8 <= end && out.length < MAX_KIDS) {
      const r = this.header(off, end);
      if (!r) break;
      out.push(r);
      off = r.start + r.len;
    }
    return out;
  }

  /** Kapsayıcının alt kayıtları. */
  kids(r: Rec | undefined, depth = 0): Rec[] {
    if (!r || r.ver !== 0xf || depth > MAX_DEPTH) return [];
    if (!r.kids) r.kids = this.list(r.start, r.start + r.len, depth);
    return r.kids;
  }

  kid(r: Rec | undefined, type: number, inst?: number): Rec | undefined {
    return this.kids(r).find((k) => k.type === type && (inst === undefined || k.inst === inst));
  }

  /** Ağaçta (derinlik öncelikli) belirtilen türdeki ilk kayıt. */
  find(r: Rec | undefined, type: number, depth = 0): Rec | undefined {
    if (!r || depth > MAX_DEPTH) return undefined;
    for (const k of this.kids(r, depth)) {
      if (k.type === type) return k;
      const f = this.find(k, type, depth + 1);
      if (f) return f;
    }
    return undefined;
  }

  slice(r: Rec): Uint8Array {
    return this.bytes.subarray(r.start, r.start + r.len);
  }

  /** UTF-16LE metin. */
  utf16(off: number, chars: number): string {
    const n = Math.max(0, Math.min(chars, Math.floor((this.bytes.length - off) / 2)));
    let s = "";
    for (let i = 0; i < n; i++) s += String.fromCharCode(this.u16(off + i * 2));
    return s;
  }

  /** Tek baytlı metin (her bayt bir UTF-16 kod biriminin alt baytı). */
  latin1(off: number, n: number): string {
    const end = Math.min(this.bytes.length, off + n);
    let s = "";
    for (let i = off; i < end; i++) s += String.fromCharCode(this.bytes[i]);
    return s;
  }
}
