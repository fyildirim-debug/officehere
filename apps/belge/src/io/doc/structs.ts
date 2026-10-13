// .doc yapıları: parça tablosu (CLX), biçim FKP'leri (CHPX/PAPX), stil sayfası (STSH), yazı tipi
// tablosu (SttbfFfn), liste tanımları (PlfLst/PlfLfo) ve satır içi resimler (PICF + OfficeArt BLIP).
import { Bytes, FC, type Fib } from "./fib";
import { applyChpx, applyPapx, newChp, newPap, type Chp, type Pap } from "./sprm";

export interface Piece {
  cpStart: number;
  cpEnd: number;
  /** Metnin WordDocument içindeki bayt konumu. */
  fc: number;
  compressed: boolean;
  /** Prm1 grpprl (hızlı kaydetme değişiklikleri). */
  prm: Uint8Array | null;
}

/** CLX: Prc dizisi + Pcdt (parça tablosu). */
export function parsePieces(t: Bytes, fib: Fib): Piece[] {
  const { fc, lcb } = fib.fcLcb(FC.Clx);
  if (!lcb) throw new Error("Bozuk .doc dosyası (parça tablosu yok).");
  const prcs: Uint8Array[] = [];
  let o = fc;
  const end = fc + lcb;
  let guard = 0;
  while (o < end && t.u8(o) === 0x01 && guard++ < 100000) {
    const cb = t.i16(o + 1);
    prcs.push(t.slice(o + 3, Math.max(0, cb)));
    o += 3 + Math.max(0, cb);
  }
  if (t.u8(o) !== 0x02) throw new Error("Bozuk .doc dosyası (parça tablosu bulunamadı).");
  const plcLen = t.u32(o + 1);
  const at = o + 5;
  const n = Math.floor((plcLen - 4) / 12);
  if (n <= 0 || n > 500000) throw new Error("Bozuk .doc dosyası (parça sayısı geçersiz).");
  const pieces: Piece[] = [];
  for (let i = 0; i < n; i++) {
    const cpStart = t.i32(at + i * 4);
    const cpEnd = t.i32(at + (i + 1) * 4);
    const pcd = at + (n + 1) * 4 + i * 8;
    const raw = t.u32(pcd + 2);
    const compressed = (raw & 0x40000000) !== 0;
    const fcv = raw & 0x3fffffff;
    const prmV = t.u16(pcd + 6);
    const prm = prmV & 1 ? (prcs[prmV >> 1] ?? null) : null;
    if (cpEnd > cpStart) pieces.push({ cpStart, cpEnd, fc: compressed ? fcv / 2 : fcv, compressed, prm });
  }
  return pieces;
}

const cp1252 = new TextDecoder("windows-1252");

/** Parçalardan tüm metin (CP dizini = karakter dizini). */
export function pieceText(wd: Bytes, pieces: Piece[], maxCp: number): string {
  const parts: string[] = [];
  let at = 0;
  for (const p of pieces) {
    if (p.cpStart >= maxCp) break;
    if (p.cpStart > at) parts.push("\u0000".repeat(Math.min(p.cpStart - at, 1_000_000)));
    const n = Math.min(p.cpEnd, maxCp) - p.cpStart;
    if (p.compressed) parts.push(cp1252.decode(wd.slice(p.fc, n)));
    else parts.push(wd.utf16(p.fc, n));
    at = p.cpStart + n;
  }
  return parts.join("");
}

/** FC (bayt) aralığındaki biçim koşusu. */
export interface FcRun {
  fcStart: number;
  fcEnd: number;
  grpprl: Uint8Array;
  istd: number;
}

/** PlcBteChpx / PlcBtePapx → FKP sayfalarındaki koşular. */
export function parseFkps(wd: Bytes, t: Bytes, fib: Fib, papx: boolean): FcRun[] {
  const { fc, lcb } = fib.fcLcb(papx ? FC.PlcfBtePapx : FC.PlcfBteChpx);
  if (!lcb) return [];
  const n = Math.floor((lcb - 4) / 8);
  const runs: FcRun[] = [];
  for (let i = 0; i < n; i++) {
    const pn = t.u32(fc + (n + 1) * 4 + i * 4) & 0x3fffff;
    const page = pn * 512;
    if (page + 512 > wd.length) continue;
    const count = wd.u8(page + 511);
    for (let j = 0; j < count; j++) {
      const fcStart = wd.u32(page + j * 4);
      const fcEnd = wd.u32(page + (j + 1) * 4);
      if (papx) {
        const bOff = wd.u8(page + (count + 1) * 4 + j * 13) * 2;
        if (!bOff) {
          runs.push({ fcStart, fcEnd, grpprl: new Uint8Array(0), istd: 0 });
          continue;
        }
        let cb = wd.u8(page + bOff);
        let dataAt = page + bOff + 1;
        if (cb === 0) {
          cb = wd.u8(page + bOff + 1);
          dataAt++;
          cb *= 2;
        } else cb = cb * 2 - 1;
        if (dataAt + cb > page + 512 || cb < 2) {
          runs.push({ fcStart, fcEnd, grpprl: new Uint8Array(0), istd: 0 });
          continue;
        }
        runs.push({ fcStart, fcEnd, istd: wd.u16(dataAt), grpprl: wd.slice(dataAt + 2, cb - 2) });
      } else {
        const off = wd.u8(page + (count + 1) * 4 + j) * 2;
        const cb = off ? wd.u8(page + off) : 0;
        runs.push({ fcStart, fcEnd, istd: 0, grpprl: off && page + off + 1 + cb <= page + 512 ? wd.slice(page + off + 1, cb) : new Uint8Array(0) });
      }
    }
  }
  return runs.sort((a, b) => a.fcStart - b.fcStart);
}

/** CP aralığındaki biçim koşusu. */
export interface CpRun {
  cpStart: number;
  cpEnd: number;
  run: FcRun;
  piece: Piece;
}

/** FC koşularını parçalar üzerinden CP koşularına çevirir (sıralı). */
export function cpRuns(pieces: Piece[], runs: FcRun[]): CpRun[] {
  const out: CpRun[] = [];
  for (const p of pieces) {
    const size = p.compressed ? 1 : 2;
    const fcEnd = p.fc + (p.cpEnd - p.cpStart) * size;
    // İkili arama: parçanın başlangıcını kapsayan ilk koşu.
    let lo = 0;
    let hi = runs.length;
    while (lo < hi) {
      const mid = (lo + hi) >> 1;
      if (runs[mid].fcEnd <= p.fc) lo = mid + 1;
      else hi = mid;
    }
    for (let k = lo; k < runs.length && runs[k].fcStart < fcEnd; k++) {
      const r = runs[k];
      const a = Math.max(r.fcStart, p.fc);
      const b = Math.min(r.fcEnd, fcEnd);
      if (b <= a) continue;
      out.push({ cpStart: p.cpStart + Math.floor((a - p.fc) / size), cpEnd: p.cpStart + Math.ceil((b - p.fc) / size), run: r, piece: p });
    }
  }
  return out.sort((x, y) => x.cpStart - y.cpStart);
}

/** CP'yi kapsayan koşuyu bulur (ikili arama). */
export function runAt(runs: CpRun[], cp: number): CpRun | null {
  let lo = 0;
  let hi = runs.length - 1;
  while (lo <= hi) {
    const mid = (lo + hi) >> 1;
    const r = runs[mid];
    if (cp < r.cpStart) hi = mid - 1;
    else if (cp >= r.cpEnd) lo = mid + 1;
    else return r;
  }
  return null;
}

/* ------------------------------------------------------------------ */
/* Stil sayfası                                                        */
/* ------------------------------------------------------------------ */

export interface DocStyle {
  sti: number;
  stk: number;
  base: number;
  name: string;
  papx: Uint8Array | null;
  chpx: Uint8Array | null;
}

export class StyleSheet {
  styles: (DocStyle | null)[] = [];
  /** Varsayılan yazı tipi (ftcAsci). */
  defaultFtc = 0;
  private chpCache = new Map<number, Chp>();
  private papCache = new Map<number, Pap>();

  constructor(t: Bytes, fib: Fib) {
    const { fc, lcb } = fib.fcLcb(FC.Stshf);
    if (!lcb) return;
    const cbStshi = t.u16(fc);
    const cstd = t.u16(fc + 2);
    const cbBase = t.u16(fc + 4) || 10;
    if (cbStshi >= 16) this.defaultFtc = t.u16(fc + 2 + 12);
    let o = fc + 2 + cbStshi;
    const end = fc + lcb;
    for (let i = 0; i < cstd && o + 2 <= end; i++) {
      const cb = t.u16(o);
      o += 2;
      if (!cb) {
        this.styles.push(null);
        continue;
      }
      try {
        this.styles.push(parseStd(t.slice(o, cb), cbBase));
      } catch {
        this.styles.push(null);
      }
      o += cb;
    }
  }

  /** Stilin karakter özellikleri (temel stil zinciriyle). */
  chp(istd: number, depth = 0): Chp {
    const hit = this.chpCache.get(istd);
    if (hit) return hit;
    const s = this.styles[istd];
    const base = s && s.base !== 0xfff && depth < 10 && s.base !== istd ? this.chp(s.base, depth + 1) : { ...newChp(), ftc: this.defaultFtc };
    const c = { ...base };
    if (s?.chpx) applyChpx(c, s.chpx, base);
    this.chpCache.set(istd, c);
    return c;
  }

  pap(istd: number, depth = 0): Pap {
    const hit = this.papCache.get(istd);
    if (hit) return hit;
    const s = this.styles[istd];
    const base = s && s.base !== 0xfff && depth < 10 && s.base !== istd ? this.pap(s.base, depth + 1) : newPap();
    const p = { ...base, istd };
    if (s?.papx) applyPapx(p, s.papx);
    p.istd = istd;
    this.papCache.set(istd, p);
    return p;
  }
}

function parseStd(b: Uint8Array, cbBase: number): DocStyle {
  const t = new Bytes(b);
  const sti = t.u16(0) & 0x0fff;
  const w1 = t.u16(2);
  const stk = w1 & 0xf;
  const base = w1 >> 4;
  const cupx = t.u16(4) & 0xf;
  let o = cbBase;
  const cch = t.u16(o);
  const name = t.utf16(o + 2, Math.min(cch, 255));
  o += 2 + cch * 2 + 2;
  const upx: Uint8Array[] = [];
  for (let i = 0; i < cupx && o + 2 <= b.length; i++) {
    if (o % 2) o++;
    const cb = t.u16(o);
    upx.push(t.slice(o + 2, Math.min(cb, b.length - o - 2)));
    o += 2 + cb;
  }
  let papx: Uint8Array | null = null;
  let chpx: Uint8Array | null = null;
  if (stk === 1) {
    papx = upx[0] && upx[0].length >= 2 ? upx[0].subarray(2) : null;
    chpx = upx[1] ?? null;
  } else if (stk === 2) chpx = upx[0] ?? null;
  return { sti, stk, base, name, papx, chpx };
}

/** Yazı tipi adları (SttbfFfn). */
export function parseFonts(t: Bytes, fib: Fib): string[] {
  const { fc, lcb } = fib.fcLcb(FC.SttbfFfn);
  const out: string[] = [];
  if (!lcb) return out;
  let n = t.u16(fc);
  let o = fc + 4;
  if (n === 0xffff) {
    n = t.u16(fc + 2);
    o = fc + 6;
  }
  for (let i = 0; i < n && o < fc + lcb; i++) {
    const len = t.u8(o);
    const ffn = new Bytes(t.slice(o + 1, len));
    let name = "";
    for (let k = 39; k + 1 < len; k += 2) {
      const ch = ffn.u16(k);
      if (!ch) break;
      name += String.fromCharCode(ch);
    }
    out.push(name);
    o += 1 + len;
  }
  return out;
}

/* ------------------------------------------------------------------ */
/* Listeler                                                            */
/* ------------------------------------------------------------------ */

export interface DocListLevel {
  nfc: number;
  start: number;
  text: string;
}

export class Lists {
  private byLsid = new Map<number, DocListLevel[]>();
  private lfos: number[] = [];

  constructor(t: Bytes, fib: Fib) {
    try {
      const lst = fib.fcLcb(FC.PlfLst);
      if (lst.lcb) {
        const n = t.i16(lst.fc);
        const defs: { lsid: number; simple: boolean }[] = [];
        for (let i = 0; i < n; i++) defs.push({ lsid: t.i32(lst.fc + 2 + i * 28), simple: (t.u8(lst.fc + 2 + i * 28 + 26) & 1) !== 0 });
        let o = lst.fc + 2 + n * 28;
        for (const d of defs) {
          const levels: DocListLevel[] = [];
          for (let k = 0; k < (d.simple ? 1 : 9); k++) {
            const start = t.i32(o);
            const nfc = t.u8(o + 4);
            const cbChpx = t.u8(o + 24);
            const cbPapx = t.u8(o + 25);
            o += 28 + cbPapx + cbChpx;
            const cch = t.u16(o);
            const text = t.utf16(o + 2, Math.min(cch, 64));
            o += 2 + cch * 2;
            levels.push({ nfc, start, text });
          }
          this.byLsid.set(d.lsid, levels);
        }
      }
      const lfo = fib.fcLcb(FC.PlfLfo);
      if (lfo.lcb) {
        const n = Math.min(t.i32(lfo.fc), 10000);
        for (let i = 0; i < n; i++) this.lfos.push(t.i32(lfo.fc + 4 + i * 16));
      }
    } catch {
      // Bozuk liste tablosu: listeler madde işaretli kabul edilir.
    }
  }

  level(ilfo: number, ilvl: number): DocListLevel | null {
    const lsid = this.lfos[ilfo - 1];
    const levels = lsid === undefined ? undefined : this.byLsid.get(lsid);
    return levels?.[Math.min(ilvl, levels.length - 1)] ?? null;
  }
}

/* ------------------------------------------------------------------ */
/* Resimler                                                            */
/* ------------------------------------------------------------------ */

/** OfficeArt BLIP kaydından PNG/JPEG baytları (diğer türlerde null). */
function blipBytes(b: Bytes, o: number): Uint8Array | null {
  if (o + 8 > b.length) return null;
  const inst = b.u16(o) >> 4;
  const type = b.u16(o + 2);
  const len = b.u32(o + 4);
  if (type !== 0xf01d && type !== 0xf01e && type !== 0xf02a) return null;
  const two = inst === 0x46b || inst === 0x6e3 || inst === 0x6e1 || inst === 0x6e5;
  const skip = 16 + (two ? 16 : 0) + 1;
  const n = Math.min(len - skip, b.length - o - 8 - skip);
  return n > 0 ? b.slice(o + 8 + skip, n) : null;
}

/** Şekil kapsayıcısındaki (SpContainer) şekil kimliği ve resim deposu sırası (pib). */
function shapeInfo(b: Bytes, start: number, end: number): { spid: number | null; pib: number | null } {
  let spid: number | null = null;
  let pib: number | null = null;
  for (let o = start, guard = 0; o + 8 <= end && guard < 200; guard++) {
    const type = b.u16(o + 2);
    const len = b.u32(o + 4);
    if (type === 0xf00a) spid = b.u32(o + 8);
    else if (type === 0xf00b || type === 0xf122) {
      const n = b.u16(o) >> 4;
      for (let k = 0; k < n && o + 8 + k * 6 + 6 <= end; k++) if ((b.u16(o + 8 + k * 6) & 0x3fff) === 0x104) pib = b.u32(o + 8 + k * 6 + 2);
    }
    o += 8 + len;
  }
  return { spid, pib };
}

/** Çizimlerdeki şekiller: şekil kimliği → pib (resim satır içiyken PICF'te pib yoksa buradan bulunur). */
function shapePibs(t: Bytes, from: number, end: number, out: Map<number, number>, depth = 0) {
  for (let o = from, guard = 0; o + 8 <= end && guard < 5000 && depth < 8; guard++) {
    const ver = t.u16(o) & 0xf;
    const type = t.u16(o + 2);
    const len = t.u32(o + 4);
    if (type === 0xf004) {
      const s = shapeInfo(t, o + 8, Math.min(end, o + 8 + len));
      if (s.spid !== null && s.pib) out.set(s.spid, s.pib);
    } else if (ver === 0xf) shapePibs(t, o + 8, Math.min(end, o + 8 + len), out, depth + 1);
    o += 8 + len;
  }
}

export interface PictureStore {
  blips: (Uint8Array | null)[];
  pibs: Map<number, number>;
  /** Kendi başvurusu olmayan satır içi resimler için sıradaki depo girdisi. */
  next: number;
}

/** Çizim grubundaki resim deposu (BStore): FBSE başına BLIP (gömülü ya da WordDocument'ta gecikmeli). */
export function parseBStore(t: Bytes, fib: Fib, wd: Bytes): PictureStore {
  const out: (Uint8Array | null)[] = [];
  const pibs = new Map<number, number>();
  try {
    const { fc, lcb } = fib.fcLcb(FC.DggInfo);
    if (!lcb || t.u16(fc + 2) !== 0xf000) return { blips: out, pibs, next: 0 };
    const end = Math.min(fc + 8 + t.u32(fc + 4), fc + lcb);
    for (let o = fc + 8, guard = 0; o + 8 <= end && guard < 1000; guard++) {
      const type = t.u16(o + 2);
      const len = t.u32(o + 4);
      if (type === 0xf001) {
        for (let p = o + 8, g2 = 0; p + 8 <= o + 8 + len && g2 < 10000; g2++) {
          const plen = t.u32(p + 4);
          if (t.u16(p + 2) === 0xf007) {
            const cbName = t.u8(p + 8 + 33);
            const foDelay = t.u32(p + 8 + 28);
            out.push(plen > 36 + cbName ? blipBytes(t, p + 8 + 36 + cbName) : foDelay < wd.length ? blipBytes(wd, foDelay) : null);
          } else out.push(null);
          p += 8 + plen;
        }
      }
      o += 8 + len;
    }
    // Çizimler: her birinin önünde 1 baytlık etiket (dgglbl) vardır.
    let o = fc + 8 + t.u32(fc + 4);
    for (let guard = 0; o + 9 <= fc + lcb && guard < 16; guard++) {
      const len = t.u32(o + 1 + 4);
      if (t.u16(o + 1 + 2) !== 0xf002) break;
      shapePibs(t, o + 1 + 8, Math.min(fc + lcb, o + 1 + 8 + len), pibs);
      o += 1 + 8 + len;
    }
  } catch {
    // Bozuk çizim grubu: resimler atlanır.
  }
  return { blips: out, pibs, next: 0 };
}

/** Data akışındaki PICF'ten PNG/JPEG baytları ve hedef boyut (px). */
export function readPicture(data: Bytes | null, at: number, store: PictureStore = { blips: [], pibs: new Map(), next: 0 }): { bytes: Uint8Array; w: number; h: number } | null {
  if (!data || at < 0 || at + 0x44 > data.length) return null;
  try {
    const lcb = data.i32(at);
    const cbHeader = data.u16(at + 4);
    if (lcb < cbHeader || at + lcb > data.length) return null;
    const mm = data.i16(at + 6);
    const dxa = data.i16(at + 28);
    const dya = data.i16(at + 30);
    const mx = data.u16(at + 32) || 1000;
    const my = data.u16(at + 34) || 1000;
    let o = at + cbHeader;
    if (mm === 0x66) o += 1 + data.u8(o);
    const end = at + lcb;
    let bytes: Uint8Array | null = null;
    let spid: number | null = null;
    // OfficeArt kayıtları: kapsayıcılara girilir; BLIP/FBSE kaydında ya da şekil özelliğindeki (pib) depo başvurusunda resim.
    for (let guard = 0; o + 8 <= end && guard < 500 && !bytes; guard++) {
      const ver = data.u16(o) & 0xf;
      const type = data.u16(o + 2);
      const len = data.u32(o + 4);
      if (ver === 0xf) {
        o += 8;
        continue;
      }
      if (type === 0xf00a) spid = data.u32(o + 8);
      if (type === 0xf007) bytes = blipBytes(data, o + 8 + 36 + data.u8(o + 8 + 33));
      else if (type === 0xf00b || type === 0xf122) {
        const n = data.u16(o) >> 4;
        for (let k = 0; k < n && o + 8 + k * 6 + 6 <= end; k++) {
          const pid = data.u16(o + 8 + k * 6) & 0x3fff;
          if (pid === 0x104) bytes = store.blips[data.u32(o + 8 + k * 6 + 2) - 1] ?? null;
        }
      } else bytes = blipBytes(data, o);
      o += 8 + len;
    }
    // Şekil özelliği çizimdeyse şekil kimliğiyle bulunur.
    if (!bytes && spid !== null && store.pibs.has(spid)) bytes = store.blips[store.pibs.get(spid)! - 1] ?? null;
    // Eşleşme yoksa depodaki sıradaki resim (bazı üreticiler satır içi resmi yalnız depoda tutar).
    if (!bytes && store.next < store.blips.length) bytes = store.blips[store.next++] ?? null;
    if (!bytes) {
      // Yedek: imza taraması
      const b = data.slice(at + cbHeader, Math.max(0, lcb - cbHeader));
      for (let i = 0; i + 8 < b.length; i++) {
        if ((b[i] === 0x89 && b[i + 1] === 0x50 && b[i + 2] === 0x4e && b[i + 3] === 0x47) || (b[i] === 0xff && b[i + 1] === 0xd8 && b[i + 2] === 0xff)) {
          bytes = b.subarray(i);
          break;
        }
      }
    }
    if (!bytes) return null;
    return { bytes, w: Math.round(((dxa * mx) / 1000) / 15), h: Math.round(((dya * my) / 1000) / 15) };
  } catch {
    return null;
  }
}
