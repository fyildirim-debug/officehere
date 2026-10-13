// .doc özellik değiştiricileri (Sprm, MS-DOC 2.6): işlenen boyutu, karakter/paragraf/tablo/bölüm uygulaması.
import type { UnderlineKind } from "../common";

/** Karakter durumu (stil + doğrudan biçim). */
export interface Chp {
  bold: boolean;
  italic: boolean;
  strike: boolean;
  dstrike: boolean;
  outline: boolean;
  shadow: boolean;
  smallCaps: boolean;
  caps: boolean;
  hidden: boolean;
  emboss: boolean;
  imprint: boolean;
  underline: UnderlineKind | null;
  /** Yarım punto. */
  hps: number;
  color: string | null;
  highlight: string | null;
  bg: string | null;
  iss: number;
  ftc: number | null;
  spacing: number;
  istd: number | null;
  /** Özel karakter (resim, dipnot başvurusu). */
  spec: boolean;
  picLoc: number | null;
  ole: boolean;
}

export interface Pap {
  istd: number;
  jc: number | null;
  left: number | null;
  right: number | null;
  first: number | null;
  before: number | null;
  after: number | null;
  line: { dya: number; mult: boolean } | null;
  keep: boolean;
  keepNext: boolean;
  pageBreakBefore: boolean;
  inTable: boolean;
  ttp: boolean;
  itap: number;
  ilvl: number;
  ilfo: number;
  outline: number | null;
  shading: string | null;
  /** sprmTDefTable işleneni (satır sonu paragrafında). */
  tdef: Uint8Array | null;
  tshd: Uint8Array | null;
  header: boolean;
}

export const newChp = (): Chp => ({ bold: false, italic: false, strike: false, dstrike: false, outline: false, shadow: false, smallCaps: false, caps: false, hidden: false, emboss: false, imprint: false, underline: null, hps: 20, color: null, highlight: null, bg: null, iss: 0, ftc: null, spacing: 0, istd: null, spec: false, picLoc: null, ole: false });
export const newPap = (): Pap => ({ istd: 0, jc: null, left: null, right: null, first: null, before: null, after: null, line: null, keep: false, keepNext: false, pageBreakBefore: false, inTable: false, ttp: false, itap: 0, ilvl: 0, ilfo: 0, outline: null, shading: null, tdef: null, tshd: null, header: false });

/** Ico renk dizini (sprmCIco, vurgu). */
export const ICO = [null, "#000000", "#0000FF", "#00FFFF", "#00FF00", "#FF00FF", "#FF0000", "#FFFF00", "#FFFFFF", "#000080", "#008080", "#008000", "#800080", "#800000", "#808000", "#808080", "#C0C0C0"];

const KUL: Record<number, UnderlineKind> = { 1: "single", 2: "words", 3: "double", 4: "dotted", 6: "thick", 7: "dashed", 9: "dashed", 10: "dashed", 11: "wave", 20: "dotted", 23: "dashed", 25: "dashed", 26: "dashed", 27: "wave", 39: "dashed", 43: "wave", 55: "dashed", 71: "dashed" };

/** COLORREF (r, g, b, fAuto) → "#RRGGBB" ya da otomatik için null. */
function colorRef(b: Uint8Array, o: number): string | null {
  if (b[o + 3] === 0xff) return null;
  return `#${[b[o], b[o + 1], b[o + 2]].map((x) => x.toString(16).padStart(2, "0")).join("").toUpperCase()}`;
}

/** Sprm işlenen boyutu (bayt; işlem kodu hariç). */
export function operandSize(sprm: number, b: Uint8Array, at: number): number {
  switch (sprm >> 13) {
    case 0:
    case 1:
      return 1;
    case 2:
    case 4:
    case 5:
      return 2;
    case 3:
      return 4;
    case 7:
      return 3;
    default: {
      if (sprm === 0xd608 || sprm === 0xd606) return at + 1 < b.length ? 2 + ((b[at] | (b[at + 1] << 8)) - 1) : 0;
      if (sprm === 0xc615 && b[at] === 255) {
        const del = b[at + 1] ?? 0;
        const add = b[at + 2 + del * 4] ?? 0;
        return 1 + 1 + del * 4 + 1 + add * 3;
      }
      return 1 + (b[at] ?? 0);
    }
  }
}

/** Sprm dizisini dolaşır: (işlem kodu, işlenen başlangıcı, işlenen boyutu). */
export function eachSprm(b: Uint8Array, fn: (sprm: number, at: number, size: number) => void) {
  let i = 0;
  let guard = 0;
  while (i + 2 <= b.length && guard++ < 10000) {
    const sprm = b[i] | (b[i + 1] << 8);
    const size = operandSize(sprm, b, i + 2);
    if (i + 2 + size > b.length) break;
    fn(sprm, i + 2, size);
    i += 2 + size;
  }
}

const u16 = (b: Uint8Array, o: number) => b[o] | (b[o + 1] << 8);
const i16 = (b: Uint8Array, o: number) => (u16(b, o) << 16) >> 16;
const u32 = (b: Uint8Array, o: number) => (b[o] | (b[o + 1] << 8) | (b[o + 2] << 16) | (b[o + 3] << 24)) >>> 0;

/** Aç/kapa işleneni: 0 kapalı, 1 açık, 0x80 stil değeri, 0x81 stil değerinin tersi. */
function toggle(v: number, styleVal: boolean): boolean {
  if (v === 0x80) return styleVal;
  if (v === 0x81) return !styleVal;
  return v !== 0;
}

/** Karakter sprm'lerini uygular; "base" stilin değerleridir (aç/kapa göreli değerler için). */
export function applyChpx(c: Chp, grpprl: Uint8Array, base: Chp) {
  eachSprm(grpprl, (sprm, at) => {
    const v = grpprl[at];
    switch (sprm) {
      case 0x0835:
        c.bold = toggle(v, base.bold);
        break;
      case 0x0836:
        c.italic = toggle(v, base.italic);
        break;
      case 0x0837:
        c.strike = toggle(v, base.strike);
        break;
      case 0x0838:
        c.outline = toggle(v, base.outline);
        break;
      case 0x0839:
        c.shadow = toggle(v, base.shadow);
        break;
      case 0x083a:
        c.smallCaps = toggle(v, base.smallCaps);
        break;
      case 0x083b:
        c.caps = toggle(v, base.caps);
        break;
      case 0x083c:
        c.hidden = toggle(v, base.hidden);
        break;
      case 0x2a53:
        c.dstrike = toggle(v, base.dstrike);
        break;
      case 0x0858:
        c.emboss = toggle(v, base.emboss);
        break;
      case 0x0854:
        c.imprint = toggle(v, base.imprint);
        break;
      case 0x2a3e:
        c.underline = v === 0 ? null : (KUL[v] ?? "single");
        break;
      case 0x2a42:
        c.color = ICO[v] ?? null;
        break;
      case 0x6870:
        c.color = colorRef(grpprl, at);
        break;
      case 0x4a43:
        c.hps = Math.max(2, Math.min(3276, u16(grpprl, at)));
        break;
      case 0x2a48:
        c.iss = v;
        break;
      case 0x2a0c:
        c.highlight = ICO[v] ?? null;
        break;
      case 0x4a4f:
        c.ftc = u16(grpprl, at);
        break;
      case 0x8840:
        c.spacing = i16(grpprl, at);
        break;
      case 0x4a30:
        c.istd = u16(grpprl, at);
        break;
      case 0x0855:
        c.spec = v !== 0;
        break;
      case 0x6a03:
        c.picLoc = u32(grpprl, at);
        break;
      case 0x080a:
        c.ole = v !== 0;
        break;
      case 0xca71:
        // SHD: cvFore, cvBack, ipat
        c.bg = colorRef(grpprl, at + 1 + 4);
        break;
      case 0x4866: {
        const back = (u16(grpprl, at) >> 5) & 0x1f;
        c.bg = ICO[back] ?? null;
        break;
      }
    }
  });
}

/** Paragraf (ve tablo satırı) sprm'lerini uygular. */
export function applyPapx(p: Pap, grpprl: Uint8Array) {
  eachSprm(grpprl, (sprm, at, size) => {
    const v = grpprl[at];
    switch (sprm) {
      case 0x4600:
        p.istd = u16(grpprl, at);
        break;
      case 0x2403:
      case 0x2461:
        p.jc = v;
        break;
      case 0x840f:
      case 0x845e:
        p.left = i16(grpprl, at);
        break;
      case 0x840e:
      case 0x845d:
        p.right = i16(grpprl, at);
        break;
      case 0x8411:
      case 0x8460:
        p.first = i16(grpprl, at);
        break;
      case 0xa413:
        p.before = u16(grpprl, at);
        break;
      case 0xa414:
        p.after = u16(grpprl, at);
        break;
      case 0x6412:
        p.line = { dya: i16(grpprl, at), mult: i16(grpprl, at + 2) !== 0 };
        break;
      case 0x2405:
        p.keep = v !== 0;
        break;
      case 0x2406:
        p.keepNext = v !== 0;
        break;
      case 0x2407:
        p.pageBreakBefore = v !== 0;
        break;
      case 0x2416:
        p.inTable = v !== 0;
        break;
      case 0x2417:
        p.ttp = v !== 0;
        break;
      case 0x6649:
        p.itap = u32(grpprl, at);
        if (p.itap > 0) p.inTable = true;
        break;
      case 0x244b:
        if (v) p.inTable = true;
        break;
      case 0x244c:
        if (v) p.ttp = true;
        break;
      case 0x260a:
        p.ilvl = Math.min(8, v);
        break;
      case 0x460b:
        p.ilfo = u16(grpprl, at);
        break;
      case 0x2640:
        p.outline = v < 9 ? v : null;
        break;
      case 0xc64d:
        p.shading = colorRef(grpprl, at + 1 + 4);
        break;
      case 0x442d: {
        const back = (u16(grpprl, at) >> 5) & 0x1f;
        p.shading = ICO[back] ?? null;
        break;
      }
      case 0xd608:
        p.tdef = grpprl.subarray(at, at + size);
        break;
      case 0xd670:
        p.tshd = grpprl.subarray(at, at + size);
        break;
      case 0x3404:
        p.header = v !== 0;
        break;
    }
  });
}

export interface SectionProps {
  width: number;
  height: number;
  left: number;
  right: number;
  top: number;
  bottom: number;
  header: number;
  footer: number;
  landscape: boolean;
  columns: number;
  colSpace: number;
  titlePage: boolean;
  /** Bölüm başlama türü (0 sürekli, 2 yeni sayfa, 3 çift, 4 tek). */
  bkc: number;
}

export const newSep = (): SectionProps => ({ width: 12240, height: 15840, left: 1800, right: 1800, top: 1440, bottom: 1440, header: 720, footer: 720, landscape: false, columns: 1, colSpace: 720, titlePage: false, bkc: 2 });

export function applySepx(s: SectionProps, grpprl: Uint8Array) {
  eachSprm(grpprl, (sprm, at) => {
    switch (sprm) {
      case 0xb01f:
        s.width = u16(grpprl, at);
        break;
      case 0xb020:
        s.height = u16(grpprl, at);
        break;
      case 0xb021:
        s.left = u16(grpprl, at);
        break;
      case 0xb022:
        s.right = u16(grpprl, at);
        break;
      case 0x9023:
        s.top = Math.abs(i16(grpprl, at));
        break;
      case 0x9024:
        s.bottom = Math.abs(i16(grpprl, at));
        break;
      case 0xb017:
        s.header = u16(grpprl, at);
        break;
      case 0xb018:
        s.footer = u16(grpprl, at);
        break;
      case 0x301d:
        s.landscape = grpprl[at] === 2;
        break;
      case 0x500b:
        s.columns = Math.min(13, u16(grpprl, at) + 1);
        break;
      case 0x900c:
        s.colSpace = u16(grpprl, at);
        break;
      case 0x300a:
        s.titlePage = grpprl[at] !== 0;
        break;
      case 0x3009:
        s.bkc = grpprl[at];
        break;
    }
  });
}

/** sprmTDefTable → hücre kenarları (twip) ve birleştirme bayrakları. */
export function parseTDef(op: Uint8Array): { edges: number[]; merge: number[]; vmerge: number[] } | null {
  if (op.length < 3) return null;
  const n = op[2];
  if (!n || 3 + (n + 1) * 2 > op.length) return null;
  const edges: number[] = [];
  for (let i = 0; i <= n; i++) edges.push(i16(op, 3 + i * 2));
  const merge: number[] = [];
  const vmerge: number[] = [];
  const tcAt = 3 + (n + 1) * 2;
  for (let i = 0; i < n; i++) {
    const o = tcAt + i * 20;
    if (o + 2 > op.length) {
      merge.push(0);
      vmerge.push(0);
      continue;
    }
    const grf = u16(op, o);
    const h = grf & 3;
    const vm = (grf >> 5) & 3;
    merge.push(h === 1 ? 1 : h >= 2 ? 2 : 0);
    vmerge.push(vm === 3 ? 1 : vm === 1 ? 2 : 0);
  }
  return { edges, merge, vmerge };
}

/** sprmTDefTableShd → hücre dolgu renkleri. */
export function parseTShd(op: Uint8Array | null): (string | null)[] {
  if (!op || !op.length) return [];
  const n = Math.floor((op[0] ?? 0) / 10);
  const out: (string | null)[] = [];
  for (let i = 0; i < n && 1 + i * 10 + 10 <= op.length; i++) {
    const o = 1 + i * 10;
    const back = op.subarray(o + 4, o + 8);
    out.push(back[3] === 0xff ? null : `#${[back[0], back[1], back[2]].map((x) => x.toString(16).padStart(2, "0")).join("").toUpperCase()}`);
  }
  return out;
}
