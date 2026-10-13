/**
 * MS-PPT metin okuyucu: TextHeaderAtom + TextCharsAtom/TextBytesAtom +
 * StyleTextPropAtom (paragraf ve karakter biçim aralıkları) ve asıl slayttaki
 * TextMasterStyleAtom düzey stilleri.
 */
import { normalizeRuns } from "../../model/text";
import type { BulletDef, Color, LevelStyle, Paragraph, ParaProps, Run, RunProps } from "../../model/types";
import { RT, type Rec, type Stream } from "./records";

const MAX_TEXT = 1_000_000;

export interface TextItem {
  /** TextHeaderAtom.textType (0 başlık, 1 gövde, 2 not, 4 diğer, 5 orta gövde, 6 orta başlık…). */
  type: number;
  text: string;
  style?: Rec;
}

export interface TextEnv {
  s: Stream;
  fonts: string[];
  /** Slayt renk düzeni (8 renk, "#RRGGBB"). */
  scheme: string[];
}

/** PF (paragraf) istisnası. */
interface PF {
  bulletFlags?: number;
  bulletChar?: number;
  bulletFont?: number;
  bulletSize?: number;
  bulletColor?: Color;
  align?: number;
  lineSpacing?: number;
  spaceBefore?: number;
  spaceAfter?: number;
  leftMargin?: number;
  indent?: number;
}

/** CF (karakter) istisnası. */
interface CF {
  styleMask: number;
  style: number;
  font?: number;
  size?: number;
  color?: Color;
  position?: number;
}

/** ColorIndexStruct → renk (0xFE: RGB, 0–7: düzen rengi). */
export function colorIndex(s: Stream, o: number, scheme: string[]): Color | undefined {
  const r = s.u8(o);
  const g = s.u8(o + 1);
  const b = s.u8(o + 2);
  const idx = s.u8(o + 3);
  if (idx === 0xfe) return { rgb: hex(r, g, b) };
  if (idx < 8 && scheme[idx]) return { rgb: scheme[idx] };
  return undefined;
}

export const hex = (r: number, g: number, b: number) => `#${[r, g, b].map((v) => v.toString(16).padStart(2, "0")).join("").toUpperCase()}`;

/** TextPFException'ı okur; okunan bayt sayısını döndürür. */
function readPF(s: Stream, o: number, end: number, scheme: string[]): [PF, number] {
  const pf: PF = {};
  const start = o;
  const m = s.u32(o);
  o += 4;
  const take16 = () => {
    const v = s.i16(o);
    o += 2;
    return v;
  };
  if (m & 0x0f) pf.bulletFlags = s.u16((o += 2) - 2);
  if (m & 0x80) pf.bulletChar = s.u16((o += 2) - 2);
  if (m & 0x10) pf.bulletFont = s.u16((o += 2) - 2);
  if (m & 0x40) pf.bulletSize = take16();
  if (m & 0x20) {
    pf.bulletColor = colorIndex(s, o, scheme);
    o += 4;
  }
  if (m & 0x800) pf.align = s.u16((o += 2) - 2);
  if (m & 0x1000) pf.lineSpacing = take16();
  if (m & 0x2000) pf.spaceBefore = take16();
  if (m & 0x4000) pf.spaceAfter = take16();
  if (m & 0x100) pf.leftMargin = take16();
  if (m & 0x400) pf.indent = take16();
  if (m & 0x8000) o += 2;
  if (m & 0x100000) {
    const n = Math.min(s.u16(o), 1000);
    o += 2 + n * 4;
  }
  if (m & 0x10000) o += 2;
  if (m & 0xe0000) o += 2;
  if (m & 0x200000) o += 2;
  return [pf, Math.min(o, end) - start];
}

/** TextCFException'ı okur. */
function readCF(s: Stream, o: number, end: number, scheme: string[]): [CF, number] {
  const start = o;
  const m = s.u32(o);
  o += 4;
  const cf: CF = { styleMask: m & 0xffff, style: 0 };
  if (m & 0x3eb7) cf.style = s.u16((o += 2) - 2);
  if (m & 0x10000) cf.font = s.u16((o += 2) - 2);
  if (m & 0x200000) o += 2;
  if (m & 0x400000) o += 2;
  if (m & 0x800000) o += 2;
  if (m & 0x20000) cf.size = s.u16((o += 2) - 2);
  if (m & 0x40000) {
    cf.color = colorIndex(s, o, scheme);
    o += 4;
  }
  if (m & 0x80000) cf.position = s.i16((o += 2) - 2);
  if (m & 0x100000) o += 4;
  if (m & 0x1000000) o += 2;
  if (m & 0x2000000) o += 2;
  if (m & 0x4000000) o += 4;
  return [cf, Math.min(o, end) - start];
}

/** Simge yazı tipi karakterleri yerine düz madde işareti. */
function bulletChar(c: number | undefined): string {
  if (c === undefined || c === 0) return "•";
  if (c >= 0xf000 && c <= 0xf0ff) {
    const lo = c & 0xff;
    return lo === 0xa7 ? "▪" : lo === 0xd8 ? "➢" : lo === 0x76 ? "❖" : lo === 0xfc ? "✓" : lo === 0x6e ? "■" : "•";
  }
  if (c >= 0xe000 && c <= 0xf8ff) return "•";
  return String.fromCharCode(c);
}

export function paraProps(pf: PF): ParaProps {
  const p: ParaProps = {};
  if (pf.bulletFlags !== undefined) {
    if (pf.bulletFlags & 1) {
      const b: BulletDef = { type: "char", char: bulletChar(pf.bulletChar) };
      if (pf.bulletFlags & 4 && pf.bulletColor) b.color = pf.bulletColor;
      if (pf.bulletFlags & 8 && pf.bulletSize !== undefined && pf.bulletSize > 0 && pf.bulletSize <= 400) b.size = pf.bulletSize;
      p.bullet = b;
    } else p.bullet = { type: "none" };
  }
  if (pf.align !== undefined) p.align = pf.align === 1 ? "ctr" : pf.align === 2 ? "r" : pf.align === 3 || pf.align === 4 ? "just" : "l";
  if (pf.lineSpacing !== undefined) {
    if (pf.lineSpacing >= 0) p.lineSpacing = Math.min(1000, pf.lineSpacing);
    else p.lineSpacingPt = -pf.lineSpacing / 8;
  }
  if (pf.spaceBefore !== undefined) p.spaceBefore = pf.spaceBefore >= 0 ? (pf.spaceBefore / 100) * 18 : -pf.spaceBefore / 8;
  if (pf.spaceAfter !== undefined) p.spaceAfter = pf.spaceAfter >= 0 ? (pf.spaceAfter / 100) * 18 : -pf.spaceAfter / 8;
  if (pf.leftMargin !== undefined) {
    p.marL = pf.leftMargin / 6;
    p.indent = ((pf.indent ?? 0) - pf.leftMargin) / 6;
  }
  return p;
}

export function runProps(cf: CF, fonts: string[]): RunProps {
  const r: RunProps = {};
  if (cf.styleMask & 1) r.bold = !!(cf.style & 1);
  if (cf.styleMask & 2) r.italic = !!(cf.style & 2);
  if (cf.styleMask & 4) r.underline = !!(cf.style & 4);
  if (cf.styleMask & 0x10) r.shadow = !!(cf.style & 0x10);
  if (cf.font !== undefined && fonts[cf.font]) r.font = fonts[cf.font];
  if (cf.size !== undefined && cf.size > 0 && cf.size < 4000) r.size = cf.size;
  if (cf.color) r.color = cf.color;
  if (cf.position) r.baseline = Math.max(-100, Math.min(100, cf.position));
  return r;
}

/** Bir istemci metin kutusu ya da anahat listesindeki metin öğeleri. */
export function textItems(s: Stream, recs: Rec[]): TextItem[] {
  const out: TextItem[] = [];
  let cur: TextItem | undefined;
  for (const r of recs) {
    if (r.type === RT.TextHeaderAtom) {
      cur = { type: s.u32(r.start), text: "" };
      out.push(cur);
    } else if (cur && r.type === RT.TextCharsAtom) cur.text = s.utf16(r.start, Math.min(MAX_TEXT, Math.floor(r.len / 2)));
    else if (cur && r.type === RT.TextBytesAtom) cur.text = s.latin1(r.start, Math.min(MAX_TEXT, r.len));
    else if (cur && r.type === RT.StyleTextPropAtom) cur.style = r;
  }
  return out;
}

/** Metin öğesini paragraflara çevirir. */
export function itemParagraphs(item: TextItem, env: TextEnv): Paragraph[] {
  const { s } = env;
  const text = item.text;
  const pfRuns: { n: number; level: number; pf: PF }[] = [];
  const cfRuns: { n: number; cf: CF }[] = [];
  if (item.style) {
    const end = item.style.start + item.style.len;
    let o = item.style.start;
    let covered = 0;
    while (o + 6 <= end && covered <= text.length && pfRuns.length < 10000) {
      const n = s.u32(o);
      const level = s.u16(o + 4);
      const [pf, used] = readPF(s, o + 6, end, env.scheme);
      if (!n) break;
      pfRuns.push({ n, level, pf });
      covered += n;
      o += 6 + used;
    }
    covered = 0;
    while (o + 4 <= end && covered <= text.length && cfRuns.length < 20000) {
      const n = s.u32(o);
      const [cf, used] = readCF(s, o + 4, end, env.scheme);
      if (!n) break;
      cfRuns.push({ n, cf });
      covered += n;
      o += 4 + used;
    }
  }
  const pfAt = (pos: number) => {
    let acc = 0;
    for (const r of pfRuns) {
      acc += r.n;
      if (pos < acc) return r;
    }
    return pfRuns[pfRuns.length - 1];
  };
  // Karakter biçimi sınırları.
  const cfBounds: { from: number; to: number; props: RunProps }[] = [];
  let acc = 0;
  for (const r of cfRuns) {
    cfBounds.push({ from: acc, to: acc + r.n, props: runProps(r.cf, env.fonts) });
    acc += r.n;
  }
  const propsAt = (pos: number) => cfBounds.find((b) => pos >= b.from && pos < b.to)?.props ?? cfBounds[cfBounds.length - 1]?.props ?? {};

  const paras: Paragraph[] = [];
  let start = 0;
  const parts = text.split("\r");
  for (const part of parts) {
    const pr = pfAt(start);
    const runs: Run[] = [];
    let i = 0;
    while (i < part.length) {
      const from = start + i;
      const b = cfBounds.find((x) => from >= x.from && from < x.to);
      const limit = b ? Math.min(part.length, b.to - start) : part.length;
      const seg = part.slice(i, Math.max(limit, i + 1));
      runs.push({ text: seg.replace(/\x0b/g, "\v").replace(/[\u0000-\u0008\u000c\u000e-\u001f]/g, ""), props: { ...propsAt(from) } });
      i += seg.length;
    }
    const para: Paragraph = { runs: normalizeRuns(runs.filter((r) => r.text)), level: Math.max(0, Math.min(8, pr?.level ?? 0)), props: pr ? paraProps(pr.pf) : {} };
    const endProps = propsAt(start + part.length);
    if (Object.keys(endProps).length) para.endProps = endProps;
    paras.push(para);
    start += part.length + 1;
    if (paras.length > 5000) break;
  }
  return paras.length ? paras : [{ runs: [], level: 0, props: {} }];
}

/** TextMasterStyleAtom → düzey stilleri. */
export function masterLevels(s: Stream, r: Rec, env: Omit<TextEnv, "s">): LevelStyle[] {
  const out: LevelStyle[] = [];
  const end = r.start + r.len;
  const n = Math.min(5, s.u16(r.start));
  let o = r.start + 2;
  for (let i = 0; i < n && o < end; i++) {
    if (r.inst >= 5) o += 2;
    const [pf, pu] = readPF(s, o, end, env.scheme);
    o += pu;
    const [cf, cu] = readCF(s, o, end, env.scheme);
    o += cu;
    const para = paraProps(pf);
    out.push({ para, run: runProps(cf, env.fonts) });
  }
  return out;
}
