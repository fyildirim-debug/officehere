// PDF son işleme: belge bilgileri ve PDF/A-2b / PDF/A-1b uyumu (XMP, sRGB OutputIntent, trailer /ID, CIDSet).
// Chromium'un PDF'i (Skia) yazı tiplerini alt küme olarak gömer ve ToUnicode yazar; eksik olanlar burada eklenir.
// Çizimi değiştirmeyen içerik düzeltmeleri: .notdef glifi yerine aynı genişlikte boşluk; 1b'de opaklığı 0
// olan (görünmez) çizimler, sRGB ICC v4 profilleri ve 32767'yi aşan ötelemeler. Electron'a bağlı değildir.
import { randomBytes } from "node:crypto";
import { readFileSync } from "node:fs";
import path from "node:path";
import { PDFArray, PDFDict, PDFDocument, PDFHexString, PDFName, PDFNumber, PDFRawStream, PDFStream, PDFString, decodePDFRawStream, type PDFObject } from "@cantoo/pdf-lib";

export type PdfStandard = "pdf" | "pdfa-2b" | "pdfa-1b";
export type PdfAfterExport = "none" | "open" | "folder";

export interface PdfMeta {
  title?: string;
  author?: string;
  subject?: string;
  keywords?: string;
  /** Oluşturan uygulama (xmp:CreatorTool). */
  creator?: string;
  /** Üretici (pdf:Producer). */
  producer?: string;
}

/** Sayfadan gelen dışa aktarma seçenekleri, denetlenmiş. */
export interface PdfExportRequest extends PdfMeta {
  landscape: boolean;
  pageSize: "A4" | "A3" | "A5" | "Letter" | "Legal" | { width: number; height: number };
  pageRanges?: string;
  standard: PdfStandard;
  outline: boolean;
  tagged: boolean;
  after: PdfAfterExport;
}

const PAGE_SIZES = new Set(["A4", "A3", "A5", "Letter", "Legal"]);

const text = (v: unknown, max = 1000) => (typeof v === "string" ? v.replace(/[\u0000-\u0008\u000b\u000c\u000e-\u001f]/g, "").trim().slice(0, max) : undefined) || undefined;

/** Sayfa sürecinden gelen seçenekleri denetler (bilinmeyen değerler varsayılana düşer). */
export function pdfExportRequest(raw: unknown): PdfExportRequest {
  const o = (raw && typeof raw === "object" ? raw : {}) as Record<string, unknown>;
  const size = o.pageSize as PdfExportRequest["pageSize"] | undefined;
  const okSize =
    typeof size === "string" ? PAGE_SIZES.has(size) : !!size && typeof size === "object" && [size.width, size.height].every((n) => typeof n === "number" && Number.isFinite(n) && n > 0 && n < 200);
  const ranges = typeof o.pageRanges === "string" ? o.pageRanges.trim() : "";
  return {
    landscape: o.landscape === true,
    pageSize: okSize ? size! : "A4",
    pageRanges: ranges && ranges.length <= 400 && /^\d+(\s*-\s*\d+)?(\s*,\s*\d+(\s*-\s*\d+)?)*$/.test(ranges) ? ranges : undefined,
    standard: o.standard === "pdfa-2b" || o.standard === "pdfa-1b" ? o.standard : "pdf",
    outline: o.outline === true,
    tagged: o.tagged === true,
    after: o.after === "open" || o.after === "folder" ? o.after : "none",
    title: text(o.title),
    author: text(o.author),
    subject: text(o.subject),
    keywords: text(o.keywords),
    creator: text(o.creator, 100),
  };
}

/* ------------------------------------------------------------------ */
/* Saydamlık (PDF/A-1b'de yasak)                                        */
/* ------------------------------------------------------------------ */

export type TransparencyKind = "gorsel" | "opaklik" | "maske" | "karisim";

export interface TransparencyFinding {
  /** 1 tabanlı sayfa numarası. */
  page: number;
  kind: TransparencyKind;
}

export const TRANSPARENCY_LABELS: Record<TransparencyKind, string> = {
  gorsel: "saydam arka planlı ya da yarı saydam görsel",
  opaklik: "yarı saydam renk ya da opaklığı düşürülmüş öğe",
  maske: "gölge, parlama ya da yumuşak kenar (saydamlık maskesi)",
  karisim: "karışım kipi (ör. çarpma, ekran)",
};

const num = (o: PDFObject | undefined) => (o instanceof PDFNumber ? o.asNumber() : undefined);
const nameOf = (o: PDFObject | undefined) => (o instanceof PDFName ? o.decodeText() : undefined);
const dictOf = (o: PDFDict | PDFStream) => (o instanceof PDFDict ? o : o.dict);

/** Grafik durumu sözlüğündeki saydamlık: opaklık < 1, yumuşak maske, Normal dışı karışım kipi. */
function scanGraphicsState(g: PDFDict, kinds: Set<TransparencyKind>) {
  if ((num(g.get(PDFName.of("ca"))) ?? 1) < 1 || (num(g.get(PDFName.of("CA"))) ?? 1) < 1) kinds.add("opaklik");
  const sm = g.get(PDFName.of("SMask"));
  if (sm && nameOf(sm) !== "None") kinds.add("maske");
  const bm = g.lookup(PDFName.of("BM"));
  const modes = bm instanceof PDFArray ? bm.asArray().map(nameOf) : [nameOf(bm)];
  if (bm && modes.some((m) => m && m !== "Normal" && m !== "Compatible")) kinds.add("karisim");
}

/** Kaynak sözlüğündeki (ve iç içe form, desen, Type3 yazı tipi kaynaklarındaki) saydamlık türleri. */
function scanResources(doc: PDFDocument, res: PDFDict | undefined, kinds: Set<TransparencyKind>, seen: Set<PDFObject>) {
  if (!res || seen.has(res)) return;
  seen.add(res);
  const ctx = doc.context;
  const each = (key: string, cb: (o: PDFDict | PDFStream) => void) => {
    const d = ctx.lookupMaybe(res.get(PDFName.of(key)), PDFDict);
    if (!d) return;
    for (const [, v] of d.entries()) {
      const o = ctx.lookup(v);
      if (o instanceof PDFDict || o instanceof PDFStream) cb(o);
    }
  };
  each("ExtGState", (gs) => scanGraphicsState(dictOf(gs), kinds));
  const nested = (o: PDFDict | PDFStream) => {
    if (seen.has(o)) return;
    seen.add(o);
    scanResources(doc, ctx.lookupMaybe(dictOf(o).get(PDFName.of("Resources")), PDFDict), kinds, seen);
  };
  each("XObject", (x) => {
    const d = dictOf(x);
    const sub = nameOf(d.get(PDFName.of("Subtype")));
    if (sub === "Image" && (d.has(PDFName.of("SMask")) || (num(d.get(PDFName.of("SMaskInData"))) ?? 0) > 0)) kinds.add("gorsel");
    if (sub === "Form") nested(x);
  });
  each("Pattern", (p) => {
    nested(p);
    const gs = ctx.lookupMaybe(dictOf(p).get(PDFName.of("ExtGState")), PDFDict);
    if (gs) scanGraphicsState(gs, kinds);
  });
  each("Font", (f) => {
    if (nameOf(dictOf(f).get(PDFName.of("Subtype"))) === "Type3") nested(f);
  });
}

/**
 * PDF/A-1b'nin izin vermediği saydamlık kaynakları, sayfa sayfa. Yalnız boş bir saydamlık grubu
 * (çizimde saydam öğe yokken) sayılmaz: 1b dönüşümünde kaldırılır.
 */
export function findTransparency(doc: PDFDocument): TransparencyFinding[] {
  const out: TransparencyFinding[] = [];
  doc.getPages().forEach((page, i) => {
    const kinds = new Set<TransparencyKind>();
    scanResources(doc, page.node.Resources(), kinds, new Set());
    const annots = page.node.Annots();
    for (const a of annots?.asArray() ?? []) {
      const d = doc.context.lookupMaybe(a, PDFDict);
      if (d && (num(d.get(PDFName.of("CA"))) ?? 1) < 1) kinds.add("opaklik");
    }
    for (const kind of kinds) out.push({ page: i + 1, kind });
  });
  return out;
}

/** "1. sayfa: …; 3. sayfa: …" (en çok 6 satır). */
export function describeTransparency(found: TransparencyFinding[]): string {
  const byPage = new Map<number, string[]>();
  for (const f of found) byPage.set(f.page, [...(byPage.get(f.page) ?? []), TRANSPARENCY_LABELS[f.kind]]);
  const lines = [...byPage].map(([p, ks]) => `• ${p}. sayfa: ${ks.join(", ")}`);
  return lines.length > 6 ? [...lines.slice(0, 6), `• … ve ${lines.length - 6} sayfa daha`].join("\n") : lines.join("\n");
}

/* ------------------------------------------------------------------ */
/* Yazı tipleri                                                         */
/* ------------------------------------------------------------------ */

/** TrueType/OpenType ('maxp') ya da çıplak CFF yazı tipindeki glif sayısı. */
export function glyphCount(font: Uint8Array): number | null {
  const v = new DataView(font.buffer, font.byteOffset, font.byteLength);
  if (font.length < 12) return null;
  const tag = String.fromCharCode(font[0], font[1], font[2], font[3]);
  if (tag === "\u0000\u0001\u0000\u0000" || tag === "true" || tag === "OTTO") {
    const n = v.getUint16(4);
    for (let i = 0; i < n; i++) {
      const o = 12 + 16 * i;
      if (o + 16 > font.length) return null;
      if (String.fromCharCode(font[o], font[o + 1], font[o + 2], font[o + 3]) === "maxp") {
        const off = v.getUint32(o + 8);
        return off + 6 <= font.length ? v.getUint16(off + 4) : null;
      }
    }
    return null;
  }
  // CFF: başlık → Name INDEX → Top DICT INDEX (ilk DICT'te CharStrings, işleç 17) → CharStrings INDEX sayısı.
  if (font[0] !== 1) return null;
  try {
    const index = (at: number) => {
      const count = v.getUint16(at);
      if (count === 0) return { count, data: at + 2, offs: [] as number[], end: at + 2 };
      const size = font[at + 2];
      const read = (k: number) => {
        let x = 0;
        for (let b = 0; b < size; b++) x = x * 256 + font[at + 3 + k * size + b];
        return x;
      };
      const offs = Array.from({ length: count + 1 }, (_, k) => read(k));
      const data = at + 3 + (count + 1) * size - 1;
      return { count, data, offs, end: data + offs[count] };
    };
    const names = index(font[2]);
    const top = index(names.end);
    let p = top.data + top.offs[0];
    const end = top.data + top.offs[1];
    const operands: number[] = [];
    while (p < end) {
      const b = font[p];
      if (b <= 21) {
        const op = b === 12 ? 1200 + font[p + 1] : b;
        p += b === 12 ? 2 : 1;
        if (op === 17) return index(operands[operands.length - 1]).count;
        operands.length = 0;
      } else if (b === 28) {
        operands.push(v.getInt16(p + 1));
        p += 3;
      } else if (b === 29) {
        operands.push(v.getInt32(p + 1));
        p += 5;
      } else if (b === 30) {
        p++;
        while (p < end && (font[p] & 0x0f) !== 0x0f && (font[p] >> 4) !== 0x0f) p++;
        p++;
        operands.push(0);
      } else if (b >= 32 && b <= 246) {
        operands.push(b - 139);
        p++;
      } else if (b >= 247 && b <= 250) {
        operands.push((b - 247) * 256 + font[p + 1] + 108);
        p += 2;
      } else if (b >= 251 && b <= 254) {
        operands.push(-(b - 251) * 256 - font[p + 1] - 108);
        p += 2;
      } else return null;
    }
  } catch {
    return null;
  }
  return null;
}

export interface FontInfo {
  name: string;
  /** Yazı tipi programı PDF'te (Type3 her zaman gömülü sayılır). */
  embedded: boolean;
  cidSet: boolean;
}

/** Sayfalarda kullanılan yazı tipleri (descendant CIDFont'lar dahil). */
export function listFonts(doc: PDFDocument): FontInfo[] {
  const ctx = doc.context;
  const out = new Map<string, FontInfo>();
  for (const [, obj] of ctx.enumerateIndirectObjects()) {
    if (!(obj instanceof PDFDict) || nameOf(obj.get(PDFName.of("Type"))) !== "Font") continue;
    const sub = nameOf(obj.get(PDFName.of("Subtype")));
    if (sub === "Type0") continue; // descendant ayrıca sayılır
    const name = nameOf(obj.get(PDFName.of("BaseFont"))) ?? (sub === "Type3" ? "Type3" : "?");
    const fd = ctx.lookupMaybe(obj.get(PDFName.of("FontDescriptor")), PDFDict);
    const embedded = sub === "Type3" || (!!fd && ["FontFile", "FontFile2", "FontFile3"].some((k) => fd.has(PDFName.of(k))));
    out.set(`${name}#${out.size}`, { name, embedded, cidSet: !!fd?.has(PDFName.of("CIDSet")) });
  }
  return [...out.values()];
}

/** PDF/A-1b: alt küme CIDFont'ların tanımlayıcısına tüm glifleri işaretleyen CIDSet (veraPDF yalnız "hepsi"ni kabul eder). */
function addCidSets(doc: PDFDocument): number {
  const ctx = doc.context;
  let n = 0;
  for (const [, obj] of ctx.enumerateIndirectObjects()) {
    if (!(obj instanceof PDFDict) || nameOf(obj.get(PDFName.of("Type"))) !== "Font") continue;
    const sub = nameOf(obj.get(PDFName.of("Subtype")));
    if (sub !== "CIDFontType2" && sub !== "CIDFontType0") continue;
    const fd = ctx.lookupMaybe(obj.get(PDFName.of("FontDescriptor")), PDFDict);
    const file = fd && (ctx.lookup(fd.get(PDFName.of("FontFile2"))) ?? ctx.lookup(fd.get(PDFName.of("FontFile3"))));
    if (!fd || !(file instanceof PDFRawStream)) continue;
    const count = glyphCount(decodePDFRawStream(file).decode());
    if (!count) continue;
    const bits = new Uint8Array(Math.ceil(count / 8));
    for (let g = 0; g < count; g++) bits[g >> 3] |= 0x80 >> (g & 7);
    fd.set(PDFName.of("CIDSet"), ctx.register(ctx.flateStream(bits)));
    n++;
  }
  return n;
}

/* ------------------------------------------------------------------ */
/* .notdef (PDF/A-2 6.2.11.8)                                           */
/* ------------------------------------------------------------------ */

/**
 * Type0 (Identity-H) yazı tipinde 0 numaralı glifin (.notdef) genişliği (binde birim); başka kodlamada null.
 * Hiçbir yazı tipinde olmayan karakter (ör. Symbol yazı tipinin özel alan madde imi) bu glifle çizilir.
 */
function notdefWidth(doc: PDFDocument, font: PDFDict): number | null {
  const ctx = doc.context;
  if (nameOf(font.get(PDFName.of("Subtype"))) !== "Type0" || nameOf(font.get(PDFName.of("Encoding"))) !== "Identity-H") return null;
  const desc = ctx.lookupMaybe(ctx.lookupMaybe(font.get(PDFName.of("DescendantFonts")), PDFArray)?.get(0), PDFDict);
  if (!desc) return null;
  const map = desc.get(PDFName.of("CIDToGIDMap"));
  if (map && nameOf(map) !== "Identity") return null;
  const w = ctx.lookupMaybe(desc.get(PDFName.of("W")), PDFArray)?.asArray() ?? [];
  for (let i = 0; i + 1 < w.length; ) {
    const c = num(ctx.lookup(w[i])) ?? 0;
    const next = ctx.lookup(w[i + 1]);
    if (next instanceof PDFArray) {
      if (c === 0 && next.size()) return num(ctx.lookup(next.get(0))) ?? 0;
      i += 2;
    } else {
      if (c === 0) return num(ctx.lookup(w[i + 2])) ?? 0;
      i += 3;
    }
  }
  return num(ctx.lookup(desc.get(PDFName.of("DW")))) ?? 1000;
}

const codes = (hex: string): string[] => hex.replace(/\s+/g, "").match(/.{4}/g) ?? [];

/** Onaltılık glif dizisi → TJ öğeleri: .notdef yerine aynı genişlikte kaydırma (sonraki glifler yerinde kalır). */
function withoutNotdef(hex: string, width: number): string {
  const parts: string[] = [];
  let run = "";
  for (const c of codes(hex)) {
    if (c !== "0000") run += c;
    else {
      if (run) parts.push(`<${run}>`);
      run = "";
      parts.push(String(-width));
    }
  }
  if (run) parts.push(`<${run}>`);
  return parts.join(" ");
}

/** İçerik akışındaki Tj/TJ metinlerinden .notdef'i çıkarır; değişiklik yoksa null. */
export function rewriteNotdef(content: string, fonts: Map<string, number | null>): string | null {
  let width: number | null = null;
  let changed = false;
  const fix = (hex: string) => {
    if (width === null || !codes(hex).includes("0000")) return null;
    changed = true;
    return withoutNotdef(hex, width);
  };
  const out = content.replace(/\/([^\s/[\]<>()%{}]+)\s+-?[\d.]+\s+Tf\b|<([0-9A-Fa-f\s]*)>\s*Tj\b|\[((?:\s*(?:<[0-9A-Fa-f\s]*>|-?[\d.]+))*)\s*\]\s*TJ\b/g, (m, name?: string, tj?: string, arr?: string) => {
    if (name !== undefined) {
      width = fonts.get(name) ?? null;
      return m;
    }
    if (tj !== undefined) {
      const r = fix(tj);
      return r === null ? m : `[${r}] TJ`;
    }
    const items = arr!.replace(/<([0-9A-Fa-f\s]*)>/g, (h, hex: string) => fix(hex) ?? h);
    return items === arr ? m : `[${items}] TJ`;
  });
  return changed ? out : null;
}

/** Sayfa ve form içerik akışlarında .notdef glifini aynı genişlikte boşlukla değiştirir. */
function dropNotdef(doc: PDFDocument): void {
  const ctx = doc.context;
  rewriteContents(doc, (text, res) => {
    const fonts = new Map<string, number | null>();
    const d = ctx.lookupMaybe(res?.get(PDFName.of("Font")), PDFDict);
    for (const [k, v] of d?.entries() ?? []) {
      const f = ctx.lookup(v);
      fonts.set(k.decodeText(), f instanceof PDFDict ? notdefWidth(doc, f) : null);
    }
    return [...fonts.values()].some((w) => w !== null) ? (rewriteNotdef(text, fonts) ?? text) : text;
  });
}

/* ------------------------------------------------------------------ */
/* Görünmez boyama (opaklık 0)                                          */
/* ------------------------------------------------------------------ */

interface Token {
  start: number;
  end: number;
  op: boolean;
  text: string;
}

const WS = /[\0\t\n\f\r ]/;
const DELIM = /[\0\t\n\f\r ()<>[\]{}/%]/;

/** İçerik akışı belirteçleri (işleç / işlenen, konumlarıyla); satır içi görselde (BI) null. */
function tokenize(src: string): Token[] | null {
  const out: Token[] = [];
  let i = 0;
  const skipString = (j: number) => {
    let depth = 0;
    for (; j < src.length; j++) {
      const c = src[j];
      if (c === "\\") j++;
      else if (c === "(") depth++;
      else if (c === ")" && --depth === 0) return j + 1;
    }
    return j;
  };
  // Dizi ve sözlük tek işlenen sayılır (içindeki dizgeler dahil).
  const skipNested = (j: number, open: string, close: string) => {
    let depth = 0;
    while (j < src.length) {
      if (src.startsWith(open, j)) {
        depth++;
        j += open.length;
      } else if (src.startsWith(close, j)) {
        j += close.length;
        if (--depth === 0) return j;
      } else if (src[j] === "(") j = skipString(j);
      else j++;
    }
    return j;
  };
  while (i < src.length) {
    const c = src[i];
    if (WS.test(c)) {
      i++;
      continue;
    }
    if (c === "%") {
      while (i < src.length && src[i] !== "\n" && src[i] !== "\r") i++;
      continue;
    }
    let end: number;
    let op = false;
    if (c === "(") end = skipString(i);
    else if (src.startsWith("<<", i)) end = skipNested(i, "<<", ">>");
    else if (c === "<") end = src.indexOf(">", i) + 1 || src.length;
    else if (c === "[") end = skipNested(i, "[", "]");
    else if (c === "/") {
      end = i + 1;
      while (end < src.length && !DELIM.test(src[end])) end++;
    } else {
      end = i + 1;
      while (end < src.length && !DELIM.test(src[end])) end++;
      const word = src.slice(i, end);
      op = !/^[-+.\d]/.test(word) && word !== "true" && word !== "false" && word !== "null";
      if (word === "BI") return null;
    }
    out.push({ start: i, end, op, text: src.slice(i, end) });
    i = end;
  }
  return out;
}

interface PaintState {
  fill: number;
  stroke: number;
  mode: number;
}

/**
 * Opaklığı 0 olan (hiç görünmeyen) boyamaları içerik akışından çıkarır: dolgu/çizgi işleçleri `n` ya da
 * görünen yarısına, metin ve görsel çizimi silinir. Çizim birebir aynı kalır. Güvenle dönüştürülemezse null.
 */
export function stripInvisiblePaint(src: string, alpha: Map<string, { ca?: number; CA?: number }>): string | null {
  const toks = tokenize(src);
  if (!toks) return null;
  const edits: { from: number; to: number; text: string }[] = [];
  const stack: PaintState[] = [];
  let st: PaintState = { fill: 1, stroke: 1, mode: 0 };
  let args: Token[] = [];
  for (const t of toks) {
    if (!t.op) {
      args.push(t);
      continue;
    }
    const from = args[0]?.start ?? t.start;
    const fillOff = st.fill === 0;
    const strokeOff = st.stroke === 0;
    const replace = (text: string) => edits.push({ from: t.start, to: t.end, text });
    const remove = () => edits.push({ from, to: t.end, text: "" });
    switch (t.text) {
      case "q":
        stack.push({ ...st });
        break;
      case "Q":
        st = stack.pop() ?? st;
        break;
      case "gs": {
        const a = alpha.get(args[0]?.text.slice(1) ?? "");
        if (a?.ca !== undefined) st = { ...st, fill: a.ca };
        if (a?.CA !== undefined) st = { ...st, stroke: a.CA };
        break;
      }
      case "Tr":
        st = { ...st, mode: Number(args[0]?.text ?? 0) };
        break;
      case "f":
      case "F":
      case "f*":
        if (fillOff) replace("n");
        break;
      case "S":
      case "s":
        if (strokeOff) replace("n");
        break;
      case "B":
      case "B*":
      case "b":
      case "b*": {
        const close = t.text.startsWith("b");
        const even = t.text.endsWith("*") ? "*" : "";
        if (fillOff && strokeOff) replace("n");
        else if (fillOff) replace(close ? "s" : "S");
        else if (strokeOff) replace(close ? `h f${even}` : `f${even}`);
        break;
      }
      case "Tj":
      case "TJ":
      case "'":
      case '"': {
        const fills = [0, 2, 4, 6].includes(st.mode);
        const strokes = [1, 2, 5, 6].includes(st.mode);
        const visible = (fills && !fillOff) || (strokes && !strokeOff);
        const partly = (fills && fillOff) || (strokes && strokeOff);
        if (!visible && partly) {
          // Satır atlayan metin işleçlerinde konum etkisi kalmalı: dönüştürülmez.
          if (t.text !== "Tj" && t.text !== "TJ") return null;
          remove();
        } else if (partly) return null;
        break;
      }
      case "Do":
      case "sh":
        if (fillOff) remove();
        break;
    }
    args = [];
  }
  if (!edits.length) return src;
  let out = src;
  for (const e of edits.sort((x, y) => y.from - x.from)) out = out.slice(0, e.from) + e.text + out.slice(e.to);
  return out;
}

/**
 * PDF/A-1b hazırlığı: opaklığı 0 olan grafik durumlarıyla yapılan görünmez çizimler (ör. kenarlıksız tablo
 * hücrelerinin saydam kenarlıkları) kaldırılır ve bu grafik durumları opak yapılır. Hepsi güvenle
 * dönüştürülemezse belge değişmez (saydamlık olarak bildirilir).
 */
export function removeInvisiblePaint(doc: PDFDocument): boolean {
  const ctx = doc.context;
  const targets = new Set<PDFDict>();
  const alphaOf = (res: PDFDict | undefined) => {
    const m = new Map<string, { ca?: number; CA?: number }>();
    const d = ctx.lookupMaybe(res?.get(PDFName.of("ExtGState")), PDFDict);
    for (const [k, v] of d?.entries() ?? []) {
      const g = ctx.lookup(v);
      if (!(g instanceof PDFDict)) continue;
      const a = { ca: num(g.get(PDFName.of("ca"))), CA: num(g.get(PDFName.of("CA"))) };
      if (a.ca === 0 || a.CA === 0) targets.add(g);
      m.set(k.decodeText(), a);
    }
    return m;
  };
  const pending: { ref: PDFObject; stream: PDFRawStream; text: string }[] = [];
  const handled = new Set<PDFObject>();
  const visit = (ref: PDFObject, res: PDFDict | undefined): boolean => {
    const s = ctx.lookup(ref);
    if (!(s instanceof PDFRawStream) || handled.has(s)) return true;
    handled.add(s);
    const alpha = alphaOf(res);
    const xo = ctx.lookupMaybe(res?.get(PDFName.of("XObject")), PDFDict);
    for (const [, v] of xo?.entries() ?? []) {
      const f = ctx.lookup(v);
      if (f instanceof PDFRawStream && nameOf(f.dict.get(PDFName.of("Subtype"))) === "Form") {
        if (!visit(v, ctx.lookupMaybe(f.dict.get(PDFName.of("Resources")), PDFDict) ?? res)) return false;
      }
    }
    if (![...alpha.values()].some((a) => a.ca === 0 || a.CA === 0)) return true;
    const text = Buffer.from(decodePDFRawStream(s).decode()).toString("latin1");
    const fixed = stripInvisiblePaint(text, alpha);
    if (fixed === null) return false;
    if (fixed !== text) pending.push({ ref, stream: s, text: fixed });
    return true;
  };
  for (const page of doc.getPages()) {
    const res = page.node.Resources();
    const c = page.node.get(PDFName.of("Contents"));
    for (const r of c instanceof PDFArray ? c.asArray() : c ? [c] : []) if (!visit(r, res)) return false;
  }
  if (!targets.size) return true;
  // Hedef grafik durumlarını kullanan ama gezilmeyen akış (desen, Type3, açıklama görünümü) varsa dokunulmaz.
  for (const [, obj] of ctx.enumerateIndirectObjects()) {
    const d = obj instanceof PDFRawStream ? obj.dict : obj instanceof PDFDict ? obj : null;
    if (!d) continue;
    // Sayfalar ve gezilen formlar dönüştürüldü.
    if (obj instanceof PDFDict ? nameOf(d.get(PDFName.of("Type"))) === "Page" : handled.has(obj)) continue;
    const res = ctx.lookupMaybe(d.get(PDFName.of("Resources")), PDFDict);
    const gs = ctx.lookupMaybe(res?.get(PDFName.of("ExtGState")), PDFDict);
    if (gs?.entries().some(([, v]) => targets.has(ctx.lookup(v) as PDFDict))) return false;
  }
  for (const p of pending) {
    const next = ctx.flateStream(Buffer.from(p.text, "latin1"));
    for (const [k, v] of p.stream.dict.entries()) if (!["Length", "Filter", "DecodeParms"].includes(k.decodeText())) next.dict.set(k, v);
    ctx.assign(p.ref as Parameters<typeof ctx.assign>[0], next);
  }
  for (const g of targets) {
    if (num(g.get(PDFName.of("ca"))) === 0) g.set(PDFName.of("ca"), PDFNumber.of(1));
    if (num(g.get(PDFName.of("CA"))) === 0) g.set(PDFName.of("CA"), PDFNumber.of(1));
  }
  return true;
}

/* ------------------------------------------------------------------ */
/* PDF/A-1 sınırları: ICC v4 profilleri ve 32767'yi aşan sayılar        */
/* ------------------------------------------------------------------ */

/** ICC profilinin sürümü ve kırmızı/yeşil/mavi birincil renkleri (XYZ, D50). */
function iccInfo(icc: Uint8Array): { major: number; primaries: number[] | null } {
  const v = new DataView(icc.buffer, icc.byteOffset, icc.byteLength);
  if (icc.length < 132) return { major: 0, primaries: null };
  const tags = new Map<string, number>();
  const count = v.getUint32(128);
  for (let i = 0; i < count && 132 + 12 * i + 12 <= icc.length; i++) {
    const o = 132 + 12 * i;
    tags.set(String.fromCharCode(icc[o], icc[o + 1], icc[o + 2], icc[o + 3]), v.getUint32(o + 4));
  }
  const xyz = (sig: string) => {
    const o = tags.get(sig);
    return o === undefined || o + 20 > icc.length ? null : [0, 1, 2].map((k) => v.getInt32(o + 8 + 4 * k) / 65536);
  };
  const rgb = [xyz("rXYZ"), xyz("gXYZ"), xyz("bXYZ")];
  return { major: icc[8], primaries: rgb.every(Boolean) ? rgb.flat() as number[] : null };
}

/**
 * PDF/A-1 (PDF 1.4) ICC v4 profilini kabul etmez. Skia renk profilsiz görsellere v4 sRGB profili gömer:
 * birincil renkleri sRGB olan v4 profiller pakettaki sRGB v2 profiliyle değiştirilir (renk aynı kalır).
 * Kalan (sRGB dışı) v4 profil sayısını döndürür.
 */
function downgradeIcc(doc: PDFDocument, srgb: Uint8Array): number {
  const ctx = doc.context;
  const ref = iccInfo(srgb).primaries!;
  const profiles = new Set<string>();
  for (const [, obj] of ctx.enumerateIndirectObjects()) {
    const scan = (o: PDFObject | undefined) => {
      const a = ctx.lookup(o);
      if (a instanceof PDFArray && nameOf(a.get(0)) === "ICCBased") profiles.add(a.get(1).toString());
      else if (a instanceof PDFArray && nameOf(a.get(0)) === "Indexed") scan(a.get(1));
    };
    const d = obj instanceof PDFRawStream ? obj.dict : obj instanceof PDFDict ? obj : null;
    if (obj instanceof PDFArray) scan(obj);
    if (!d) continue;
    scan(d.get(PDFName.of("ColorSpace")));
    const cs = ctx.lookupMaybe(ctx.lookupMaybe(d.get(PDFName.of("Resources")), PDFDict)?.get(PDFName.of("ColorSpace")), PDFDict);
    for (const [, v] of cs?.entries() ?? []) scan(v);
  }
  let left = 0;
  for (const [r, obj] of ctx.enumerateIndirectObjects()) {
    if (!profiles.has(r.toString()) || !(obj instanceof PDFRawStream)) continue;
    const info = iccInfo(decodePDFRawStream(obj).decode());
    if (info.major < 4) continue;
    if (num(obj.dict.get(PDFName.of("N"))) === 3 && info.primaries?.every((x, i) => Math.abs(x - ref[i]) < 0.002)) ctx.assign(r, ctx.flateStream(srgb, { N: 3 }));
    else left++;
  }
  return left;
}

const LIMIT = 32767;

/** `cm` ötelemesi sınırı aşıyorsa öteleme parçalara bölünür (dönüşüm aynı kalır). */
export function splitLargeMatrices(src: string): string | null {
  const toks = tokenize(src);
  if (!toks) return null;
  const edits: { from: number; to: number; text: string }[] = [];
  let args: Token[] = [];
  for (const t of toks) {
    if (!t.op) {
      args.push(t);
      continue;
    }
    if (t.text === "cm" && args.length === 6) {
      const [a, b, c, d, e, f] = args.map((x) => Number(x.text));
      if ([a, b, c, d].every((x) => Math.abs(x) <= LIMIT) && (Math.abs(e) > LIMIT || Math.abs(f) > LIMIT)) {
        // M = L × T(e, f): önce öteleme parçaları, sonra doğrusal kısım.
        const n = Math.ceil(Math.max(Math.abs(e), Math.abs(f)) / (LIMIT - 1));
        const fmt = (x: number) => String(Math.round(x * 1e6) / 1e6);
        const step = `1 0 0 1 ${fmt(e / n)} ${fmt(f / n)} cm\n`;
        edits.push({ from: args[0].start, to: t.end, text: `${step.repeat(n)}${a} ${b} ${c} ${d} 0 0 cm` });
      }
    }
    args = [];
  }
  let out = src;
  for (const e of edits.reverse()) out = out.slice(0, e.from) + e.text + out.slice(e.to);
  return out;
}

/** Sayfa ve form içerik akışlarına (kendi kaynaklarıyla) bir metin dönüşümü uygular; dönüştürülemeyen akış sayısı. */
function rewriteContents(doc: PDFDocument, fn: (text: string, res: PDFDict | undefined) => string | null): number {
  const ctx = doc.context;
  const seen = new Set<PDFObject>();
  let failed = 0;
  const visit = (ref: PDFObject, res: PDFDict | undefined) => {
    const s = ctx.lookup(ref);
    if (!(s instanceof PDFRawStream) || seen.has(s)) return;
    seen.add(s);
    const xo = ctx.lookupMaybe(res?.get(PDFName.of("XObject")), PDFDict);
    for (const [, v] of xo?.entries() ?? []) {
      const f = ctx.lookup(v);
      if (f instanceof PDFRawStream && nameOf(f.dict.get(PDFName.of("Subtype"))) === "Form") visit(v, ctx.lookupMaybe(f.dict.get(PDFName.of("Resources")), PDFDict) ?? res);
    }
    const text = Buffer.from(decodePDFRawStream(s).decode()).toString("latin1");
    const fixed = fn(text, res);
    if (fixed === null) failed++;
    if (fixed === null || fixed === text) return;
    const next = ctx.flateStream(Buffer.from(fixed, "latin1"));
    for (const [k, v] of s.dict.entries()) if (!["Length", "Filter", "DecodeParms"].includes(k.decodeText())) next.dict.set(k, v);
    ctx.assign(ref as Parameters<typeof ctx.assign>[0], next);
  };
  for (const page of doc.getPages()) {
    const c = page.node.get(PDFName.of("Contents"));
    for (const r of c instanceof PDFArray ? c.asArray() : c ? [c] : []) visit(r, page.node.Resources());
  }
  return failed;
}

const REAL = /^[-+]?(\d+\.\d*|\.\d+)$/;

/** İçerik akışında |değer| > 32767 olan ondalıklı sayı var mı (PDF/A-1 sınırı; tamsayı sınırı 2³¹). */
export function hasLargeReal(src: string): boolean {
  for (const t of tokenize(src) ?? []) {
    if (t.op) continue;
    // Dizilerde (TJ kaydırmaları) dizgeler atlanır; sözlükler (işaretli içerik) tamsayı taşır.
    const parts = t.text.startsWith("[") ? t.text.replace(/<[^>]*>|\((?:\\.|[^\\)])*\)/g, " ").replace(/[[\]]/g, " ").split(/\s+/) : [t.text];
    if (parts.some((x) => REAL.test(x) && Math.abs(Number(x)) > LIMIT)) return true;
  }
  return false;
}

function exceedsRealLimit(doc: PDFDocument): boolean {
  let over = false;
  rewriteContents(doc, (text) => {
    over ||= hasLargeReal(text);
    return text;
  });
  return over;
}

/**
 * PDF/A-1b hazırlığı: görünmez çizimler, sRGB v4 profilleri ve büyük ötelemeler dönüştürülür; kalan engeller
 * (gerçek saydamlık, sRGB dışı v4 profil, sayı sınırını aşan konumlar) kullanıcıya gösterilecek satırlar olarak döner.
 */
export function prepareForPdfA1(doc: PDFDocument, srgb: Uint8Array): string[] {
  removeInvisiblePaint(doc);
  const icc = downgradeIcc(doc, srgb);
  rewriteContents(doc, (text) => splitLargeMatrices(text));
  const out: string[] = [];
  const t = findTransparency(doc);
  if (t.length) out.push(describeTransparency(t));
  if (icc) out.push(`• ${icc} görselde sRGB dışı (ICC v4) renk profili var.`);
  if (exceedsRealLimit(doc)) out.push("• Sayfa içi konumlar PDF/A-1'in sayı sınırını aşıyor (çok sayfalı uzun belge).");
  return out;
}

/* ------------------------------------------------------------------ */
/* Meta veri ve PDF/A                                                   */
/* ------------------------------------------------------------------ */

const xml = (s: string) => s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
const isoDate = (d: Date) => d.toISOString().replace(/\.\d+Z$/, "Z");

/** XMP paketi: Info sözlüğüyle birebir aynı değerler (PDF/A bunları karşılaştırır). */
export function xmpPacket(meta: PdfMeta, date: Date, part?: 1 | 2): string {
  const alt = (v: string) => `<rdf:Alt><rdf:li xml:lang="x-default">${xml(v)}</rdf:li></rdf:Alt>`;
  const lines = [
    `<dc:format>application/pdf</dc:format>`,
    meta.title ? `<dc:title>${alt(meta.title)}</dc:title>` : "",
    meta.author ? `<dc:creator><rdf:Seq><rdf:li>${xml(meta.author)}</rdf:li></rdf:Seq></dc:creator>` : "",
    meta.subject ? `<dc:description>${alt(meta.subject)}</dc:description>` : "",
    meta.keywords ? `<pdf:Keywords>${xml(meta.keywords)}</pdf:Keywords>` : "",
    meta.producer ? `<pdf:Producer>${xml(meta.producer)}</pdf:Producer>` : "",
    meta.creator ? `<xmp:CreatorTool>${xml(meta.creator)}</xmp:CreatorTool>` : "",
    `<xmp:CreateDate>${isoDate(date)}</xmp:CreateDate>`,
    `<xmp:ModifyDate>${isoDate(date)}</xmp:ModifyDate>`,
    `<xmp:MetadataDate>${isoDate(date)}</xmp:MetadataDate>`,
    part ? `<pdfaid:part>${part}</pdfaid:part><pdfaid:conformance>B</pdfaid:conformance>` : "",
  ].filter(Boolean);
  return [
    `<?xpacket begin="﻿" id="W5M0MpCehiHzreSzNTczkc9d"?>`,
    `<x:xmpmeta xmlns:x="adobe:ns:meta/"><rdf:RDF xmlns:rdf="http://www.w3.org/1999/02/22-rdf-syntax-ns#">`,
    `<rdf:Description rdf:about="" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:pdf="http://ns.adobe.com/pdf/1.3/" xmlns:xmp="http://ns.adobe.com/xap/1.0/"${part ? ` xmlns:pdfaid="http://www.aiim.org/pdfa/ns/id/"` : ""}>`,
    ...lines,
    `</rdf:Description></rdf:RDF></x:xmpmeta>`,
    // Yerinde düzenleme için boşluk (XMP önerisi).
    " ".repeat(1024),
    `<?xpacket end="w"?>`,
  ].join("\n");
}

/** sRGB IEC61966-2.1 ICC v2 profili (International Color Consortium; değiştirilmeden dağıtılabilir, bkz. HUKUKI.md). */
export function loadSrgbProfile(): Uint8Array {
  // Derlemede dist-electron/main/srgb altına kopyalanır; geliştirmede ve testte build/srgb'den okunur.
  const candidates = [path.join(__dirname, "srgb", "sRGB2014.icc"), path.join(__dirname, "..", "..", "build", "srgb", "sRGB2014.icc")];
  for (const p of candidates) {
    try {
      const icc = readFileSync(p);
      if (icc.length > 128 && icc.toString("latin1", 36, 40) === "acsp") return icc;
    } catch {
      // sonraki aday
    }
  }
  throw Object.assign(new Error("sRGB renk profili bulunamadı"), { code: "ENOENT" });
}

export interface ApplyOptions {
  standard: PdfStandard;
  meta: PdfMeta;
  /** PDF/A için zorunlu (sRGB ICC). */
  icc?: Uint8Array;
  date?: Date;
  /** Trailer /ID (16 bayt); verilmezse rastgele. */
  id?: Uint8Array;
}

function setInfo(doc: PDFDocument, meta: PdfMeta, date: Date) {
  const ctx = doc.context;
  let info = ctx.lookupMaybe(ctx.trailerInfo.Info, PDFDict);
  if (!info) {
    info = ctx.obj({});
    ctx.trailerInfo.Info = ctx.register(info);
  }
  for (const k of ["Title", "Author", "Subject", "Keywords", "Creator", "Producer", "CreationDate", "ModDate", "Trapped"]) info.delete(PDFName.of(k));
  const put = (k: string, v?: string) => v && info.set(PDFName.of(k), PDFHexString.fromText(v));
  put("Title", meta.title);
  put("Author", meta.author);
  put("Subject", meta.subject);
  put("Keywords", meta.keywords);
  put("Creator", meta.creator);
  put("Producer", meta.producer);
  doc.setCreationDate(date);
  doc.setModificationDate(date);
  // Pencere başlığında dosya adı yerine belge başlığı.
  if (meta.title) doc.catalog.getOrCreateViewerPreferences().setDisplayDocTitle(true);
}

/** Belge bilgilerini yazar; PDF/A seçiliyse XMP, OutputIntent, /ID ve (1b) CIDSet ekler. */
export function applyPdfStandard(doc: PDFDocument, o: ApplyOptions): void {
  const date = o.date ?? new Date();
  setInfo(doc, o.meta, date);
  if (o.standard === "pdf") return;
  if (!o.icc) throw new Error("PDF/A için renk profili gerekli");
  const ctx = doc.context;
  const part = o.standard === "pdfa-1b" ? 1 : 2;
  const meta = ctx.stream(new TextEncoder().encode(xmpPacket(o.meta, date, part)), { Type: "Metadata", Subtype: "XML" });
  doc.catalog.set(PDFName.of("Metadata"), ctx.register(meta));
  const profile = ctx.register(ctx.flateStream(o.icc, { N: 3 }));
  const intent = ctx.obj({
    Type: "OutputIntent",
    S: "GTS_PDFA1",
    OutputConditionIdentifier: PDFString.of("sRGB IEC61966-2.1"),
    RegistryName: PDFString.of("http://www.color.org"),
    Info: PDFString.of("sRGB IEC61966-2.1"),
    DestOutputProfile: profile,
  });
  doc.catalog.set(PDFName.of("OutputIntents"), ctx.obj([intent]));
  const id = PDFHexString.of(Buffer.from(o.id ?? randomBytes(16)).toString("hex"));
  ctx.trailerInfo.ID = ctx.obj([id, id]);
  // Açıklamalar (bağlantılar) yazdırılır ve görünür olmalı; görsellerde yumuşatma (Interpolate) kapalı.
  for (const page of doc.getPages()) {
    for (const a of page.node.Annots()?.asArray() ?? []) {
      const d = ctx.lookupMaybe(a, PDFDict);
      if (!d) continue;
      const f = num(d.get(PDFName.of("F"))) ?? 0;
      d.set(PDFName.of("F"), PDFNumber.of((f | 4) & ~(1 | 2 | 32)));
    }
  }
  for (const [, obj] of ctx.enumerateIndirectObjects()) {
    if (!(obj instanceof PDFStream)) continue;
    const d = obj.dict;
    if (d.get(PDFName.of("Interpolate"))?.toString() === "true") d.set(PDFName.of("Interpolate"), ctx.obj(false));
    // 1b: saydam içerik yoksa kalan boş saydamlık grupları da kaldırılır (çizim değişmez).
    if (part === 1 && nameOf(ctx.lookupMaybe(d.get(PDFName.of("Group")), PDFDict)?.get(PDFName.of("S"))) === "Transparency") d.delete(PDFName.of("Group"));
  }
  dropNotdef(doc);
  if (part === 1) {
    prepareForPdfA1(doc, o.icc);
    for (const page of doc.getPages()) {
      const g = ctx.lookupMaybe(page.node.get(PDFName.of("Group")), PDFDict);
      if (nameOf(g?.get(PDFName.of("S"))) === "Transparency") page.node.delete(PDFName.of("Group"));
    }
    addCidSets(doc);
  }
}

export function loadPdf(bytes: Uint8Array): Promise<PDFDocument> {
  return PDFDocument.load(bytes, { updateMetadata: false });
}

/** Nesne akışı yok (PDF/A-1 PDF 1.4 tabanlıdır; çapraz başvuru tablosu klasik kalır). */
export function savePdf(doc: PDFDocument): Promise<Uint8Array> {
  return doc.save({ useObjectStreams: false, addDefaultPage: false, updateFieldAppearances: false });
}
