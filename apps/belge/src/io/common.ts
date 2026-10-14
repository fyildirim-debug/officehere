// Biçim okuyucu/yazıcılarının ortak yardımcıları: karakter özellikleri ↔ işaretler, liste katlama,
// resim verisi, uzunluk birimleri ve düğüm gezintisi (ODT, RTF, Markdown, DOC).
import type { JSONContent } from "@tiptap/core";
import { indexLine, type IndexItem } from "../editor/indexing";
import { effectiveProps, resolveStyles, type StyleDef, type StyleProps } from "../model/styles";
import type { DocSettings } from "../model/document";

export type Mark = { type: string; attrs?: Record<string, unknown> };

/* ------------------------------------------------------------------ */
/* Karakter özellikleri                                                */
/* ------------------------------------------------------------------ */

export type UnderlineKind = "single" | "double" | "thick" | "dotted" | "dashed" | "wave" | "words";

/** Biçimlerden bağımsız karakter özellikleri (okuyucular doldurur, yazıcılar işaretlerden çıkarır). */
export interface CharProps {
  bold?: boolean;
  italic?: boolean;
  underline?: UnderlineKind | null;
  underlineColor?: string | null;
  strike?: boolean;
  dstrike?: boolean;
  sup?: boolean;
  sub?: boolean;
  font?: string | null;
  /** Yazı boyutu (pt). */
  size?: number | null;
  color?: string | null;
  highlight?: string | null;
  /** Karakter gölgelendirmesi (vurgudan ayrı). */
  bg?: string | null;
  caps?: boolean;
  smallCaps?: boolean;
  hidden?: boolean;
  /** Karakter aralığı (pt). */
  spacing?: number | null;
  shadow?: boolean;
  outline?: boolean;
  emboss?: boolean;
  imprint?: boolean;
  link?: string | null;
  charStyle?: string | null;
}

/** Karakter özelliklerinden ProseMirror işaretleri. */
export function charMarks(p: CharProps): Mark[] {
  const m: Mark[] = [];
  if (p.link) m.push({ type: "link", attrs: { href: p.link } });
  if (p.bold) m.push({ type: "bold" });
  if (p.italic) m.push({ type: "italic" });
  if (p.underline) m.push({ type: "underline", attrs: { type: p.underline, color: p.underlineColor ?? null } });
  if (p.strike) m.push({ type: "strike" });
  if (p.sup) m.push({ type: "superscript" });
  else if (p.sub) m.push({ type: "subscript" });
  if (p.highlight) m.push({ type: "highlight", attrs: { color: p.highlight } });
  if (p.charStyle) m.push({ type: "charStyle", attrs: { styleId: p.charStyle } });
  const ts: Record<string, unknown> = {};
  if (p.font) ts.fontFamily = p.font;
  if (p.size) ts.fontSize = `${+p.size.toFixed(2)}pt`;
  if (p.color) ts.color = p.color;
  if (p.bg) ts.backgroundColor = p.bg;
  if (p.caps) ts.caps = true;
  if (p.smallCaps) ts.smallCaps = true;
  if (p.hidden) ts.hidden = true;
  if (p.dstrike) ts.doubleStrike = true;
  if (p.spacing) ts.letterSpacing = p.spacing;
  for (const k of ["shadow", "outline", "emboss", "imprint"] as const) if (p[k]) ts[k] = true;
  if (Object.keys(ts).length) m.push({ type: "textStyle", attrs: ts });
  return m;
}

/** İşaretlerden karakter özellikleri (yazıcılar için). */
export function marksToProps(marks: Mark[] | undefined): CharProps {
  const p: CharProps = {};
  for (const m of marks ?? []) {
    const a = m.attrs ?? {};
    switch (m.type) {
      case "bold":
        p.bold = true;
        break;
      case "italic":
        p.italic = true;
        break;
      case "underline":
        p.underline = ((a.type as UnderlineKind | undefined) ?? "single") || "single";
        p.underlineColor = (a.color as string | null) ?? null;
        break;
      case "strike":
        p.strike = true;
        break;
      case "superscript":
        p.sup = true;
        break;
      case "subscript":
        p.sub = true;
        break;
      case "highlight":
        p.highlight = (a.color as string | null) ?? "#FFFF00";
        break;
      case "link":
        p.link = (a.href as string | null) ?? null;
        break;
      case "charStyle":
        p.charStyle = (a.styleId as string | null) ?? null;
        break;
      case "textStyle": {
        if (a.fontFamily) p.font = String(a.fontFamily);
        const fs = parseFloat(String(a.fontSize ?? ""));
        if (Number.isFinite(fs) && fs > 0) p.size = String(a.fontSize).endsWith("px") ? fs * 0.75 : fs;
        if (a.color) p.color = normColor(String(a.color));
        if (a.backgroundColor) p.bg = normColor(String(a.backgroundColor));
        if (a.caps) p.caps = true;
        if (a.smallCaps) p.smallCaps = true;
        if (a.hidden) p.hidden = true;
        if (a.doubleStrike) p.dstrike = true;
        if (a.letterSpacing) p.spacing = Number(a.letterSpacing) || null;
        for (const k of ["shadow", "outline", "emboss", "imprint"] as const) if (a[k]) p[k] = true;
        break;
      }
    }
  }
  return p;
}

/** Eş genişlikli yazı tipi mi (Markdown kod parçası karşılığı). */
export function isMonoFont(f: string | null | undefined): boolean {
  return !!f && /^(courier( new)?|consolas|liberation mono|dejavu sans mono|cascadia (code|mono)|menlo|monaco|lucida console|source code pro|noto sans mono|monospace)$/i.test(f.trim());
}

export const MONO_FONT = "Courier New";

/* ------------------------------------------------------------------ */
/* Stil temelleri                                                      */
/* ------------------------------------------------------------------ */

let defaultStyles: Record<string, StyleDef> | null = null;

/** Belgenin stil tablosu (ayar verilmezse varsayılan küme). */
export function stylesOf(s?: DocSettings): Record<string, StyleDef> {
  if (s) return resolveStyles(s.styleSet, s.theme, s.fontTheme, s.paragraphSpacing, s.customStyles);
  defaultStyles ??= resolveStyles("default", "OfficeHere");
  return defaultStyles;
}

/** Paragraf düğümünün stil kimliği (başlık → HeadingN). */
export function nodeStyleId(n: JSONContent): string {
  if (n.type === "heading") return `Heading${Math.min(6, Math.max(1, Number(n.attrs?.level) || 1))}`;
  return (n.attrs?.styleId as string | null) || "Normal";
}

/** Stilin karakter özellikleri (yazıcılar paragraf temeli olarak açıkça yazar). */
export function styleChar(styles: Record<string, StyleDef>, id: string): CharProps & { base: StyleProps } {
  const e = effectiveProps(styles, styles[id] ? id : "Normal");
  return { font: e.font ?? null, size: e.size ?? 11, bold: !!e.bold, italic: !!e.italic, color: e.color ?? null, caps: !!e.caps, base: e };
}

/** Okuyucular: stilin zaten verdiği karakter özellikleri doğrudan biçimden atılır. */
export function stripStyleProps(p: CharProps, base: StyleProps): CharProps {
  const out = { ...p };
  const norm = (f: string | null | undefined) => (f ?? "").toLowerCase().replace(/\s+/g, " ").trim();
  if (out.font && norm(out.font) === norm(base.font)) delete out.font;
  if (out.size && Math.abs(out.size - (base.size ?? 11)) < 0.01) delete out.size;
  if (out.bold && base.bold) delete out.bold;
  if (out.italic && base.italic) delete out.italic;
  if (out.caps && base.caps) delete out.caps;
  if (out.color && base.color && out.color.toUpperCase() === base.color.toUpperCase()) delete out.color;
  return out;
}

/* ------------------------------------------------------------------ */
/* Renk ve birimler                                                    */
/* ------------------------------------------------------------------ */

/** CSS renk → "#RRGGBB" (geçersizse null). */
export function normColor(c: string | null | undefined): string | null {
  if (!c) return null;
  const s = c.trim();
  let m = /^#?([0-9a-f]{6})$/i.exec(s);
  if (m) return `#${m[1].toUpperCase()}`;
  m = /^#([0-9a-f]{3})$/i.exec(s);
  if (m) return `#${m[1].split("").map((x) => x + x).join("").toUpperCase()}`;
  m = /^rgba?\((\d+)\s*,\s*(\d+)\s*,\s*(\d+)/i.exec(s);
  if (m) return `#${[m[1], m[2], m[3]].map((x) => Math.min(255, Number(x)).toString(16).padStart(2, "0")).join("").toUpperCase()}`;
  return null;
}

/** Uzunluk dizgesi ("2.5cm", "12pt", "0.5in", "10mm", "16px", "1pc") → pt. */
export function lengthPt(v: string | null | undefined): number | null {
  if (!v) return null;
  const m = /^\s*(-?[\d.]+)\s*(cm|mm|in|pt|px|pc)?\s*$/i.exec(v);
  if (!m) return null;
  const n = parseFloat(m[1]);
  if (!Number.isFinite(n)) return null;
  switch ((m[2] ?? "pt").toLowerCase()) {
    case "cm":
      return (n / 2.54) * 72;
    case "mm":
      return (n / 25.4) * 72;
    case "in":
      return n * 72;
    case "px":
      return n * 0.75;
    case "pc":
      return n * 12;
    default:
      return n;
  }
}

export const round2 = (n: number) => Math.round(n * 100) / 100;
/** pt → twip */
export const ptTw = (pt: number) => Math.round(pt * 20);
/** pt → "1.234cm" */
export const ptCm = (pt: number) => `${+((pt / 72) * 2.54).toFixed(4)}cm`;
/** twip → "1.234cm" */
export const twCm = (tw: number) => `${+((tw / 1440) * 2.54).toFixed(4)}cm`;
/** px → "1.234cm" */
export const pxCm = (px: number) => `${+((px / 96) * 2.54).toFixed(4)}cm`;

/* ------------------------------------------------------------------ */
/* Resim verisi                                                        */
/* ------------------------------------------------------------------ */

export function toBase64(bytes: Uint8Array): string {
  let bin = "";
  for (let i = 0; i < bytes.length; i += 0x8000) bin += String.fromCharCode(...bytes.subarray(i, i + 0x8000));
  return btoa(bin);
}

export function fromBase64(s: string): Uint8Array {
  const bin = atob(s.replace(/\s+/g, ""));
  const out = new Uint8Array(bin.length);
  for (let i = 0; i < bin.length; i++) out[i] = bin.charCodeAt(i);
  return out;
}

/** "data:" adresi → MIME türü ve baytlar. */
export function parseDataUrl(src: string): { mime: string; bytes: Uint8Array } | null {
  const m = /^data:([^;,]*)((?:;[^;,]*)*),(.*)$/s.exec(src);
  if (!m) return null;
  try {
    const bytes = /;base64/i.test(m[2]) ? fromBase64(m[3]) : new TextEncoder().encode(decodeURIComponent(m[3]));
    return { mime: (m[1] || "application/octet-stream").toLowerCase(), bytes };
  } catch {
    return null;
  }
}

/** Baytların imzasından resim türü. */
export function sniffImage(b: Uint8Array): { mime: string; ext: string } | null {
  if (b.length > 8 && b[0] === 0x89 && b[1] === 0x50 && b[2] === 0x4e && b[3] === 0x47) return { mime: "image/png", ext: "png" };
  if (b.length > 3 && b[0] === 0xff && b[1] === 0xd8 && b[2] === 0xff) return { mime: "image/jpeg", ext: "jpg" };
  if (b.length > 6 && b[0] === 0x47 && b[1] === 0x49 && b[2] === 0x46) return { mime: "image/gif", ext: "gif" };
  if (b.length > 2 && b[0] === 0x42 && b[1] === 0x4d) return { mime: "image/bmp", ext: "bmp" };
  if (b.length > 12 && b[0] === 0x52 && b[1] === 0x49 && b[8] === 0x57 && b[9] === 0x45) return { mime: "image/webp", ext: "webp" };
  const head = new TextDecoder("latin1").decode(b.subarray(0, Math.min(256, b.length)));
  if (/<svg[\s>]/i.test(head) || /^\s*<\?xml[^>]*>\s*(<!--[\s\S]*?-->\s*)*<svg/i.test(head)) return { mime: "image/svg+xml", ext: "svg" };
  return null;
}

/** Resmin piksel boyutu (PNG, JPEG, GIF, BMP, WebP); bilinmiyorsa null. */
export function imageSize(b: Uint8Array): { w: number; h: number } | null {
  const dv = new DataView(b.buffer, b.byteOffset, b.byteLength);
  const t = sniffImage(b);
  try {
    if (t?.ext === "png" && b.length >= 24) return { w: dv.getUint32(16), h: dv.getUint32(20) };
    if (t?.ext === "gif" && b.length >= 10) return { w: dv.getUint16(6, true), h: dv.getUint16(8, true) };
    if (t?.ext === "bmp" && b.length >= 26) return { w: Math.abs(dv.getInt32(18, true)), h: Math.abs(dv.getInt32(22, true)) };
    if (t?.ext === "jpg") {
      let o = 2;
      // SOF işaretçisini bulana dek kesimleri atla (sınırlı adım).
      for (let guard = 0; o + 9 < b.length && guard < 1000; guard++) {
        if (b[o] !== 0xff) return null;
        const marker = b[o + 1];
        const len = dv.getUint16(o + 2);
        if (marker >= 0xc0 && marker <= 0xcf && marker !== 0xc4 && marker !== 0xc8 && marker !== 0xcc) return { w: dv.getUint16(o + 7), h: dv.getUint16(o + 5) };
        o += 2 + len;
      }
    }
    if (t?.ext === "webp" && b.length >= 30) {
      const kind = String.fromCharCode(b[12], b[13], b[14], b[15]);
      if (kind === "VP8X") return { w: 1 + (b[24] | (b[25] << 8) | (b[26] << 16)), h: 1 + (b[27] | (b[28] << 8) | (b[29] << 16)) };
      if (kind === "VP8 ") return { w: dv.getUint16(26, true) & 0x3fff, h: dv.getUint16(28, true) & 0x3fff };
      if (kind === "VP8L") {
        const bits = dv.getUint32(21, true);
        return { w: (bits & 0x3fff) + 1, h: ((bits >> 14) & 0x3fff) + 1 };
      }
    }
  } catch {
    return null;
  }
  return null;
}

/** Görüntü boyutunu sayfa genişliğine sığdırır (px). */
export function fitSize(w: number, h: number, max = 600): { width: number; height: number } {
  if (!(w > 0) || !(h > 0)) return { width: 200, height: 150 };
  const k = w > max ? max / w : 1;
  return { width: Math.max(1, Math.round(w * k)), height: Math.max(1, Math.round(h * k)) };
}

/** Bayt dizisinden resim düğümü (boyut verilmezse dosyadan okunur). */
export function imageNode(bytes: Uint8Array, opts: { mime?: string; width?: number | null; height?: number | null; alt?: string | null } = {}): JSONContent | null {
  const t = sniffImage(bytes);
  const mime = t?.mime ?? opts.mime;
  if (!mime || !mime.startsWith("image/")) return null;
  let width = opts.width ?? null;
  let height = opts.height ?? null;
  if (!width || !height) {
    const s = imageSize(bytes);
    const f = s ? fitSize(s.w, s.h) : { width: 200, height: 150 };
    width = f.width;
    height = f.height;
  }
  return { type: "image", attrs: { src: `data:${mime};base64,${toBase64(bytes)}`, width: Math.round(width), height: Math.round(height), alt: opts.alt || null } };
}

/* ------------------------------------------------------------------ */
/* Düğüm yardımcıları                                                  */
/* ------------------------------------------------------------------ */

/**
 * Noktaya bağlı açıklamanın işaretini sözcüğün tamamına yayar: sözcük birden çok koşuya bölünmüşse
 * (ör. RTF \\u kaçışları, biçim değişikliği) işaretli metne boşluksuz bitişen parçalar da işaretlenir.
 */
export function extendCommentToWord(nodes: JSONContent[], ids: Iterable<string>): void {
  const has = (n: JSONContent | undefined, id: string) => !!n && n.type === "text" && !!n.marks?.some((m) => m.type === "comment" && m.attrs?.id === id);
  const add = (n: JSONContent, id: string): JSONContent => ({ ...n, marks: [...(n.marks ?? []), { type: "comment", attrs: { id } }] });
  for (const id of ids) {
    const first = nodes.findIndex((n) => has(n, id));
    if (first < 0) continue;
    let last = first;
    while (has(nodes[last + 1], id)) last++;
    // İleri: işaretli metin boşlukla bitmiyorsa sonraki metnin baştaki sözcük parçası.
    for (let j = last; j + 1 < nodes.length; j++) {
      const cur = nodes[j].text ?? "";
      const next = nodes[j + 1];
      if (/\s$/.test(cur) || next.type !== "text" || !next.text || /^\s/.test(next.text) || has(next, id)) break;
      const m = /^\S+/.exec(next.text)!;
      if (m[0].length === next.text.length) {
        nodes[j + 1] = add(next, id);
        continue;
      }
      nodes.splice(j + 1, 1, add({ ...next, text: m[0] }, id), { ...next, text: next.text.slice(m[0].length) });
      break;
    }
    // Geri: işaretli metin boşlukla başlamıyorsa önceki metnin sondaki sözcük parçası.
    for (let j = first; j > 0; j--) {
      const cur = nodes[j].text ?? "";
      const prev = nodes[j - 1];
      if (/^\s/.test(cur) || prev.type !== "text" || !prev.text || /\s$/.test(prev.text) || has(prev, id)) break;
      const m = /\S+$/.exec(prev.text)!;
      if (m[0].length === prev.text.length) {
        nodes[j - 1] = add(prev, id);
        continue;
      }
      nodes.splice(j - 1, 1, { ...prev, text: prev.text.slice(0, prev.text.length - m[0].length) }, add({ ...prev, text: m[0] }, id));
      break;
    }
  }
}

/** Aynı işaretli bitişik metin düğümlerini birleştirir, boş metinleri atar. */
export function mergeText(nodes: JSONContent[]): JSONContent[] {
  const out: JSONContent[] = [];
  for (const n of nodes) {
    const last = out[out.length - 1];
    if (n.type === "text" && last?.type === "text" && JSON.stringify(last.marks ?? []) === JSON.stringify(n.marks ?? [])) last.text = (last.text ?? "") + (n.text ?? "");
    else if (n.type !== "text" || n.text) out.push({ ...n });
  }
  return out;
}

/** Paragraf düğümü (boş içerik alanı yazılmaz). */
export function paragraphNode(attrs: Record<string, unknown>, content: JSONContent[], type: "paragraph" | "heading" = "paragraph"): JSONContent {
  const merged = mergeText(content);
  return { type, attrs, ...(merged.length ? { content: merged } : {}) };
}

/** Liste öğesi bilgisi (katlama için). */
export interface ListInfo {
  /** Liste kimliği: aynı kimlik ve düzeydeki ardışık öğeler aynı listeye girer. */
  key: string;
  level: number;
  ordered: boolean;
  listStyle?: string | null;
  multilevel?: string | null;
  start?: number | null;
  /** Önceki öğenin devamı (ikinci paragraf, tablo): yeni madde açmaz. */
  cont?: boolean;
}

export interface FlowItem {
  node: JSONContent;
  list?: ListInfo;
}

/** Düz paragraf akışını iç içe listelere katlar (başlıklar listeden çıkarılır). */
export function foldLists(items: FlowItem[]): JSONContent[] {
  const out: JSONContent[] = [];
  type Frame = { key: string; level: number; node: JSONContent; last: JSONContent | null };
  let stack: Frame[] = [];
  for (const it of items) {
    if (!it.list || ((it.node.type !== "paragraph" || it.list.cont) && !stack.length)) {
      stack = [];
      out.push(it.node);
      continue;
    }
    if (it.node.type !== "paragraph" || it.list.cont) {
      // Liste öğesinin devamı (ek paragraf, tablo vb.): düzeydeki son öğeye eklenir.
      while (stack.length > 1 && stack[stack.length - 1].level > it.list.level) stack.pop();
      const top = stack[stack.length - 1];
      if (top?.last) top.last.content!.push(it.node);
      else {
        stack = [];
        out.push(it.node);
      }
      continue;
    }
    const { key, ordered } = it.list;
    const level = Math.max(0, Math.min(8, it.list.level));
    while (stack.length && (stack[stack.length - 1].level > level || (stack[stack.length - 1].level === level && stack[stack.length - 1].key !== key))) stack.pop();
    let top = stack[stack.length - 1];
    // Atlanan düzeyler: ara düzey için boş öğe açılmaz, liste bir üst öğeye bağlanır.
    if (!top || top.level < level) {
      const attrs: Record<string, unknown> = { listStyle: it.list.listStyle ?? null, multilevel: it.list.multilevel ?? null };
      if (ordered) attrs.start = it.list.start ?? 1;
      const listNode: JSONContent = { type: ordered ? "orderedList" : "bulletList", attrs, content: [] };
      if (!top || !top.last) out.push(listNode);
      else top.last.content!.push(listNode);
      stack.push({ key, level, node: listNode, last: null });
      top = stack[stack.length - 1];
    }
    const li: JSONContent = { type: "listItem", content: [it.node] };
    top.node.content!.push(li);
    top.last = li;
  }
  return out;
}

/** Not içeriğinden not düğümü (ilk paragrafın baştaki boşluğu ve sondaki boş paragraflar atılır). */
export function noteNode(blocks: JSONContent[], kind: "footnote" | "endnote", id: string, noteText: (b: JSONContent[]) => string): JSONContent {
  const first = blocks[0];
  if (first?.content?.[0]?.type === "text") {
    first.content[0] = { ...first.content[0], text: (first.content[0].text ?? "").replace(/^\s+/, "") };
    if (!first.content[0].text) first.content.shift();
    if (!first.content.length) delete first.content;
  }
  while (blocks.length > 1 && blocks[blocks.length - 1].type === "paragraph" && !blocks[blocks.length - 1].content?.length) blocks.pop();
  const plain = blocks.every((b) => b.type === "paragraph" && (b.content ?? []).every((x) => x.type === "text" && !x.marks?.length));
  return { type: "note", attrs: { kind, text: noteText(blocks), content: plain ? null : blocks, id } };
}

/** Madde işareti stilinin düzey varsayılanı mı (varsayılan stil null olarak tutulur). */
export function bulletStyleAt(style: string | null, level: number): string | null {
  return style === ["disc", "circle", "square"][level % 3] ? null : style;
}

/** Düğümün düz metni. */
export function textOf(n: JSONContent): string {
  if (n.type === "text") return n.text ?? "";
  if (n.type === "hardBreak") return "\n";
  const a = atomText(n);
  if (a !== null) return a;
  return (n.content ?? []).map(textOf).join("");
}

/** Satır içi atom düğümlerin metin karşılığı (metin olmayan düğümde null). */
export function atomText(n: JSONContent): string | null {
  const a = n.attrs ?? {};
  switch (n.type) {
    case "seqField":
      return String(a.num ?? "1");
    case "crossRef":
    case "citation":
      return String(a.text ?? "");
    case "mergeField":
      return `«${String(a.name ?? "")}»`;
    case "equation":
      return String(a.latex ?? "");
    case "indexEntry":
    case "bookmark":
    case "chart":
      return "";
    default:
      return null;
  }
}

/** Belgenin bölümleri (bölümsüz içerik tek bölüm sayılır). */
export function sectionsOf(doc: JSONContent): JSONContent[] {
  const kids = doc.content ?? [];
  if (kids.some((k) => k.type === "section")) return kids.filter((k) => k.type === "section");
  return [{ type: "section", content: kids }];
}

/** Kaynakça ve dizin bloklarının paragraf karşılığı (yazıcılar için). */
export function plainBlocks(n: JSONContent): JSONContent[] {
  const a = n.attrs ?? {};
  const p = (text: string, attrs: Record<string, unknown> = {}): JSONContent => ({ type: "paragraph", attrs, ...(text ? { content: [{ type: "text", text }] } : {}) });
  if (n.type === "bibliography") {
    const entries = (a.entries as { before: string; italic: string; after: string }[] | undefined) ?? [];
    return [
      { type: "heading", attrs: { level: 1 }, content: [{ type: "text", text: String(a.title || "Kaynakça") }] },
      ...entries.map((e) => ({
        type: "paragraph",
        attrs: {},
        content: mergeText([{ type: "text", text: e.before }, { type: "text", text: e.italic, marks: [{ type: "italic" }] }, { type: "text", text: e.after }]),
      })),
    ];
  }
  if (n.type === "indexBlock") {
    const entries = (a.entries as IndexItem[] | undefined) ?? [];
    return entries.map((e) => p(indexLine(e), e.sub ? { indentLeft: 18 } : {}));
  }
  return [];
}
