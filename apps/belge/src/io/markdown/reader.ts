// Markdown okuyucu: CommonMark blokları (başlık, paragraf, kod, alıntı, liste, yatay çizgi, HTML) +
// GFM tabloları, görev listeleri ve dipnotlar → ProseMirror JSON.
import type { JSONContent } from "@tiptap/core";
import { defaultSettings, normalizeDoc, type BelgeDocument } from "../../model/document";
import { noteText } from "../../editor/notes";
import { MONO_FONT, imageNode, mergeText, parseDataUrl, type Mark } from "../common";
import { normLabel, parseInline, unescapeMd, type Inl, type InlineCtx } from "./inline";

type MdBlock =
  | { t: "heading"; level: number; text: string }
  | { t: "para"; text: string }
  | { t: "code"; text: string; lang: string }
  | { t: "hr" }
  | { t: "quote"; c: MdBlock[] }
  | { t: "list"; ordered: boolean; delim: string; start: number; items: { c: MdBlock[]; task: null | boolean }[] }
  | { t: "table"; align: (string | null)[]; head: string[]; rows: string[][] }
  | { t: "html"; text: string };

interface BlockCtx extends InlineCtx {
  noteBlocks: Map<string, MdBlock[]>;
  depth: number;
}

export interface MarkdownOptions {
  /** Göreli resim yolunu baytlara çevirir (masaüstünde belgenin klasöründen); çözülemezse null. */
  resolveImage?: (path: string) => Promise<Uint8Array | null>;
  /** HTML bloklarını belge bloklarına çevirir (tarayıcıda DOM ile). */
  htmlToBlocks?: (html: string) => JSONContent[];
}

/** Code blokları için paragraf gölgesi (yazıcı bu gölge + eş genişlikli yazı tipiyle kod bloğunu tanır). */
export const CODE_SHADE = "#F2F2F2";

const isBlank = (l: string) => /^[ \t]*$/.test(l);

/** Satır başındaki sekmeleri 4'lük duraklara göre boşluğa çevirir. */
function expandTabs(line: string): string {
  if (!line.includes("\t")) return line;
  let out = "";
  let i = 0;
  for (; i < line.length && (line[i] === " " || line[i] === "\t"); i++) out += line[i] === "\t" ? " ".repeat(4 - (out.length % 4)) : " ";
  return out + line.slice(i);
}

const indentOf = (l: string) => l.length - l.trimStart().length;
const stripIndent = (l: string, n: number) => {
  let k = 0;
  while (k < n && l[k] === " ") k++;
  return l.slice(k);
};

const ATX = /^(#{1,6})(?:[ \t]+(.*?))?[ \t]*$/;
const HR = /^(?:(?:\*[ \t]*){3,}|(?:-[ \t]*){3,}|(?:_[ \t]*){3,})$/;
const FENCE = /^(`{3,}|~{3,})(.*)$/;
const LIST = /^([-+*]|\d{1,9}[.)])(?=[ \t]|$)/;
const DELIM_ROW = /^\|?\s*:?-+:?\s*(\|\s*:?-+:?\s*)*\|?\s*$/;
const HTML_BLOCK = /^<(?:!--|\/?(?:address|article|aside|blockquote|body|caption|center|col|colgroup|dd|details|dialog|div|dl|dt|fieldset|figcaption|figure|footer|form|h[1-6]|header|hr|html|iframe|legend|li|main|menu|nav|ol|p|pre|section|summary|table|tbody|td|tfoot|th|thead|tr|ul)(?:[\s/>]|$))/i;

interface Marker {
  ordered: boolean;
  delim: string;
  start: number;
  width: number;
  first: string;
}

function listMarker(line: string): Marker | null {
  const ind = indentOf(line);
  if (ind > 3) return null;
  const rest = line.slice(ind);
  const m = LIST.exec(rest);
  if (!m) return null;
  const after = rest.slice(m[0].length);
  const sp = after.length - after.trimStart().length;
  const body = after.trimStart();
  const ordered = /\d/.test(m[1][0]);
  const pad = !body ? 1 : sp >= 5 ? 1 : sp;
  return { ordered, delim: ordered ? m[1].slice(-1) : m[1], start: ordered ? parseInt(m[1], 10) : 1, width: ind + m[0].length + pad, first: !body ? "" : sp >= 5 ? after.slice(1) : body };
}

/** Satır yeni bir blok başlatıyor mu (tembel devam satırı olamaz). */
function startsBlock(line: string): boolean {
  if (indentOf(line) > 3) return false;
  const r = line.trimStart();
  return ATX.test(r) || HR.test(r) || FENCE.test(r) || r.startsWith(">") || (!!listMarker(line) && !!listMarker(line)!.first);
}

const splitRow = (l: string): string[] => {
  let s = l.trim();
  if (s.startsWith("|")) s = s.slice(1);
  if (s.endsWith("|") && !s.endsWith("\\|")) s = s.slice(0, -1);
  const cells: string[] = [];
  let cur = "";
  for (let i = 0; i < s.length; i++) {
    if (s[i] === "\\" && s[i + 1] === "|") {
      cur += "|";
      i++;
    } else if (s[i] === "|") {
      cells.push(cur.trim());
      cur = "";
    } else cur += s[i];
  }
  cells.push(cur.trim());
  return cells;
};

const REF_DEF = /^ {0,3}\[([^\]]+)\]:[ \t]*(?:<([^>\n]*)>|(\S+))(?:[ \t]+(?:"([^"]*)"|'([^']*)'|\(([^)]*)\)))?[ \t]*$/;

function parseBlocks(input: string[], ctx: BlockCtx): MdBlock[] {
  if (ctx.depth > 40) return [{ t: "para", text: input.join("\n") }];
  ctx.depth++;
  const lines = input.map(expandTabs);
  const out: MdBlock[] = [];
  let para: string[] | null = null;
  const flushPara = () => {
    if (!para) return;
    // Paragraf başındaki başvuru tanımları
    while (para.length) {
      const m = REF_DEF.exec(para[0]);
      if (!m || m[1].startsWith("^")) break;
      const key = normLabel(m[1]);
      if (!ctx.refs.has(key)) ctx.refs.set(key, { href: unescapeMd(m[2] ?? m[3] ?? ""), title: m[4] ?? m[5] ?? m[6] });
      para.shift();
    }
    if (para.length) out.push({ t: "para", text: para.join("\n").trim() });
    para = null;
  };
  let i = 0;
  while (i < lines.length) {
    const line = lines[i];
    if (isBlank(line)) {
      flushPara();
      i++;
      continue;
    }
    const ind = indentOf(line);
    if (ind >= 4 && !para) {
      const buf: string[] = [];
      while (i < lines.length && (isBlank(lines[i]) || indentOf(lines[i]) >= 4)) buf.push(stripIndent(lines[i++], 4));
      while (buf.length && isBlank(buf[buf.length - 1])) buf.pop();
      out.push({ t: "code", text: buf.join("\n"), lang: "" });
      continue;
    }
    const rest = ind < 4 ? line.slice(ind) : line;
    let m: RegExpExecArray | null;
    if (ind < 4 && (m = FENCE.exec(rest)) && !(m[1][0] === "`" && m[2].includes("`"))) {
      flushPara();
      const ch = m[1][0];
      const len = m[1].length;
      const buf: string[] = [];
      i++;
      while (i < lines.length) {
        const l = lines[i];
        const close = new RegExp(`^ {0,3}${ch === "`" ? "`" : "~"}{${len},}[ \\t]*$`).exec(l);
        i++;
        if (close) break;
        buf.push(stripIndent(l, ind));
      }
      out.push({ t: "code", text: buf.join("\n"), lang: unescapeMd(m[2].trim()) });
      continue;
    }
    if (ind < 4 && (m = ATX.exec(rest))) {
      flushPara();
      const text = (m[2] ?? "").replace(/(?:^|[ \t]+)#+[ \t]*$/, "").trim();
      out.push({ t: "heading", level: m[1].length, text });
      i++;
      continue;
    }
    if (para && ind < 4 && /^(=+|-+)[ \t]*$/.test(rest)) {
      const text: string = (para as string[]).join("\n").trim();
      para = null;
      out.push({ t: "heading", level: rest[0] === "=" ? 1 : 2, text });
      i++;
      continue;
    }
    if (ind < 4 && HR.test(rest)) {
      flushPara();
      out.push({ t: "hr" });
      i++;
      continue;
    }
    if (ind < 4 && rest.startsWith(">")) {
      flushPara();
      const buf: string[] = [];
      while (i < lines.length) {
        const l = lines[i];
        const li = indentOf(l);
        const r = li < 4 ? l.slice(li) : l;
        if (li < 4 && r.startsWith(">")) {
          buf.push(r.slice(1).replace(/^ /, ""));
          i++;
        } else if (!isBlank(l) && buf.length && !isBlank(buf[buf.length - 1]) && !startsBlock(l)) {
          buf.push(l);
          i++;
        } else break;
      }
      out.push({ t: "quote", c: parseBlocks(buf, ctx) });
      continue;
    }
    const mk = listMarker(line);
    if (mk && (!para || (mk.first && (!mk.ordered || mk.start === 1)))) {
      flushPara();
      const list: Extract<MdBlock, { t: "list" }> = { t: "list", ordered: mk.ordered, delim: mk.delim, start: mk.start, items: [] };
      while (i < lines.length) {
        const m2 = listMarker(lines[i]);
        if (!m2 || m2.ordered !== list.ordered || m2.delim !== list.delim || HR.test(lines[i].trim())) break;
        const buf: string[] = [m2.first];
        i++;
        while (i < lines.length) {
          const l = lines[i];
          if (isBlank(l)) {
            buf.push("");
            i++;
          } else if (indentOf(l) >= m2.width) {
            buf.push(stripIndent(l, m2.width));
            i++;
          } else if (!isBlank(buf[buf.length - 1]) && !startsBlock(l) && !HR.test(l.trim())) {
            buf.push(l.trim());
            i++;
          } else break;
        }
        while (buf.length > 1 && isBlank(buf[buf.length - 1])) buf.pop();
        let task: boolean | null = null;
        const tm = /^\[([ xX])\][ \t]+/.exec(buf[0]);
        if (tm) {
          task = tm[1] !== " ";
          buf[0] = buf[0].slice(tm[0].length);
        }
        list.items.push({ c: parseBlocks(buf, ctx), task });
      }
      out.push(list);
      continue;
    }
    if (ind < 4 && (m = /^\[\^([^\]]+)\]:[ \t]?(.*)$/.exec(rest))) {
      flushPara();
      const label = normLabel(m[1]);
      const buf = [m[2]];
      i++;
      while (i < lines.length && (indentOf(lines[i]) >= 4 || (isBlank(lines[i]) && i + 1 < lines.length && indentOf(lines[i + 1]) >= 4))) buf.push(stripIndent(lines[i++], 4));
      ctx.notes.add(label);
      ctx.noteBlocks.set(label, parseBlocks(buf, ctx));
      continue;
    }
    if (ind < 4 && rest.includes("|") && i + 1 < lines.length && DELIM_ROW.test(lines[i + 1].trim()) && lines[i + 1].includes("-")) {
      const head = splitRow(rest);
      const delimCells = splitRow(lines[i + 1]);
      if (head.length === delimCells.length) {
        flushPara();
        const align = delimCells.map((d) => (d.startsWith(":") && d.endsWith(":") ? "center" : d.endsWith(":") ? "right" : d.startsWith(":") ? "left" : null));
        const rows: string[][] = [];
        i += 2;
        while (i < lines.length && !isBlank(lines[i]) && !startsBlock(lines[i])) {
          const cells = splitRow(lines[i++]);
          rows.push(Array.from({ length: head.length }, (_, k) => cells[k] ?? ""));
        }
        out.push({ t: "table", align, head, rows });
        continue;
      }
    }
    if (ind < 4 && !para && HTML_BLOCK.test(rest)) {
      const buf: string[] = [];
      const comment = rest.startsWith("<!--");
      while (i < lines.length && !isBlank(lines[i])) {
        buf.push(lines[i]);
        if (comment && lines[i].includes("-->")) {
          i++;
          break;
        }
        i++;
      }
      if (!comment) out.push({ t: "html", text: buf.join("\n") });
      continue;
    }
    if (para) para.push(rest.trimStart());
    else para = [rest];
    i++;
  }
  flushPara();
  ctx.depth--;
  return out;
}

/* ------------------------------------------------------------------ */
/* JSON'a dönüştürme                                                   */
/* ------------------------------------------------------------------ */

const HTML_MARKS: Record<string, Mark> = {
  u: { type: "underline", attrs: { type: "single", color: null } },
  ins: { type: "underline", attrs: { type: "single", color: null } },
  sub: { type: "subscript" },
  sup: { type: "superscript" },
  mark: { type: "highlight", attrs: { color: "#FFFF00" } },
  s: { type: "strike" },
  del: { type: "strike" },
  strike: { type: "strike" },
  b: { type: "bold" },
  strong: { type: "bold" },
  i: { type: "italic" },
  em: { type: "italic" },
};

class Converter {
  images = new Map<string, Uint8Array | null>();
  noteNodes = new Map<string, JSONContent>();
  noteSeq = 0;

  constructor(
    readonly ctx: BlockCtx,
    readonly opts: MarkdownOptions,
  ) {}

  /** Satır içi düğümler → ProseMirror satır içi JSON. */
  inline(nodes: Inl[], marks: Mark[] = [], html: Map<string, number> = new Map()): JSONContent[] {
    const out: JSONContent[] = [];
    const withHtml = (ms: Mark[]) => {
      const extra: Mark[] = [];
      for (const [tag, n] of html) if (n > 0 && HTML_MARKS[tag] && !ms.some((x) => x.type === HTML_MARKS[tag].type)) extra.push(HTML_MARKS[tag]);
      return [...ms, ...extra];
    };
    const text = (s: string, ms: Mark[]) => {
      if (!s) return;
      const all = withHtml(ms);
      out.push({ type: "text", text: s, ...(all.length ? { marks: all } : {}) });
    };
    for (const n of nodes) {
      switch (n.k) {
        case "text":
          text(n.s, marks);
          break;
        case "soft":
          text(" ", marks);
          break;
        case "br":
          out.push({ type: "hardBreak" });
          break;
        case "code":
          text(n.s, [...marks.filter((m) => m.type !== "textStyle"), { type: "textStyle", attrs: { fontFamily: MONO_FONT } }]);
          break;
        case "em":
        case "strong":
        case "del": {
          const type = n.k === "em" ? "italic" : n.k === "strong" ? "bold" : "strike";
          out.push(...this.inline(n.c, marks.some((m) => m.type === type) ? marks : [...marks, { type }], html));
          break;
        }
        case "link":
          out.push(...this.inline(n.c, [...marks.filter((m) => m.type !== "link"), { type: "link", attrs: { href: n.href } }], html));
          break;
        case "img": {
          const img = this.image(n.src, n.alt);
          if (img) out.push(img);
          else if (/^https?:/i.test(n.src)) text(n.alt || n.src, [...marks.filter((m) => m.type !== "link"), { type: "link", attrs: { href: n.src } }]);
          else text(n.alt, marks);
          break;
        }
        case "html":
          if (n.tag === "br") out.push({ type: "hardBreak" });
          else html.set(n.tag, Math.max(0, (html.get(n.tag) ?? 0) + (n.close ? -1 : 1)));
          break;
        case "fn": {
          const note = this.note(n.label);
          if (note) out.push(note);
          break;
        }
      }
    }
    return out;
  }

  image(src: string, alt: string): JSONContent | null {
    const data = parseDataUrl(src);
    const bytes = data?.bytes ?? this.images.get(src) ?? null;
    if (!bytes) return null;
    return imageNode(bytes, { mime: data?.mime, alt: alt || null });
  }

  note(label: string): JSONContent | null {
    const cached = this.noteNodes.get(label);
    if (cached) return { ...cached, attrs: { ...cached.attrs, id: `nf${++this.noteSeq}` } };
    const blocks = this.ctx.noteBlocks.get(label);
    if (!blocks) return null;
    this.ctx.noteBlocks.delete(label);
    const content = this.blocks(blocks);
    const plain = content.every((b) => b.type === "paragraph" && (b.content ?? []).every((c) => c.type === "text" && !c.marks?.length));
    const node: JSONContent = { type: "note", attrs: { kind: "footnote", text: noteText(content), content: plain ? null : content, id: `nf${++this.noteSeq}` } };
    this.noteNodes.set(label, node);
    return node;
  }

  para(text: string, attrs: Record<string, unknown> = {}): JSONContent {
    const content = mergeText(this.inline(parseInline(text, this.ctx)));
    return { type: "paragraph", attrs, ...(content.length ? { content } : {}) };
  }

  blocks(bs: MdBlock[], quote = false): JSONContent[] {
    const out: JSONContent[] = [];
    for (const b of bs) {
      switch (b.t) {
        case "heading": {
          const content = mergeText(this.inline(parseInline(b.text, this.ctx)));
          out.push({ type: "heading", attrs: { level: b.level }, ...(content.length ? { content } : {}) });
          break;
        }
        case "para":
          out.push(this.para(b.text, quote ? { styleId: "Quote" } : {}));
          break;
        case "code": {
          const lines = b.text.split("\n");
          // Son satır normal paragraf aralığıyla biter (sonraki blokla arasında boşluk kalsın).
          lines.forEach((l, k) => {
            const attrs = { shading: CODE_SHADE, spaceBefore: 0, ...(k < lines.length - 1 ? { spaceAfter: 0 } : {}), lineHeight: 1 };
            out.push({ type: "paragraph", attrs, ...(l ? { content: [{ type: "text", text: l, marks: [{ type: "textStyle", attrs: { fontFamily: MONO_FONT } }] }] } : {}) });
          });
          break;
        }
        case "hr":
          out.push({ type: "horizontalRule" });
          break;
        case "quote":
          out.push(...this.blocks(b.c, true));
          break;
        case "list": {
          const attrs: Record<string, unknown> = { listStyle: b.ordered && b.delim === ")" ? "decimal-paren" : null };
          if (b.ordered) attrs.start = b.start;
          out.push({
            type: b.ordered ? "orderedList" : "bulletList",
            attrs,
            content: b.items.map((it) => {
              let c = this.blocks(it.c, quote);
              if (!c.length || c[0].type !== "paragraph") {
                const first = c[0]?.type === "heading" ? { ...c[0], type: "paragraph", attrs: {} } : null;
                c = first ? [first, ...c.slice(1)] : [{ type: "paragraph" }, ...c];
              }
              if (it.task !== null) {
                const p = c[0];
                p.content = mergeText([{ type: "text", text: it.task ? "☒ " : "☐ " }, ...(p.content ?? [])]);
              }
              return { type: "listItem", content: c };
            }),
          });
          break;
        }
        case "table": {
          const cell = (s: string, k: number): JSONContent => ({
            type: "tableCell",
            attrs: { colspan: 1, rowspan: 1, colwidth: null },
            content: [this.para(s, b.align[k] ? { textAlign: b.align[k] } : {})],
          });
          out.push({
            type: "table",
            attrs: { tableStyle: "grid", headerRow: true },
            content: [b.head, ...b.rows].map((r) => ({ type: "tableRow", content: r.map(cell) })),
          });
          break;
        }
        case "html": {
          const conv = this.opts.htmlToBlocks?.(b.text);
          if (conv?.length) out.push(...conv);
          else {
            const plain = b.text.replace(/<[^>]*>/g, " ").replace(/\s+/g, " ").trim();
            if (plain) out.push({ type: "paragraph", content: [{ type: "text", text: unescapeMd(plain) }] });
          }
          break;
        }
      }
    }
    return out;
  }

  /** Belgedeki göreli resim yollarını önceden çözer. */
  async preload(bs: MdBlock[]) {
    if (!this.opts.resolveImage) return;
    const srcs = new Set<string>();
    const scan = (text: string) => {
      for (const n of walk(parseInline(text, this.ctx))) if (n.k === "img" && !/^(data|https?):/i.test(n.src)) srcs.add(n.src);
    };
    const visit = (list: MdBlock[]) => {
      for (const b of list) {
        if (b.t === "para" || b.t === "heading") scan(b.text);
        else if (b.t === "quote") visit(b.c);
        else if (b.t === "list") b.items.forEach((it) => visit(it.c));
        else if (b.t === "table") [b.head, ...b.rows].forEach((r) => r.forEach(scan));
      }
    };
    visit(bs);
    for (const v of this.ctx.noteBlocks.values()) visit(v);
    for (const src of [...srcs].slice(0, 200)) {
      try {
        this.images.set(src, await this.opts.resolveImage(decodeURI(src)));
      } catch {
        this.images.set(src, null);
      }
    }
  }
}

function* walk(nodes: Inl[]): Generator<Inl> {
  for (const n of nodes) {
    yield n;
    if (n.k === "em" || n.k === "strong" || n.k === "del" || n.k === "link") yield* walk(n.c);
  }
}

/** Markdown metni → belge. */
export async function readMarkdown(text: string, opts: MarkdownOptions = {}): Promise<BelgeDocument> {
  const ctx: BlockCtx = { refs: new Map(), notes: new Set(), noteBlocks: new Map(), depth: 0 };
  const src = text.replace(/^﻿/, "").replace(/\r\n?/g, "\n").replace(/\u0000/g, "�");
  // YAML ön bilgisi (--- … ---) başlık için okunur, belgeye girmez.
  let title = "";
  let body = src;
  const fm = /^---\n([\s\S]*?)\n(?:---|\.\.\.)\n/.exec(src);
  if (fm) {
    body = src.slice(fm[0].length);
    title = /^title:\s*["']?(.*?)["']?\s*$/m.exec(fm[1])?.[1] ?? "";
  }
  const blocks = parseBlocks(body.split("\n"), ctx);
  const conv = new Converter(ctx, opts);
  await conv.preload(blocks);
  const content = conv.blocks(blocks);
  const settings = defaultSettings();
  if (title) settings.properties.title = title;
  return { content: normalizeDoc({ type: "doc", content: content.length ? content : [{ type: "paragraph" }] }), settings };
}
