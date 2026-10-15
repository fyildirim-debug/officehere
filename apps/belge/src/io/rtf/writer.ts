// RTF yazıcı: belge → RTF 1.9.1 (yazı tipi/renk/stil/liste tabloları, biçimli koşular, tablolar,
// resimler, köprüler, dipnotlar, üst/alt bilgi). Türkçe ve diğer ASCII dışı karakterler \uN ile yazılır.
import type { JSONContent } from "@tiptap/core";
import type { BelgeDocument, PageSetup } from "../../model/document";
import type { StyleDef } from "../../model/styles";
import { atomText, marksToProps, nodeStyleId, parseDataUrl, plainBlocks, sectionsOf, sniffImage, styleChar, stylesOf, type CharProps } from "../common";
import { fieldNowText, fieldPicture } from "../../editor/nodes";
import { opaquePlaceholder } from "../../editor/opaque";
import { ansiByte } from "./codepage";
import { LINK_COLOR, isoDttm } from "./state";

const STYLE_INDEX: Record<string, [number, string]> = {
  Normal: [0, "Normal"],
  Heading1: [1, "heading 1"],
  Heading2: [2, "heading 2"],
  Heading3: [3, "heading 3"],
  Heading4: [4, "heading 4"],
  Heading5: [5, "heading 5"],
  Heading6: [6, "heading 6"],
  Title: [7, "Title"],
  Subtitle: [8, "Subtitle"],
  Quote: [9, "Quote"],
  IntenseQuote: [10, "Intense Quote"],
  NoSpacing: [11, "No Spacing"],
};

const BULLET_CHARS: Record<string, string> = { disc: "●", circle: "○", square: "■", diamond: "❖", arrow: "➢", check: "✓", dash: "–" };
const LEVEL_BULLETS = ["●", "○", "■"];

/** RTF metin kaçışı: denetim karakterleri, ASCII dışı → \uN + ANSI yedek. */
export function rtfEscape(s: string): string {
  let out = "";
  for (const ch of s) {
    const code = ch.codePointAt(0)!;
    if (ch === "\\" || ch === "{" || ch === "}") out += `\\${ch}`;
    else if (ch === "\t") out += "\\tab ";
    else if (ch === "\n") out += "\\line ";
    else if (code === 0xa0) out += "\\~";
    else if (code === 0xad) out += "\\-";
    else if (code === 0x2011) out += "\\_";
    else if (code < 0x20) continue;
    else if (code < 0x80) out += ch;
    else {
      for (const unit of code > 0xffff ? [0xd800 + ((code - 0x10000) >> 10), 0xdc00 + ((code - 0x10000) & 0x3ff)] : [code]) {
        const fb = ansiByte(String.fromCharCode(unit), 1254);
        out += `\\u${unit > 32767 ? unit - 65536 : unit}${fb !== null && fb >= 0x80 ? `\\'${fb.toString(16)}` : "?"}`;
      }
    }
  }
  return out;
}

const tw = (pt: number) => Math.round(pt * 20);

function toRoman(n: number): string {
  const r: [number, string][] = [[1000, "M"], [900, "CM"], [500, "D"], [400, "CD"], [100, "C"], [90, "XC"], [50, "L"], [40, "XL"], [10, "X"], [9, "IX"], [5, "V"], [4, "IV"], [1, "I"]];
  let out = "";
  for (const [v, s] of r) while (n >= v) (out += s), (n -= v);
  return out;
}

const alpha = (n: number) => {
  let s = "";
  for (let k = n; k > 0; k = Math.floor((k - 1) / 26)) s = String.fromCharCode(65 + ((k - 1) % 26)) + s;
  return s;
};

interface LevelDef {
  ordered: boolean;
  style: string | null;
  start: number;
}

function nfcOf(l: LevelDef): number {
  if (!l.ordered) return 23;
  switch (l.style) {
    case "upper-roman":
      return 1;
    case "lower-roman":
      return 2;
    case "upper-alpha":
      return 3;
    case "lower-alpha":
    case "lower-alpha-paren":
      return 4;
    default:
      return 0;
  }
}

function markerText(l: LevelDef, n: number, level: number): string {
  if (!l.ordered) return BULLET_CHARS[l.style ?? ""] ?? LEVEL_BULLETS[level % 3];
  const suffix = l.style === "decimal-paren" || l.style === "lower-alpha-paren" ? ")" : ".";
  switch (nfcOf(l)) {
    case 1:
      return toRoman(n) + suffix;
    case 2:
      return toRoman(n).toLowerCase() + suffix;
    case 3:
      return alpha(n) + suffix;
    case 4:
      return alpha(n).toLowerCase() + suffix;
    default:
      return `${n}${suffix}`;
  }
}

class RtfWriter {
  fonts: string[] = [];
  colors: string[] = [];
  lists: LevelDef[][] = [];
  styles: Record<string, StyleDef>;
  listDefs = new Map<JSONContent, number>();
  /** Değişiklik izleme yazarları (\revtbl; 0 "Unknown") ve açık açıklama aralıkları (gövdede). */
  revAuthors: string[] = ["Unknown"];
  private openAtn = new Map<string, number>();
  private atnSeq = 0;
  private inBody = false;

  constructor(readonly doc: BelgeDocument) {
    this.styles = stylesOf(doc.settings);
    this.font(styleChar(this.styles, "Normal").font ?? "Aptos");
  }

  font(name: string): number {
    let i = this.fonts.indexOf(name);
    if (i < 0) i = this.fonts.push(name) - 1;
    return i;
  }

  color(hex: string | null | undefined): number {
    const h = (hex ?? "").replace(/^#/, "").toUpperCase();
    if (!/^[0-9A-F]{6}$/.test(h)) return 0;
    let i = this.colors.indexOf(h);
    if (i < 0) i = this.colors.push(h) - 1;
    return i + 1;
  }

  /** Koşu biçimi: stil temeli + işaretler. */
  runProps(p: CharProps, base: CharProps): string {
    const cs = p.charStyle ? this.styles[p.charStyle]?.props : undefined;
    const bold = p.bold || base.bold || cs?.bold;
    const italic = p.italic || base.italic || cs?.italic;
    let x = `\\f${this.font(p.font ?? base.font ?? "Aptos")}\\fs${Math.round((p.size ?? base.size ?? 11) * 2)}`;
    if (bold) x += "\\b";
    if (italic) x += "\\i";
    if (p.underline) x += { single: "\\ul", double: "\\uldb", thick: "\\ulth", dotted: "\\uld", dashed: "\\uldash", wave: "\\ulwave", words: "\\ulw" }[p.underline] ?? "\\ul";
    if (p.strike) x += "\\strike";
    if (p.dstrike) x += "\\striked1";
    if (p.sup) x += "\\super";
    else if (p.sub) x += "\\sub";
    const color = p.link && !p.color ? LINK_COLOR : (p.color ?? cs?.color ?? base.color);
    const ci = this.color(color);
    if (ci) x += `\\cf${ci}`;
    if (p.highlight) x += `\\highlight${this.color(p.highlight)}`;
    if (p.bg) x += `\\chcbpat${this.color(p.bg)}`;
    if (p.caps || base.caps || cs?.caps) x += "\\caps";
    if (p.smallCaps) x += "\\scaps";
    if (p.hidden) x += "\\v";
    if (p.spacing) x += `\\expndtw${tw(p.spacing)}`;
    return x;
  }

  /** Satır içi içerik (köprüler alan olarak gruplanır). */
  inline(nodes: JSONContent[], base: CharProps): string {
    let out = "";
    let i = 0;
    while (i < nodes.length) {
      const n = nodes[i];
      const link = marksToProps(n.marks).link;
      if (link) {
        let inner = "";
        while (i < nodes.length && marksToProps(nodes[i].marks).link === link) {
          out += this.atnEdges(nodes[i]);
          inner += this.inlineNode(nodes[i++], base);
        }
        out += `{\\field{\\*\\fldinst{HYPERLINK "${rtfEscape(link.replace(/"/g, "%22"))}"}}{\\fldrslt{${inner}}}}`;
        continue;
      }
      out += this.atnEdges(n);
      out += this.inlineNode(n, base);
      i++;
    }
    return out;
  }

  revAuthor(name: string): number {
    let i = this.revAuthors.indexOf(name || "Unknown");
    if (i < 0) i = this.revAuthors.push(name) - 1;
    return i;
  }

  /** Değişiklik izleme koşu sözcükleri (\revised / \deleted, yazar ve tarih). */
  revProps(n: JSONContent): string {
    const m = n.marks?.find((x) => x.type === "insertion" || x.type === "deletion");
    if (!m) return "";
    const a = this.revAuthor(String(m.attrs?.author ?? ""));
    const dt = isoDttm(String(m.attrs?.date ?? ""));
    return m.type === "insertion" ? `\\revised\\revauth${a}${dt ? `\\revdttm${dt}` : ""}` : `\\deleted\\revauthdel${a}${dt ? `\\revdttmdel${dt}` : ""}`;
  }

  /** Açıklama aralıklarının açılış/kapanışları (\atrfstart … \atrfend + \annotation); yalnız gövdede. */
  private atnEdges(n: JSONContent): string {
    if (!this.inBody) return "";
    const ids = new Set((n.marks ?? []).filter((m) => m.type === "comment").map((m) => String(m.attrs?.id ?? "")));
    let out = "";
    for (const id of [...this.openAtn.keys()]) if (!ids.has(id)) out += this.atnEnd(id);
    for (const id of ids) {
      if (this.openAtn.has(id) || !this.doc.settings.comments?.[id]) continue;
      const num = this.atnSeq++;
      this.openAtn.set(id, num);
      out += `{\\*\\atrfstart ${num}}`;
    }
    return out;
  }

  private atnEnd(id: string): string {
    const num = this.openAtn.get(id)!;
    this.openAtn.delete(id);
    const c = this.doc.settings.comments[id];
    const dt = isoDttm(c.date ?? "");
    const text = String(c.text ?? "")
      .split("\n")
      .map((t) => rtfEscape(t))
      .join("\\par ");
    return `{\\*\\atrfend ${num}}{\\*\\atnid ${rtfEscape(c.initials ?? "")}}{\\*\\atnauthor ${rtfEscape(c.author ?? "")}}\\chatn{\\*\\annotation{\\*\\atnref ${num}}${dt ? `{\\*\\atndate ${dt}}` : ""}\\pard\\plain ${text}}`;
  }

  inlineNode(n: JSONContent, base: CharProps): string {
    const p = marksToProps(n.marks);
    // İzlenen değişiklik: metin \revised / \deleted ile yazılır; silinmiş diğer öğeler yazılmaz.
    if (n.type === "text") return n.text ? `{${this.runProps(p, base)}${this.revProps(n)} ${rtfEscape(n.text)}}` : "";
    if (n.marks?.some((m) => m.type === "deletion")) return "";
    switch (n.type) {
      case "hardBreak":
        return "\\line ";
      case "pageField": {
        const kind = n.attrs?.field;
        const pic = fieldPicture(kind, n.attrs?.format);
        // Alan yönergesinde ters eğik çizgi kaçışlanır: DATE \\@ "dd.MM.yyyy".
        const f = kind === "NUMPAGES" ? "NUMPAGES" : (kind === "DATE" || kind === "TIME") && pic ? `${kind} \\\\@ "${rtfEscape(pic)}"` : "PAGE";
        // Sonuç: tarih/saat alanında kayıt anındaki değer (alanı güncellemeyen okuyucular da doğru gösterir).
        return `{\\field{\\*\\fldinst{${this.runProps(p, base)} ${f}}}{\\fldrslt{${this.runProps(p, base)} ${rtfEscape(fieldNowText(kind, new Date(), n.attrs?.format) ?? "1")}}}}`;
      }
      case "image":
        return this.image(n);
      case "note":
        return this.note(n);
      case "bookmark": {
        const name = rtfEscape(String(n.attrs?.name ?? ""));
        return name ? `{\\*\\bkmkstart ${name}}{\\*\\bkmkend ${name}}` : "";
      }
      case "opaqueInline":
        // .docx'e özgü öğe RTF'de yazılamaz: yer tutucu metin (kaydederken uyarılır).
        return `{${this.runProps(p, base)} ${rtfEscape(opaquePlaceholder(n.attrs))}}`;
      default: {
        const t = atomText(n);
        return t ? `{${this.runProps(p, base)} ${rtfEscape(t)}}` : "";
      }
    }
  }

  image(n: JSONContent): string {
    const a = n.attrs ?? {};
    const data = parseDataUrl(String(a.src ?? ""));
    if (!data) return "";
    const kind = sniffImage(data.bytes)?.ext;
    if (kind !== "png" && kind !== "jpg") return a.alt ? rtfEscape(String(a.alt)) : "";
    const w = Math.max(1, Math.round(Number(a.width) || 200));
    const h = Math.max(1, Math.round(Number(a.height) || 150));
    let hex = "";
    const parts: string[] = [];
    for (let i = 0; i < data.bytes.length; i++) {
      hex += data.bytes[i].toString(16).padStart(2, "0");
      if (hex.length >= 128) {
        parts.push(hex);
        hex = "";
      }
    }
    if (hex) parts.push(hex);
    return `{\\pict\\${kind === "png" ? "pngblip" : "jpegblip"}\\picw${w}\\pich${h}\\picwgoal${w * 15}\\pichgoal${h * 15}\n${parts.join("\n")}\n}`;
  }

  note(n: JSONContent): string {
    const wasBody = this.inBody;
    this.inBody = false;
    try {
      return this.noteBody(n);
    } finally {
      this.inBody = wasBody;
    }
  }

  private noteBody(n: JSONContent): string {
    const a = n.attrs ?? {};
    const blocks: JSONContent[] = (a.content as JSONContent[] | null)?.length
      ? (a.content as JSONContent[])
      : String(a.text ?? "")
          .split("\n")
          .map((t) => ({ type: "paragraph", ...(t ? { content: [{ type: "text", text: t }] } : {}) }));
    let body = "";
    blocks.forEach((b, i) => {
      const nb = styleChar(this.styles, "Normal");
      const head = i === 0 ? `{\\super\\fs${Math.round((nb.size ?? 11) * 2)}\\chftn}` : "";
      const text = b.type === "paragraph" || b.type === "heading" ? this.inline(b.content ?? [], nb) : "";
      body += `\\pard\\plain${head ? "" : " "}${head}${text}${i < blocks.length - 1 ? "\\par\n" : ""}`;
    });
    return `{\\super\\chftn}{\\footnote${a.kind === "endnote" ? "\\ftnalt" : ""} ${body}}`;
  }

  /** Paragraf biçimi (stilin paragraf özellikleri açıkça yazılır). */
  paraProps(n: JSONContent, list?: { ls: number; level: number }, cont?: number, inTable = false): string {
    const a = (n.attrs ?? {}) as Record<string, unknown>;
    const id = nodeStyleId(n);
    const sp = styleChar(this.styles, id).base;
    let x = `\\pard\\plain${inTable ? "\\intbl" : ""}\\s${STYLE_INDEX[id]?.[0] ?? 0}`;
    if (n.type === "heading") x += `\\outlinelevel${Math.min(5, Number(a.level ?? 1) - 1)}`;
    const align = (a.textAlign as string | null) ?? sp.align ?? null;
    if (align) x += { left: "\\ql", center: "\\qc", right: "\\qr", justify: "\\qj" }[align] ?? "";
    if (list) x += `\\ls${list.ls}\\ilvl${list.level}\\li${720 * (list.level + 1)}\\fi-360`;
    else if (cont) x += `\\li${cont}`;
    else {
      const li = (a.indentLeft as number | null) ?? sp.indentLeft ?? null;
      if (li) x += `\\li${tw(li)}`;
      if (a.indentFirst != null) x += `\\fi${tw(a.indentFirst as number)}`;
    }
    const ri = (a.indentRight as number | null) ?? sp.indentRight ?? null;
    if (ri) x += `\\ri${tw(ri)}`;
    const sb = (a.spaceBefore as number | null) ?? sp.spaceBefore ?? 0;
    const sa = (a.spaceAfter as number | null) ?? sp.spaceAfter ?? 0;
    x += `\\sb${tw(sb)}\\sa${tw(sa)}`;
    const rule = (a.lineRule as string | null) ?? "auto";
    const lh = (a.lineHeight as number | null) ?? (rule === "auto" ? (sp.line ?? 1) : null);
    if (lh) x += rule === "auto" ? `\\sl${Math.round(lh * 240)}\\slmult1` : `\\sl${rule === "exact" ? -tw(lh) : tw(lh)}\\slmult0`;
    if (a.keepNext || sp.keepNext) x += "\\keepn";
    if (a.keepLines) x += "\\keep";
    if (a.pageBreakBefore) x += "\\pagebb";
    if (a.shading) x += `\\cbpat${this.color(a.shading as string)}`;
    return x;
  }

  paragraph(n: JSONContent, opts: { list?: { ls: number; level: number; marker: string }; cont?: number; inTable?: boolean; end?: string } = {}): string {
    const base = styleChar(this.styles, nodeStyleId(n));
    let x = this.paraProps(n, opts.list, opts.cont, opts.inTable);
    // Liste işareti eski okuyucular için (\listtext); yoksa sözcük ayıracı boşluk.
    x += opts.list ? `{\\listtext\\pard\\plain\\f${this.font(base.font ?? "Aptos")}\\fs${Math.round((base.size ?? 11) * 2)} ${rtfEscape(opts.list.marker)}\\tab}` : " ";
    x += `${this.inline(n.content ?? [], base)}${opts.end ?? "\\par"}\n`;
    return x;
  }

  blocks(nodes: JSONContent[] | undefined, inTable = false): string {
    return (nodes ?? []).map((n) => this.block(n, inTable)).join("");
  }

  block(n: JSONContent, inTable = false): string {
    switch (n.type) {
      case "paragraph":
      case "heading":
        return this.paragraph(n, { inTable });
      case "bulletList":
      case "orderedList":
        return inTable ? this.flatList(n, 0) : this.list(n);
      case "table":
        return inTable ? this.blocks((n.content ?? []).flatMap((r) => (r.content ?? []).flatMap((c) => c.content ?? [])), true) : this.table(n);
      case "horizontalRule":
        return `\\pard\\plain${inTable ? "\\intbl" : ""}\\brdrb\\brdrs\\brdrw10\\brsp20 \\par\n`;
      case "pageBreak":
        return inTable ? "" : "\\pard\\plain\\page\n";
      case "columnBreak":
        return inTable ? "" : "\\pard\\plain\\column\n";
      case "toc": {
        const entries = (n.attrs?.entries as { level: number; text: string; page: number }[] | undefined) ?? [];
        const title = n.attrs?.title ? this.paragraph({ type: "paragraph", content: [{ type: "text", text: String(n.attrs.title), marks: [{ type: "bold" }] }] }, { inTable }) : "";
        return title + entries.map((e) => this.paragraph({ type: "paragraph", attrs: { indentLeft: (e.level - 1) * 11 }, content: [{ type: "text", text: `${e.text}\t${e.page}` }] }, { inTable })).join("");
      }
      case "bibliography":
      case "indexBlock":
        return this.blocks(plainBlocks(n), inTable);
      case "opaqueBlock":
        return this.paragraph({ type: "paragraph", content: [{ type: "text", text: opaquePlaceholder(n.attrs) }] }, { inTable });
      default:
        return n.content ? this.blocks(n.content, inTable) : "";
    }
  }

  /** Liste: her üst düzey liste için ayrı \ls (numaralama yeniden başlar). */
  list(n: JSONContent): string {
    const levels: LevelDef[] = [];
    const scan = (l: JSONContent, depth: number) => {
      if (depth > 8) return;
      levels[depth] ??= { ordered: l.type === "orderedList", style: (l.attrs?.listStyle as string | null) ?? null, start: Number(l.attrs?.start) || 1 };
      for (const li of l.content ?? []) for (const c of li.content ?? []) if (c.type === "bulletList" || c.type === "orderedList") scan(c, depth + 1);
    };
    scan(n, 0);
    for (let k = 0; k < 9; k++) levels[k] ??= { ordered: levels[k - 1]?.ordered ?? false, style: null, start: 1 };
    const ls = this.lists.push(levels);
    const counters = new Array(9).fill(0);
    const walk = (l: JSONContent, depth: number): string => {
      let out = "";
      const lv = levels[Math.min(8, depth)];
      counters[depth] = (Number(l.attrs?.start) || lv.start) - 1;
      for (const li of l.content ?? []) {
        counters[depth]++;
        (li.content ?? []).forEach((c, i) => {
          if (c.type === "bulletList" || c.type === "orderedList") out += depth < 8 ? walk(c, depth + 1) : this.block(c);
          else if (i === 0 && (c.type === "paragraph" || c.type === "heading")) out += this.paragraph(c, { list: { ls, level: depth, marker: markerText(lv, counters[depth], depth) } });
          else if (c.type === "paragraph" || c.type === "heading") out += this.paragraph(c, { cont: 720 * (depth + 1) });
          else out += this.block(c);
        });
      }
      return out;
    };
    return walk(n, 0);
  }

  /** Tablo içindeki liste: işaret metniyle düz paragraflar. */
  flatList(n: JSONContent, depth: number): string {
    let out = "";
    let num = Number(n.attrs?.start) || 1;
    for (const li of n.content ?? []) {
      (li.content ?? []).forEach((c, i) => {
        if (c.type === "bulletList" || c.type === "orderedList") out += this.flatList(c, depth + 1);
        else if (c.type === "paragraph" || c.type === "heading") {
          const marker = i === 0 ? (n.type === "orderedList" ? `${num++}. ` : "• ") : "";
          out += this.paragraph({ ...c, attrs: { ...c.attrs, indentLeft: 18 * (depth + 1) }, content: [{ type: "text", text: marker }, ...(c.content ?? [])] }, { inTable: true });
        }
      });
    }
    return out;
  }

  table(n: JSONContent): string {
    const rows = n.content ?? [];
    const noBorder = n.attrs?.tableStyle === "none";
    const page = this.doc.settings.page;
    const contentTw = page.width - page.margin.left - page.margin.right - page.margin.gutter;
    // Izgara sütun sayısı ve genişlikleri (ilk satırdan).
    const widths: number[] = [];
    for (const c of rows[0]?.content ?? []) {
      const span = Math.max(1, Number(c.attrs?.colspan) || 1);
      const cw = c.attrs?.colwidth as number[] | null | undefined;
      for (let k = 0; k < span; k++) widths.push(cw?.[k] ? Math.round(cw[k] * 15) : 0);
    }
    const cols = Math.max(1, widths.length);
    const free = widths.filter((w) => !w).length;
    const rest = Math.max(0, contentTw - widths.reduce((s, w) => s + w, 0));
    for (let k = 0; k < cols; k++) if (!widths[k]) widths[k] = Math.max(300, Math.round(free ? rest / free : contentTw / cols));
    // Soldan girinti (\trleft): hücre sağ kenarları da o kadar kayar.
    const left = n.attrs?.align === "center" || n.attrs?.align === "right" ? 0 : Math.round((Number(n.attrs?.indent) || 0) * 15);
    const edges = widths.map((_, k) => left + widths.slice(0, k + 1).reduce((s, w) => s + w, 0));
    const border = noBorder ? "" : "\\clbrdrt\\brdrs\\brdrw10\\clbrdrl\\brdrs\\brdrw10\\clbrdrb\\brdrs\\brdrw10\\clbrdrr\\brdrs\\brdrw10";
    // Dikey birleştirmenin kapladığı sütunlar: başlangıç sütunu → kalan satır ve genişlik.
    const covered: { rows: number; span: number }[] = Array.from({ length: cols }, () => ({ rows: 0, span: 1 }));
    const align = n.attrs?.align === "center" ? "\\trqc" : n.attrs?.align === "right" ? "\\trqr" : "\\trql";
    let out = "";
    rows.forEach((row, r) => {
      let def = `\\trowd\\trgaph108\\trleft${left}${align}${r === 0 && n.attrs?.repeatHeader ? "\\trhdr" : ""}`;
      let body = "";
      const cells = row.content ?? [];
      let col = 0;
      let ci = 0;
      while (col < cols) {
        if (covered[col].rows > 0) {
          covered[col].rows--;
          const span = covered[col].span;
          def += `\\clvmrg${border}\\cellx${edges[Math.min(cols, col + span) - 1]}`;
          body += "\\pard\\plain\\intbl\\cell\n";
          col += span;
          continue;
        }
        const c = cells[ci++];
        if (!c) {
          def += `${border}\\cellx${edges[col]}`;
          body += "\\pard\\plain\\intbl\\cell\n";
          col++;
          continue;
        }
        const span = Math.max(1, Math.min(cols - col, Number(c.attrs?.colspan) || 1));
        const rs = Math.max(1, Number(c.attrs?.rowspan) || 1);
        if (rs > 1) covered[col] = { rows: rs - 1, span };
        const bg = c.attrs?.background ? `\\clcbpat${this.color(c.attrs.background as string)}` : "";
        const va = c.attrs?.verticalAlign === "middle" ? "\\clvertalc" : c.attrs?.verticalAlign === "bottom" ? "\\clvertalb" : "";
        def += `${rs > 1 ? "\\clvmgf" : ""}${va}${bg}${border}\\cellx${edges[col + span - 1]}`;
        const blocks = c.content?.length ? c.content : [{ type: "paragraph" }];
        const inner = this.blocks(blocks, true).replace(/\\par\n$/, "");
        body += `${inner}\\cell\n`;
        col += span;
      }
      out += `${def}\n${body}${def}\\row\n`;
    });
    return `${out}\\pard\\plain\n`;
  }

  listTable(): string {
    if (!this.lists.length) return "";
    let lt = "{\\*\\listtable";
    let lo = "{\\*\\listoverridetable";
    this.lists.forEach((levels, i) => {
      const id = 1000 + i;
      lt += `{\\list\\listtemplateid${id + 5000}${levels.every((l) => !l.ordered) ? "\\listsimple0" : ""}`;
      levels.forEach((l, k) => {
        const nfc = nfcOf(l);
        const suffix = l.style === "decimal-paren" || l.style === "lower-alpha-paren" ? ")" : ".";
        const text = nfc === 23 ? `{\\leveltext\\'01${rtfEscape(BULLET_CHARS[l.style ?? ""] ?? LEVEL_BULLETS[k % 3])};}{\\levelnumbers;}` : `{\\leveltext\\'02\\'0${k}${suffix};}{\\levelnumbers\\'01;}`;
        lt += `{\\listlevel\\levelnfc${nfc}\\levelnfcn${nfc}\\leveljc0\\leveljcn0\\levelfollow0\\levelstartat${l.start}${text}\\fi-360\\li${720 * (k + 1)}\\lin${720 * (k + 1)}}`;
      });
      lt += `{\\listname ;}\\listid${id}}`;
      lo += `{\\listoverride\\listid${id}\\listoverridecount0\\ls${i + 1}}`;
    });
    return `${lt}}\n${lo}}\n`;
  }

  styleSheet(): string {
    let x = "{\\stylesheet";
    for (const [id, [idx, name]] of Object.entries(STYLE_INDEX)) {
      const c = styleChar(this.styles, id);
      const sp = c.base;
      const props = `\\f${this.font(c.font ?? "Aptos")}\\fs${Math.round((c.size ?? 11) * 2)}${c.bold ? "\\b" : ""}${c.italic ? "\\i" : ""}${this.color(c.color) ? `\\cf${this.color(c.color)}` : ""}`;
      const para = `\\sb${tw(sp.spaceBefore ?? 0)}\\sa${tw(sp.spaceAfter ?? 0)}\\sl${Math.round((sp.line ?? 1) * 240)}\\slmult1${sp.keepNext ? "\\keepn" : ""}${id.startsWith("Heading") ? `\\outlinelevel${idx - 1}` : ""}`;
      x += `{\\s${idx}${idx ? "\\sbasedon0\\snext0" : "\\snext0"}${para}${props} ${name};}`;
    }
    return `${x}}\n`;
  }

  hf(key: "header" | "footer" | "headerFirst" | "footerFirst" | "headerEven" | "footerEven", word: string): string {
    const d = this.doc.settings[key];
    if (!d?.content?.length) return "";
    return `{\\${word} ${this.blocks(d.content).replace(/\\par\n$/, "")}}\n`;
  }

  pageProps(p: PageSetup, sect = false): string {
    const m = p.margin;
    if (sect) return `\\pgwsxn${p.width}\\pghsxn${p.height}\\marglsxn${m.left + m.gutter}\\margrsxn${m.right}\\margtsxn${m.top}\\margbsxn${m.bottom}${p.orientation === "landscape" ? "\\lndscpsxn" : ""}${p.columns > 1 ? `\\cols${p.columns}\\colsx${p.columnSpace}` : ""}`;
    return `\\paperw${p.width}\\paperh${p.height}\\margl${m.left + m.gutter}\\margr${m.right}\\margt${m.top}\\margb${m.bottom}${p.orientation === "landscape" ? "\\landscape" : ""}`;
  }

  write(): string {
    const s = this.doc.settings;
    let body = "";
    this.inBody = true;
    sectionsOf(this.doc.content).forEach((sec, i) => {
      if (i > 0) {
        const kind = sec.attrs?.kind;
        body += `\\sect\\sectd${kind === "continuous" ? "\\sbknone" : "\\sbkpage"}${this.pageProps((sec.attrs?.page as PageSetup | null) ?? s.page, true)}\n`;
      }
      body += this.blocks(sec.content);
    });
    this.inBody = false;
    body = body.replace(/\\par\n$/, "\n");
    // Belge sonunda açık kalan açıklama aralıkları son paragrafta kapanır.
    for (const id of [...this.openAtn.keys()]) body = body.replace(/\n$/, "") + this.atnEnd(id) + "\n";
    const hf = this.hf("header", "header") + this.hf("footer", "footer") + this.hf("headerFirst", "headerf") + this.hf("footerFirst", "footerf") + (s.oddEven ? this.hf("headerEven", "headerl") + this.hf("footerEven", "footerl") : "");
    const p = s.properties;
    const d = new Date(p.created || Date.now());
    const info = `{\\info${p.title ? `{\\title ${rtfEscape(p.title)}}` : ""}${p.subject ? `{\\subject ${rtfEscape(p.subject)}}` : ""}${p.author ? `{\\author ${rtfEscape(p.author)}}` : ""}${p.keywords ? `{\\keywords ${rtfEscape(p.keywords)}}` : ""}${Number.isFinite(d.getTime()) ? `{\\creatim\\yr${d.getUTCFullYear()}\\mo${d.getUTCMonth() + 1}\\dy${d.getUTCDate()}\\hr${d.getUTCHours()}\\min${d.getUTCMinutes()}}` : ""}}\n`;
    const styles = this.styleSheet();
    const lists = this.listTable();
    const fonts = `{\\fonttbl${this.fonts.map((f, i) => `{\\f${i}\\fnil\\fcharset162 ${rtfEscape(f)};}`).join("")}}\n`;
    const colors = `{\\colortbl;${this.colors.map((c) => `\\red${parseInt(c.slice(0, 2), 16)}\\green${parseInt(c.slice(2, 4), 16)}\\blue${parseInt(c.slice(4, 6), 16)};`).join("")}}\n`;
    const doc = `\\deftab708${this.pageProps(s.page)}${s.page.titlePage ? "\\titlepg" : ""}${s.oddEven ? "\\facingp" : ""}\\headery${s.page.margin.header}\\footery${s.page.margin.footer}\\sectd${this.pageProps(s.page, true)}${s.page.titlePage ? "\\titlepg" : ""}\n`;
    const revtbl = this.revAuthors.length > 1 ? `{\\*\\revtbl ${this.revAuthors.map((a) => `{${rtfEscape(a)};}`).join("")}}\n` : "";
    return `{\\rtf1\\ansi\\ansicpg1254\\deff0\\uc1\\deflang1055\n${fonts}${colors}${styles}${lists}${revtbl}{\\*\\generator OfficeHere Belge;}\n${info}${doc}${hf}${body}}\n`;
  }
}

/** Belge → RTF metni. */
export function writeRtf(doc: BelgeDocument): string {
  return new RtfWriter(doc).write();
}
