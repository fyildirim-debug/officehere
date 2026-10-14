// ODT yazıcı: belge → OpenDocument Metin 1.3 paketi (mimetype sıkıştırmasız ve ilk, manifest,
// content.xml, styles.xml, meta.xml, Pictures/). Paragraf/başlık biçimleri, listeler, tablolar
// (birleştirilmiş hücreler), resimler, köprüler, dipnotlar, metin kutuları, İçindekiler, üst/alt bilgi.
import JSZip from "jszip";
import type { JSONContent } from "@tiptap/core";
import type { BelgeDocument, CommentData, PageSetup } from "../../model/document";
import { samePage } from "../../model/document";
import type { StyleDef } from "../../model/styles";
import { fieldNowText, fieldPicture } from "../../editor/nodes";
import { normalizeDropCap } from "../../editor/dropCap";
import { atomText, marksToProps, nodeStyleId, parseDataUrl, plainBlocks, sectionsOf, sniffImage, stylesOf, textOf, type CharProps } from "../common";
import { cleanText, esc } from "../docx/xml";
import { ODF_CHAR, ODF_NS, ODF_PARA, StyleBook, XML_DECL, fontDecls, hasContent, metaXml, namedStyles, pageLayout, paraPropsAttrs, type ListLevelDef } from "./writerStyles";

const MIME_ODT = "application/vnd.oasis.opendocument.text";
const pxCm = (px: number) => `${+((px / 96) * 2.54).toFixed(4)}cm`;
const twCm = (tw: number) => `${+((tw / 1440) * 2.54).toFixed(4)}cm`;

/** Metin: sınırdaki ve ardışık boşluklar text:s, sekme ve satır sonu öğe olarak. */
export function odtText(s: string): string {
  let out = "";
  const parts = cleanText(s).split(/(\t|\n| +)/);
  parts.forEach((p, i) => {
    if (!p) return;
    if (p === "\t") out += "<text:tab/>";
    else if (p === "\n") out += "<text:line-break/>";
    else if (p[0] === " ") {
      const inner = i > 0 && i < parts.length - 1 && parts[i - 1] !== "" && parts[i + 1] !== "" && !/[\t\n]/.test(parts[i - 1]) && !/[\t\n]/.test(parts[i + 1]);
      if (inner) out += p.length > 1 ? ` <text:s text:c="${p.length - 1}"/>` : " ";
      else out += p.length > 1 ? `<text:s text:c="${p.length}"/>` : "<text:s/>";
    } else out += esc(p);
  });
  return out;
}

const isDeleted = (n: JSONContent) => !!n.marks?.some((m) => m.type === "deletion");

class OdtWriter {
  fonts = new Set<string>();
  body: StyleBook;
  hfBook: StyleBook;
  styles: Record<string, StyleDef>;
  media: { path: string; mime: string; data: Uint8Array }[] = [];
  pendingBreak: "page" | "column" | null = null;
  pendingMaster: string | null = null;
  masters: { name: string; layout: string; page: PageSetup }[] = [];
  notes = { footnote: 0, endnote: 0 };
  tables = 0;
  frames = 0;
  /** Sonraki paragrafa bağlanacak metin kutuları. */
  pendingFrames = "";
  /** Açık açıklama aralıkları (office:annotation … office:annotation-end) ve not içi yazım durumu. */
  openComments = new Set<string>();
  private commentNames = new Map<string, string>();
  private inNote = false;
  /** Değişiklik izleme bölgeleri (text:tracked-changes), açık ekleme aralığı ve son silme bölgesi. */
  private changes: { id: string; kind: "insertion" | "deletion"; author: string; date: string; text: string }[] = [];
  private openIns: { key: string; id: string } | null = null;
  private lastDel: { key: string; region: { text: string } } | null = null;

  constructor(readonly doc: BelgeDocument) {
    this.body = new StyleBook("", this.fonts);
    this.hfBook = new StyleBook("M", this.fonts);
    this.styles = stylesOf(doc.settings);
  }

  /* -------------------------- satır içi -------------------------- */

  inline(nodes: JSONContent[], book: StyleBook): string {
    let out = "";
    let i = 0;
    // Açıklamalar ve değişiklik izleme yalnız belge gövdesinde (üst/alt bilgi ve not metni dışında) yazılır.
    const edges = (n: JSONContent) => (book === this.body && !this.inNote ? this.trackEdges(n) + this.commentEdges(n) : "");
    while (i < nodes.length) {
      const link = marksToProps(nodes[i].marks).link;
      if (link) {
        let inner = "";
        while (i < nodes.length && marksToProps(nodes[i].marks).link === link) {
          inner += edges(nodes[i]);
          inner += this.inlineNode(nodes[i++], book);
        }
        out += `<text:a xlink:type="simple" xlink:href="${esc(link)}" text:style-name="Internet_20_link" text:visited-style-name="Visited_20_Internet_20_Link">${inner}</text:a>`;
        continue;
      }
      out += edges(nodes[i]);
      out += this.inlineNode(nodes[i++], book);
    }
    if (book === this.body && !this.inNote) out += this.closeTrack();
    return out;
  }

  /** Ekleme aralığının başı/sonu ve silinen metnin konumu (silinen metin bölgede saklanır). */
  private trackEdges(n: JSONContent): string {
    const ins = n.marks?.find((m) => m.type === "insertion");
    const del = n.marks?.find((m) => m.type === "deletion");
    const who = (m: { attrs?: Record<string, unknown> }) => ({ author: String(m.attrs?.author ?? ""), date: String(m.attrs?.date ?? "") });
    let out = "";
    const key = ins && !del ? JSON.stringify(who(ins)) : null;
    if (this.openIns && this.openIns.key !== key) {
      out += `<text:change-end text:change-id="${this.openIns.id}"/>`;
      this.openIns = null;
    }
    if (key && !this.openIns) {
      const id = `ct${this.changes.length + 1}`;
      this.changes.push({ id, kind: "insertion", ...who(ins!), text: "" });
      this.openIns = { key, id };
      out += `<text:change-start text:change-id="${id}"/>`;
    }
    if (!del) {
      this.lastDel = null;
      return out;
    }
    const text = n.type === "text" ? (n.text ?? "") : n.type === "hardBreak" ? " " : (atomText(n) ?? "");
    const dkey = JSON.stringify(who(del));
    if (this.lastDel && this.lastDel.key === dkey) {
      this.lastDel.region.text += text;
      return out;
    }
    const id = `ct${this.changes.length + 1}`;
    const region = { id, kind: "deletion" as const, ...who(del), text };
    this.changes.push(region);
    this.lastDel = { key: dkey, region };
    return `${out}<text:change text:change-id="${id}"/>`;
  }

  private closeTrack(): string {
    this.lastDel = null;
    if (!this.openIns) return "";
    const id = this.openIns.id;
    this.openIns = null;
    return `<text:change-end text:change-id="${id}"/>`;
  }

  /** text:tracked-changes: bölgeler (yazar, tarih; silmelerde silinen metin). */
  private trackedChangesXml(): string {
    if (!this.changes.length) return "";
    const info = (c: { author: string; date: string }) => `<office:change-info><dc:creator>${esc(c.author)}</dc:creator>${c.date ? `<dc:date>${esc(c.date.replace(/\.\d+Z$|Z$/, ""))}</dc:date>` : ""}</office:change-info>`;
    const regions = this.changes
      .map((c) => `<text:changed-region xml:id="${c.id}" text:id="${c.id}">${c.kind === "insertion" ? `<text:insertion>${info(c)}</text:insertion>` : `<text:deletion>${info(c)}<text:p>${odtText(c.text)}</text:p></text:deletion>`}</text:changed-region>`)
      .join("");
    return `<text:tracked-changes text:track-changes="false">${regions}</text:tracked-changes>`;
  }

  /** Düğümün açıklama işaretlerine göre kapanan ve açılan açıklama aralıkları. */
  private commentEdges(n: JSONContent): string {
    if (isDeleted(n)) return "";
    const ids = new Set((n.marks ?? []).filter((m) => m.type === "comment").map((m) => String(m.attrs?.id ?? "")));
    let out = "";
    for (const id of [...this.openComments]) {
      if (ids.has(id)) continue;
      out += `<office:annotation-end office:name="${this.commentName(id)}"/>`;
      this.openComments.delete(id);
    }
    for (const id of ids) {
      if (this.openComments.has(id)) continue;
      const c = this.doc.settings.comments?.[id];
      if (!c) continue;
      this.openComments.add(id);
      out += this.annotation(id, c);
    }
    return out;
  }

  private commentName(id: string): string {
    let name = this.commentNames.get(id);
    if (!name) {
      name = `__Annotation__${this.commentNames.size + 1}`;
      this.commentNames.set(id, name);
    }
    return name;
  }

  /** office:annotation: yazar, tarih, baş harfler ve metin paragrafları. */
  private annotation(id: string, c: CommentData): string {
    const paras = String(c.text ?? "")
      .split("\n")
      .map((t) => `<text:p>${odtText(t)}</text:p>`)
      .join("");
    const resolved = c.resolved ? ' loext:resolved="true"' : "";
    const date = c.date ? `<dc:date>${esc(c.date.replace(/Z$/, ""))}</dc:date>` : "";
    const initials = c.initials ? `<meta:creator-initials>${esc(c.initials)}</meta:creator-initials>` : "";
    return `<office:annotation office:name="${this.commentName(id)}"${resolved}><dc:creator>${esc(c.author ?? "")}</dc:creator>${date}${initials}${paras}</office:annotation>`;
  }

  private span(text: string, p: CharProps, book: StyleBook): string {
    const { link: _l, charStyle, ...rest } = p;
    const parent = charStyle && ODF_CHAR[charStyle] ? ODF_CHAR[charStyle][0] : null;
    const name = book.text(rest, parent);
    return name ? `<text:span text:style-name="${name}">${text}</text:span>` : text;
  }

  inlineNode(n: JSONContent, book: StyleBook): string {
    if (isDeleted(n)) return "";
    const p = marksToProps(n.marks);
    switch (n.type) {
      case "text":
        return n.text ? this.span(odtText(n.text), p, book) : "";
      case "hardBreak":
        return "<text:line-break/>";
      case "pageField": {
        const f = n.attrs?.field;
        let el = '<text:page-number text:select-page="current">1</text:page-number>';
        if (f === "NUMPAGES") el = "<text:page-count>1</text:page-count>";
        else if (f === "DATE" || f === "TIME") {
          // Güncellenen (sabit olmayan) alan; biçim veri stiliyle korunur.
          const now = new Date();
          const pic = fieldPicture(f, n.attrs?.format)!;
          const time = f === "TIME";
          const ds = book.dataStyle(pic, time);
          const text = odtText(fieldNowText(f, now, n.attrs?.format) ?? "");
          const pad = (x: number) => String(x).padStart(2, "0");
          const value = `${now.getFullYear()}-${pad(now.getMonth() + 1)}-${pad(now.getDate())}T${pad(now.getHours())}:${pad(now.getMinutes())}:${pad(now.getSeconds())}`;
          el = time
            ? `<text:time style:data-style-name="${ds}" text:time-value="${value}">${text}</text:time>`
            : `<text:date style:data-style-name="${ds}" text:date-value="${value}">${text}</text:date>`;
        }
        return this.span(el, p, book);
      }
      case "image":
        return this.image(n, book);
      case "note":
        return this.note(n);
      case "bookmark": {
        const name = String(n.attrs?.name ?? "");
        return name ? `<text:bookmark text:name="${esc(name)}"/>` : "";
      }
      default: {
        const t = atomText(n);
        return t ? this.span(odtText(t), p, book) : "";
      }
    }
  }

  image(n: JSONContent, book: StyleBook): string {
    const a = n.attrs ?? {};
    const data = parseDataUrl(String(a.src ?? ""));
    if (!data) return "";
    const kind = sniffImage(data.bytes);
    const path = `Pictures/resim${this.media.length + 1}.${kind?.ext ?? "png"}`;
    this.media.push({ path, mime: kind?.mime ?? data.mime, data: data.bytes });
    const w = Math.max(1, Number(a.width) || 200);
    const h = Math.max(1, Number(a.height) || 150);
    const style = book.raw("graphic", "fr", '<style:graphic-properties style:vertical-pos="top" style:vertical-rel="baseline" fo:padding="0cm" fo:border="none"/>', ' style:parent-style-name="Graphics"');
    const id = ++this.frames;
    const alt = a.alt ? `<svg:desc>${esc(String(a.alt))}</svg:desc>` : "";
    return `<draw:frame draw:style-name="${style}" draw:name="Resim${id}" text:anchor-type="as-char" svg:width="${pxCm(w)}" svg:height="${pxCm(h)}" draw:z-index="${id}"><draw:image xlink:href="${path}" xlink:type="simple" xlink:show="embed" xlink:actuate="onLoad" draw:mime-type="${kind?.mime ?? data.mime}"/>${alt}</draw:frame>`;
  }

  note(n: JSONContent): string {
    const a = n.attrs ?? {};
    const end = a.kind === "endnote";
    const num = end ? ++this.notes.endnote : ++this.notes.footnote;
    const blocks: JSONContent[] = (a.content as JSONContent[] | null)?.length
      ? (a.content as JSONContent[])
      : String(a.text ?? "")
          .split("\n")
          .map((t) => ({ type: "paragraph", ...(t ? { content: [{ type: "text", text: t }] } : {}) }));
    const wasInNote = this.inNote;
    this.inNote = true;
    const body = blocks.map((b) => (b.type === "paragraph" || b.type === "heading" ? `<text:p text:style-name="${end ? "Endnote" : "Footnote"}">${this.inline(b.content ?? [], this.body)}</text:p>` : this.block(b, this.body))).join("");
    this.inNote = wasInNote;
    const roman = ["i", "ii", "iii", "iv", "v", "vi", "vii", "viii", "ix", "x"][num - 1] ?? String(num);
    return `<text:note text:id="${end ? "edn" : "ftn"}${num}" text:note-class="${end ? "endnote" : "footnote"}"><text:note-citation>${end ? roman : num}</text:note-citation><text:note-body>${body}</text:note-body></text:note>`;
  }

  /* -------------------------- bloklar -------------------------- */

  blocks(nodes: JSONContent[] | undefined, book: StyleBook): string {
    return (nodes ?? []).map((n) => this.block(n, book)).join("");
  }

  /** Bekleyen sayfa sonu / asıl sayfa değişimi paragraf stiline eklenir. */
  private takeBreak(): { brk: "page" | "column" | null; master: string | null } {
    const r = { brk: this.pendingBreak, master: this.pendingMaster };
    this.pendingBreak = null;
    this.pendingMaster = null;
    return r;
  }

  paragraph(n: JSONContent, book: StyleBook, opts: { inList?: boolean; style?: string } = {}): string {
    const a = (n.attrs ?? {}) as Record<string, unknown>;
    const id = nodeStyleId(n);
    const parent = opts.style ?? ODF_PARA[id]?.[0] ?? "Standard";
    const { brk, master } = this.takeBreak();
    const direct = { ...a };
    if (opts.inList) {
      delete direct.indentLeft;
      delete direct.indentFirst;
    }
    // Büyük ilk harf (ODF'de kenar boşluğu kipi yoktur: metin içinde yazılır).
    const dc = opts.inList ? null : normalizeDropCap(a.dropCap);
    const dcFont = dc?.font ? book.text({ font: dc.font } as CharProps, null) : null;
    const dropXml = dc ? `<style:drop-cap style:lines="${dc.lines}" style:length="1" style:distance="${(dc.distance / (72 / 2.54)).toFixed(3)}cm"${dcFont ? ` style:style-name="${dcFont}"` : ""}/>` : "";
    const style = book.para(parent, paraPropsAttrs(direct, master ? null : brk), master ? { "style:master-page-name": master } : {}, dropXml);
    const frames = this.pendingFrames;
    this.pendingFrames = "";
    const inner = frames + this.inline(n.content ?? [], book);
    if (n.type === "heading") {
      const level = Math.min(6, Math.max(1, Number(a.level) || 1));
      return `<text:h text:style-name="${style}" text:outline-level="${level}">${inner}</text:h>`;
    }
    return `<text:p text:style-name="${style}">${inner}</text:p>`;
  }

  block(n: JSONContent, book: StyleBook, inList = false): string {
    switch (n.type) {
      case "paragraph":
      case "heading":
        return this.paragraph(n, book, { inList });
      case "bulletList":
      case "orderedList":
        return this.list(n, book);
      case "table":
        return inList ? this.blocks(textRows(n), book) : this.table(n, book);
      case "horizontalRule":
        return this.paragraph({ type: "paragraph" }, book, { style: "Horizontal_20_Line" });
      case "pageBreak":
        this.pendingBreak = "page";
        return "";
      case "columnBreak":
        this.pendingBreak = "column";
        return "";
      case "textBox":
        this.pendingFrames += this.textBox(n, book);
        return "";
      case "toc":
        return this.toc(n, book);
      case "bibliography":
      case "indexBlock":
        return this.blocks(plainBlocks(n), book);
      default:
        return n.content ? this.blocks(n.content, book) : "";
    }
  }

  textBox(n: JSONContent, book: StyleBook): string {
    const a = (n.attrs ?? {}) as Record<string, unknown>;
    const border = a.border ? `fo:border="${Number(a.borderWidth) || 0.75}pt solid ${esc(String(a.border))}"` : 'fo:border="none"';
    const fill = a.fill ? ` draw:fill="solid" draw:fill-color="${esc(String(a.fill))}" fo:background-color="${esc(String(a.fill))}"` : ' draw:fill="none"';
    const style = book.raw("graphic", "fr", `<style:graphic-properties ${border}${fill} fo:padding="0.1cm" style:wrap="run-through" style:run-through="foreground" style:vertical-pos="from-top" style:vertical-rel="paragraph" style:horizontal-pos="from-left" style:horizontal-rel="paragraph"/>`);
    const id = ++this.frames;
    const saved = this.pendingFrames;
    this.pendingFrames = "";
    const inner = this.blocks(n.content, book) || "<text:p/>";
    this.pendingFrames = saved;
    return `<draw:frame draw:style-name="${style}" draw:name="Metin Kutusu ${id}" text:anchor-type="paragraph" svg:x="${pxCm(Number(a.x) || 0)}" svg:y="${pxCm(Number(a.y) || 0)}" svg:width="${pxCm(Number(a.w) || 240)}" svg:height="${pxCm(Number(a.h) || 120)}" draw:z-index="${id}"><draw:text-box>${inner}</draw:text-box></draw:frame>`;
  }

  list(n: JSONContent, book: StyleBook): string {
    const levels: ListLevelDef[] = [];
    const scan = (l: JSONContent, d: number) => {
      if (d > 9) return;
      levels[d] ??= { ordered: l.type === "orderedList", style: (l.attrs?.listStyle as string | null) ?? null, start: Number(l.attrs?.start) || 1, multilevel: (l.attrs?.multilevel as string | null) ?? null };
      for (const li of l.content ?? []) for (const c of li.content ?? []) if (c.type === "bulletList" || c.type === "orderedList") scan(c, d + 1);
    };
    scan(n, 0);
    for (let k = 0; k < 10; k++) levels[k] ??= { ordered: levels[k - 1]?.ordered ?? false, style: null, start: 1, multilevel: levels[k - 1]?.multilevel ?? null };
    const name = book.list(levels);
    const walk = (l: JSONContent, d: number): string => {
      let out = d === 0 ? `<text:list text:style-name="${name}">` : "<text:list>";
      const start = Number(l.attrs?.start) || 1;
      (l.content ?? []).forEach((li, i) => {
        const sv = i === 0 && l.type === "orderedList" && start !== levels[Math.min(9, d)].start ? ` text:start-value="${start}"` : "";
        out += `<text:list-item${sv}>`;
        const kids = li.content?.length ? li.content : [{ type: "paragraph" }];
        for (const c of kids) {
          if (c.type === "bulletList" || c.type === "orderedList") out += d < 9 ? walk(c, d + 1) : "";
          else if (c.type === "paragraph" || c.type === "heading") out += this.paragraph(c.type === "heading" ? { ...c, type: "paragraph" } : c, book, { inList: true });
          else out += this.block(c, book, true);
        }
        out += "</text:list-item>";
      });
      return `${out}</text:list>`;
    };
    return walk(n, 0);
  }

  table(n: JSONContent, book: StyleBook): string {
    const rows = n.content ?? [];
    const id = ++this.tables;
    const page = this.doc.settings.page;
    const contentTw = page.width - page.margin.left - page.margin.right - page.margin.gutter;
    const widths: number[] = [];
    for (const c of rows[0]?.content ?? []) {
      const span = Math.max(1, Number(c.attrs?.colspan) || 1);
      const cw = c.attrs?.colwidth as number[] | null | undefined;
      for (let k = 0; k < span; k++) widths.push(cw?.[k] ? cw[k] * 15 : 0);
    }
    const cols = Math.max(1, widths.length);
    const free = widths.filter((w) => !w).length;
    const rest = Math.max(0, contentTw - widths.reduce((s, w) => s + w, 0));
    for (let k = 0; k < cols; k++) if (!widths[k]) widths[k] = Math.max(300, free ? rest / free : contentTw / cols);
    const total = widths.reduce((s, w) => s + w, 0);
    const { brk, master } = this.takeBreak();
    const align = n.attrs?.align === "center" ? "center" : n.attrs?.align === "right" ? "right" : "left";
    // Soldan girinti (px → twip): sola yaslı tabloda fo:margin-left.
    const indent = align === "left" ? Math.round((Number(n.attrs?.indent) || 0) * 15) : 0;
    const margin = indent ? ` fo:margin-left="${twCm(indent)}"` : "";
    const tStyle = book.raw("table", `Tablo`, `<style:table-properties style:width="${twCm(total)}" table:align="${align}"${margin}${brk === "page" ? ' fo:break-before="page"' : ""}/>`, master ? ` style:master-page-name="${master}"` : "");
    let x = `<table:table table:name="Tablo${id}" table:style-name="${tStyle}">`;
    for (const w of widths) x += `<table:table-column table:style-name="${book.raw("table-column", "Sutun", `<style:table-column-properties style:column-width="${twCm(w)}"/>`)}"/>`;
    const border = n.attrs?.tableStyle === "none" ? "none" : "0.5pt solid #000000";
    const covered: number[] = new Array(cols).fill(0);
    rows.forEach((row, r) => {
      let rx = "<table:table-row>";
      let col = 0;
      const cells = row.content ?? [];
      let ci = 0;
      while (col < cols) {
        if (covered[col] > 0) {
          covered[col]--;
          rx += "<table:covered-table-cell/>";
          col++;
          continue;
        }
        const c = cells[ci++];
        if (!c) {
          rx += `<table:table-cell office:value-type="string"><text:p text:style-name="Table_20_Contents"/></table:table-cell>`;
          col++;
          continue;
        }
        const span = Math.max(1, Math.min(cols - col, Number(c.attrs?.colspan) || 1));
        const rs = Math.max(1, Math.min(rows.length - r, Number(c.attrs?.rowspan) || 1));
        for (let k = 0; k < span; k++) covered[col + k] = rs - 1;
        const bg = c.attrs?.background ? ` fo:background-color="${esc(String(c.attrs.background))}"` : "";
        const va = c.attrs?.verticalAlign === "middle" ? ' style:vertical-align="middle"' : c.attrs?.verticalAlign === "bottom" ? ' style:vertical-align="bottom"' : "";
        const cs = book.raw("table-cell", "Hucre", `<style:table-cell-properties fo:padding="0.097cm" fo:border="${border}"${bg}${va}/>`);
        const inner = this.blocks(c.content?.length ? c.content : [{ type: "paragraph" }], book);
        rx += `<table:table-cell table:style-name="${cs}" office:value-type="string"${span > 1 ? ` table:number-columns-spanned="${span}"` : ""}${rs > 1 ? ` table:number-rows-spanned="${rs}"` : ""}>${inner}</table:table-cell>`;
        for (let k = 1; k < span; k++) rx += "<table:covered-table-cell/>";
        col += span;
      }
      rx += "</table:table-row>";
      x += r === 0 && n.attrs?.repeatHeader ? `<table:table-header-rows>${rx}</table:table-header-rows>` : rx;
    });
    return `${x}</table:table>`;
  }

  toc(n: JSONContent, book: StyleBook): string {
    const entries = (n.attrs?.entries as { level: number; text: string; page: number }[] | undefined) ?? [];
    const title = String(n.attrs?.title ?? "İçindekiler");
    let src = `<text:table-of-content-source text:outline-level="3"><text:index-title-template text:style-name="Contents_20_Heading">${esc(title)}</text:index-title-template>`;
    for (let i = 1; i <= 3; i++) src += `<text:table-of-content-entry-template text:outline-level="${i}" text:style-name="Contents_20_${i}"><text:index-entry-text/><text:index-entry-tab-stop style:type="right" style:leader-char="."/><text:index-entry-page-number/></text:table-of-content-entry-template>`;
    src += "</text:table-of-content-source>";
    const { brk } = this.takeBreak();
    const head = book.para("Contents_20_Heading", paraPropsAttrs({}, brk));
    let body = `<text:index-body>${title ? `<text:index-title text:name="İçindekiler_baslik"><text:p text:style-name="${head}">${esc(title)}</text:p></text:index-title>` : ""}`;
    for (const e of entries) body += `<text:p text:style-name="Contents_20_${Math.min(9, Math.max(1, e.level))}">${odtText(e.text)}<text:tab/>${e.page}</text:p>`;
    return `<text:table-of-content text:name="İçindekiler1" text:protected="true">${src}${body}</text:index-body></text:table-of-content>`;
  }

  /* -------------------------- paket -------------------------- */

  async write(): Promise<Uint8Array> {
    const s = this.doc.settings;
    const hasHeader = hasContent(s.header) || hasContent(s.headerFirst) || hasContent(s.headerEven);
    const hasFooter = hasContent(s.footer) || hasContent(s.footerFirst) || hasContent(s.footerEven);
    let text = "";
    sectionsOf(this.doc.content).forEach((sec, i) => {
      if (i > 0) {
        const page = (sec.attrs?.page as PageSetup | null) ?? s.page;
        if (!samePage(page, s.page)) {
          const name = `Bolum${i + 1}`;
          this.masters.push({ name, layout: `pm${this.masters.length + 2}`, page });
          this.pendingMaster = name;
        } else if (sec.attrs?.kind !== "continuous") this.pendingBreak = "page";
      }
      text += this.blocks(sec.content, this.body);
    });
    if (this.pendingBreak || this.pendingMaster || this.pendingFrames) text += this.paragraph({ type: "paragraph" }, this.body);
    // Belge sonunda açık kalan açıklama aralıkları son paragrafın içinde kapanır.
    if (this.openComments.size) {
      const ends = [...this.openComments].map((id) => `<office:annotation-end office:name="${this.commentName(id)}"/>`).join("");
      this.openComments.clear();
      const at = Math.max(text.lastIndexOf("</text:p>"), text.lastIndexOf("</text:h>"));
      text = at >= 0 ? text.slice(0, at) + ends + text.slice(at) : text + `<text:p>${ends}</text:p>`;
    }
    // Üst/alt bilgiler (styles.xml kapsamı)
    const hf = (key: "header" | "footer" | "headerFirst" | "footerFirst" | "headerEven" | "footerEven", tag: string) => {
      const d = s[key];
      if (!hasContent(d)) return "";
      const style = key.startsWith("header") ? "Header" : "Footer";
      const inner = (d!.content ?? []).map((b) => (b.type === "paragraph" || b.type === "heading" ? this.paragraph({ ...b, type: "paragraph" }, this.hfBook, { style }) : this.block(b, this.hfBook))).join("");
      return `<style:${tag}>${inner}</style:${tag}>`;
    };
    const hfXml = hf("header", "header") + (s.oddEven ? hf("headerEven", "header-left") : "") + (s.page.titlePage ? hf("headerFirst", "header-first") : "") + hf("footer", "footer") + (s.oddEven ? hf("footerEven", "footer-left") : "") + (s.page.titlePage ? hf("footerFirst", "footer-first") : "");
    const named = namedStyles(this.styles, this.fonts);
    const layouts = pageLayout("pm1", s.page, hasHeader, hasFooter, s.pageColor) + this.masters.map((m) => pageLayout(m.layout, m.page, hasHeader, hasFooter, s.pageColor)).join("");
    const masters = `<style:master-page style:name="Standard" style:page-layout-name="pm1">${hfXml}</style:master-page>` + this.masters.map((m) => `<style:master-page style:name="${m.name}" style:page-layout-name="${m.layout}">${hfXml}</style:master-page>`).join("");
    const fonts = fontDecls(this.fonts);
    const content = `${XML_DECL}<office:document-content ${ODF_NS} office:version="1.3">${fonts}<office:automatic-styles>${this.body.toXml()}</office:automatic-styles><office:body><office:text>${this.trackedChangesXml()}${text || '<text:p text:style-name="Standard"/>'}</office:text></office:body></office:document-content>`;
    const styles = `${XML_DECL}<office:document-styles ${ODF_NS} office:version="1.3">${fonts}${named}<office:automatic-styles>${layouts}${this.hfBook.toXml()}</office:automatic-styles><office:master-styles>${masters}</office:master-styles></office:document-styles>`;
    const zip = new JSZip();
    zip.file("mimetype", MIME_ODT, { compression: "STORE" });
    zip.file("content.xml", content);
    zip.file("styles.xml", styles);
    zip.file("meta.xml", metaXml(s));
    for (const m of this.media) zip.file(m.path, m.data, { compression: "STORE" });
    const manifest =
      `${XML_DECL}<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0" manifest:version="1.3">` +
      `<manifest:file-entry manifest:full-path="/" manifest:version="1.3" manifest:media-type="${MIME_ODT}"/>` +
      ["content.xml", "styles.xml", "meta.xml"].map((f) => `<manifest:file-entry manifest:full-path="${f}" manifest:media-type="text/xml"/>`).join("") +
      this.media.map((m) => `<manifest:file-entry manifest:full-path="${m.path}" manifest:media-type="${m.mime}"/>`).join("") +
      "</manifest:manifest>";
    zip.file("META-INF/manifest.xml", manifest);
    return zip.generateAsync({ type: "uint8array", compression: "DEFLATE", compressionOptions: { level: 6 } });
  }
}

/** Liste içindeki tablo: satırlar sekmeyle ayrılmış paragraflara düşer (ODF liste öğesi tablo içeremez). */
function textRows(n: JSONContent): JSONContent[] {
  return (n.content ?? []).map((r) => ({ type: "paragraph", content: [{ type: "text", text: (r.content ?? []).map((c) => textOf(c)).join("\t") }].filter((t) => t.text) }));
}

/** Belge → ODT baytları. */
export function writeOdt(doc: BelgeDocument): Promise<Uint8Array> {
  return new OdtWriter(doc).write();
}
