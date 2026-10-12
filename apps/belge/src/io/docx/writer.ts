// DOCX yazıcı: ProseMirror JSON + belge ayarları → OOXML paketi (ECMA-376).
import JSZip from "jszip";
import type { JSONContent } from "@tiptap/core";
import type { BelgeDocument, CommentData, DocSettings } from "../../model/document";
import { resolveStyles, type BorderSpec } from "../../model/styles";
import type { ParaBorders } from "../../editor/paragraphFormat";
import type { TocEntry } from "../../editor/nodes";
import { NS, REL, XML_HEAD, cleanText, esc, hex6, highlightName } from "./xml";
import { borderXml, buildStylesXml, fontsXml, tableStyleId } from "./styleXml";
import { levelDef, sigKey, type ListSig } from "./numbering";

type Mark = { type: string; attrs?: Record<string, unknown> };

const DOC_NS = `xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas" xmlns:mc="${NS.mc}" xmlns:o="${NS.o}" xmlns:r="${NS.r}" xmlns:v="${NS.v}" xmlns:w10="${NS.w10}" xmlns:wp="${NS.wp}" xmlns:a="${NS.a}" xmlns:pic="${NS.pic}" xmlns:w="${NS.w}"`;

const EMU_PER_PX = 9525;

interface Media {
  name: string;
  data: Uint8Array;
  ext: string;
}

/** Bir parça (belge, üst/alt bilgi) için ilişki tablosu. */
class Rels {
  private list: { id: string; type: string; target: string; external?: boolean }[] = [];
  add(type: string, target: string, external = false): string {
    const hit = this.list.find((r) => r.type === type && r.target === target);
    if (hit) return hit.id;
    const id = `rId${this.list.length + 1}`;
    this.list.push({ id, type, target, external });
    return id;
  }
  xml(): string {
    return `${XML_HEAD}<Relationships xmlns="${NS.rel}">${this.list
      .map((r) => `<Relationship Id="${r.id}" Type="${r.type}" Target="${esc(r.target)}"${r.external ? ' TargetMode="External"' : ""}/>`)
      .join("")}</Relationships>`;
  }
  get size() {
    return this.list.length;
  }
}

function dataUrlToBytes(src: string): { data: Uint8Array; ext: string } | null {
  const m = /^data:([^;,]+)(;base64)?,(.*)$/s.exec(src);
  if (!m) return null;
  const mime = m[1].toLowerCase();
  const ext = mime.includes("png") ? "png" : mime.includes("jpeg") || mime.includes("jpg") ? "jpeg" : mime.includes("gif") ? "gif" : mime.includes("bmp") ? "bmp" : mime.includes("svg") ? "svg" : mime.includes("webp") ? "webp" : "png";
  let bytes: Uint8Array;
  if (m[2]) {
    const bin = atob(m[3]);
    bytes = new Uint8Array(bin.length);
    for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
  } else bytes = new TextEncoder().encode(decodeURIComponent(m[3]));
  return { data: bytes, ext };
}

const ptTw = (pt: number) => Math.round(pt * 20);

/** Yazım bağlamı: belge genelinde sayaçlar. */
class Ctx {
  media: Media[] = [];
  numbering = new Map<string, { abstractId: number; sig: ListSig }>();
  nums: { numId: number; abstractId: number; start: number }[] = [];
  bookmarkId = 0;
  drawingId = 1;
  commentIds = new Map<string, number>();
  openComments = new Set<string>();
  usedComments = new Set<string>();
  revId = 1;
  footnotes: string[] = [];
  endnotes: string[] = [];
  constructor(public settings: DocSettings) {}

  commentNum(id: string): number {
    let n = this.commentIds.get(id);
    if (n === undefined) {
      n = this.commentIds.size;
      this.commentIds.set(id, n);
    }
    return n;
  }

  numFor(sig: ListSig, start: number): number {
    const key = sigKey(sig);
    let abs = this.numbering.get(key);
    if (!abs) {
      abs = { abstractId: this.numbering.size, sig };
      this.numbering.set(key, abs);
    }
    const numId = this.nums.length + 1;
    this.nums.push({ numId, abstractId: abs.abstractId, start });
    return numId;
  }
}

/** Bir parçanın (gövde/üst bilgi/alt bilgi) bloklarını yazar. */
class PartWriter {
  constructor(
    private ctx: Ctx,
    private rels: Rels,
    private contentWidthTw: number,
  ) {}

  blocks(nodes: JSONContent[] | undefined, list?: { numId: number; ilvl: number }): string {
    let out = "";
    for (const n of nodes ?? []) out += this.block(n, list);
    return out;
  }

  block(n: JSONContent, list?: { numId: number; ilvl: number }): string {
    switch (n.type) {
      case "paragraph":
      case "heading":
        return this.paragraph(n, list);
      case "bulletList":
      case "orderedList":
        return this.list(n, list ? list.ilvl + 1 : 0);
      case "table":
        return this.table(n);
      case "horizontalRule":
        return '<w:p><w:pPr><w:pBdr><w:bottom w:val="single" w:sz="6" w:space="1" w:color="auto"/></w:pBdr></w:pPr></w:p>';
      case "pageBreak":
        return '<w:p><w:r><w:br w:type="page"/></w:r></w:p>';
      case "toc":
        return this.toc(n);
      default:
        return n.content ? this.blocks(n.content, list) : "";
    }
  }

  list(n: JSONContent, ilvl: number): string {
    const sig: ListSig = {
      type: n.type === "bulletList" ? "bullet" : "ordered",
      style: (n.attrs?.listStyle as string | null) ?? null,
      multilevel: (n.attrs?.multilevel as string | null) ?? null,
    };
    const numId = this.ctx.numFor(sig, (n.attrs?.start as number | undefined) ?? 1);
    let out = "";
    for (const li of n.content ?? []) {
      (li.content ?? []).forEach((child, i) => {
        if (i === 0 && (child.type === "paragraph" || child.type === "heading")) out += this.paragraph(child, { numId, ilvl });
        else if (child.type === "paragraph" || child.type === "heading") out += this.paragraph(child, undefined, 720 * (ilvl + 1));
        else out += this.block(child, { numId, ilvl });
      });
    }
    return out;
  }

  pPr(n: JSONContent, list?: { numId: number; ilvl: number }, contIndent?: number, extra = ""): string {
    const a = (n.attrs ?? {}) as Record<string, unknown>;
    let x = "";
    let style: string | null = null;
    if (n.type === "heading") style = `Heading${a.level ?? 1}`;
    else if (a.styleId) style = a.styleId as string;
    else if (list || contIndent) style = "ListParagraph";
    if (style) x += `<w:pStyle w:val="${esc(style)}"/>`;
    if (a.keepNext) x += "<w:keepNext/>";
    if (a.keepLines) x += "<w:keepLines/>";
    if (a.pageBreakBefore) x += "<w:pageBreakBefore/>";
    if (list) x += `<w:numPr><w:ilvl w:val="${list.ilvl}"/><w:numId w:val="${list.numId}"/></w:numPr>`;
    const b = a.borders as ParaBorders | null;
    if (b && (b.top || b.left || b.bottom || b.right)) {
      x += `<w:pBdr>${b.top ? borderXml("top", b.top) : ""}${b.left ? borderXml("left", b.left as BorderSpec) : ""}${b.bottom ? borderXml("bottom", b.bottom) : ""}${b.right ? borderXml("right", b.right as BorderSpec) : ""}</w:pBdr>`;
    }
    if (a.shading) x += `<w:shd w:val="clear" w:color="auto" w:fill="${hex6(a.shading as string) ?? "auto"}"/>`;
    x += extra;
    const sp: string[] = [];
    if (a.spaceBefore != null) sp.push(`w:before="${ptTw(a.spaceBefore as number)}"`);
    if (a.spaceAfter != null) sp.push(`w:after="${ptTw(a.spaceAfter as number)}"`);
    if (a.lineHeight != null) {
      const rule = (a.lineRule as string | null) ?? "auto";
      if (rule === "auto") sp.push(`w:line="${Math.round((a.lineHeight as number) * 240)}" w:lineRule="auto"`);
      else sp.push(`w:line="${ptTw(a.lineHeight as number)}" w:lineRule="${rule === "exact" ? "exact" : "atLeast"}"`);
    }
    if (sp.length) x += `<w:spacing ${sp.join(" ")}/>`;
    const ind: string[] = [];
    if (contIndent) ind.push(`w:left="${contIndent}"`);
    else if (a.indentLeft != null && !list) ind.push(`w:left="${ptTw(a.indentLeft as number)}"`);
    if (a.indentRight != null) ind.push(`w:right="${ptTw(a.indentRight as number)}"`);
    if (a.indentFirst != null && !list) {
      const f = a.indentFirst as number;
      ind.push(f >= 0 ? `w:firstLine="${ptTw(f)}"` : `w:hanging="${ptTw(-f)}"`);
    }
    if (ind.length) x += `<w:ind ${ind.join(" ")}/>`;
    const al = a.textAlign as string | null;
    if (al) x += `<w:jc w:val="${al === "justify" ? "both" : al === "left" ? "left" : al}"/>`;
    return x ? `<w:pPr>${x}</w:pPr>` : "";
  }

  paragraph(n: JSONContent, list?: { numId: number; ilvl: number }, contIndent?: number): string {
    return `<w:p>${this.pPr(n, list, contIndent)}${this.inline(n.content ?? [])}</w:p>`;
  }

  /** Satır içi içerik: köprü, izlenen değişiklik ve açıklama sarmalayıcılarıyla koşular. */
  inline(nodes: JSONContent[]): string {
    let out = "";
    let link: string | null = null;
    let rev: "ins" | "del" | null = null;
    const closeRev = () => {
      if (rev) out += rev === "ins" ? "</w:ins>" : "</w:del>";
      rev = null;
    };
    const closeLink = () => {
      closeRev();
      if (link !== null) out += "</w:hyperlink>";
      link = null;
    };
    for (const node of nodes) {
      const marks = (node.marks ?? []) as Mark[];
      // Açıklama aralıkları
      const cids = new Set(marks.filter((m) => m.type === "comment").map((m) => String(m.attrs?.id ?? "")));
      for (const id of [...this.ctx.openComments]) {
        if (!cids.has(id)) {
          closeLink();
          const num = this.ctx.commentNum(id);
          out += `<w:commentRangeEnd w:id="${num}"/><w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr><w:commentReference w:id="${num}"/></w:r>`;
          this.ctx.openComments.delete(id);
        }
      }
      for (const id of cids) {
        if (!this.ctx.openComments.has(id)) {
          closeLink();
          out += `<w:commentRangeStart w:id="${this.ctx.commentNum(id)}"/>`;
          this.ctx.openComments.add(id);
          this.ctx.usedComments.add(id);
        }
      }
      if (node.type === "bookmark") {
        const id = this.ctx.bookmarkId++;
        out += `<w:bookmarkStart w:id="${id}" w:name="${esc(String(node.attrs?.name ?? ""))}"/><w:bookmarkEnd w:id="${id}"/>`;
        continue;
      }
      const lm = marks.find((m) => m.type === "link");
      const href = lm ? String(lm.attrs?.href ?? "") : null;
      if (href !== link) {
        closeLink();
        if (href !== null) {
          if (href.startsWith("#")) out += `<w:hyperlink w:anchor="${esc(href.slice(1))}" w:history="1">`;
          else out += `<w:hyperlink r:id="${this.rels.add(REL.hyperlink, href, true)}" w:history="1">`;
          link = href;
        }
      }
      const r: "ins" | "del" | null = marks.some((m) => m.type === "insertion") ? "ins" : marks.some((m) => m.type === "deletion") ? "del" : null;
      if (r !== rev) {
        closeRev();
        if (r) {
          const m = marks.find((x) => x.type === (r === "ins" ? "insertion" : "deletion"))!;
          out += `<w:${r} w:id="${this.ctx.revId++}" w:author="${esc(String(m.attrs?.author || "Kullanıcı"))}" w:date="${esc(String(m.attrs?.date || new Date().toISOString().replace(/\.\d+Z$/, "Z")))}">`;
          rev = r;
        }
      }
      out += this.run(node, marks, link !== null, rev === "del");
    }
    closeLink();
    return out;
  }

  rPr(marks: Mark[], isLink: boolean): string {
    let x = "";
    const has = (t: string) => marks.find((m) => m.type === t);
    const ts = (has("textStyle")?.attrs ?? {}) as Record<string, unknown>;
    const cs = has("charStyle");
    if (cs) x += `<w:rStyle w:val="${esc(String(cs.attrs?.styleId))}"/>`;
    else if (isLink) x += '<w:rStyle w:val="Hyperlink"/>';
    if (ts.fontFamily) x += fontsXml(String(ts.fontFamily).replace(/["']/g, "").split(",")[0].trim());
    if (has("bold")) x += "<w:b/><w:bCs/>";
    if (has("italic")) x += "<w:i/><w:iCs/>";
    if (ts.caps) x += "<w:caps/>";
    if (ts.smallCaps) x += "<w:smallCaps/>";
    if (has("strike")) x += "<w:strike/>";
    if (ts.doubleStrike) x += "<w:dstrike/>";
    if (ts.hidden) x += "<w:vanish/>";
    const color = hex6(ts.color as string | undefined);
    if (color) x += `<w:color w:val="${color}"/>`;
    if (ts.letterSpacing) x += `<w:spacing w:val="${ptTw(Number(ts.letterSpacing))}"/>`;
    if (ts.fontSize) {
      const pt = parseFloat(String(ts.fontSize));
      if (pt > 0) x += `<w:sz w:val="${Math.round(pt * 2)}"/><w:szCs w:val="${Math.round(pt * 2)}"/>`;
    }
    const hl = has("highlight");
    let shdFill: string | null = hex6(ts.backgroundColor as string | undefined);
    if (hl) {
      const c = hex6((hl.attrs?.color as string | undefined) ?? "#FFFF00") ?? "FFFF00";
      const name = highlightName(`#${c}`);
      if (name) x += `<w:highlight w:val="${name}"/>`;
      else shdFill = c;
    }
    const u = has("underline");
    if (u) {
      const t = String(u.attrs?.type ?? "single");
      const val = t === "dashed" ? "dash" : t === "wave" ? "wave" : t;
      const uc = hex6(u.attrs?.color as string | undefined);
      x += `<w:u w:val="${val}"${uc ? ` w:color="${uc}"` : ""}/>`;
    }
    if (shdFill) x += `<w:shd w:val="clear" w:color="auto" w:fill="${shdFill}"/>`;
    if (has("superscript")) x += '<w:vertAlign w:val="superscript"/>';
    else if (has("subscript")) x += '<w:vertAlign w:val="subscript"/>';
    return x ? `<w:rPr>${x}</w:rPr>` : "";
  }

  run(node: JSONContent, marks: Mark[], isLink: boolean, deleted: boolean): string {
    const rpr = this.rPr(marks, isLink);
    switch (node.type) {
      case "text": {
        const text = cleanText(node.text ?? "");
        let body = "";
        const tTag = deleted ? "w:delText" : "w:t";
        for (const [i, seg] of text.split("\t").entries()) {
          if (i > 0) body += "<w:tab/>";
          if (seg) body += `<${tTag} xml:space="preserve">${esc(seg)}</${tTag}>`;
        }
        return `<w:r>${rpr}${body}</w:r>`;
      }
      case "hardBreak":
        return `<w:r>${rpr}<w:br/></w:r>`;
      case "pageField": {
        const f = node.attrs?.field === "NUMPAGES" ? "NUMPAGES" : "PAGE";
        return `<w:fldSimple w:instr=" ${f} \\* MERGEFORMAT "><w:r>${rpr}<w:t>1</w:t></w:r></w:fldSimple>`;
      }
      case "image":
        return this.image(node, rpr);
      case "note": {
        const end = node.attrs?.kind === "endnote";
        const list = end ? this.ctx.endnotes : this.ctx.footnotes;
        list.push(String(node.attrs?.text ?? ""));
        const id = list.length;
        return end
          ? `<w:r><w:rPr><w:rStyle w:val="EndnoteReference"/></w:rPr><w:endnoteReference w:id="${id}"/></w:r>`
          : `<w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteReference w:id="${id}"/></w:r>`;
      }
      default:
        return "";
    }
  }

  image(node: JSONContent, rpr: string): string {
    const a = (node.attrs ?? {}) as Record<string, unknown>;
    const src = String(a.src ?? "");
    const bin = dataUrlToBytes(src);
    if (!bin) return "";
    const name = `image${this.ctx.media.length + 1}.${bin.ext}`;
    this.ctx.media.push({ name, data: bin.data, ext: bin.ext });
    const rid = this.rels.add(REL.image, `media/${name}`);
    const w = Math.max(1, Math.round(Number(a.width) || 200));
    const h = Math.max(1, Math.round(Number(a.height) || 150));
    const cx = w * EMU_PER_PX;
    const cy = h * EMU_PER_PX;
    const id = this.ctx.drawingId++;
    const descr = a.alt ? ` descr="${esc(String(a.alt))}"` : "";
    const rot = Number(a.rotate) || 0;
    const xfrmAttr = `${rot ? ` rot="${rot * 60000}"` : ""}${a.flipH ? ' flipH="1"' : ""}${a.flipV ? ' flipV="1"' : ""}`;
    return (
      `<w:r>${rpr}<w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="${cx}" cy="${cy}"/><wp:effectExtent l="0" t="0" r="0" b="0"/>` +
      `<wp:docPr id="${id}" name="Resim ${id}"${descr}/><wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/></wp:cNvGraphicFramePr>` +
      `<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:nvPicPr><pic:cNvPr id="${id}" name="${name}"${descr}/><pic:cNvPicPr/></pic:nvPicPr>` +
      `<pic:blipFill><a:blip r:embed="${rid}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>` +
      `<pic:spPr><a:xfrm${xfrmAttr}><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>`
    );
  }

  table(n: JSONContent): string {
    const a = (n.attrs ?? {}) as Record<string, unknown>;
    const rows = n.content ?? [];
    // Izgara: sütun sayısı ve genişlikler (ilk satırdan).
    const colWidths: (number | null)[] = [];
    for (const cell of rows[0]?.content ?? []) {
      const span = (cell.attrs?.colspan as number) ?? 1;
      const cw = (cell.attrs?.colwidth as number[] | null) ?? null;
      for (let j = 0; j < span; j++) colWidths.push(cw?.[j] ?? null);
    }
    const ncols = colWidths.length || 1;
    const known = colWidths.filter((w): w is number => !!w);
    const autoW = Math.round(this.contentWidthTw / ncols);
    const gridTw = colWidths.map((w) => (w ? Math.round(w * 15) : known.length === ncols ? autoW : autoW));
    const style = tableStyleId(a.tableStyle as string | null);
    const fr = a.headerRow === false ? 0 : 1;
    const lr = a.totalRow ? 1 : 0;
    const fc = a.firstColumn ? 1 : 0;
    const nh = a.bandedRows === false ? 1 : 0;
    const lookVal = (fr * 0x20) | (lr * 0x40) | (fc * 0x80) | (nh * 0x200) | 0x400;
    const look = `<w:tblLook w:val="${lookVal.toString(16).toUpperCase().padStart(4, "0")}" w:firstRow="${fr}" w:lastRow="${lr}" w:firstColumn="${fc}" w:lastColumn="0" w:noHBand="${nh}" w:noVBand="1"/>`;
    const allFixed = known.length === ncols;
    const tblW = allFixed ? `<w:tblW w:w="${gridTw.reduce((s, x) => s + x, 0)}" w:type="dxa"/>` : '<w:tblW w:w="5000" w:type="pct"/>';
    const jc = a.align ? `<w:jc w:val="${a.align}"/>` : "";
    let x = `<w:tbl><w:tblPr><w:tblStyle w:val="${style}"/>${tblW}${jc}<w:tblLayout w:type="fixed"/>${look}</w:tblPr>`;
    x += `<w:tblGrid>${gridTw.map((w) => `<w:gridCol w:w="${w}"/>`).join("")}</w:tblGrid>`;
    // rowspan → vMerge: her satırda dolu ızgara sütunlarını izle.
    const pending: { col: number; span: number; left: number; width: number }[] = [];
    rows.forEach((row) => {
      x += "<w:tr>";
      let col = 0;
      const cells = row.content ?? [];
      let ci = 0;
      const emitContinue = (p: { col: number; span: number; width: number }) => {
        x += `<w:tc><w:tcPr><w:tcW w:w="${p.width}" w:type="dxa"/>${p.span > 1 ? `<w:gridSpan w:val="${p.span}"/>` : ""}<w:vMerge/></w:tcPr><w:p/></w:tc>`;
      };
      while (ci < cells.length || pending.some((p) => p.col >= col && p.left > 0)) {
        const cont = pending.find((p) => p.col === col && p.left > 0);
        if (cont) {
          emitContinue(cont);
          cont.left--;
          col += cont.span;
          continue;
        }
        if (ci >= cells.length) break;
        const cell = cells[ci++];
        const ca = (cell.attrs ?? {}) as Record<string, unknown>;
        const span = (ca.colspan as number) ?? 1;
        const rowspan = (ca.rowspan as number) ?? 1;
        const width = gridTw.slice(col, col + span).reduce((s, w) => s + w, 0);
        let tcPr = `<w:tcW w:w="${width}" w:type="dxa"/>`;
        if (span > 1) tcPr += `<w:gridSpan w:val="${span}"/>`;
        if (rowspan > 1) {
          tcPr += '<w:vMerge w:val="restart"/>';
          pending.push({ col, span, left: rowspan - 1, width });
        }
        const bg = hex6(ca.background as string | undefined);
        if (bg) tcPr += `<w:shd w:val="clear" w:color="auto" w:fill="${bg}"/>`;
        if (ca.verticalAlign) tcPr += `<w:vAlign w:val="${ca.verticalAlign === "middle" ? "center" : ca.verticalAlign}"/>`;
        let inner = this.blocks(cell.content);
        if (!inner.endsWith("</w:p>")) inner += "<w:p/>";
        x += `<w:tc><w:tcPr>${tcPr}</w:tcPr>${inner}</w:tc>`;
        col += span;
      }
      for (let i = pending.length - 1; i >= 0; i--) if (pending[i].left <= 0) pending.splice(i, 1);
      x += "</w:tr>";
    });
    return `${x}</w:tbl>`;
  }

  toc(n: JSONContent): string {
    const entries = (n.attrs?.entries as TocEntry[]) ?? [];
    const title = String(n.attrs?.title ?? "İçindekiler");
    const tab = `<w:tabs><w:tab w:val="right" w:leader="dot" w:pos="${this.contentWidthTw}"/></w:tabs>`;
    let x = `<w:sdt><w:sdtPr><w:docPartObj><w:docPartGallery w:val="Table of Contents"/><w:docPartUnique/></w:docPartObj></w:sdtPr><w:sdtContent>`;
    x += `<w:p><w:pPr><w:pStyle w:val="TOCHeading"/></w:pPr><w:r><w:t xml:space="preserve">${esc(title)}</w:t></w:r></w:p>`;
    const begin = '<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> TOC \\o "1-3" \\h \\z \\u </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r>';
    const end = '<w:r><w:fldChar w:fldCharType="end"/></w:r>';
    if (!entries.length) {
      x += `<w:p>${begin}<w:r><w:t>İçindekiler girdisi bulunamadı.</w:t></w:r>${end}</w:p>`;
    } else {
      entries.forEach((e, i) => {
        const lvl = Math.min(3, Math.max(1, e.level));
        x += `<w:p><w:pPr><w:pStyle w:val="TOC${lvl}"/>${tab}</w:pPr>${i === 0 ? begin : ""}<w:r><w:t xml:space="preserve">${esc(cleanText(e.text))}</w:t></w:r><w:r><w:tab/></w:r><w:r><w:t>${e.page}</w:t></w:r>${i === entries.length - 1 ? end : ""}</w:p>`;
      });
    }
    return `${x}</w:sdtContent></w:sdt>`;
  }
}

/** Filigran: üst bilgiye VML WordArt şekli (ECMA-376 Bölüm 4, VML). */
function watermarkXml(s: DocSettings, pageWTw: number): string {
  const wm = s.watermark;
  if (!wm) return "";
  const widthPt = Math.round((pageWTw / 20) * 0.75);
  const heightPt = Math.round(widthPt * (wm.layout === "diagonal" ? 0.4 : 0.3));
  const color = hex6(wm.color) ?? "C0C0C0";
  const rot = wm.layout === "diagonal" ? "rotation:315;" : "";
  return (
    `<w:p><w:pPr><w:pStyle w:val="Header"/></w:pPr><w:r><w:rPr><w:noProof/></w:rPr><w:pict>` +
    `<v:shapetype id="_x0000_t136" coordsize="21600,21600" o:spt="136" adj="10800" path="m@7,l@8,m@5,21600l@6,21600e"><v:formulas><v:f eqn="sum #0 0 10800"/><v:f eqn="prod #0 2 1"/><v:f eqn="sum 21600 0 @1"/><v:f eqn="sum 0 0 @2"/><v:f eqn="sum 21600 0 @3"/><v:f eqn="if @0 @3 0"/><v:f eqn="if @0 21600 @1"/><v:f eqn="if @0 0 @2"/><v:f eqn="if @0 @4 21600"/><v:f eqn="mid @5 @6"/><v:f eqn="mid @8 @5"/><v:f eqn="mid @7 @8"/><v:f eqn="mid @6 @7"/><v:f eqn="sum @6 0 @5"/></v:formulas><v:path textpathok="t" o:connecttype="custom" o:connectlocs="@9,0;@10,10800;@11,21600;@12,10800" o:connectangles="270,180,90,0"/><v:textpath on="t" fitshape="t"/><v:handles><v:h position="#0,bottomRight" xrange="6629,14971"/></v:handles><o:lock v:ext="edit" text="t" shapetype="t"/></v:shapetype>` +
    `<v:shape id="PowerPlusWaterMarkObject1" o:spid="_x0000_s1025" type="#_x0000_t136" style="position:absolute;margin-left:0;margin-top:0;width:${widthPt}pt;height:${heightPt}pt;${rot}z-index:-251657216;mso-position-horizontal:center;mso-position-horizontal-relative:margin;mso-position-vertical:center;mso-position-vertical-relative:margin" o:allowincell="f" fillcolor="#${color}" stroked="f">` +
    `<v:fill opacity="${wm.semitransparent ? ".5" : "1"}"/><v:textpath style="font-family:&quot;${esc(wm.font)}&quot;;font-size:1pt" string="${esc(wm.text)}"/><w10:wrap anchorx="margin" anchory="margin"/></v:shape></w:pict></w:r></w:p>`
  );
}

function commentsXml(ctx: Ctx, comments: Record<string, CommentData>): string {
  let x = `${XML_HEAD}<w:comments xmlns:w="${NS.w}" xmlns:r="${NS.r}">`;
  for (const [id, num] of ctx.commentIds) {
    const c = comments[id];
    if (!c) continue;
    const paras = (c.text || "").split("\n");
    x += `<w:comment w:id="${num}" w:author="${esc(c.author)}" w:date="${esc(c.date.replace(/\.\d+Z$/, "Z"))}" w:initials="${esc(c.initials)}">`;
    paras.forEach((t, i) => {
      x += `<w:p><w:pPr><w:pStyle w:val="CommentText"/></w:pPr>${i === 0 ? '<w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr><w:annotationRef/></w:r>' : ""}<w:r><w:t xml:space="preserve">${esc(cleanText(t))}</w:t></w:r></w:p>`;
    });
    x += "</w:comment>";
  }
  return `${x}</w:comments>`;
}

/** Dipnot / son not parçası (ayraçlar + notlar). */
function notesXml(kind: "footnote" | "endnote", texts: string[]): string {
  const tag = kind === "footnote" ? "footnote" : "endnote";
  const root = kind === "footnote" ? "footnotes" : "endnotes";
  const style = kind === "footnote" ? "Footnote" : "Endnote";
  const sep = (id: number, type: string, el: string) => `<w:${tag} w:type="${type}" w:id="${id}"><w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:r><w:${el}/></w:r></w:p></w:${tag}>`;
  let x = `${XML_HEAD}<w:${root} xmlns:w="${NS.w}" xmlns:r="${NS.r}">${sep(-1, "separator", "separator")}${sep(0, "continuationSeparator", "continuationSeparator")}`;
  texts.forEach((t, i) => {
    const paras = t.split("\n");
    x += `<w:${tag} w:id="${i + 1}">`;
    paras.forEach((line, k) => {
      x += `<w:p><w:pPr><w:pStyle w:val="${style}Text"/></w:pPr>${k === 0 ? `<w:r><w:rPr><w:rStyle w:val="${style}Reference"/></w:rPr><w:${tag}Ref/></w:r><w:r><w:t xml:space="preserve"> </w:t></w:r>` : ""}<w:r><w:t xml:space="preserve">${esc(cleanText(line))}</w:t></w:r></w:p>`;
    });
    x += `</w:${tag}>`;
  });
  return `${x}</w:${root}>`;
}

function numberingXml(ctx: Ctx): string {
  let x = `${XML_HEAD}<w:numbering xmlns:w="${NS.w}">`;
  for (const { abstractId, sig } of ctx.numbering.values()) {
    x += `<w:abstractNum w:abstractNumId="${abstractId}"><w:multiLevelType w:val="${sig.multilevel ? "multilevel" : "hybridMultilevel"}"/>`;
    for (let l = 0; l < 9; l++) {
      const d = levelDef(sig, l);
      const left = 720 * (l + 1);
      x += `<w:lvl w:ilvl="${l}"><w:start w:val="1"/><w:numFmt w:val="${d.numFmt}"/><w:lvlText w:val="${esc(d.lvlText)}"/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="${left}" w:hanging="360"/></w:pPr>${d.font ? `<w:rPr>${fontsXml(d.font)}</w:rPr>` : ""}</w:lvl>`;
    }
    x += "</w:abstractNum>";
  }
  for (const n of ctx.nums) {
    x += `<w:num w:numId="${n.numId}"><w:abstractNumId w:val="${n.abstractId}"/>${n.start !== 1 ? `<w:lvlOverride w:ilvl="0"><w:startOverride w:val="${n.start}"/></w:lvlOverride>` : ""}</w:num>`;
  }
  return `${x}</w:numbering>`;
}

function sectPr(s: DocSettings, headerRid: string | null, footerRid: string | null): string {
  const p = s.page;
  const m = p.margin;
  let x = "<w:sectPr>";
  if (headerRid) x += `<w:headerReference w:type="default" r:id="${headerRid}"/>`;
  if (footerRid) x += `<w:footerReference w:type="default" r:id="${footerRid}"/>`;
  x += `<w:pgSz w:w="${p.width}" w:h="${p.height}"${p.orientation === "landscape" ? ' w:orient="landscape"' : ""}/>`;
  x += `<w:pgMar w:top="${m.top}" w:right="${m.right}" w:bottom="${m.bottom}" w:left="${m.left}" w:header="${m.header}" w:footer="${m.footer}" w:gutter="${m.gutter}"/>`;
  if (s.pageBorder) {
    const b: BorderSpec = { style: s.pageBorder.style, size: s.pageBorder.width, color: s.pageBorder.color, space: 24 };
    x += `<w:pgBorders w:offsetFrom="page">${borderXml("top", b)}${borderXml("left", b)}${borderXml("bottom", b)}${borderXml("right", b)}</w:pgBorders>`;
  }
  x += `<w:cols w:space="${p.columnSpace}"${p.columns > 1 ? ` w:num="${p.columns}"` : ""}/><w:docGrid w:linePitch="360"/></w:sectPr>`;
  return x;
}

function hasHfContent(json: JSONContent | null): boolean {
  if (!json?.content) return false;
  const walk = (n: JSONContent): boolean => (n.type === "text" && !!n.text) || n.type === "pageField" || n.type === "image" || !!n.content?.some(walk);
  return json.content.some(walk);
}

export async function writeDocx(doc: BelgeDocument): Promise<Uint8Array> {
  const { settings } = doc;
  const zip = new JSZip();
  const ctx = new Ctx(settings);
  const docRels = new Rels();
  const contentWidthTw = settings.page.width - settings.page.margin.left - settings.page.margin.right - settings.page.margin.gutter;
  docRels.add(REL.styles, "styles.xml");
  docRels.add(REL.settings, "settings.xml");
  docRels.add(REL.numbering, "numbering.xml");

  const body = new PartWriter(ctx, docRels, contentWidthTw);
  let bodyXml = body.blocks(doc.content.content);
  // Kapanmamış açıklama aralıkları
  for (const id of [...ctx.openComments]) {
    const num = ctx.commentNum(id);
    bodyXml += `<w:p><w:commentRangeEnd w:id="${num}"/><w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr><w:commentReference w:id="${num}"/></w:r></w:p>`;
  }
  ctx.openComments.clear();
  const lastIsTable = /<\/w:tbl>$/.test(bodyXml) || /<\/w:sdt>$/.test(bodyXml);
  if (!bodyXml || lastIsTable) bodyXml += "<w:p/>";

  // Üst / alt bilgi
  let headerRid: string | null = null;
  let footerRid: string | null = null;
  const hdrNs = `<w:hdr ${DOC_NS}>`;
  if (hasHfContent(settings.header) || settings.watermark) {
    const rels = new Rels();
    const pw = new PartWriter(ctx, rels, contentWidthTw);
    let inner = hasHfContent(settings.header) ? pw.blocks(settings.header?.content) : "";
    inner = watermarkXml(settings, settings.page.width) + (inner || '<w:p><w:pPr><w:pStyle w:val="Header"/></w:pPr></w:p>');
    zip.file("word/header1.xml", `${XML_HEAD}${hdrNs}${inner}</w:hdr>`);
    if (rels.size) zip.file("word/_rels/header1.xml.rels", rels.xml());
    headerRid = docRels.add(REL.header, "header1.xml");
  }
  if (hasHfContent(settings.footer)) {
    const rels = new Rels();
    const pw = new PartWriter(ctx, rels, contentWidthTw);
    const inner = pw.blocks(settings.footer?.content) || "<w:p/>";
    zip.file("word/footer1.xml", `${XML_HEAD}<w:ftr ${DOC_NS}>${inner}</w:ftr>`);
    if (rels.size) zip.file("word/_rels/footer1.xml.rels", rels.xml());
    footerRid = docRels.add(REL.footer, "footer1.xml");
  }

  const bg = settings.pageColor ? `<w:background w:color="${hex6(settings.pageColor)}"/>` : "";
  zip.file("word/document.xml", `${XML_HEAD}<w:document ${DOC_NS} mc:Ignorable="w10">${bg}<w:body>${bodyXml}${sectPr(settings, headerRid, footerRid)}</w:body></w:document>`);

  const styles = resolveStyles(settings.styleSet, settings.theme, settings.fontTheme, settings.paragraphSpacing);
  zip.file("word/styles.xml", buildStylesXml(styles));
  zip.file("word/numbering.xml", numberingXml(ctx));
  zip.file(
    "word/settings.xml",
    `${XML_HEAD}<w:settings xmlns:w="${NS.w}" xmlns:r="${NS.r}"><w:zoom w:percent="100"/>${settings.pageColor ? "<w:displayBackgroundShape/>" : ""}<w:defaultTabStop w:val="708"/><w:hyphenationZone w:val="425"/><w:characterSpacingControl w:val="doNotCompress"/>${ctx.footnotes.length ? '<w:footnotePr><w:footnote w:id="-1"/><w:footnote w:id="0"/></w:footnotePr>' : ""}${ctx.endnotes.length ? '<w:endnotePr><w:endnote w:id="-1"/><w:endnote w:id="0"/></w:endnotePr>' : ""}<w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat><w:themeFontLang w:val="tr-TR"/></w:settings>`,
  );
  if (ctx.usedComments.size) {
    zip.file("word/comments.xml", commentsXml(ctx, settings.comments));
    docRels.add(REL.comments, "comments.xml");
  }
  if (ctx.footnotes.length) {
    zip.file("word/footnotes.xml", notesXml("footnote", ctx.footnotes));
    docRels.add(REL.footnotes, "footnotes.xml");
  }
  if (ctx.endnotes.length) {
    zip.file("word/endnotes.xml", notesXml("endnote", ctx.endnotes));
    docRels.add(REL.endnotes, "endnotes.xml");
  }
  for (const m of ctx.media) zip.file(`word/media/${m.name}`, m.data);
  zip.file("word/_rels/document.xml.rels", docRels.xml());

  // Paket ilişkileri ve içerik türleri
  zip.file(
    "_rels/.rels",
    `${XML_HEAD}<Relationships xmlns="${NS.rel}"><Relationship Id="rId1" Type="${REL.officeDocument}" Target="word/document.xml"/><Relationship Id="rId2" Type="${REL.core}" Target="docProps/core.xml"/><Relationship Id="rId3" Type="${REL.app}" Target="docProps/app.xml"/><Relationship Id="rId4" Type="${REL.custom}" Target="docProps/custom.xml"/></Relationships>`,
  );
  const exts = new Set(ctx.media.map((m) => m.ext));
  const mime: Record<string, string> = { png: "image/png", jpeg: "image/jpeg", gif: "image/gif", bmp: "image/bmp", svg: "image/svg+xml", webp: "image/webp" };
  const ov = (part: string, type: string) => `<Override PartName="/${part}" ContentType="application/vnd.openxmlformats-officedocument.${type}"/>`;
  zip.file(
    "[Content_Types].xml",
    `${XML_HEAD}<Types xmlns="${NS.ct}"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>${[...exts]
      .map((e) => `<Default Extension="${e}" ContentType="${mime[e] ?? "application/octet-stream"}"/>`)
      .join("")}` +
      ov("word/document.xml", "wordprocessingml.document.main+xml") +
      ov("word/styles.xml", "wordprocessingml.styles+xml") +
      ov("word/settings.xml", "wordprocessingml.settings+xml") +
      ov("word/numbering.xml", "wordprocessingml.numbering+xml") +
      (headerRid ? ov("word/header1.xml", "wordprocessingml.header+xml") : "") +
      (footerRid ? ov("word/footer1.xml", "wordprocessingml.footer+xml") : "") +
      (ctx.usedComments.size ? ov("word/comments.xml", "wordprocessingml.comments+xml") : "") +
      (ctx.footnotes.length ? ov("word/footnotes.xml", "wordprocessingml.footnotes+xml") : "") +
      (ctx.endnotes.length ? ov("word/endnotes.xml", "wordprocessingml.endnotes+xml") : "") +
      `<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>` +
      ov("docProps/app.xml", "extended-properties+xml") +
      ov("docProps/custom.xml", "custom-properties+xml") +
      `</Types>`,
  );
  const pr = settings.properties;
  const now = new Date().toISOString().replace(/\.\d+Z$/, "Z");
  zip.file(
    "docProps/core.xml",
    `${XML_HEAD}<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">` +
      `${pr.title ? `<dc:title>${esc(pr.title)}</dc:title>` : ""}${pr.subject ? `<dc:subject>${esc(pr.subject)}</dc:subject>` : ""}<dc:creator>${esc(pr.author || "OfficeHere")}</dc:creator>${pr.keywords ? `<cp:keywords>${esc(pr.keywords)}</cp:keywords>` : ""}<cp:lastModifiedBy>${esc(pr.author || "OfficeHere")}</cp:lastModifiedBy>` +
      `<dcterms:created xsi:type="dcterms:W3CDTF">${esc((pr.created || now).replace(/\.\d+Z$/, "Z"))}</dcterms:created><dcterms:modified xsi:type="dcterms:W3CDTF">${now}</dcterms:modified></cp:coreProperties>`,
  );
  zip.file(
    "docProps/app.xml",
    `${XML_HEAD}<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"><Application>OfficeHere Belge</Application><DocSecurity>0</DocSecurity><AppVersion>01.0000</AppVersion></Properties>`,
  );
  const custom = { styleSet: settings.styleSet, theme: settings.theme, fontTheme: settings.fontTheme, paragraphSpacing: settings.paragraphSpacing, watermark: settings.watermark };
  zip.file(
    "docProps/custom.xml",
    `${XML_HEAD}<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"><property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="2" name="OfficeHereBelge"><vt:lpwstr>${esc(JSON.stringify(custom))}</vt:lpwstr></property></Properties>`,
  );
  return zip.generateAsync({ type: "uint8array", compression: "DEFLATE", mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document" });
}
