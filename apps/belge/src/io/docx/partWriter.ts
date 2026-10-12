// DOCX parça yazıcısı: ProseMirror JSON blokları → WordprocessingML (paragraflar, koşular, izlenen
// değişiklikler, resimler, metin kutuları, tablolar, İçindekiler). İlişkiler (Rels) ve ortak sayaçlar (Ctx).
import type { JSONContent } from "@tiptap/core";
import type { DocSettings } from "../../model/document";
import { type BorderSpec } from "../../model/styles";
import type { ParaBorders } from "../../editor/paragraphFormat";
import type { TocEntry } from "../../editor/nodes";
import { NS, REL, XML_HEAD, cleanText, esc, hex6, highlightName } from "./xml";
import { borderXml, fontsXml, tableStyleId } from "./styleXml";
import { sigKey, type ListSig } from "./numbering";

export type Mark = { type: string; attrs?: Record<string, unknown> };

export const DOC_NS = `xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas" xmlns:mc="${NS.mc}" xmlns:o="${NS.o}" xmlns:r="${NS.r}" xmlns:v="${NS.v}" xmlns:w10="${NS.w10}" xmlns:wp="${NS.wp}" xmlns:a="${NS.a}" xmlns:pic="${NS.pic}" xmlns:w="${NS.w}" xmlns:w14="${NS.w14}" xmlns:wp14="${NS.wp14}" xmlns:wps="${NS.wps}"`;

const EMU_PER_PX = 9525;

export interface NoteData {
  text: string;
  content: JSONContent[] | null;
}

interface Media {
  name: string;
  data: Uint8Array;
  ext: string;
}

/** Bir parça (belge, üst/alt bilgi) için ilişki tablosu. */
export class Rels {
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
export class Ctx {
  media: Media[] = [];
  numbering = new Map<string, { abstractId: number; sig: ListSig }>();
  nums: { numId: number; abstractId: number; start: number }[] = [];
  bookmarkId = 0;
  drawingId = 1;
  commentIds = new Map<string, number>();
  openComments = new Set<string>();
  usedComments = new Set<string>();
  revId = 1;
  footnotes: NoteData[] = [];
  endnotes: NoteData[] = [];
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
export class PartWriter {
  /** Bölümün son paragrafı ve ona eklenecek bölüm özellikleri (OOXML: sectPr son paragrafın pPr'ındadır). */
  sectTarget: JSONContent | null = null;
  sectXml = "";
  constructor(
    private ctx: Ctx,
    private rels: Rels,
    public contentWidthTw: number,
  ) {}

  /** Bir sonraki paragrafın ilk koşusuna yazılacak bağlı nesneler (metin kutuları). */
  private pendingDrawings = "";

  blocks(nodes: JSONContent[] | undefined, list?: { numId: number; ilvl: number }): string {
    let out = "";
    const arr = nodes ?? [];
    arr.forEach((n, i) => {
      if (n.type === "textBox") {
        // Metin kutusu sonraki paragrafa bağlanır (OOXML: çapa bir koşu içindedir); yoksa kendi paragrafına.
        this.pendingDrawings += `<w:r>${this.textBox(n)}</w:r>`;
        const next = arr[i + 1];
        if (!next || (next.type !== "paragraph" && next.type !== "heading")) {
          out += `<w:p>${this.pendingDrawings}</w:p>`;
          this.pendingDrawings = "";
        }
        return;
      }
      out += this.block(n, list);
    });
    return out;
  }

  /** Metin kutusu: wps şekli (Word 2010+) ve VML yedeği. */
  textBox(n: JSONContent): string {
    const a = (n.attrs ?? {}) as Record<string, unknown>;
    const w = Number(a.w) || 240;
    const h = Number(a.h) || 120;
    const x = Number(a.x) || 0;
    const y = Number(a.y) || 0;
    const emu = (px: number) => Math.round(px * EMU_PER_PX);
    const pt = (px: number) => +(px * 0.75).toFixed(2);
    const border = hex6(a.border as string | null | undefined);
    const fill = hex6(a.fill as string | null | undefined);
    const bw = Number(a.borderWidth) || 0.75;
    const id = this.ctx.drawingId++;
    const inner = new PartWriter(this.ctx, this.rels, Math.round((w - 19) * 15)).blocks(n.content) || "<w:p/>";
    const ln = border ? `<a:ln w="${Math.round(bw * 12700)}"><a:solidFill><a:srgbClr val="${border}"/></a:solidFill></a:ln>` : "<a:ln><a:noFill/></a:ln>";
    const fl = fill ? `<a:solidFill><a:srgbClr val="${fill}"/></a:solidFill>` : "<a:noFill/>";
    const choice =
      `<w:drawing><wp:anchor distT="0" distB="0" distL="114300" distR="114300" simplePos="0" relativeHeight="${251659264 + id}" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">` +
      `<wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="column"><wp:posOffset>${emu(x)}</wp:posOffset></wp:positionH><wp:positionV relativeFrom="paragraph"><wp:posOffset>${emu(y)}</wp:posOffset></wp:positionV>` +
      `<wp:extent cx="${emu(w)}" cy="${emu(h)}"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:wrapNone/><wp:docPr id="${id}" name="Metin Kutusu ${id}"/><wp:cNvGraphicFramePr/>` +
      `<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp><wps:cNvSpPr txBox="1"/><wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${emu(w)}" cy="${emu(h)}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom>${fl}${ln}</wps:spPr>` +
      `<wps:txbx><w:txbxContent>${inner}</w:txbxContent></wps:txbx><wps:bodyPr rot="0" vert="horz" wrap="square" lIns="91440" tIns="45720" rIns="91440" bIns="45720" anchor="t" anchorCtr="0"><a:noAutofit/></wps:bodyPr></wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing>`;
    const style = `position:absolute;margin-left:${pt(x)}pt;margin-top:${pt(y)}pt;width:${pt(w)}pt;height:${pt(h)}pt;z-index:${id};mso-position-horizontal-relative:text;mso-position-vertical-relative:text`;
    const vml =
      `<w:pict><v:shapetype id="_x0000_t202" coordsize="21600,21600" o:spt="202" path="m,l,21600r21600,l21600,xe"><v:stroke joinstyle="miter"/><v:path gradientshapeok="t" o:connecttype="rect"/></v:shapetype>` +
      `<v:shape id="Metin Kutusu ${id}" o:spid="_x0000_s${1024 + id}" type="#_x0000_t202" style="${style}"${fill ? ` fillcolor="#${fill}"` : ' filled="f"'}${border ? ` strokecolor="#${border}" strokeweight="${bw}pt"` : ' stroked="f"'}><v:textbox><w:txbxContent>${inner}</w:txbxContent></v:textbox></v:shape></w:pict>`;
    return `<mc:AlternateContent><mc:Choice Requires="wps">${choice}</mc:Choice><mc:Fallback>${vml}</mc:Fallback></mc:AlternateContent>`;
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
      case "columnBreak":
        return '<w:p><w:r><w:br w:type="column"/></w:r></w:p>';
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
    if (n === this.sectTarget) x += this.sectXml;
    return x ? `<w:pPr>${x}</w:pPr>` : "";
  }

  paragraph(n: JSONContent, list?: { numId: number; ilvl: number }, contIndent?: number): string {
    const drawings = this.pendingDrawings;
    this.pendingDrawings = "";
    return `<w:p>${this.pPr(n, list, contIndent)}${drawings}${this.inline(n.content ?? [])}</w:p>`;
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
    // İzlenen biçim değişikliği: önceki biçim w:rPrChange içinde (rPr'nin son öğesi).
    const fc = has("formatChange");
    if (fc) {
      let old: Mark[] = [];
      try {
        old = JSON.parse(String(fc.attrs?.old ?? "[]")) as Mark[];
      } catch {
        old = [];
      }
      const inner = this.rPr(old, isLink) || "<w:rPr/>";
      x += `<w:rPrChange w:id="${this.ctx.revId++}" w:author="${esc(String(fc.attrs?.author || "Kullanıcı"))}" w:date="${esc(String(fc.attrs?.date || new Date().toISOString().replace(/\.\d+Z$/, "Z")))}">${inner}</w:rPrChange>`;
    }
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
        list.push({ text: String(node.attrs?.text ?? ""), content: (node.attrs?.content as JSONContent[] | null) ?? null });
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
