// DOCX okuyucu: OOXML (JSZip + DOMParser) → ProseMirror JSON + belge ayarları.
import JSZip from "jszip";
import type { JSONContent } from "@tiptap/core";
import { defaultSettings, type BelgeDocument, type CommentData, type DocSettings } from "../../model/document";
import { effectiveProps, resolveStyles, type StyleDef } from "../../model/styles";
import type { TocEntry } from "../../editor/nodes";
import { attr, attrNum, descendants, kid, kids, parseXml, path, relsTarget, NS } from "./xml";
import { Numbering, StyleSheet, parsePPr, parseRPr, parseTheme, stripUndef, type ParaProps, type RunProps, type ThemeFonts } from "./readerStyles";
import { tableStyleFromId } from "./styleXml";
import type { ListSig } from "./numbering";

type Mark = { type: string; attrs?: Record<string, unknown> };

interface Rel {
  type: string;
  target: string;
  external: boolean;
}

const KNOWN_BY_NAME: Record<string, string> = {
  normal: "Normal",
  "no spacing": "NoSpacing",
  title: "Title",
  subtitle: "Subtitle",
  quote: "Quote",
  "intense quote": "IntenseQuote",
  "list paragraph": "ListParagraph",
  "subtle emphasis": "SubtleEmphasis",
  emphasis: "Emphasis",
  "intense emphasis": "IntenseEmphasis",
  strong: "Strong",
  "subtle reference": "SubtleReference",
  "intense reference": "IntenseReference",
  "book title": "BookTitle",
};

const MIME: Record<string, string> = { png: "image/png", jpg: "image/jpeg", jpeg: "image/jpeg", gif: "image/gif", bmp: "image/bmp", svg: "image/svg+xml", webp: "image/webp", tif: "image/tiff", tiff: "image/tiff", emf: "image/emf", wmf: "image/wmf" };

function toBase64(bytes: Uint8Array): string {
  let bin = "";
  for (let i = 0; i < bytes.length; i += 0x8000) bin += String.fromCharCode(...bytes.subarray(i, i + 0x8000));
  return btoa(bin);
}

async function readRels(zip: JSZip, part: string): Promise<Map<string, Rel>> {
  const map = new Map<string, Rel>();
  const dir = part.includes("/") ? part.slice(0, part.lastIndexOf("/")) : "";
  const relPath = `${dir ? dir + "/" : ""}_rels/${part.slice(part.lastIndexOf("/") + 1)}.rels`;
  const f = zip.file(relPath);
  if (!f) return map;
  const doc = parseXml(await f.async("string"));
  for (const r of Array.from(doc.getElementsByTagName("*")).filter((e) => e.localName === "Relationship")) {
    const external = r.getAttribute("TargetMode") === "External";
    const target = r.getAttribute("Target") ?? "";
    map.set(r.getAttribute("Id") ?? "", { type: r.getAttribute("Type") ?? "", target: external ? target : relsTarget(part, target), external });
  }
  return map;
}

interface FieldState {
  instr: string;
  phase: "instr" | "result";
  kind: "link" | "page" | "numpages" | "other";
  href?: string;
  emitted?: boolean;
}

/** Bir parçanın (gövde / üst bilgi / alt bilgi) dönüştürücüsü. */
class PartReader {
  private fields: FieldState[] = [];
  private openComments = new Set<string>();
  private listCounts = new Map<string, number>();
  watermark: DocSettings["watermark"] = null;

  constructor(
    private r: DocxReader,
    private rels: Map<string, Rel>,
  ) {}

  blocks(container: Element): JSONContent[] {
    const items: { node: JSONContent; list?: { numId: string; ilvl: number; sig: ListSig; start: number } }[] = [];
    const visit = (el: Element) => {
      for (const c of kids(el)) {
        switch (c.localName) {
          case "p":
            for (const it of this.paragraph(c)) items.push(it);
            break;
          case "tbl":
            items.push({ node: this.table(c) });
            break;
          case "sdt": {
            const gallery = attr(descendants(kid(c, "sdtPr") ?? c, "docPartGallery")[0], "val");
            const content = kid(c, "sdtContent");
            if (gallery === "Table of Contents" && content) items.push({ node: this.toc(content) });
            else if (content) visit(content);
            break;
          }
          case "customXml":
          case "ins":
          case "smartTag":
            visit(c);
            break;
          case "AlternateContent": {
            const ch = kid(c, "Choice") ?? kid(c, "Fallback");
            if (ch) visit(ch);
            break;
          }
        }
      }
    };
    visit(container);
    return this.foldLists(items);
  }

  private foldLists(items: { node: JSONContent; list?: { numId: string; ilvl: number; sig: ListSig; start: number } }[]): JSONContent[] {
    const out: JSONContent[] = [];
    type Frame = { numId: string; level: number; node: JSONContent; last: JSONContent | null };
    let stack: Frame[] = [];
    for (const it of items) {
      if (!it.list) {
        stack = [];
        out.push(it.node);
        continue;
      }
      const { numId, ilvl, sig, start } = it.list;
      while (stack.length && (stack[stack.length - 1].level > ilvl || (stack[stack.length - 1].level === ilvl && stack[stack.length - 1].numId !== numId))) stack.pop();
      const top = stack[stack.length - 1];
      if (!top || top.level < ilvl) {
        const key = `${numId}:${ilvl}`;
        const seen = this.listCounts.get(key) ?? 0;
        const attrs: Record<string, unknown> = { listStyle: sig.style, multilevel: sig.multilevel };
        if (sig.type === "ordered") attrs.start = seen ? start + seen : start;
        const listNode: JSONContent = { type: sig.type === "bullet" ? "bulletList" : "orderedList", attrs, content: [] };
        if (!top || !top.last) out.push(listNode);
        else top.last.content!.push(listNode);
        stack.push({ numId, level: ilvl, node: listNode, last: null });
      }
      const frame = stack[stack.length - 1];
      const li: JSONContent = { type: "listItem", content: [it.node] };
      frame.node.content!.push(li);
      frame.last = li;
      const key = `${numId}:${ilvl}`;
      this.listCounts.set(key, (this.listCounts.get(key) ?? 0) + 1);
    }
    return out;
  }

  private paragraph(p: Element): { node: JSONContent; list?: { numId: string; ilvl: number; sig: ListSig; start: number } }[] {
    const pPr = kid(p, "pPr");
    const styleId = attr(kid(pPr, "pStyle"), "val");
    const direct = parsePPr(pPr);
    const mapped = this.r.mapParagraphStyle(styleId);
    const styleDiff = this.r.styleDiff(styleId, mapped.ourId);
    // Liste: doğrudan numPr ya da stil üzerinden.
    const stylePP = this.r.sheet.effective(styleId).p;
    const numId = direct.numId ?? stylePP.numId;
    const ilvl = direct.ilvl ?? stylePP.ilvl ?? 0;
    const list = numId ? this.r.numbering.resolve(numId, ilvl) : null;
    const pp: ParaProps = { ...styleDiff.p, ...stripUndef(direct) };
    const attrs: Record<string, unknown> = {};
    if (mapped.type === "heading") attrs.level = mapped.level;
    else attrs.styleId = list && mapped.ourId === "ListParagraph" ? null : mapped.ourId === "Normal" ? null : mapped.ourId;
    if (pp.align) attrs.textAlign = pp.align;
    if (!list) {
      if (pp.indentLeft) attrs.indentLeft = pp.indentLeft;
      if (pp.indentFirst) attrs.indentFirst = pp.indentFirst;
    }
    if (pp.indentRight) attrs.indentRight = pp.indentRight;
    if (pp.spaceBefore !== undefined) attrs.spaceBefore = pp.spaceBefore;
    if (pp.spaceAfter !== undefined) attrs.spaceAfter = pp.spaceAfter;
    if (pp.lineHeight !== undefined) {
      attrs.lineHeight = pp.lineHeight;
      attrs.lineRule = pp.lineRule && pp.lineRule !== "auto" ? pp.lineRule : null;
    }
    if (pp.shading) attrs.shading = pp.shading;
    if (pp.borders) attrs.borders = pp.borders;
    if (direct.keepNext) attrs.keepNext = true;
    if (direct.keepLines) attrs.keepLines = true;
    if (direct.pageBreakBefore) attrs.pageBreakBefore = true;

    const extra: JSONContent[] = [];
    const segments: JSONContent[][] = [[]];
    const wmBefore = this.watermark;
    this.inline(p, styleDiff.r, [], segments, extra);
    // Yalnız filigran taşıyan üst bilgi paragrafı atlanır.
    if (this.watermark !== wmBefore && segments.every((x) => x.length === 0)) return [];
    const type = mapped.type === "heading" ? "heading" : "paragraph";
    const out: { node: JSONContent; list?: { numId: string; ilvl: number; sig: ListSig; start: number } }[] = [];
    // Yatay çizgi: metni olmayan, yalnız alt kenarlıklı paragraf.
    const onlyBottom = pp.borders && pp.borders.bottom && !pp.borders.top && !pp.borders.left && !pp.borders.right;
    if (onlyBottom && segments.length === 1 && segments[0].length === 0 && mapped.ourId === "Normal" && !list) {
      return [{ node: { type: "horizontalRule" } }];
    }
    segments.forEach((content, i) => {
      if (i > 0) out.push({ node: { type: "pageBreak" } });
      const empty = content.length === 0;
      // Sayfa sonunun önündeki/arkasındaki boş parçalar paragraf oluşturmaz.
      if (empty && segments.length > 1 && (i === 0 || i === segments.length - 1)) return;
      const node: JSONContent = { type, attrs: { ...attrs }, ...(empty ? {} : { content: mergeText(content) }) };
      out.push({ node, list: i === 0 && list && numId ? { numId, ilvl, sig: list.sig, start: list.start } : undefined });
    });
    for (const e of extra) out.push({ node: e });
    return out;
  }

  /** Satır içi öğeleri dolaşır; sayfa sonu yeni bir parça başlatır. */
  private inline(el: Element, base: RunProps, marks: Mark[], segs: JSONContent[][], extra: JSONContent[]) {
    const push = (n: JSONContent) => segs[segs.length - 1].push(n);
    for (const c of kids(el)) {
      switch (c.localName) {
        case "r":
          this.run(c, base, marks, segs, extra);
          break;
        case "hyperlink": {
          const rid = c.getAttributeNS(NS.r, "id") ?? attr(c, "id");
          const anchor = attr(c, "anchor");
          let href: string | null = null;
          if (rid && this.rels.get(rid)) href = this.rels.get(rid)!.target;
          if (anchor) href = (href ?? "") + `#${anchor}`;
          this.inline(c, base, href ? [...marks.filter((m) => m.type !== "link"), { type: "link", attrs: { href } }] : marks, segs, extra);
          break;
        }
        case "ins":
        case "moveTo":
          this.inline(c, base, [...marks, { type: "insertion", attrs: { author: attr(c, "author") ?? "", date: attr(c, "date") ?? "" } }], segs, extra);
          break;
        case "del":
        case "moveFrom":
          this.inline(c, base, [...marks, { type: "deletion", attrs: { author: attr(c, "author") ?? "", date: attr(c, "date") ?? "" } }], segs, extra);
          break;
        case "fldSimple": {
          const instr = (attr(c, "instr") ?? "").trim();
          const kw = instr.split(/\s+/)[0]?.toUpperCase();
          if (kw === "PAGE" || kw === "NUMPAGES") push({ type: "pageField", attrs: { field: kw } });
          else if (kw === "HYPERLINK") {
            const href = parseHyperlinkInstr(instr);
            this.inline(c, base, href ? [...marks, { type: "link", attrs: { href } }] : marks, segs, extra);
          } else this.inline(c, base, marks, segs, extra);
          break;
        }
        case "bookmarkStart": {
          const name = attr(c, "name");
          if (name && name !== "_GoBack") push({ type: "bookmark", attrs: { name } });
          break;
        }
        case "commentRangeStart":
          this.openComments.add(`c${attr(c, "id")}`);
          break;
        case "commentRangeEnd":
          this.openComments.delete(`c${attr(c, "id")}`);
          break;
        case "smartTag":
        case "customXml":
          this.inline(c, base, marks, segs, extra);
          break;
        case "sdt": {
          const content = kid(c, "sdtContent");
          if (content) this.inline(content, base, marks, segs, extra);
          break;
        }
        case "AlternateContent": {
          const ch = kid(c, "Choice") ?? kid(c, "Fallback");
          if (ch) this.inline(ch, base, marks, segs, extra);
          break;
        }
      }
    }
  }

  private run(r: Element, base: RunProps, outer: Mark[], segs: JSONContent[][], extra: JSONContent[]) {
    const rPr = kid(r, "rPr");
    const rStyle = attr(kid(rPr, "rStyle"), "val");
    const charStyle = this.r.mapCharStyle(rStyle);
    const props: RunProps = { ...base, ...(charStyle ? {} : stripUndef(this.r.sheet.charProps(rStyle))), ...stripUndef(parseRPr(rPr, this.r.theme)) };
    if (rStyle && /hyperlink/i.test(rStyle)) {
      delete props.color;
      delete props.underline;
    }
    const field = this.fields[this.fields.length - 1];
    let marks = [...outer, ...runMarks(props)];
    if (charStyle) marks.push({ type: "charStyle", attrs: { styleId: charStyle } });
    for (const id of this.openComments) marks.push({ type: "comment", attrs: { id } });
    if (field?.phase === "result" && field.kind === "link" && field.href) marks = [...marks.filter((m) => m.type !== "link"), { type: "link", attrs: { href: field.href } }];
    const push = (n: JSONContent) => segs[segs.length - 1].push(marks.length && n.type !== "pageBreak" ? { ...n, marks } : n);
    for (const c of kids(r)) {
      const f = this.fields[this.fields.length - 1];
      const suppress = f && (f.phase === "instr" || f.kind === "page" || f.kind === "numpages");
      switch (c.localName) {
        case "fldChar": {
          const t = attr(c, "fldCharType");
          if (t === "begin") this.fields.push({ instr: "", phase: "instr", kind: "other" });
          else if (t === "separate" && f) {
            f.phase = "result";
            const kw = f.instr.trim().split(/\s+/)[0]?.toUpperCase();
            if (kw === "HYPERLINK") {
              f.kind = "link";
              f.href = parseHyperlinkInstr(f.instr) ?? undefined;
              if (f.href) marks = [...marks.filter((m) => m.type !== "link"), { type: "link", attrs: { href: f.href } }];
            } else if (kw === "PAGE" || kw === "NUMPAGES") {
              f.kind = kw === "PAGE" ? "page" : "numpages";
              push({ type: "pageField", attrs: { field: kw } });
              f.emitted = true;
            }
          } else if (t === "end" && f) {
            const kw = f.instr.trim().split(/\s+/)[0]?.toUpperCase();
            if (!f.emitted && f.phase === "instr" && (kw === "PAGE" || kw === "NUMPAGES")) push({ type: "pageField", attrs: { field: kw } });
            this.fields.pop();
            if (f.kind === "link") marks = marks.filter((m) => m.type !== "link");
          }
          break;
        }
        case "instrText":
          if (f) f.instr += c.textContent ?? "";
          break;
        case "t":
        case "delText":
          if (!suppress && c.textContent) push({ type: "text", text: c.textContent });
          break;
        case "tab":
        case "ptab":
          if (!suppress) push({ type: "text", text: "\t" });
          break;
        case "br":
          if (suppress) break;
          if (attr(c, "type") === "page") segs.push([]);
          else push({ type: "hardBreak" });
          break;
        case "cr":
          if (!suppress) push({ type: "hardBreak" });
          break;
        case "noBreakHyphen":
          if (!suppress) push({ type: "text", text: "‑" });
          break;
        case "softHyphen":
          if (!suppress) push({ type: "text", text: "­" });
          break;
        case "sym": {
          const code = parseInt(attr(c, "char") ?? "", 16);
          if (!suppress && code) push({ type: "text", text: String.fromCodePoint(code >= 0xf000 ? code - 0xf000 : code) });
          break;
        }
        case "footnoteReference":
        case "endnoteReference": {
          if (suppress) break;
          const kind = c.localName === "footnoteReference" ? "footnote" : "endnote";
          const nid = attr(c, "id") ?? "";
          const text = (kind === "footnote" ? this.r.footnotes : this.r.endnotes).get(nid) ?? "";
          segs[segs.length - 1].push({ type: "note", attrs: { kind, text, id: `n${kind[0]}${nid}` } });
          break;
        }
        case "drawing":
          if (!suppress) this.drawing(c, push, extra);
          break;
        case "pict":
          if (!suppress) this.pict(c, push, extra);
          break;
        case "AlternateContent": {
          const ch = kid(c, "Choice") ?? kid(c, "Fallback");
          const dr = ch ? descendants(ch, "drawing")[0] : null;
          if (dr) this.drawing(dr, push, extra);
          else if (ch) {
            const pc = descendants(ch, "pict")[0];
            if (pc) this.pict(pc, push, extra);
          }
          break;
        }
      }
    }
  }

  private drawing(d: Element, push: (n: JSONContent) => void, extra: JSONContent[]) {
    const blip = descendants(d, "blip")[0];
    const txbx = descendants(d, "txbxContent")[0];
    if (txbx && !blip) {
      for (const b of this.blocks(txbx)) extra.push(b);
      return;
    }
    if (!blip) return;
    const rid = blip.getAttributeNS(NS.r, "embed") ?? attr(blip, "embed");
    const rel = rid ? this.rels.get(rid) : undefined;
    const src = rel ? this.r.media.get(rel.target) : undefined;
    if (!src) return;
    const ext = descendants(d, "extent")[0];
    const cx = attrNum(ext, "cx") ?? 1905000;
    const cy = attrNum(ext, "cy") ?? 1428750;
    const docPr = descendants(d, "docPr")[0];
    const xfrm = descendants(d, "xfrm")[0];
    const rot = Math.round((attrNum(xfrm, "rot") ?? 0) / 60000);
    push({
      type: "image",
      attrs: {
        src,
        alt: docPr?.getAttribute("descr") || null,
        width: Math.round(cx / 9525),
        height: Math.round(cy / 9525),
        rotate: rot || 0,
        flipH: xfrm?.getAttribute("flipH") === "1",
        flipV: xfrm?.getAttribute("flipV") === "1",
      },
    });
  }

  private pict(pc: Element, push: (n: JSONContent) => void, extra: JSONContent[]) {
    const tp = descendants(pc, "textpath").find((x) => x.hasAttribute("string"));
    if (tp) {
      const shape = tp.parentElement;
      const style = shape?.getAttribute("style") ?? "";
      const fill = shape?.getAttribute("fillcolor") ?? "#C0C0C0";
      const fontM = /font-family:\s*&?quot;?"?([^";&]+)/.exec(tp.getAttribute("style") ?? "");
      const opacity = descendants(shape ?? pc, "fill")[0]?.getAttribute("opacity");
      this.watermark = {
        text: tp.getAttribute("string") ?? "",
        color: fill.startsWith("#") ? fill.toUpperCase() : fill === "silver" ? "#C0C0C0" : fill === "gray" ? "#808080" : "#C0C0C0",
        layout: /rotation:\s*(315|-45)/.test(style) ? "diagonal" : "horizontal",
        font: fontM?.[1]?.trim() || "Aptos",
        semitransparent: !!opacity && parseFloat(opacity) < 1,
      };
      return;
    }
    const img = descendants(pc, "imagedata")[0];
    if (img) {
      const rid = img.getAttributeNS(NS.r, "id") ?? attr(img, "id");
      const rel = rid ? this.rels.get(rid) : undefined;
      const src = rel ? this.r.media.get(rel.target) : undefined;
      if (!src) return;
      const style = img.parentElement?.getAttribute("style") ?? "";
      const w = /width:\s*([\d.]+)pt/.exec(style);
      const h = /height:\s*([\d.]+)pt/.exec(style);
      push({ type: "image", attrs: { src, width: w ? Math.round(parseFloat(w[1]) * 1.333) : 200, height: h ? Math.round(parseFloat(h[1]) * 1.333) : 150 } });
      return;
    }
    const txbx = descendants(pc, "txbxContent")[0];
    if (txbx) for (const b of this.blocks(txbx)) extra.push(b);
  }

  private table(tbl: Element): JSONContent {
    const tblPr = kid(tbl, "tblPr");
    const styleId = attr(kid(tblPr, "tblStyle"), "val");
    const grid = kids(kid(tbl, "tblGrid"), "gridCol").map((g) => (attrNum(g, "w") ?? 0) / 15);
    const look = kid(tblPr, "tblLook");
    const lookVal = parseInt(attr(look, "val") ?? "0", 16) || 0;
    const flag = (name: string, bit: number) => {
      const v = attr(look, name);
      return v !== null ? v === "1" || v === "true" : (lookVal & bit) !== 0;
    };
    const ownBorders = kid(tblPr, "tblBorders");
    const bordersNone = ownBorders ? kids(ownBorders).every((b) => ["none", "nil"].includes(attr(b, "val") ?? "none")) : this.r.sheet.styles.get(styleId ?? "")?.tblBordersNone ?? !styleId;
    let tableStyle = tableStyleFromId(styleId);
    if (!tableStyle || (tableStyle === "none" && !bordersNone)) tableStyle = bordersNone ? "none" : "grid";
    const jc = attr(kid(tblPr, "jc"), "val");
    const tblW = kid(tblPr, "tblW");
    const fitWindow = attr(tblW, "type") === "pct" && (attrNum(tblW, "w") ?? 0) >= 4900;
    // Satırları ve hücre ızgara konumlarını topla.
    type Cell = { el: Element; col: number; span: number; vm: "restart" | "continue" | null };
    const rows: Cell[][] = kids(tbl, "tr").map((tr) => {
      let col = 0;
      const cells: Cell[] = [];
      const gridBefore = attrNum(path(tr, "trPr", "gridBefore"), "val") ?? 0;
      col += gridBefore;
      for (const tc of kids(tr, "tc")) {
        const tcPr = kid(tc, "tcPr");
        const span = attrNum(kid(tcPr, "gridSpan"), "val") ?? 1;
        const vmEl = kid(tcPr, "vMerge");
        const vm = vmEl ? (attr(vmEl, "val") === "restart" ? "restart" : "continue") : null;
        cells.push({ el: tc, col, span, vm });
        col += span;
      }
      return cells;
    });
    const content: JSONContent[] = rows.map((cells, ri) => ({
      type: "tableRow",
      content: cells
        .filter((c) => c.vm !== "continue")
        .map((c) => {
          let rowspan = 1;
          if (c.vm === "restart") {
            for (let k = ri + 1; k < rows.length; k++) {
              if (rows[k].some((x) => x.col === c.col && x.vm === "continue")) rowspan++;
              else break;
            }
          }
          const tcPr = kid(c.el, "tcPr");
          const bg = attr(kid(tcPr, "shd"), "fill");
          const va = attr(kid(tcPr, "vAlign"), "val");
          const widths = grid.length >= c.col + c.span ? grid.slice(c.col, c.col + c.span).map((w) => Math.round(w)) : null;
          let blocks = this.blocks(c.el);
          if (!blocks.length) blocks = [{ type: "paragraph" }];
          return {
            type: "tableCell",
            attrs: {
              colspan: c.span,
              rowspan,
              colwidth: !fitWindow && widths && widths.every((w) => w > 0) ? widths : null,
              background: bg && /^[0-9a-f]{6}$/i.test(bg) ? `#${bg.toUpperCase()}` : null,
              verticalAlign: va === "center" ? "middle" : va === "bottom" ? "bottom" : null,
            },
            content: blocks,
          };
        }),
    }));
    return {
      type: "table",
      attrs: { tableStyle, headerRow: flag("firstRow", 0x20), bandedRows: !flag("noHBand", 0x200), firstColumn: flag("firstColumn", 0x80), totalRow: flag("lastRow", 0x40), align: jc === "center" || jc === "right" ? jc : jc === "end" ? "right" : null },
      content: content.filter((r) => (r.content?.length ?? 0) > 0),
    };
  }

  private toc(content: Element): JSONContent {
    const entries: TocEntry[] = [];
    let title = "İçindekiler";
    for (const p of descendants(content, "p")) {
      const sid = attr(path(p, "pPr", "pStyle"), "val") ?? "";
      const info = this.r.sheet.styles.get(sid);
      const name = (info?.name ?? sid).toLowerCase();
      const text = descendants(p, "t")
        .map((t) => t.textContent ?? "")
        .join("");
      const lm = /^toc\s*(\d)$/.exec(name) ?? /^TOC(\d)$/i.exec(sid);
      if (lm) {
        // Metin ve sayfa numarası son sekmeyle ayrılır.
        const parts: string[] = [""];
        for (const r of descendants(p, "r")) {
          for (const c of kids(r)) {
            if (c.localName === "tab") parts.push("");
            else if (c.localName === "t") parts[parts.length - 1] += c.textContent ?? "";
          }
        }
        const page = parseInt(parts.length > 1 ? parts[parts.length - 1] : "", 10);
        const label = (parts.length > 1 ? parts.slice(0, -1).join(" ") : parts[0]).trim();
        if (label) entries.push({ level: Number(lm[1]), text: label, page: Number.isFinite(page) ? page : 0, pos: 0 });
      } else if (text.trim() && !entries.length && /toc heading|tocheading/i.test(name + sid)) {
        title = text.trim();
      } else if (text.trim() && !entries.length) title = text.trim();
    }
    return { type: "toc", attrs: { entries, title } };
  }
}

function parseHyperlinkInstr(instr: string): string | null {
  const anchor = /\\l\s+"([^"]+)"/.exec(instr);
  const url = /HYPERLINK\s+"([^"]+)"/i.exec(instr);
  if (url) return url[1] + (anchor ? `#${anchor[1]}` : "");
  if (anchor) return `#${anchor[1]}`;
  return null;
}

/** Koşu özelliklerinden işaretler. */
export function runMarks(p: RunProps): Mark[] {
  const m: Mark[] = [];
  if (p.bold) m.push({ type: "bold" });
  if (p.italic) m.push({ type: "italic" });
  if (p.underline) m.push({ type: "underline", attrs: { type: p.underline.type, color: p.underline.color } });
  if (p.strike) m.push({ type: "strike" });
  if (p.vertAlign === "superscript") m.push({ type: "superscript" });
  if (p.vertAlign === "subscript") m.push({ type: "subscript" });
  if (p.highlight) m.push({ type: "highlight", attrs: { color: p.highlight } });
  const ts: Record<string, unknown> = {};
  if (p.font) ts.fontFamily = p.font;
  if (p.size) ts.fontSize = `${p.size}pt`;
  if (p.color) ts.color = p.color;
  if (p.shd) ts.backgroundColor = p.shd;
  if (p.caps) ts.caps = true;
  if (p.smallCaps) ts.smallCaps = true;
  if (p.vanish) ts.hidden = true;
  if (p.dstrike) ts.doubleStrike = true;
  if (p.spacing) ts.letterSpacing = p.spacing;
  if (Object.keys(ts).length) m.push({ type: "textStyle", attrs: ts });
  return m;
}

/** Aynı işaretli bitişik metin düğümlerini birleştirir. */
function mergeText(nodes: JSONContent[]): JSONContent[] {
  const out: JSONContent[] = [];
  for (const n of nodes) {
    const last = out[out.length - 1];
    if (n.type === "text" && last?.type === "text" && JSON.stringify(last.marks ?? []) === JSON.stringify(n.marks ?? [])) last.text = (last.text ?? "") + (n.text ?? "");
    else if (n.type !== "text" || n.text) out.push({ ...n });
  }
  return out;
}

export class DocxReader {
  sheet!: StyleSheet;
  numbering!: Numbering;
  theme: ThemeFonts = { major: "Aptos Display", minor: "Aptos" };
  media = new Map<string, string>();
  footnotes = new Map<string, string>();
  endnotes = new Map<string, string>();
  ours: Record<string, StyleDef> = resolveStyles("default", "OfficeHere");

  mapParagraphStyle(id: string | null): { type: "heading" | "paragraph"; level?: number; ourId: string } {
    const info = id ? this.sheet.styles.get(id) : undefined;
    const name = (info?.name ?? id ?? "").toLowerCase().trim();
    const hm = /^heading\s*([1-9])$/.exec(name) ?? (id ? /^Heading([1-9])$/.exec(id) : null);
    if (hm) {
      const level = Math.min(6, Number(hm[1]));
      return { type: "heading", level, ourId: `Heading${level}` };
    }
    const known = KNOWN_BY_NAME[name] ?? (id && this.ours[id] && this.ours[id].type === "paragraph" ? id : undefined);
    if (known) return { type: "paragraph", ourId: known };
    if (!id || id === this.sheet.defaultParagraph) return { type: "paragraph", ourId: "Normal" };
    // Bilinmeyen stil: temel stili zincirde bilinen bir stile dayanıyorsa onu kullan.
    for (const s of [...this.sheet.chain(id)].reverse()) {
      const n = s.name.toLowerCase();
      const h = /^heading\s*([1-9])$/.exec(n);
      if (h) return { type: "heading", level: Math.min(6, Number(h[1])), ourId: `Heading${Math.min(6, Number(h[1]))}` };
      if (KNOWN_BY_NAME[n] && KNOWN_BY_NAME[n] !== "Normal") return { type: "paragraph", ourId: KNOWN_BY_NAME[n] };
    }
    return { type: "paragraph", ourId: "Normal" };
  }

  mapCharStyle(id: string | null): string | null {
    if (!id) return null;
    const info = this.sheet.styles.get(id);
    const name = (info?.name ?? id).toLowerCase();
    const k = KNOWN_BY_NAME[name];
    return k && this.ours[k]?.type === "character" ? k : null;
  }

  private diffCache = new Map<string, { r: RunProps; p: ParaProps }>();

  /** Kaynak stilin etkin özellikleriyle bizim karşılık stilimiz arasındaki fark (doğrudan biçim olarak uygulanır). */
  styleDiff(sourceId: string | null, ourId: string): { r: RunProps; p: ParaProps } {
    const key = `${sourceId}|${ourId}`;
    const hit = this.diffCache.get(key);
    if (hit) return hit;
    const src = this.sheet.effective(sourceId);
    const our = effectiveProps(this.ours, ourId);
    const r: RunProps = {};
    const p: ParaProps = {};
    const norm = (f?: string) => (f ?? "").toLowerCase().replace(/\s+/g, " ").trim();
    if (src.r.font && norm(src.r.font) !== norm(our.font)) r.font = src.r.font;
    if (src.r.size && Math.abs(src.r.size - (our.size ?? 11)) > 0.01) r.size = src.r.size;
    if (src.r.color !== undefined && (src.r.color ?? null) !== (our.color ? our.color.toUpperCase() : null) && src.r.color) r.color = src.r.color;
    if (!!src.r.bold !== !!our.bold && src.r.bold) r.bold = true;
    if (!!src.r.italic !== !!our.italic && src.r.italic) r.italic = true;
    if (src.r.caps && !our.caps) r.caps = true;
    if (src.p.spaceBefore !== undefined && Math.abs(src.p.spaceBefore - (our.spaceBefore ?? 0)) > 0.01) p.spaceBefore = src.p.spaceBefore;
    if (src.p.spaceAfter !== undefined && Math.abs(src.p.spaceAfter - (our.spaceAfter ?? 0)) > 0.01) p.spaceAfter = src.p.spaceAfter;
    if (src.p.lineHeight !== undefined) {
      const ourLine = our.line ?? 1;
      if ((src.p.lineRule ?? "auto") !== "auto" || Math.abs(src.p.lineHeight - ourLine) > 0.015) {
        p.lineHeight = src.p.lineHeight;
        p.lineRule = src.p.lineRule;
      }
    }
    if (src.p.align && src.p.align !== (our.align ?? "left")) p.align = src.p.align;
    if (src.p.indentLeft !== undefined && Math.abs(src.p.indentLeft - (our.indentLeft ?? 0)) > 0.01) p.indentLeft = src.p.indentLeft;
    if (src.p.indentFirst) p.indentFirst = src.p.indentFirst;
    if (src.p.shading && src.p.shading !== (our.shading ?? null)) p.shading = src.p.shading;
    const out = { r, p };
    this.diffCache.set(key, out);
    return out;
  }

  async read(data: ArrayBuffer | Uint8Array): Promise<BelgeDocument> {
    const zip = await JSZip.loadAsync(data);
    const rootRels = await readRels(zip, "");
    let main = "word/document.xml";
    for (const r of rootRels.values()) if (r.type.endsWith("/officeDocument")) main = r.target;
    const mainFile = zip.file(main);
    if (!mainFile) throw new Error("Geçerli bir DOCX dosyası değil (document.xml yok).");
    const docRels = await readRels(zip, main);
    const loadXml = async (p: string | undefined) => {
      if (!p) return null;
      const f = zip.file(p);
      return f ? parseXml(await f.async("string")) : null;
    };
    const relOf = (suffix: string) => [...docRels.values()].find((r) => r.type.endsWith(suffix))?.target;

    // Ayarlar (OfficeHere özel özellikleri)
    const settings = defaultSettings();
    const custom = await loadXml("docProps/custom.xml");
    if (custom) {
      for (const pr of descendants(custom.documentElement, "property")) {
        if (pr.getAttribute("name") === "OfficeHereBelge") {
          try {
            const j = JSON.parse(pr.textContent ?? "{}") as Partial<DocSettings>;
            if (j.styleSet) settings.styleSet = j.styleSet;
            if (j.theme) settings.theme = j.theme;
            if (j.fontTheme !== undefined) settings.fontTheme = j.fontTheme ?? null;
            if (j.paragraphSpacing) settings.paragraphSpacing = j.paragraphSpacing;
          } catch {
            // yok say
          }
        }
      }
    }
    this.ours = resolveStyles(settings.styleSet, settings.theme, settings.fontTheme, settings.paragraphSpacing);
    this.theme = parseTheme(await loadXml(relOf("/theme")));
    this.sheet = new StyleSheet(await loadXml(relOf("/styles")), this.theme);
    this.numbering = new Numbering(await loadXml(relOf("/numbering")));
    // Dipnotlar ve son notlar (metin olarak)
    for (const [suffix, map] of [["/footnotes", this.footnotes], ["/endnotes", this.endnotes]] as const) {
      const x = await loadXml(relOf(suffix));
      if (!x) continue;
      for (const n of kids(x.documentElement)) {
        if (n.localName !== "footnote" && n.localName !== "endnote") continue;
        const type = attr(n, "type");
        if (type && type !== "normal") continue;
        const text = kids(n, "p")
          .map((p) =>
            descendants(p, "t")
              .map((t) => t.textContent ?? "")
              .join(""),
          )
          .join("\n")
          .replace(/^\s+/, "");
        map.set(attr(n, "id") ?? "", text);
      }
    }
    // Medya
    const mediaFiles = Object.keys(zip.files).filter((n) => !zip.files[n].dir).filter((n) => /\/media\//.test(n) || /^word\/[^/]+\.(png|jpe?g|gif|bmp)$/i.test(n));
    for (const name of mediaFiles) {
      const ext = name.slice(name.lastIndexOf(".") + 1).toLowerCase();
      const bytes = await zip.file(name)!.async("uint8array");
      this.media.set(name, `data:${MIME[ext] ?? "application/octet-stream"};base64,${toBase64(bytes)}`);
    }

    const docXml = parseXml(await mainFile.async("string"));
    const docEl = docXml.documentElement;
    const body = kid(docEl, "body");
    if (!body) throw new Error("Belge gövdesi bulunamadı.");
    const reader = new PartReader(this, docRels);
    const blocks = reader.blocks(body);

    // Bölüm özellikleri
    const sect = kid(body, "sectPr") ?? descendants(body, "sectPr").pop() ?? null;
    if (sect) {
      const pgSz = kid(sect, "pgSz");
      const pgMar = kid(sect, "pgMar");
      const w = attrNum(pgSz, "w");
      const h = attrNum(pgSz, "h");
      if (w && h) {
        settings.page.width = w;
        settings.page.height = h;
      }
      settings.page.orientation = attr(pgSz, "orient") === "landscape" || (w && h && w > h) ? "landscape" : "portrait";
      const m = settings.page.margin;
      const num = (n: string, def: number) => Math.abs(attrNum(pgMar, n) ?? def);
      settings.page.margin = { top: num("top", m.top), bottom: num("bottom", m.bottom), left: num("left", m.left), right: num("right", m.right), header: num("header", m.header), footer: num("footer", m.footer), gutter: num("gutter", 0) };
      const cols = kid(sect, "cols");
      settings.page.columns = Math.max(1, Math.min(3, attrNum(cols, "num") ?? 1));
      settings.page.columnSpace = attrNum(cols, "space") ?? 708;
      const pb = kid(sect, "pgBorders");
      const top = kid(pb, "top");
      if (top && attr(top, "val") !== "none") {
        const val = attr(top, "val") ?? "single";
        settings.pageBorder = {
          style: val === "double" ? "double" : val === "dotted" ? "dotted" : val.startsWith("dash") ? "dashed" : val === "thick" ? "thick" : "single",
          color: /^[0-9a-f]{6}$/i.test(attr(top, "color") ?? "") ? `#${attr(top, "color")!.toUpperCase()}` : "#000000",
          width: (attrNum(top, "sz") ?? 8) / 8,
        };
      }
      const hfRef = async (tag: "headerReference" | "footerReference") => {
        const refs = kids(sect, tag);
        const ref = refs.find((r) => attr(r, "type") === "default") ?? refs[0];
        const rid = ref ? (ref.getAttributeNS(NS.r, "id") ?? attr(ref, "id")) : null;
        const target = rid ? docRels.get(rid)?.target : undefined;
        const xml = await loadXml(target);
        if (!xml || !target) return null;
        const pr = new PartReader(this, await readRels(zip, target));
        const content = pr.blocks(xml.documentElement);
        if (pr.watermark) settings.watermark = pr.watermark;
        const hasText = JSON.stringify(content).match(/"text":"[^"]*\S|pageField|"image"/);
        return hasText ? ({ type: "doc", content } as JSONContent) : null;
      };
      settings.header = await hfRef("headerReference");
      settings.footer = await hfRef("footerReference");
    }
    const bg = kid(docEl, "background");
    const bgc = attr(bg, "color");
    if (bgc && /^[0-9a-f]{6}$/i.test(bgc) && bgc.toUpperCase() !== "FFFFFF") settings.pageColor = `#${bgc.toUpperCase()}`;

    // Açıklamalar
    const commentsXml = await loadXml(relOf("/comments"));
    if (commentsXml) {
      const comments: Record<string, CommentData> = {};
      for (const c of kids(commentsXml.documentElement, "comment")) {
        const id = `c${attr(c, "id")}`;
        const text = kids(c, "p")
          .map((p) =>
            descendants(p, "t")
              .map((t) => t.textContent ?? "")
              .join(""),
          )
          .join("\n");
        const author = attr(c, "author") ?? "";
        comments[id] = { id, author, initials: attr(c, "initials") ?? initialsOf(author), date: attr(c, "date") ?? new Date().toISOString(), text };
      }
      settings.comments = comments;
    }

    // Belge özellikleri
    const core = await loadXml("docProps/core.xml");
    if (core) {
      const get = (name: string) => descendants(core.documentElement, name)[0]?.textContent ?? "";
      settings.properties = { title: get("title"), subject: get("subject"), author: get("creator"), keywords: get("keywords"), created: get("created") || settings.properties.created, modified: get("modified") || settings.properties.modified };
    }

    return { content: { type: "doc", content: blocks.length ? blocks : [{ type: "paragraph" }] }, settings };
  }
}

export function initialsOf(name: string): string {
  return (
    name
      .split(/\s+/)
      .filter(Boolean)
      .map((w) => w[0]!.toLocaleUpperCase("tr-TR"))
      .join("")
      .slice(0, 3) || "K"
  );
}

export async function readDocx(data: ArrayBuffer | Uint8Array): Promise<BelgeDocument> {
  return new DocxReader().read(data);
}
