// DOCX parça dönüştürücüsü: gövde / üst bilgi / alt bilgi / not içeriğini (paragraflar, koşular,
// alanlar, çizimler, metin kutuları, tablolar, İçindekiler) ProseMirror JSON'a çevirir.
import type { JSONContent } from "@tiptap/core";
import type { DocSettings } from "../../model/document";
import type { TocEntry } from "../../editor/nodes";
import { attr, attrNum, descendants, kid, kids, path, NS } from "./xml";
import { parsePPr, parseRPr, stripUndef, type ParaProps, type RunProps } from "./readerStyles";
import { tableStyleFromId } from "./styleXml";
import type { ListSig } from "./numbering";
import type { DocxReader } from "./reader";
import { readDrawing, readPict } from "./readerDrawing";

export type Mark = { type: string; attrs?: Record<string, unknown> };

export interface Rel {
  type: string;
  target: string;
  external: boolean;
}

interface FieldState {
  instr: string;
  phase: "instr" | "result";
  kind: "link" | "page" | "numpages" | "other";
  href?: string;
  emitted?: boolean;
}

/** Bir parçanın (gövde / üst bilgi / alt bilgi) dönüştürücüsü. */
export class PartReader {
  private fields: FieldState[] = [];
  /** Paragraf içindeki kesmelerin türü (sayfa / sütun), parça sırasıyla. */
  private segKinds: ("pageBreak" | "columnBreak")[] = [];
  /** Gövdede bölüm sonları: sectPr taşıyan paragrafların ardından öğe sırası. */
  sectionEnds: { index: number; sectPr: Element }[] = [];
  private openComments = new Set<string>();
  private listCounts = new Map<string, number>();
  watermark: DocSettings["watermark"] = null;

  constructor(
    readonly r: DocxReader,
    readonly rels: Map<string, Rel>,
  ) {}

  /** Not metni okunuyor: not stilinin (Dipnot Metni) biçimi doğrudan biçime çevrilmez. */
  noteMode = false;

  blocks(container: Element): JSONContent[] {
    return this.foldLists(this.collect(container));
  }

  private collect(container: Element) {
    const items: { node: JSONContent; list?: { numId: string; ilvl: number; sig: ListSig; start: number } }[] = [];
    const visit = (el: Element) => {
      for (const c of kids(el)) {
        switch (c.localName) {
          case "p": {
            for (const it of this.paragraph(c)) items.push(it);
            const sp = kid(kid(c, "pPr"), "sectPr");
            if (sp && container.localName === "body") this.sectionEnds.push({ index: items.length, sectPr: sp });
            break;
          }
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
    return items;
  }

  /** Gövdeyi bölümlere ayırır (her bölümün listeleri ayrı katlanır). */
  sections(body: Element): { blocks: JSONContent[]; sectPr: Element | null }[] {
    this.sectionEnds = [];
    const items = this.collect(body);
    const out: { blocks: JSONContent[]; sectPr: Element | null }[] = [];
    let start = 0;
    for (const e of this.sectionEnds) {
      out.push({ blocks: this.foldLists(items.slice(start, e.index)), sectPr: e.sectPr });
      start = e.index;
    }
    out.push({ blocks: this.foldLists(items.slice(start)), sectPr: kid(body, "sectPr") });
    return out;
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
    const noteStyle = this.noteMode && /(foot|end)note|dipnot|son ?not/i.test(`${styleId ?? ""} ${this.r.sheet.styles.get(styleId ?? "")?.name ?? ""}`);
    const mapped = noteStyle ? { type: "paragraph" as const, ourId: "Normal" } : this.r.mapParagraphStyle(styleId);
    const styleDiff = noteStyle ? { r: {}, p: {} } : this.r.styleDiff(styleId, mapped.ourId);
    // Liste: doğrudan numPr ya da stil üzerinden.
    const stylePP = this.r.sheet.effective(styleId).p;
    const numId = direct.numId ?? stylePP.numId;
    const ilvl = direct.ilvl ?? stylePP.ilvl ?? (numId ? listStyleLevel(this.r.numbering.styleLevel(numId, styleId), this.r.sheet.styles.get(styleId ?? "")?.name) : 0);
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
    if (direct.tabs?.length) attrs.tabs = direct.tabs;
    if (direct.outlineLevel && mapped.type !== "heading") attrs.outlineLevel = direct.outlineLevel;
    if (direct.keepNext) attrs.keepNext = true;
    if (direct.keepLines) attrs.keepLines = true;
    if (direct.pageBreakBefore) attrs.pageBreakBefore = true;

    const extra: JSONContent[] = [];
    const segments: JSONContent[][] = [[]];
    this.segKinds = [];
    this.boxes = [];
    const wmBefore = this.watermark;
    this.inline(p, styleDiff.r, [], segments, extra);
    // Yalnız filigran taşıyan üst bilgi paragrafı atlanır.
    if (this.watermark !== wmBefore && segments.every((x) => x.length === 0) && !this.boxes.length) return [];
    const type = mapped.type === "heading" ? "heading" : "paragraph";
    const out: { node: JSONContent; list?: { numId: string; ilvl: number; sig: ListSig; start: number } }[] = [];
    // Yatay çizgi: metni olmayan, yalnız alt kenarlıklı paragraf.
    const onlyBottom = pp.borders && pp.borders.bottom && !pp.borders.top && !pp.borders.left && !pp.borders.right;
    if (onlyBottom && segments.length === 1 && segments[0].length === 0 && mapped.ourId === "Normal" && !list) {
      return [{ node: { type: "horizontalRule" } }];
    }
    segments.forEach((content, i) => {
      if (i > 0) out.push({ node: { type: this.segKinds[i - 1] ?? "pageBreak" } });
      const empty = content.length === 0;
      // Sayfa sonunun önündeki/arkasındaki boş parçalar paragraf oluşturmaz.
      if (empty && segments.length > 1 && (i === 0 || i === segments.length - 1)) return;
      const node: JSONContent = { type, attrs: { ...attrs }, ...(empty ? {} : { content: mergeText(content) }) };
      out.push({ node, list: i === 0 && list && numId ? { numId, ilvl, sig: list.sig, start: list.start } : undefined });
    });
    for (const e of extra) out.push({ node: e });
    // Paragrafa bağlı metin kutuları paragraftan önce yer alır.
    if (this.boxes.length) out.unshift(...this.boxes.map((node) => ({ node })));
    return out;
  }

  /** Geçerli paragrafa bağlı metin kutuları. */
  boxes: JSONContent[] = [];

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
    const change = kid(rPr, "rPrChange");
    if (change) marks.push(this.formatChange(change, base));
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
          if (attr(c, "type") === "page" || attr(c, "type") === "column") {
            this.segKinds.push(attr(c, "type") === "column" ? "columnBreak" : "pageBreak");
            segs.push([]);
          } else push({ type: "hardBreak" });
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
        case "footnoteRef":
        case "endnoteRef":
          // Not metninin başındaki kendi başvuru işareti atlanır.
          break;
        case "footnoteReference":
        case "endnoteReference": {
          if (suppress) break;
          const kind = c.localName === "footnoteReference" ? "footnote" : "endnote";
          const nid = attr(c, "id") ?? "";
          const note = (kind === "footnote" ? this.r.footnotes : this.r.endnotes).get(nid);
          segs[segs.length - 1].push({ type: "note", attrs: { kind, text: note?.text ?? "", content: note?.content ?? null, id: `n${kind[0]}${nid}` } });
          break;
        }
        case "drawing":
          if (!suppress) readDrawing(this, c, push, extra);
          break;
        case "pict":
          if (!suppress) readPict(this, c, push, extra);
          break;
        case "AlternateContent": {
          const ch = kid(c, "Choice") ?? kid(c, "Fallback");
          const dr = ch ? descendants(ch, "drawing")[0] : null;
          if (dr) readDrawing(this, dr, push, extra);
          else if (ch) {
            const pc = descendants(ch, "pict")[0];
            if (pc) readPict(this, pc, push, extra);
          }
          break;
        }
      }
    }
  }

  /** w:rPrChange: değişiklik öncesi biçim, "formatChange" işaretinin "old" özniteliğine. */
  private formatChange(change: Element, base: RunProps): Mark {
    const oldPr = kid(change, "rPr");
    const oldStyle = attr(kid(oldPr, "rStyle"), "val");
    const oldChar = this.r.mapCharStyle(oldStyle);
    const props: RunProps = { ...base, ...(oldChar ? {} : stripUndef(this.r.sheet.charProps(oldStyle))), ...stripUndef(parseRPr(oldPr, this.r.theme)) };
    const old = runMarks(props);
    if (oldChar) old.push({ type: "charStyle", attrs: { styleId: oldChar } });
    return { type: "formatChange", attrs: { author: attr(change, "author") ?? "", date: attr(change, "date") ?? "", old: JSON.stringify(old) } };
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

/** Not metni paragraflarından not stili (Dipnot Metni) temizlenir. */
export function stripNoteStyle(b: JSONContent): JSONContent {
  if (b.type !== "paragraph" || !b.attrs) return b;
  const attrs = { ...b.attrs };
  if (attrs.styleId && /note|not/i.test(String(attrs.styleId))) attrs.styleId = null;
  return { ...b, attrs };
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

/**
 * Stilden gelen liste düzeyi. "List Bullet 2" / "Liste Madde İşareti 3" gibi stiller kendi numaralandırma
 * tanımlarında 0. düzeyde olsa da Word'de bir alt düzey gibi girintili görünür: adındaki sayıdan düzey çıkarılır.
 */
export function listStyleLevel(level: number | null, name: string | undefined): number {
  if (level && level > 0) return level;
  const m = /^(?:list|liste)\b.*?\s(\d)$/i.exec((name ?? "").trim());
  return m ? Math.max(0, Math.min(8, Number(m[1]) - 1)) : (level ?? 0);
}
