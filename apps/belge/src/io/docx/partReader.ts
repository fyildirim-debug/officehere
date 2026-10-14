// DOCX parça dönüştürücüsü: gövde / üst bilgi / alt bilgi / not içeriğini (paragraflar, koşular,
// alanlar, çizimler, metin kutuları, tablolar, İçindekiler) ProseMirror JSON'a çevirir.
import { toLinear } from "../../editor/math/ast";
import { fromOmml } from "../../editor/math/render";
import { indexNode, xeNode } from "./indexXml";
import type { JSONContent } from "@tiptap/core";
import type { DocSettings } from "../../model/document";
import { fieldPicture, type TocEntry } from "../../editor/nodes";
import { pictureOfInstr } from "../../editor/dateFormat";
import { normalizeDropCap } from "../../editor/dropCap";
import { attr, attrNum, descendants, kid, kids, path, NS } from "./xml";
import { parsePPr, parseRPr, stripUndef, type ParaProps, type RunProps } from "./readerStyles";
import { tableStyleFromId } from "./styleXml";
import type { ListSig } from "./numbering";
import type { DocxReader } from "./reader";
import { readDrawing, readPict } from "./readerDrawing";
import { extendCommentToWord } from "../common";

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
/** Sayfa/tarih alanı olarak okunan alan anahtar sözcükleri. */
const PAGE_FIELDS = new Set(["PAGE", "NUMPAGES", "DATE", "TIME"]);

/** Büyük ilk harf paragrafının harfi: çerçeve boyutu (yazı tipi boyutu) atılır, yazı tipi ayrıca döner. */
function dropLetter(n: JSONContent): { node: JSONContent; font: string | null } {
  let font: string | null = null;
  const marks = (n.marks ?? [])
    .map((m) => {
      if (m.type !== "textStyle") return m;
      font = (m.attrs?.fontFamily as string | null) ?? null;
      const { fontSize: _s, fontFamily: _f, ...rest } = m.attrs ?? {};
      return { ...m, attrs: rest };
    })
    .filter((m) => m.type !== "textStyle" || Object.values(m.attrs ?? {}).some((v) => v !== null && v !== undefined && v !== false));
  return { node: { ...n, ...(marks.length ? { marks } : { marks: undefined }) }, font };
}

/** Sayfa/tarih alanı düğümü; tarih ve saatte \@ biçim resmi korunur (varsayılansa yazılmaz). */
function pageFieldNode(kw: string, instr: string): JSONContent {
  if (kw !== "DATE" && kw !== "TIME") return { type: "pageField", attrs: { field: kw } };
  const pic = pictureOfInstr(instr);
  return { type: "pageField", attrs: { field: kw, format: pic && pic !== fieldPicture(kw) ? pic : null } };
}

export class PartReader {
  private fields: FieldState[] = [];
  /** Paragraf içindeki kesmelerin türü (sayfa / sütun), parça sırasıyla. */
  private segKinds: ("pageBreak" | "columnBreak")[] = [];
  /** Gövdede bölüm sonları: sectPr taşıyan paragrafların ardından öğe sırası. */
  sectionEnds: { index: number; sectPr: Element }[] = [];
  private openComments = new Set<string>();
  /** Aralığı açılmış ve en az bir metne uygulanmış açıklamalar. */
  private rangedComments = new Set<string>();
  private markedComments = new Set<string>();
  /** Aralıksız (nokta) açıklamalar: ardından gelen ilk sözcüğe, yoksa paragrafın son metnine bağlanır. */
  pendingPoints: string[] = [];
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
    const items: { node: JSONContent; list?: { numId: string; ilvl: number; sig: ListSig; start: number }; ctx?: string }[] = [];
    // Büyük ilk harf: Word harfi çerçeveli ayrı bir paragrafa (w:framePr w:dropCap) yazar; sonraki paragrafa katılır.
    let drop: { items: typeof items; dc: Record<string, unknown> } | null = null;
    const flushDrop = () => {
      if (drop) items.push(...drop.items);
      drop = null;
    };
    const visit = (el: Element) => {
      for (const c of kids(el)) {
        if (c.localName !== "p" && c.localName !== "bookmarkStart" && c.localName !== "bookmarkEnd") flushDrop();
        switch (c.localName) {
          case "p": {
            const frame = kid(kid(c, "pPr"), "framePr");
            const dcMode = attr(frame, "dropCap");
            const its = this.paragraph(c);
            if ((dcMode === "drop" || dcMode === "margin") && its.length === 1 && its[0].node.content?.length) {
              flushDrop();
              drop = { items: its, dc: { mode: dcMode, lines: attrNum(frame, "lines") ?? 3, distance: Math.round(((attrNum(frame, "hSpace") ?? 0) / 20) * 100) / 100 } };
              break;
            }
            const pending = drop as { items: typeof items; dc: Record<string, unknown> } | null;
            const host = pending && its.length && !its[0].list && (its[0].node.type === "paragraph" || its[0].node.type === "heading") ? its[0].node : null;
            if (pending && host) {
              const letters = (pending.items[0].node.content ?? []).map(dropLetter);
              const font = letters.map((l) => l.font).find(Boolean) ?? null;
              host.content = mergeText([...letters.map((l) => l.node), ...(host.content ?? [])]);
              host.attrs = { ...host.attrs, dropCap: normalizeDropCap({ ...pending.dc, font }) };
              drop = null;
            } else flushDrop();
            for (const it of its) items.push(it);
            const sp = kid(kid(c, "pPr"), "sectPr");
            if (sp && container.localName === "body") this.sectionEnds.push({ index: items.length, sectPr: sp });
            break;
          }
          case "tbl":
            items.push({ node: this.table(c) });
            break;
          case "altChunk":
          case "subDoc":
            // Başka biçimde gömülü içerik parçası / alt belge okunamaz.
            this.r.warnings.add("objects", `w:${c.localName}`);
            break;
          case "sdt": {
            const pr = kid(c, "sdtPr") ?? c;
            const gallery = attr(descendants(pr, "docPartGallery")[0], "val");
            const tag = attr(kid(pr, "tag"), "val");
            const content = kid(c, "sdtContent");
            const instr = content ? descendants(content, "instrText").map((t) => t.textContent ?? "").join("") : "";
            if (content && (tag === "OfficeHereIndex" || /^\s*INDEX\b/i.test(instr))) items.push({ node: indexNode(content, instr) });
            else if (content && (tag === "OfficeHereTableOfFigures" || /\\c\s+"/.test(instr))) items.push({ node: tofNode(this.toc(content), instr) });
            else if (gallery === "Table of Contents" && content) items.push({ node: this.toc(content) });
            else if (content && (gallery === "Bibliographies" || tag === "OfficeHereBibliography")) items.push({ node: bibliographyNode(content) });
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
    flushDrop();
    // Bağlamsal aralık (w:contextualSpacing): aynı stildeki ardışık paragraflar arasındaki aralık yok sayılır.
    // Modelde karşılığı olmadığından aradaki aralıklar doğrudan sıfırlanır (görünüm ve sayfa düzeni korunur).
    for (let i = 0; i + 1 < items.length; i++) {
      const a = items[i];
      const b = items[i + 1];
      if (!a.ctx || a.ctx !== b.ctx) continue;
      a.node.attrs = { ...a.node.attrs, spaceAfter: 0 };
      b.node.attrs = { ...b.node.attrs, spaceBefore: 0 };
    }
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

  private paragraph(p: Element): { node: JSONContent; list?: { numId: string; ilvl: number; sig: ListSig; start: number }; ctx?: string }[] {
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
    if (direct.suppressLineNumbers) attrs.suppressLineNumbers = true;

    const extra: JSONContent[] = [];
    const segments: JSONContent[][] = [[]];
    this.segKinds = [];
    this.boxes = [];
    const wmBefore = this.watermark;
    this.inline(p, styleDiff.r, [], segments, extra);
    if (this.pendingPoints.length) this.attachPointsToLast(segments);
    if (this.pointIds.size) {
      for (const seg of segments) extendCommentToWord(seg, this.pointIds);
      this.pointIds.clear();
    }
    // Yalnız filigran taşıyan üst bilgi paragrafı atlanır.
    if (this.watermark !== wmBefore && segments.every((x) => x.length === 0) && !this.boxes.length) return [];
    const type = mapped.type === "heading" ? "heading" : "paragraph";
    const out: { node: JSONContent; list?: { numId: string; ilvl: number; sig: ListSig; start: number }; ctx?: string }[] = [];
    const ctx = (direct.contextual ?? stylePP.contextual) ? (styleId ?? this.r.sheet.defaultParagraph) : undefined;
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
      // Başlıklar (anahat numaralandırması olsa da) listeye girmez: liste öğesi paragrafla başlamalıdır.
      out.push({ node, list: i === 0 && list && numId && type !== "heading" ? { numId, ilvl, sig: list.sig, start: list.start } : undefined, ctx });
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
        case "oMath":
        case "oMathPara": {
          // Denklem: OMML → doğrusal giriş.
          const list = c.localName === "oMath" ? [c] : kids(c, "oMath");
          for (const m of list) push({ type: "equation", attrs: { latex: toLinear(fromOmml(m)), display: c.localName === "oMathPara" } });
          break;
        }
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
          const fnode = fieldNode(instr, kids(c).map((r) => descendants(r, "t").map((t) => t.textContent ?? "").join("")).join(""));
          if (fnode) push(fnode);
          else if (PAGE_FIELDS.has(kw)) push(pageFieldNode(kw, instr));
          else if (kw === "HYPERLINK") {
            const href = parseHyperlinkInstr(instr);
            this.inline(c, base, href ? [...marks, { type: "link", attrs: { href } }] : marks, segs, extra);
          } else {
            // Tanınmayan alan: görünen sonucu metin olarak okunur (alan kodu kaydedilmez).
            if (kw && kw !== "SHAPE") this.r.warnings.add("fields", `alan: ${kw}`);
            this.inline(c, base, marks, segs, extra);
          }
          break;
        }
        case "bookmarkStart": {
          const name = attr(c, "name");
          if (name && name !== "_GoBack") push({ type: "bookmark", attrs: { name } });
          break;
        }
        case "commentRangeStart":
          this.openComments.add(`c${attr(c, "id")}`);
          this.rangedComments.add(`c${attr(c, "id")}`);
          break;
        case "commentRangeEnd": {
          const id = `c${attr(c, "id")}`;
          this.openComments.delete(id);
          // Boş aralık (hiçbir metni kapsamayan açıklama) nokta açıklaması gibi bağlanır.
          if (this.rangedComments.has(id) && !this.markedComments.has(id) && !this.pendingPoints.includes(id)) this.pendingPoints.push(id);
          break;
        }
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
    const push = (n: JSONContent) => {
      if (n.type === "text" && n.text) {
        for (const id of this.openComments) this.markedComments.add(id);
        if (this.pendingPoints.length) {
          // Bekleyen nokta açıklaması bu metnin ilk sözcüğüne bağlanır.
          const m = /^\s*\S+/.exec(n.text);
          const head = m ? m[0] : n.text;
          const points = this.takePoints();
          segs[segs.length - 1].push({ ...n, text: head, marks: [...marks, ...points] });
          if (head.length < n.text.length) segs[segs.length - 1].push(marks.length ? { ...n, text: n.text.slice(head.length), marks } : { ...n, text: n.text.slice(head.length) });
          return;
        }
      }
      segs[segs.length - 1].push(marks.length && n.type !== "pageBreak" ? { ...n, marks } : n);
    };
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
            } else if (PAGE_FIELDS.has(kw)) {
              f.kind = kw === "NUMPAGES" ? "numpages" : "page";
              push(pageFieldNode(kw, f.instr));
              f.emitted = true;
            } else {
              // SEQ / REF / PAGEREF / CITATION / MERGEFIELD: düğüm olarak alınır, sonuç metni atlanır.
              const fnode = fieldNode(f.instr, "");
              if (fnode) {
                f.kind = "page";
                push(fnode);
                f.emitted = true;
              }
            }
          } else if (t === "end" && f) {
            const kw = f.instr.trim().split(/\s+/)[0]?.toUpperCase();
            if (!f.emitted && f.phase === "instr" && PAGE_FIELDS.has(kw)) push(pageFieldNode(kw, f.instr));
            if (!f.emitted && f.phase === "instr" && kw === "XE") {
              const xe = xeNode(f.instr);
              if (xe) push(xe);
            } else if (!f.emitted && f.kind !== "link" && kw && kw !== "SHAPE" && !PAGE_FIELDS.has(kw)) this.r.warnings.add("fields", `alan: ${kw}`);
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
        case "commentReference": {
          // Aralığı olmayan (yalnız başvurusu bulunan) açıklama: konumundaki sözcüğe bağlanır.
          const id = `c${attr(c, "id")}`;
          if (!this.rangedComments.has(id) && !this.pendingPoints.includes(id)) this.pendingPoints.push(id);
          break;
        }
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
        case "object":
          // Gömülü (OLE) nesne: önizleme resmi okunur, nesnenin kendisi kaydedilmez.
          if (!suppress) {
            this.r.warnings.add("objects", "w:object");
            readPict(this, c, push, extra);
          }
          break;
        case "control":
          this.r.warnings.add("objects", "w:control");
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

  /** Bu paragrafta noktaya bağlanan açıklamalar (sözcüğün tamamına yayılır). */
  private pointIds = new Set<string>();

  private takePoints(): Mark[] {
    const out = this.pendingPoints.map((id) => ({ type: "comment", attrs: { id } }));
    for (const id of this.pendingPoints) {
      this.markedComments.add(id);
      this.pointIds.add(id);
    }
    this.pendingPoints = [];
    return out;
  }

  /** Paragraf sonunda bekleyen nokta açıklamaları paragrafın son metnine bağlanır (metin yoksa sonraki paragrafa kalır). */
  private attachPointsToLast(segs: JSONContent[][]) {
    for (let s = segs.length - 1; s >= 0; s--) {
      for (let i = segs[s].length - 1; i >= 0; i--) {
        const n = segs[s][i];
        if (n.type === "text" && n.text) {
          segs[s][i] = { ...n, marks: [...(n.marks ?? []), ...this.takePoints()] };
          return;
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
        // Geçersiz/dev birleştirme (bozuk dosya) sütun sınırına (63) indirgenir.
        const span = Math.max(1, Math.min(63, Math.round(attrNum(kid(tcPr, "gridSpan"), "val") ?? 1) || 1));
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
    const firstTr = kids(tbl, "tr")[0];
    const hdr = path(firstTr, "trPr", "tblHeader");
    const repeatHeader = !!hdr && !["0", "false", "off"].includes(attr(hdr, "val") ?? "");
    const mar = kid(tblPr, "tblCellMar");
    const pt = (name: string, alt: string, def: number) => {
      const el = kid(mar, name) ?? kid(mar, alt);
      const w = attrNum(el, "w");
      return el && w !== null && attr(el, "type") !== "pct" ? w / 20 : def;
    };
    const cellMargins = mar ? { top: pt("top", "top", 0), bottom: pt("bottom", "bottom", 0), left: pt("left", "start", 5.4), right: pt("right", "end", 5.4) } : null;
    const align = jc === "center" || jc === "right" ? jc : jc === "end" ? "right" : null;
    // Soldan girinti (tblInd, twip → px); yüzde ve "nil" girintisi yok sayılır.
    const ind = kid(tblPr, "tblInd");
    const indType = attr(ind, "type");
    const indent = ind && (!indType || indType === "dxa") ? Math.round((attrNum(ind, "w") ?? 0) / 15) || 0 : 0;
    return {
      type: "table",
      attrs: {
        tableStyle,
        headerRow: flag("firstRow", 0x20),
        bandedRows: !flag("noHBand", 0x200),
        firstColumn: flag("firstColumn", 0x80),
        totalRow: flag("lastRow", 0x40),
        align,
        ...(indent && !align ? { indent } : {}),
        repeatHeader,
        cellMargins,
      },
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
  for (const k of ["shadow", "outline", "emboss", "imprint"] as const) if (p[k]) ts[k] = true;
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

/** Alan yönergesinden düğüm (SEQ, REF, PAGEREF, CITATION, MERGEFIELD); tanınmayan alan için null. */
export function fieldNode(instr: string, result: string): JSONContent | null {
  const t = instr.trim();
  const parts = t.match(/"[^"]*"|\S+/g) ?? [];
  const kw = parts[0]?.toUpperCase();
  const arg = (parts[1] ?? "").replace(/^"|"$/g, "");
  const sw = (s: string) => new RegExp(`\\\\${s}(\\s|$)`, "i").test(t);
  switch (kw) {
    case "SEQ":
      return { type: "seqField", attrs: { label: arg || "Şekil", num: parseInt(result, 10) || 1 } };
    case "REF":
      return arg ? { type: "crossRef", attrs: { ref: arg, kind: sw("p") ? "position" : "full", link: sw("h"), text: result } } : null;
    case "PAGEREF":
      return arg ? { type: "crossRef", attrs: { ref: arg, kind: "page", link: sw("h"), text: result } } : null;
    case "CITATION": {
      const pm = /\\p\s+("[^"]*"|\S+)/i.exec(t);
      return arg ? { type: "citation", attrs: { tag: arg, pages: pm ? pm[1].replace(/"/g, "") : "", text: result } } : null;
    }
    case "MERGEFIELD":
      return arg ? { type: "mergeField", attrs: { name: arg } } : null;
    default:
      return null;
  }
}

/** Şekiller Tablosu: İçindekiler ayrıştırıcısının girdileri + \\c etiketi. */
function tofNode(toc: JSONContent, instr: string): JSONContent {
  const label = /\\c\s+"([^"]+)"/.exec(instr)?.[1] ?? "Şekil";
  const entries = ((toc.attrs?.entries as TocEntry[]) ?? []).map((e) => ({ ...e, level: 1 }));
  return { type: "toc", attrs: { kind: "tof", label, title: "", entries } };
}

/** Kaynakça bloğu (girdiler kaynaklardan yeniden üretilir; başlık korunur). */
/** Kaynakça bloğu: başlık ve girdiler (italik bölüm ayrı; kaynaklar belgede varsa alan güncellemesi yeniler). */
function bibliographyNode(content: Element): JSONContent {
  const [first, ...rest] = descendants(content, "p");
  const title = first ? descendants(first, "t").map((t) => t.textContent ?? "").join("").trim() : "";
  const entries = rest
    .map((p) => {
      const e = { before: "", italic: "", after: "" };
      for (const r of descendants(p, "r")) {
        const text = kids(r, "t").map((t) => t.textContent ?? "").join("");
        if (!text) continue;
        const i = path(r, "rPr", "i");
        const italic = !!i && !["0", "false"].includes(attr(i, "val") ?? "");
        if (italic && !e.after) e.italic += text;
        else if (e.italic) e.after += text;
        else e.before += text;
      }
      return e;
    })
    .filter((e) => e.before || e.italic || e.after);
  return { type: "bibliography", attrs: { title: title || "Kaynakça", entries } };
}
