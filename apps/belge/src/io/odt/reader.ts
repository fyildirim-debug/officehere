// ODT okuyucu (OpenDocument Metin 1.3): content.xml + styles.xml + meta.xml → ProseMirror JSON.
// Paragraflar, başlıklar, biçimler (stil zinciri), listeler, tablolar (birleştirilmiş hücreler), resimler,
// köprüler, dipnotlar, metin kutuları, İçindekiler, sayfa sonları, üst/alt bilgi ve sayfa yapısı.
import JSZip from "jszip";
import type { JSONContent } from "@tiptap/core";
import { defaultSettings, normalizeDoc, samePage, type BelgeDocument, type CommentData, type DocSettings, type PageSetup } from "../../model/document";
import { fieldPicture, type TocEntry } from "../../editor/nodes";
import { noteText } from "../../editor/notes";
import { normalizeDropCap } from "../../editor/dropCap";
import { bulletStyleAt, charMarks, extendCommentToWord, foldLists, fromBase64, imageNode, noteNode, paragraphNode, stripStyleProps, styleChar, stylesOf, type CharProps, type FlowItem, type ListInfo } from "../common";
import { attr, descendants, flattenedDocs, kid, kids, oversizedEntry, parseXml, zipText } from "../docx/xml";
import type { StyleDef, StyleProps } from "../../model/styles";
import { BelgeOpenError, LoadWarnings } from "../warnings";
import { OdtStyles, decodeName, type ListLevel, type OdtStyle, type PageLayout, type ParaProps } from "./styles";

const KNOWN_PARA: Record<string, string> = {
  standard: "Normal",
  "default paragraph style": "Normal",
  "text body": "Normal",
  "body text": "Normal",
  default: "Normal",
  title: "Title",
  subtitle: "Subtitle",
  quotations: "Quote",
  quote: "Quote",
  "intense quote": "IntenseQuote",
  "no spacing": "NoSpacing",
  "list paragraph": "ListParagraph",
  caption: "Caption",
  illustration: "Caption",
  figure: "Caption",
  drawing: "Caption",
};

const KNOWN_CHAR: Record<string, string> = {
  emphasis: "Emphasis",
  "strong emphasis": "Strong",
  strong: "Strong",
  "subtle emphasis": "SubtleEmphasis",
  "intense emphasis": "IntenseEmphasis",
  "subtle reference": "SubtleReference",
  "intense reference": "IntenseReference",
  "book title": "BookTitle",
};

const BULLETS: Record<string, string> = { "○": "circle", "◦": "circle", "■": "square", "▪": "square", "–": "dash", "-": "dash", "✓": "check", "✔": "check", "➢": "arrow", "►": "arrow", "❖": "diamond" };

/** İçerdikleri metnin biçimini belirleyen kapsayıcı paragraf stilleri. */
const CONTAINER = new Set(["footnote", "endnote", "header", "footer", "table contents", "table heading", "contents heading", "horizontal line"]);

const MAX_REPEAT = 64;
const MAX_ROWS = 5000;

type Scope = "content" | "styles";

interface Ctx {
  props: CharProps;
  charStyle: string | null;
  link: string | null;
}

interface ListCtx {
  style: string | null;
  level: number;
  key: string;
}

const ptPx = (pt: number) => Math.round((pt * 4) / 3);
const ptTw = (pt: number) => Math.round(pt * 20);
const lower = (s: OdtStyle) => s.display.toLowerCase().trim();

export class OdtReader {
  st = new OdtStyles();
  media = new Map<string, Uint8Array>();
  scope: Scope = "content";
  private noteSeq = 0;
  private listSeq = 0;
  private counts = new Map<string, number>();
  private lastListOfStyle = new Map<string, string>();
  /** Paragraf içi beyaz boşluk daraltma durumu. */
  private lastSpace = true;
  private extra: JSONContent[] = [];
  /** Açıklamalar (office:annotation): okunanlar, açık aralıklar ve aralıksız (nokta) açıklama. */
  private comments: Record<string, CommentData> = {};
  private openComments = new Set<string>();
  private commentByName = new Map<string, string>();
  private pointComment: string | null = null;
  /** Bu paragrafta noktaya bağlanan açıklamalar (sözcüğün tamamına yayılır). */
  private pointIds = new Set<string>();
  /** Değişiklik izleme bölgeleri (text:tracked-changes) ve açık ekleme aralıkları. */
  private changes = new Map<string, { kind: "insertion" | "deletion" | "format"; author: string; date: string; text: string }>();
  private openInsertions = new Map<string, { author: string; date: string }>();
  /** Okunamayan (kaydedince kaybolacak) içerik uyarıları. */
  warnings = new LoadWarnings();
  /** Belgenin stil tablosu (kaynağın varsayılanları Normal stile taşındıktan sonra). */
  private styles: Record<string, StyleDef> = stylesOf();

  async read(data: ArrayBuffer | Uint8Array): Promise<BelgeDocument> {
    let zip: JSZip;
    try {
      zip = await JSZip.loadAsync(data);
    } catch (err) {
      throw new BelgeOpenError("Geçerli bir ODT dosyası değil (ZIP paketi açılamadı).", (err as Error).message);
    }
    const contentFile = zip.file("content.xml");
    if (!contentFile) throw new Error("Geçerli bir ODT dosyası değil (content.xml yok).");
    const mime = await zip.file("mimetype")?.async("string");
    if (mime && !/opendocument\.text/.test(mime)) throw new Error(notTextMessage(mime));
    for (const name of Object.keys(zip.files)) {
      const f = zip.files[name];
      if (f.dir || /\.xml$/i.test(name) || name === "mimetype" || this.media.size > 500) continue;
      if (/^(\.\/)?(Object |ObjectReplacements\/)/.test(name)) this.warnings.add("objects", name);
      if (/^(\.\/)?Basic\/|^(\.\/)?Scripts\//.test(name)) this.warnings.add("macros");
      if (oversizedEntry(f)) {
        this.warnings.add("images", `olağandışı büyük: ${name}`);
        continue;
      }
      this.media.set(name.replace(/^\.\//, ""), await f.async("uint8array"));
    }
    const stylesFile = zip.file("styles.xml");
    const stylesXml = stylesFile ? parseXml(await zipText(stylesFile)) : null;
    const contentXml = parseXml(await zipText(contentFile));
    if (stylesXml) this.st.load(stylesXml.documentElement, "styles");
    this.st.load(contentXml.documentElement, "content");
    let meta: Element | null = null;
    const metaFile = zip.file("meta.xml");
    if (metaFile) {
      try {
        meta = parseXml(await zipText(metaFile)).documentElement;
      } catch {
        // Bozuk meta.xml yok sayılır.
      }
    }
    if (flattenedDocs.has(contentXml)) this.warnings.add("deep");
    return this.readParsed(contentXml.documentElement, meta);
  }

  /** Düz XML OpenDocument (.fodt): stiller, içerik, meta ve resimler tek belgede. */
  readFlat(xml: string): BelgeDocument {
    const doc = parseXml(xml);
    const root = doc.documentElement;
    if (root.localName !== "document") throw new Error("Geçerli bir düz XML OpenDocument (.fodt) dosyası değil.");
    const mime = attr(root, "mimetype");
    if (mime && !/opendocument\.text/.test(mime)) throw new Error(notTextMessage(mime));
    if (flattenedDocs.has(doc)) this.warnings.add("deep");
    if (kid(root, "scripts") && descendants(kid(root, "scripts")!, "*").length) this.warnings.add("macros");
    // Tek belgede otomatik stiller hem içerik hem üst/alt bilgi kapsamında geçerlidir.
    this.st.load(root, "styles");
    const masters = this.st.masterOrder.length;
    this.st.load(root, "content");
    this.st.masterOrder.length = masters;
    return this.readParsed(root, kid(root, "meta"));
  }

  private readParsed(contentRoot: Element, meta: Element | null): BelgeDocument {
    const settings = defaultSettings();
    this.adoptDefaults(settings);
    this.readTrackedChanges(contentRoot);
    const text = descendants(contentRoot, "text").find((e) => e.parentElement?.localName === "body");
    if (!text) throw new Error("Belge gövdesi bulunamadı.");
    this.scope = "content";
    const items = this.blocks(text);
    // Bölümler: asıl sayfa değişen paragraflarda yeni bölüm başlar.
    const folded = foldLists(items);
    const firstMaster = this.firstMaster(text) ?? this.st.masterOrder[0] ?? "Standard";
    const defPage = this.page(firstMaster, settings.page);
    settings.page = defPage;
    const firstLayout = this.st.masters.get(firstMaster)?.layout;
    settings.pageColor = (firstLayout ? this.st.layouts.get(firstLayout)?.background : null) ?? null;
    const sections: JSONContent[] = [{ type: "section", attrs: { kind: "nextPage", page: null }, content: [] }];
    for (const n of folded) {
      if (n.type === "__master") {
        const page = this.page(String(n.attrs?.name), defPage);
        if (sections[sections.length - 1].content!.length) sections.push({ type: "section", attrs: { kind: "nextPage", page: samePage(page, defPage) ? null : page }, content: [] });
        continue;
      }
      sections[sections.length - 1].content!.push(n);
    }
    // Okunan açıklamalardan yalnız belgede bağlı olanlar (aralığı ya da noktası metne düşenler) tutulur.
    const used = new Set<string>();
    const walk = (n: JSONContent) => {
      for (const m of n.marks ?? []) if (m.type === "comment") used.add(String(m.attrs?.id));
      n.content?.forEach(walk);
    };
    sections.forEach(walk);
    settings.comments = Object.fromEntries(Object.entries(this.comments).filter(([id]) => used.has(id)));
    this.openComments.clear();
    this.readHeaders(firstMaster, settings);
    if (meta) this.readMeta(meta, settings);
    if (Object.keys(this.comments).length > Object.keys(settings.comments).length) this.warnings.add("comments", "bağlanamayan açıklama");
    return { content: normalizeDoc({ type: "doc", content: sections }), settings };
  }

  /**
   * Kaynağın varsayılan paragraf stili (Standart; yoksa varsayılan stil) bizim Normal stilimize taşınır:
   * belge kendi yazı tipi, boyutu ve aralıklarıyla görünür, kaydedilince sayfa düzeni kaymaz.
   */
  private adoptDefaults(settings: DocSettings) {
    const chain = this.st.chain("Standard", "paragraph", "styles");
    const t = this.st.mergeText(this.st.defaultText(), chain);
    const pp = this.st.mergePara(chain);
    const ours = stylesOf().Normal;
    const props = { ...ours.props };
    let changed = false;
    const norm = (f?: string | null) => (f ?? "").toLowerCase().replace(/\s+/g, " ").trim();
    if (t.font && norm(t.font) !== norm(props.font)) {
      props.font = t.font;
      changed = true;
    }
    const set = (k: "size" | "spaceBefore" | "spaceAfter" | "line", v: number, d: number) => {
      if (Math.abs(v - ((props[k] as number | undefined) ?? d)) > 0.01) {
        props[k] = round(v);
        changed = true;
      }
    };
    set("size", t.size ?? 12, 11);
    set("spaceBefore", pp.spaceBefore ?? 0, 0);
    set("spaceAfter", pp.spaceAfter ?? 0, 0);
    set("line", pp.lineRule === "auto" || pp.lineRule === undefined ? (pp.lineHeight ?? 1) : 1, 1);
    if (changed) settings.customStyles = { ...settings.customStyles, Normal: { ...ours, props } };
    this.styles = stylesOf(settings);
  }

  /** text:tracked-changes → değişiklik bölgeleri (silinen metin bölgenin içindedir). */
  private readTrackedChanges(root: Element) {
    for (const tc of descendants(root, "tracked-changes")) {
      for (const region of kids(tc, "changed-region")) {
        const id = attr(region, "id") ?? region.getAttribute("xml:id");
        const ch = kids(region).find((k) => k.localName === "insertion" || k.localName === "deletion" || k.localName === "format-change");
        if (!id || !ch) continue;
        const info = kid(ch, "change-info");
        const author = kid(info, "creator")?.textContent?.trim() ?? "";
        const raw = kid(info, "date")?.textContent?.trim() ?? "";
        const date = raw && !/Z|[+-]\d\d:\d\d$/.test(raw) ? `${raw.replace(/\.\d+$/, "")}Z` : raw;
        const text = ch.localName === "deletion" ? kids(ch).filter((k) => k.localName === "p" || k.localName === "h" || k.localName === "list").map((k) => (k.textContent ?? "").replace(/\s+/g, " ")).join(" ").trim() : "";
        const kind = ch.localName === "insertion" ? "insertion" : ch.localName === "deletion" ? "deletion" : "format";
        if (kind === "format") this.warnings.add("trackFormat", "text:format-change");
        if (kind === "deletion" && kids(ch).some((k) => k.localName === "table" || descendants(k, "frame").length)) this.warnings.add("trackChanges", "silinen tablo/çerçeve");
        this.changes.set(id, { kind, author, date, text });
      }
    }
  }

  /* -------------------------- bloklar -------------------------- */

  blocks(container: Element, list?: ListCtx): FlowItem[] {
    const out: FlowItem[] = [];
    for (const c of kids(container)) out.push(...this.block(c, list));
    return out;
  }

  private block(c: Element, list?: ListCtx): FlowItem[] {
    switch (c.localName) {
      case "p":
      case "h":
        return this.paragraph(c, c.localName === "h", list);
      case "list":
        return this.list(c, list);
      case "table":
        return [...this.breakItems(this.st.get(attr(c, "style-name"), "table", this.scope)?.table.breakBefore ? "page" : undefined), { node: this.table(c) }];
      case "section":
      case "index-body":
        return this.blocks(c, list);
      case "table-of-content":
        return [{ node: this.toc(c) }];
      case "alphabetical-index":
      case "illustration-index":
      case "table-index":
      case "object-index":
      case "user-index":
      case "bibliography":
        return this.blocks(kid(c, "index-body") ?? c, list);
      default:
        return [];
    }
  }

  private breakItems(kind: "page" | "column" | undefined): FlowItem[] {
    return kind ? [{ node: { type: kind === "page" ? "pageBreak" : "columnBreak" } }] : [];
  }

  /** Paragraf stilinin bizim stilimiz ve başlık düzeyi. */
  private mapPara(chain: OdtStyle[], heading: boolean, outline: number | null): { id: string; level: number | null } {
    let level: number | null = heading ? Math.max(1, Math.min(6, outline ?? 1)) : null;
    let id = "Normal";
    for (let i = chain.length - 1; i >= 0; i--) {
      const n = lower(chain[i]);
      const h = /^(?:heading|başlık)\s*(\d)$/.exec(n);
      if (h && !heading) {
        level = Math.min(6, Number(h[1]));
        break;
      }
      if (KNOWN_PARA[n]) {
        id = KNOWN_PARA[n];
        break;
      }
    }
    return { id: level ? `Heading${level}` : id, level };
  }

  private paragraph(p: Element, heading: boolean, list?: ListCtx, cont = false): FlowItem[] {
    const chain = this.st.chain(attr(p, "style-name"), "paragraph", this.scope);
    const outline = Number(attr(p, "outline-level")) || null;
    const map = this.mapPara(chain, heading, outline);
    let base: StyleProps = styleChar(this.styles, map.id).base;
    const pp = this.st.mergePara(chain);
    // Kapsayıcı stiller (dipnot, üst/alt bilgi, tablo içeriği): kendi biçimleri doğrudan biçim sayılmaz.
    let firstBase: number | undefined;
    const ci = chain.findLastIndex((s) => !s.auto && CONTAINER.has(lower(s)));
    if (ci >= 0) {
      // Yalnız kapsayıcı stillerin kendi tanımladıkları (belge varsayılanından gelenler değil).
      const own = chain.slice(chain.findIndex((s) => !s.auto && CONTAINER.has(lower(s))), ci + 1);
      const cp: ParaProps = {};
      for (const s of own) for (const [k, v] of Object.entries(s.para)) if (v !== undefined) (cp as Record<string, unknown>)[k] = v;
      const ct = this.st.mergeText({}, own);
      const pick = <T extends object>(o: T) => Object.fromEntries(Object.entries(o).filter(([, v]) => v !== undefined && v !== null)) as Partial<T>;
      base = { ...base, ...pick({ font: ct.font ?? undefined, size: ct.size ?? undefined, bold: ct.bold, italic: ct.italic, color: ct.color ?? undefined, caps: ct.caps }), ...pick({ spaceBefore: cp.spaceBefore, spaceAfter: cp.spaceAfter, line: cp.lineRule === "auto" ? cp.lineHeight : undefined, align: cp.align as StyleProps["align"], indentLeft: cp.indentLeft, indentRight: cp.indentRight, keepNext: cp.keepNext }) };
      firstBase = cp.indentFirst;
    }
    const out: FlowItem[] = [];
    const master = chain.find((s) => s.auto && s.masterPage)?.masterPage;
    if (master && this.scope === "content") out.push({ node: { type: "__master", attrs: { name: master } } });
    else out.push(...this.breakItems(pp.breakBefore));
    // Satır içi içerik
    this.lastSpace = true;
    this.extra = [];
    const content = this.inline(p, { props: this.st.mergeText(this.st.defaultText(), chain), charStyle: null, link: null }, base, []);
    // Paragrafın sonundaki aralıksız açıklama son metne bağlanır.
    if (this.pointComment) {
      const lastText = [...content].reverse().find((n) => n.type === "text");
      if (lastText) {
        lastText.marks = [...(lastText.marks ?? []), { type: "comment", attrs: { id: this.pointComment } }];
        this.pointIds.add(this.pointComment);
        this.pointComment = null;
      }
    }
    if (this.pointIds.size) {
      extendCommentToWord(content, this.pointIds);
      this.pointIds.clear();
    }
    const last = content[content.length - 1];
    if (last?.type === "text" && last.text?.endsWith(" ") && this.lastSpace) last.text = last.text.replace(/ +$/, "");
    const boxes = this.extra;
    this.extra = [];
    const attrs: Record<string, unknown> = {};
    if (map.level) attrs.level = map.level;
    else if (map.id !== "Normal" && !(list && map.id === "ListParagraph")) attrs.styleId = map.id;
    if (pp.align && pp.align !== (base.align ?? "left")) attrs.textAlign = pp.align;
    const near = (x: number | undefined, y: number | undefined) => x === undefined || (y !== undefined && Math.abs(x - y) < 0.06) || (y === undefined && Math.abs(x) < 0.06);
    if (!list) {
      if (pp.indentLeft && !near(pp.indentLeft, base.indentLeft)) attrs.indentLeft = round(pp.indentLeft);
      if (pp.indentFirst && !near(pp.indentFirst, firstBase)) attrs.indentFirst = round(pp.indentFirst);
    }
    if (pp.indentRight && !near(pp.indentRight, base.indentRight)) attrs.indentRight = round(pp.indentRight);
    if (!near(pp.spaceBefore, base.spaceBefore)) attrs.spaceBefore = round(pp.spaceBefore!);
    if (!near(pp.spaceAfter, base.spaceAfter)) attrs.spaceAfter = round(pp.spaceAfter!);
    if (pp.lineHeight !== undefined) {
      if (pp.lineRule === "auto") {
        if (Math.abs(pp.lineHeight - (base.line ?? 1)) > 0.015) attrs.lineHeight = round(pp.lineHeight);
      } else {
        attrs.lineHeight = round(pp.lineHeight);
        attrs.lineRule = pp.lineRule;
      }
    }
    if (pp.shading) attrs.shading = pp.shading;
    if (pp.dropCap && !list) {
      const font = pp.dropCap.textStyle ? this.st.mergeText({}, this.st.chain(pp.dropCap.textStyle, "text", this.scope)).font : null;
      attrs.dropCap = normalizeDropCap({ mode: "drop", lines: pp.dropCap.lines, distance: Math.round(pp.dropCap.distance * 100) / 100, font: font ?? null });
    }
    if (pp.keepNext && !base.keepNext) attrs.keepNext = true;
    if (pp.keepLines) attrs.keepLines = true;
    const hr = !content.length && !map.level && ((pp.borderBottom && !pp.borderOther) || chain.some((s) => lower(s) === "horizontal line"));
    const node = hr ? { type: "horizontalRule" } : paragraphNode(attrs, content, map.level ? "heading" : "paragraph");
    for (const b of boxes) out.push({ node: b });
    out.push({ node, list: list && !map.level ? this.listInfo(list, cont) : undefined });
    out.push(...this.breakItems(pp.breakAfter));
    return out;
  }

  /* -------------------------- listeler -------------------------- */

  private listInfo(l: ListCtx, cont: boolean): ListInfo {
    const lv: ListLevel | undefined = this.st.list(l.style, this.scope)?.[l.level];
    const ordered = !!lv && !lv.bullet && lv.format !== "";
    let listStyle: string | null = null;
    if (lv?.bullet) listStyle = bulletStyleAt(BULLETS[lv.char] ?? null, l.level);
    else if (lv) {
      const paren = lv.suffix.includes(")");
      listStyle = lv.format === "a" ? (paren ? "lower-alpha-paren" : "lower-alpha") : lv.format === "A" ? "upper-alpha" : lv.format === "i" ? "lower-roman" : lv.format === "I" ? "upper-roman" : paren ? "decimal-paren" : null;
    }
    const ck = `${l.key}:${l.level}`;
    const n = this.counts.get(ck) ?? (lv?.start ?? 1) - 1;
    if (!cont) this.counts.set(ck, n + 1);
    return { key: l.key, level: l.level, ordered, listStyle, multilevel: lv && lv.display > 1 ? "outline" : null, start: n + 1, cont };
  }

  private list(el: Element, parent?: ListCtx): FlowItem[] {
    const style = attr(el, "style-name") ?? parent?.style ?? null;
    let key: string;
    if (parent) key = parent.key;
    else {
      const cont = attr(el, "continue-numbering") === "true" || !!attr(el, "continue-list");
      const prev = style ? this.lastListOfStyle.get(style) : undefined;
      key = cont && prev ? prev : `l${++this.listSeq}`;
      if (style) this.lastListOfStyle.set(style, key);
    }
    const level = parent ? Math.min(8, parent.level + 1) : 0;
    const ctx: ListCtx = { style, level, key };
    // Yeni alt liste düzeyinde numara yeniden başlar.
    if (parent) this.counts.delete(`${key}:${level}`);
    const out: FlowItem[] = [];
    for (const item of kids(el)) {
      if (item.localName !== "list-item" && item.localName !== "list-header") continue;
      const sv = Number(attr(item, "start-value"));
      if (sv > 0) this.counts.set(`${key}:${level}`, sv - 1);
      let first = true;
      for (const c of kids(item)) {
        if (c.localName === "p" || c.localName === "h") {
          const s = ctx.style ? ctx : { ...ctx, style: this.st.chain(attr(c, "style-name"), "paragraph", this.scope).reverse().find((x) => x.listStyle)?.listStyle ?? null };
          out.push(...this.paragraph(c, c.localName === "h", s, !first));
          first = false;
        } else if (c.localName === "list") {
          out.push(...this.list(c, ctx));
          first = false;
        } else {
          for (const it of this.block(c)) out.push({ ...it, list: this.listInfo(ctx, true) });
          first = false;
        }
      }
    }
    return out;
  }

  /* -------------------------- satır içi -------------------------- */

  private textNode(s: string, ctx: Ctx, base: ReturnType<typeof styleChar>["base"], out: JSONContent[]) {
    if (!s) return;
    const props: CharProps = { ...stripStyleProps(ctx.props, base), link: ctx.link, charStyle: ctx.charStyle };
    // Bağlantı stillerinin varsayılan renkleri biçim sayılmaz.
    if (ctx.link && (props.color === "#0563C1" || props.color === "#000080")) props.color = null;
    const marks = charMarks(props);
    for (const id of this.openComments) marks.push({ type: "comment", attrs: { id } });
    // Değişiklik izleme: açık ekleme aralığındaki metin.
    const ins = [...this.openInsertions.values()].pop();
    if (ins && this.scope === "content") marks.push({ type: "insertion", attrs: { author: ins.author, date: ins.date } });
    // Aralıksız açıklama ardından gelen ilk sözcüğe bağlanır.
    const point = this.pointComment;
    if (point) {
      const m = /^\s*\S+/.exec(s);
      const head = m ? m[0] : s;
      this.pointComment = null;
      this.pointIds.add(point);
      out.push({ type: "text", text: head, marks: [...marks, { type: "comment", attrs: { id: point } }] });
      if (head.length < s.length) out.push({ type: "text", text: s.slice(head.length), ...(marks.length ? { marks } : {}) });
      return;
    }
    out.push({ type: "text", text: s, ...(marks.length ? { marks } : {}) });
  }

  /** office:annotation → açıklama verisi; aralıklı açıklama annotation-end'e kadar açık kalır. */
  private annotation(e: Element) {
    if (this.scope !== "content") return;
    const id = `c${Object.keys(this.comments).length}`;
    const author = kid(e, "creator")?.textContent?.trim() ?? "";
    const initials = kid(e, "creator-initials")?.textContent?.trim() || author.split(/\s+/).map((w) => w[0] ?? "").join("").toLocaleUpperCase("tr-TR").slice(0, 3);
    const raw = kid(e, "date")?.textContent?.trim() ?? "";
    const date = raw && !/Z|[+-]\d\d:\d\d$/.test(raw) ? `${raw.replace(/\.\d+$/, "")}Z` : raw;
    const text = kids(e)
      .flatMap((k) => (k.localName === "p" || k.localName === "h" ? [k] : k.localName === "list" ? descendants(k, "p") : []))
      .map((k) => k.textContent ?? "")
      .join("\n");
    this.comments[id] = { id, author, initials, date, text, ...(attr(e, "resolved") === "true" ? { resolved: true } : {}) };
    const name = attr(e, "name");
    if (name) {
      this.commentByName.set(name, id);
      this.openComments.add(id);
    } else this.pointComment = id;
  }

  private inline(el: Element, ctx: Ctx, base: ReturnType<typeof styleChar>["base"], out: JSONContent[]): JSONContent[] {
    for (let c = el.firstChild; c; c = c.nextSibling) {
      if (c.nodeType === 3) {
        let t = (c.nodeValue ?? "").replace(/[ \t\r\n]+/g, " ");
        if (this.lastSpace) t = t.replace(/^ /, "");
        if (t) {
          this.textNode(t, ctx, base, out);
          this.lastSpace = t.endsWith(" ");
        }
        continue;
      }
      if (c.nodeType !== 1) continue;
      const e = c as Element;
      switch (e.localName) {
        case "span":
        case "a": {
          const chain = this.st.chain(attr(e, "style-name"), "text", this.scope);
          let charStyle = ctx.charStyle;
          let from = 0;
          for (let i = chain.length - 1; i >= 0; i--) {
            const k = KNOWN_CHAR[lower(chain[i])];
            if (k) {
              charStyle = k;
              from = i + 1;
              break;
            }
            if (/internet link|hyperlink/.test(lower(chain[i]))) {
              from = i + 1;
              break;
            }
          }
          const props = this.st.mergeText(ctx.props, chain.slice(from));
          const href = e.localName === "a" ? attr(e, "href") : null;
          this.inline(e, { props, charStyle, link: href ?? ctx.link }, base, out);
          break;
        }
        case "s": {
          const n = Math.max(1, Math.min(1000, Number(attr(e, "c")) || 1));
          this.textNode(" ".repeat(n), ctx, base, out);
          this.lastSpace = false;
          break;
        }
        case "tab":
          this.textNode("\t", ctx, base, out);
          this.lastSpace = false;
          break;
        case "line-break":
          out.push({ type: "hardBreak" });
          this.lastSpace = true;
          break;
        case "bookmark":
        case "bookmark-start": {
          const name = attr(e, "name");
          if (name) out.push({ type: "bookmark", attrs: { name } });
          break;
        }
        case "note":
          out.push(this.note(e));
          this.lastSpace = false;
          break;
        case "frame":
          this.frame(e, out);
          this.lastSpace = false;
          break;
        case "date":
        case "time": {
          // Sabit tarih/saat (text:fixed="true") yazıldığı günün değerini korur: metin olarak okunur.
          if (attr(e, "fixed") === "true") {
            const txt = e.textContent ?? "";
            if (txt) this.textNode(txt, ctx, base, out);
          } else {
            const ds = attr(e, "data-style-name");
            const pic = ds ? (this.st.dataStyles.get(ds) ?? null) : null;
            const field = e.localName === "date" ? "DATE" : "TIME";
            out.push({ type: "pageField", attrs: { field, format: pic && pic !== fieldPicture(field) ? pic : null } });
          }
          this.lastSpace = false;
          break;
        }
        case "page-number":
        case "page-count":
          out.push({ type: "pageField", attrs: { field: e.localName === "page-count" ? "NUMPAGES" : "PAGE" } });
          this.lastSpace = false;
          break;
        case "annotation":
          this.annotation(e);
          break;
        case "change-start":
        case "change-end":
        case "change": {
          const id = attr(e, "change-id") ?? "";
          const ch = this.changes.get(id);
          if (!ch || this.scope !== "content") break;
          if (ch.kind === "insertion") {
            if (e.localName === "change-start") this.openInsertions.set(id, ch);
            else this.openInsertions.delete(id);
          } else if (ch.kind === "deletion" && e.localName === "change" && ch.text) {
            // Silinen metin bulunduğu konuma silme işaretiyle geri konur (kabul/ret düzenleyicide yapılır).
            const props: CharProps = { ...stripStyleProps(ctx.props, base), link: ctx.link, charStyle: ctx.charStyle };
            out.push({ type: "text", text: ch.text, marks: [...charMarks(props), { type: "deletion", attrs: { author: ch.author, date: ch.date } }] });
            this.lastSpace = false;
          }
          break;
        }
        case "annotation-end": {
          const id = this.commentByName.get(attr(e, "name") ?? "");
          if (id) this.openComments.delete(id);
          break;
        }
        case "meta":
        case "meta-field":
        case "ruby":
        case "ruby-base":
        case "hidden-text":
        case "sequence":
        case "sequence-ref":
        case "bookmark-ref":
        case "reference-ref":
        case "note-ref":
        case "chapter":
        case "title":
        case "subject":
        case "author-name":
        case "author-initials":
        case "file-name":
        case "user-defined":
        case "variable-get":
        case "variable-set":
        case "user-field-get":
        case "expression":
        case "text-input":
        case "drop-down":
        case "conditional-text":
        case "initial-creator":
        case "creation-date":
        case "description":
        case "keywords":
          this.inline(e, ctx, base, out);
          break;
      }
    }
    return out;
  }

  private note(e: Element): JSONContent {
    const kind = attr(e, "note-class") === "endnote" ? "endnote" : "footnote";
    const body = kid(e, "note-body");
    const saved = { lastSpace: this.lastSpace, extra: this.extra };
    const blocks = body ? foldLists(this.blocks(body)).filter((b) => !b.type?.startsWith("__")) : [];
    this.lastSpace = saved.lastSpace;
    this.extra = saved.extra;
    return noteNode(blocks, kind, `no${++this.noteSeq}`, noteText);
  }

  private frame(e: Element, out: JSONContent[]) {
    const w = lengthPx(attr(e, "width"));
    const h = lengthPx(attr(e, "height"));
    const img = kid(e, "image");
    const box = kid(e, "text-box");
    if (box) {
      const saved = { lastSpace: this.lastSpace, extra: this.extra };
      const content = foldLists(this.blocks(box)).filter((b) => ["paragraph", "heading", "bulletList", "orderedList"].includes(b.type ?? ""));
      this.lastSpace = saved.lastSpace;
      this.extra = saved.extra;
      this.extra.push({ type: "textBox", attrs: { x: lengthPx(attr(e, "x")) ?? 0, y: lengthPx(attr(e, "y")) ?? 0, w: w ?? 240, h: h ?? 120 }, content: content.length ? content : [{ type: "paragraph" }] });
      return;
    }
    // Resim: paketteki dosya ya da gömülü base64 veri; nesne yedek resimleri de denenir.
    for (const im of img ? [img, ...kids(e, "image").slice(1)] : []) {
      const href = (attr(im, "href") ?? "").replace(/^\.\//, "");
      const bin = kid(im, "binary-data");
      let bytes: Uint8Array | null = null;
      if (bin) {
        try {
          bytes = fromBase64(bin.textContent ?? "");
        } catch {
          bytes = null;
        }
      } else if (href) bytes = this.media.get(href) ?? this.media.get(decodeURIComponent(href)) ?? null;
      if (!bytes) continue;
      const alt = kid(e, "desc")?.textContent ?? kid(e, "title")?.textContent ?? null;
      const node = imageNode(bytes, { width: w, height: h, alt });
      if (node) {
        out.push(node);
        return;
      }
    }
  }

  /* -------------------------- tablolar -------------------------- */

  private table(t: Element): JSONContent {
    const tStyle = this.st.get(attr(t, "style-name"), "table", this.scope);
    const cols: { width?: number; rel?: number }[] = [];
    const rows: Element[] = [];
    let header = false;
    const collect = (el: Element) => {
      for (const c of kids(el)) {
        if (c.localName === "table-column") {
          const n = Math.max(1, Math.min(MAX_REPEAT, Number(attr(c, "number-columns-repeated")) || 1));
          const cs = this.st.get(attr(c, "style-name"), "table-column", this.scope)?.column ?? {};
          for (let k = 0; k < n && cols.length < MAX_REPEAT; k++) cols.push(cs);
        } else if (c.localName === "table-columns" || c.localName === "table-header-columns" || c.localName === "table-column-group") collect(c);
        else if (c.localName === "table-header-rows") {
          header = true;
          collect(c);
        } else if (c.localName === "table-rows" || c.localName === "table-row-group") collect(c);
        else if (c.localName === "table-row") {
          const n = Math.max(1, Math.min(100, Number(attr(c, "number-rows-repeated")) || 1));
          for (let k = 0; k < n && rows.length < MAX_ROWS; k++) rows.push(c);
        }
      }
    };
    collect(t);
    // Sütun genişlikleri (px): mutlak ya da göreli.
    const totalRel = cols.reduce((s, c) => s + (c.rel ?? 0), 0);
    const tableW = tStyle?.table.width ?? 481.9;
    const colPx = cols.map((c) => (c.width ? ptPx(c.width) : c.rel && totalRel ? ptPx((tableW * c.rel) / totalRel) : 0));
    let anyBorder = false;
    const content: JSONContent[] = [];
    for (const r of rows) {
      const cells: JSONContent[] = [];
      let col = 0;
      for (const c of kids(r)) {
        if (c.localName !== "table-cell" && c.localName !== "covered-table-cell") continue;
        const rep = Math.max(1, Math.min(MAX_REPEAT, Number(attr(c, "number-columns-repeated")) || 1));
        if (c.localName === "covered-table-cell") {
          col += rep;
          continue;
        }
        const span = Math.max(1, Math.min(MAX_REPEAT, Number(attr(c, "number-columns-spanned")) || 1));
        const rowspan = Math.max(1, Math.min(MAX_ROWS, Number(attr(c, "number-rows-spanned")) || 1));
        const cs = this.st.get(attr(c, "style-name"), "table-cell", this.scope)?.cell ?? {};
        if (cs.border) anyBorder = true;
        for (let k = 0; k < rep && col < MAX_REPEAT; k++) {
          const blocks = foldLists(this.blocks(c)).filter((b) => !b.type?.startsWith("__"));
          const widths = colPx.slice(col, col + span);
          cells.push({
            type: "tableCell",
            attrs: { colspan: span, rowspan, colwidth: widths.length === span && widths.every((x) => x > 0) ? widths : null, background: cs.bg ?? null, verticalAlign: cs.valign ?? null },
            content: blocks.length ? blocks : [{ type: "paragraph" }],
          });
          col += 1;
        }
      }
      if (cells.length) content.push({ type: "tableRow", content: cells });
    }
    if (!content.length) content.push({ type: "tableRow", content: [{ type: "tableCell", attrs: { colspan: 1, rowspan: 1 }, content: [{ type: "paragraph" }] }] });
    const align = tStyle?.table.align;
    // Soldan girinti: sola yaslı ("left") tabloda fo:margin-left ("margins" kenarlara yayılır).
    const indent = align === "left" && tStyle?.table.marginLeft ? Math.round(ptPx(tStyle.table.marginLeft)) : 0;
    return {
      type: "table",
      attrs: { tableStyle: anyBorder ? "grid" : "none", headerRow: true, repeatHeader: header, align: align === "center" ? "center" : align === "right" ? "right" : null, ...(indent ? { indent } : {}) },
      content,
    };
  }

  private toc(el: Element): JSONContent {
    const body = kid(el, "index-body");
    const entries: TocEntry[] = [];
    let title = "";
    const titleEl = kid(body, "index-title");
    if (titleEl) title = (titleEl.textContent ?? "").trim();
    for (const p of body ? descendants(body, "p") : []) {
      if (titleEl?.contains(p)) continue;
      const chain = this.st.chain(attr(p, "style-name"), "paragraph", this.scope);
      const lv = chain.map((s) => /^contents\s*(\d+)$/.exec(lower(s))).find(Boolean);
      // Metin ve sayfa numarası son sekmeyle ayrılır.
      const parts: string[] = [""];
      const walk = (n: Node) => {
        for (let c = n.firstChild; c; c = c.nextSibling) {
          if (c.nodeType === 3) parts[parts.length - 1] += c.nodeValue ?? "";
          else if ((c as Element).localName === "tab") parts.push("");
          else if ((c as Element).localName === "s") parts[parts.length - 1] += " ";
          else walk(c);
        }
      };
      walk(p);
      const page = parseInt(parts.length > 1 ? parts[parts.length - 1] : "", 10);
      const label = (parts.length > 1 ? parts.slice(0, -1).join(" ") : parts[0]).replace(/\s+/g, " ").trim();
      if (label) entries.push({ level: lv ? Math.min(9, Number(lv[1])) : 1, text: label, page: Number.isFinite(page) ? page : 0, pos: 0 });
    }
    return { type: "toc", attrs: { entries, title: title || "İçindekiler" } };
  }

  /* -------------------------- sayfa, üst/alt bilgi, meta -------------------------- */

  private firstMaster(text: Element): string | null {
    for (const p of descendants(text, "*").filter((e) => e.localName === "p" || e.localName === "h").slice(0, 50)) {
      const m = this.st.chain(attr(p, "style-name"), "paragraph", "content").find((s) => s.masterPage)?.masterPage;
      if (m) return m;
    }
    return null;
  }

  private page(master: string, def: PageSetup): PageSetup {
    const mp = this.st.masters.get(master);
    const lay: PageLayout | undefined = mp?.layout ? this.st.layouts.get(mp.layout) : undefined;
    if (!lay) return { ...def, margin: { ...def.margin } };
    const p: PageSetup = { ...def, margin: { ...def.margin } };
    if (lay.width && lay.height) {
      p.width = ptTw(lay.width);
      p.height = ptTw(lay.height);
    }
    p.orientation = lay.landscape || p.width > p.height ? "landscape" : "portrait";
    const m = lay.margin;
    if (m.left !== undefined) p.margin.left = ptTw(m.left);
    if (m.right !== undefined) p.margin.right = ptTw(m.right);
    p.margin.gutter = 0;
    // ODF'de üst kenar boşluğu üst bilgiye kadardır; gövde üst bilgi yüksekliği + aralık kadar aşağıdadır.
    const top = m.top ?? 70.85;
    const bottom = m.bottom ?? 70.85;
    if (lay.header && mp?.header) {
      p.margin.header = ptTw(top);
      p.margin.top = ptTw(top + lay.header.minHeight + lay.header.spacing);
    } else {
      p.margin.top = ptTw(top);
      p.margin.header = Math.min(p.margin.top, def.margin.header);
    }
    if (lay.footer && mp?.footer) {
      p.margin.footer = ptTw(bottom);
      p.margin.bottom = ptTw(bottom + lay.footer.minHeight + lay.footer.spacing);
    } else {
      p.margin.bottom = ptTw(bottom);
      p.margin.footer = Math.min(p.margin.bottom, def.margin.footer);
    }
    p.columns = lay.columns ?? 1;
    if (lay.columnGap) p.columnSpace = ptTw(lay.columnGap);
    return p;
  }

  private readHeaders(master: string, settings: BelgeDocument["settings"]) {
    const mp = this.st.masters.get(master);
    if (!mp) return;
    this.scope = "styles";
    const read = (el: Element | null): JSONContent | null => {
      if (!el || attr(el, "display") === "false") return null;
      const blocks = foldLists(this.blocks(el)).filter((b) => !b.type?.startsWith("__"));
      return JSON.stringify(blocks).match(/"text":"[^"]*\S|pageField|"image"/) ? { type: "doc", content: blocks } : null;
    };
    settings.header = read(mp.header);
    settings.footer = read(mp.footer);
    settings.headerFirst = read(mp.headerFirst);
    settings.footerFirst = read(mp.footerFirst);
    if (settings.headerFirst || settings.footerFirst) settings.page.titlePage = true;
    settings.headerEven = read(mp.headerLeft);
    settings.footerEven = read(mp.footerLeft);
    if (settings.headerEven || settings.footerEven) settings.oddEven = true;
    this.scope = "content";
  }

  private readMeta(x: Element, settings: BelgeDocument["settings"]) {
    try {
      const get = (n: string) => descendants(x, n)[0]?.textContent?.trim() ?? "";
      const p = settings.properties;
      p.title = get("title");
      p.subject = get("subject");
      p.keywords = descendants(x, "keyword")
        .map((k) => k.textContent?.trim())
        .filter(Boolean)
        .join("; ");
      p.author = get("initial-creator") || get("creator");
      p.created = get("creation-date") || p.created;
      p.modified = get("date") || p.modified;
    } catch {
      // Bozuk meta.xml yok sayılır.
    }
  }
}

const round = (n: number) => Math.round(n * 100) / 100;

function lengthPx(v: string | null): number | null {
  const m = /^\s*(-?[\d.]+)\s*(cm|mm|in|pt|px|pc)?\s*$/i.exec(v ?? "");
  if (!m) return null;
  const n = parseFloat(m[1]);
  const k: Record<string, number> = { cm: 96 / 2.54, mm: 96 / 25.4, in: 96, pt: 4 / 3, px: 1, pc: 16 };
  return Math.round(n * (k[(m[2] ?? "px").toLowerCase()] ?? 1));
}

export { decodeName };

/** OpenDocument'ın metin belgesi olmayan türü için anlaşılır ileti. */
function notTextMessage(mime: string): string {
  if (/spreadsheet/.test(mime)) return "Bu bir OpenDocument hesap tablosu; OfficeHere Tablo ile açın.";
  if (/presentation/.test(mime)) return "Bu bir OpenDocument sunusu; OfficeHere Sunum ile açın.";
  return "Bu OpenDocument dosyası bir metin belgesi değil.";
}

/** ODT baytları → belge. */
export async function readOdt(data: ArrayBuffer | Uint8Array, warnings?: LoadWarnings): Promise<BelgeDocument> {
  const r = new OdtReader();
  if (warnings) r.warnings = warnings;
  return r.read(data);
}

/** Düz XML OpenDocument (.fodt) metni → belge. */
export function readFodt(xml: string, warnings?: LoadWarnings): BelgeDocument {
  const r = new OdtReader();
  if (warnings) r.warnings = warnings;
  return r.readFlat(xml);
}
