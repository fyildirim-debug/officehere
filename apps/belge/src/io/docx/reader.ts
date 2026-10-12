// DOCX okuyucu: OOXML (JSZip + DOMParser) → ProseMirror JSON + belge ayarları.
import JSZip from "jszip";
import type { JSONContent } from "@tiptap/core";
import { defaultSettings, normalizeDoc, pageMetrics, samePage, type BelgeDocument, type CommentData, type DocSettings, type PageSetup } from "../../model/document";
import { noteText } from "../../editor/notes";
import { effectiveProps, resolveStyles, type StyleDef } from "../../model/styles";
import { attr, attrNum, descendants, kid, kids, parseXml, relsTarget, NS } from "./xml";
import { Numbering, StyleSheet, parseTheme, type ParaProps, type RunProps, type ThemeFonts } from "./readerStyles";
import { PartReader, stripNoteStyle, type Rel } from "./partReader";
import { readCustomXml } from "./customXml";
import type { HfKey } from "../../editor/sectionHf";

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
  caption: "Caption",
  bibliography: "Bibliography",
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


export class DocxReader {
  sheet!: StyleSheet;
  numbering!: Numbering;
  theme: ThemeFonts = { major: "Aptos Display", minor: "Aptos" };
  media = new Map<string, string>();
  /** Grafik parçaları (yol → XML). */
  charts = new Map<string, Document>();
  footnotes = new Map<string, { text: string; content: JSONContent[] | null }>();
  endnotes = new Map<string, { text: string; content: JSONContent[] | null }>();
  ours: Record<string, StyleDef> = resolveStyles("default", "OfficeHere");
  /** Belge (son bölüm) metin alanı genişliği ve sol kenar boşluğu (px; kaydırmalı resim konumu için). */
  contentWPx = 605;
  marginLPx = 94;

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

  /** sectPr → sayfa yapısı. */
  pageOf(sect: Element, def: PageSetup): PageSetup {
    const page: PageSetup = { ...def, margin: { ...def.margin } };
    const pgSz = kid(sect, "pgSz");
    const pgMar = kid(sect, "pgMar");
    const w = attrNum(pgSz, "w");
    const h = attrNum(pgSz, "h");
    if (w && h) {
      page.width = w;
      page.height = h;
    }
    page.orientation = attr(pgSz, "orient") === "landscape" || (w && h && w > h) ? "landscape" : "portrait";
    const m = page.margin;
    const num = (n: string, d: number) => Math.abs(attrNum(pgMar, n) ?? d);
    page.margin = { top: num("top", m.top), bottom: num("bottom", m.bottom), left: num("left", m.left), right: num("right", m.right), header: num("header", m.header), footer: num("footer", m.footer), gutter: num("gutter", 0) };
    const cols = kid(sect, "cols");
    page.columns = Math.max(1, Math.min(13, attrNum(cols, "num") ?? 1));
    page.columnSpace = attrNum(cols, "space") ?? 708;
    page.columnLine = page.columns > 1 && (attr(cols, "sep") === "1" || attr(cols, "sep") === "true");
    const ln = kid(sect, "lnNumType");
    const countBy = attrNum(ln, "countBy");
    if (ln && countBy && countBy > 0) {
      const restart = attr(ln, "restart");
      const start = attrNum(ln, "start");
      const distance = attrNum(ln, "distance");
      page.lineNumbers = { restart: restart === "continuous" || restart === "newSection" ? restart : "newPage", countBy, ...(start ? { start: start + 1 } : {}), ...(distance ? { distance } : {}) };
    } else delete page.lineNumbers;
    const tp = kid(sect, "titlePg");
    page.titlePage = !!tp && attr(tp, "val") !== "0" && attr(tp, "val") !== "false";
    if (!page.columnLine) delete page.columnLine;
    if (!page.titlePage) delete page.titlePage;
    return page;
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
            if (j.customStyles && typeof j.customStyles === "object") settings.customStyles = j.customStyles;
            if (Array.isArray(j.sources)) settings.sources = j.sources;
            if (j.bibStyle === "apa" || j.bibStyle === "ieee") settings.bibStyle = j.bibStyle;
          } catch {
            // yok say
          }
        }
      }
    }
    // Özel XML: kaynakça kaynakları (OfficeHere ayarı yoksa) ve birleştirme veri kaynağı.
    const cx = await readCustomXml(zip);
    if (cx.sources && !settings.sources?.length) {
      settings.sources = cx.sources;
      if (cx.bibStyle && !settings.bibStyle) settings.bibStyle = cx.bibStyle;
    }
    if (cx.mailMerge) settings.mailMerge = cx.mailMerge;
    this.ours = resolveStyles(settings.styleSet, settings.theme, settings.fontTheme, settings.paragraphSpacing, settings.customStyles);
    this.theme = parseTheme(await loadXml(relOf("/theme")));
    this.sheet = new StyleSheet(await loadXml(relOf("/styles")), this.theme);
    this.numbering = new Numbering(await loadXml(relOf("/numbering")));
    // Medya
    const mediaFiles = Object.keys(zip.files).filter((n) => !zip.files[n].dir).filter((n) => /\/media\//.test(n) || /^word\/[^/]+\.(png|jpe?g|gif|bmp)$/i.test(n));
    for (const name of mediaFiles) {
      const ext = name.slice(name.lastIndexOf(".") + 1).toLowerCase();
      const bytes = await zip.file(name)!.async("uint8array");
      this.media.set(name, `data:${MIME[ext] ?? "application/octet-stream"};base64,${toBase64(bytes)}`);
    }

    for (const name of Object.keys(zip.files).filter((n) => /\/charts\/chart\d+\.xml$/i.test(n))) {
      try {
        this.charts.set(name, parseXml(await zip.file(name)!.async("string")));
      } catch {
        // bozuk grafik parçası atlanır
      }
    }

    // Dipnotlar ve son notlar (zengin içerik)
    for (const [suffix, map] of [["/footnotes", this.footnotes], ["/endnotes", this.endnotes]] as const) {
      const target = relOf(suffix);
      const x = await loadXml(target);
      if (!x || !target) continue;
      const nr = new PartReader(this, await readRels(zip, target));
      nr.noteMode = true;
      for (const n of kids(x.documentElement)) {
        if (n.localName !== "footnote" && n.localName !== "endnote") continue;
        const type = attr(n, "type");
        if (type && type !== "normal") continue;
        const blocks = nr.blocks(n).map(stripNoteStyle);
        // İlk paragrafın başındaki boşluk (başvuru işaretinden sonra) atılır.
        const first = blocks[0];
        if (first?.content?.[0]?.type === "text") {
          first.content[0] = { ...first.content[0], text: (first.content[0].text ?? "").replace(/^\s+/, "") };
          if (!first.content[0].text) first.content.shift();
          if (!first.content.length) delete first.content;
        }
        const text = noteText(blocks);
        const plain = blocks.every((b) => b.type === "paragraph" && (b.content ?? []).every((c) => c.type === "text" && !c.marks?.length));
        map.set(attr(n, "id") ?? "", { text, content: plain ? null : blocks });
      }
    }

    const docXml = parseXml(await mainFile.async("string"));
    const docEl = docXml.documentElement;
    const body = kid(docEl, "body");
    if (!body) throw new Error("Belge gövdesi bulunamadı.");
    const reader = new PartReader(this, docRels);
    const bodySect = kid(body, "sectPr");
    if (bodySect) {
      const m = pageMetrics(this.pageOf(bodySect, settings.page));
      this.contentWPx = m.contentW;
      this.marginLPx = m.ml;
    }
    const parts = reader.sections(body);
    const pages = parts.map((p) => (p.sectPr ? this.pageOf(p.sectPr, settings.page) : { ...settings.page }));
    const last = pages[pages.length - 1];
    settings.page = last;
    const sections: JSONContent[] = parts.map((p, i) => {
      // Bölümün başlama türü kendi sectPr'ındadır.
      const t = attr(kid(p.sectPr, "type"), "val");
      const kind = t === "continuous" || t === "evenPage" || t === "oddPage" ? t : "nextPage";
      const page = samePage(pages[i], last) ? null : pages[i];
      return { type: "section", attrs: { kind, page }, content: p.blocks.length ? p.blocks : [{ type: "paragraph" }] };
    });
    // Sayfa kenarlığı ve üst/alt bilgiler (ilk bölümden; sonraki bölümler devralır).
    const sects = parts.map((p) => p.sectPr).filter((x): x is Element => !!x);
    const firstSect = sects[0] ?? null;
    if (firstSect) {
      const pb = kid(firstSect, "pgBorders");
      const top = kid(pb, "top");
      if (top && attr(top, "val") !== "none") {
        const val = attr(top, "val") ?? "single";
        settings.pageBorder = {
          style: val === "double" ? "double" : val === "dotted" ? "dotted" : val.startsWith("dash") ? "dashed" : val === "thick" ? "thick" : "single",
          color: /^[0-9a-f]{6}$/i.test(attr(top, "color") ?? "") ? `#${attr(top, "color")!.toUpperCase()}` : "#000000",
          width: (attrNum(top, "sz") ?? 8) / 8,
        };
      }
    }
    // Üst/alt bilgiler: ilk bölümünküler belge ayarlarına; sonraki bölümde farklı parçaya başvuru varsa
    // bölümün kendi içeriği (Öncekine Bağla kapalı), yoksa öncekine bağlı.
    const partCache = new Map<string, JSONContent | null>();
    const readPart = async (target: string): Promise<JSONContent | null> => {
      if (partCache.has(target)) return partCache.get(target)!;
      const xml = await loadXml(target);
      let out: JSONContent | null = null;
      if (xml) {
        const pr = new PartReader(this, await readRels(zip, target));
        const content = pr.blocks(xml.documentElement);
        if (pr.watermark) settings.watermark = pr.watermark;
        const hasText = JSON.stringify(content).match(/"text":"[^"]*\S|pageField|"image"/);
        out = hasText ? ({ type: "doc", content } as JSONContent) : null;
      }
      partCache.set(target, out);
      return out;
    };
    const refTarget = (sp: Element | null, tag: string, type: string) => {
      const ref = sp ? kids(sp, tag).find((r) => (attr(r, "type") ?? "default") === type) : undefined;
      const rid = ref ? (ref.getAttributeNS(NS.r, "id") ?? attr(ref, "id")) : null;
      return rid ? docRels.get(rid)?.target : undefined;
    };
    const HF: [HfKey, "headerReference" | "footerReference", string][] = [
      ["header", "headerReference", "default"],
      ["footer", "footerReference", "default"],
      ["headerFirst", "headerReference", "first"],
      ["footerFirst", "footerReference", "first"],
      ["headerEven", "headerReference", "even"],
      ["footerEven", "footerReference", "even"],
    ];
    for (const [key, tag, type] of HF) {
      let eff: string | undefined;
      for (let i = 0; i < parts.length; i++) {
        const t = refTarget(parts[i].sectPr, tag, type);
        if (i === 0) {
          eff = t;
          settings[key] = t ? await readPart(t) : null;
          continue;
        }
        if (!t || t === eff) continue;
        eff = t;
        const sec = sections[i];
        const sid = `s${i}`;
        sec.attrs = { ...sec.attrs, sid };
        settings.sectionHf = { ...settings.sectionHf, [sid]: { ...settings.sectionHf?.[sid], [key]: await readPart(t) } };
      }
    }
    const settingsXml = await loadXml(relOf("/settings"));
    const prot = settingsXml ? descendants(settingsXml.documentElement, "documentProtection")[0] : undefined;
    const edit = attr(prot, "edit");
    if (prot && (edit === "readOnly" || edit === "comments" || edit === "trackedChanges")) {
      const enf = attr(prot, "enforcement");
      settings.protection = {
        mode: edit,
        enforced: enf === "1" || enf === "true" || enf === "on",
        ...(attr(prot, "hashValue") ? { algorithm: attr(prot, "algorithmName") ?? "SHA-512", hash: attr(prot, "hashValue")!, salt: attr(prot, "saltValue") ?? "", spinCount: attrNum(prot, "spinCount") ?? 100000 } : {}),
      };
    }
    if (settingsXml && descendants(settingsXml.documentElement, "evenAndOddHeaders").some((e) => attr(e, "val") !== "0" && attr(e, "val") !== "false")) settings.oddEven = true;
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

    return { content: normalizeDoc({ type: "doc", content: sections }), settings };
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
