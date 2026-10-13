// ODT yazıcı stil defteri: otomatik stillerin tekilleştirilmesi, karakter/paragraf özelliklerinin ODF
// karşılıkları, adlandırılmış stiller (styles.xml), liste stilleri, sayfa düzeni ve asıl sayfa.
import type { JSONContent } from "@tiptap/core";
import type { DocSettings, PageSetup } from "../../model/document";
import type { StyleDef } from "../../model/styles";
import { styleChar, type CharProps } from "../common";
import { esc } from "../docx/xml";
import { pictureToOdf } from "./dataStyles";

export const ODF_NS =
  'xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" ' +
  'xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" ' +
  'xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0" ' +
  'xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:meta="urn:oasis:names:tc:opendocument:xmlns:meta:1.0" ' +
  'xmlns:number="urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0" xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0" ' +
  'xmlns:loext="urn:org:documentfoundation:names:experimental:office:xmlns:loext:1.0"';

export const XML_DECL = '<?xml version="1.0" encoding="UTF-8"?>\n';

/** Bizim stil kimliğimiz → ODF adı ve görünen adı. */
export const ODF_PARA: Record<string, [string, string]> = {
  Normal: ["Standard", "Standard"],
  Heading1: ["Heading_20_1", "Heading 1"],
  Heading2: ["Heading_20_2", "Heading 2"],
  Heading3: ["Heading_20_3", "Heading 3"],
  Heading4: ["Heading_20_4", "Heading 4"],
  Heading5: ["Heading_20_5", "Heading 5"],
  Heading6: ["Heading_20_6", "Heading 6"],
  Title: ["Title", "Title"],
  Subtitle: ["Subtitle", "Subtitle"],
  Quote: ["Quotations", "Quotations"],
  IntenseQuote: ["Intense_20_Quote", "Intense Quote"],
  NoSpacing: ["No_20_Spacing", "No Spacing"],
  ListParagraph: ["List_20_Paragraph", "List Paragraph"],
  Caption: ["Caption", "Caption"],
};

export const ODF_CHAR: Record<string, [string, string]> = {
  Emphasis: ["Emphasis", "Emphasis"],
  Strong: ["Strong_20_Emphasis", "Strong Emphasis"],
  SubtleEmphasis: ["Subtle_20_Emphasis", "Subtle Emphasis"],
  IntenseEmphasis: ["Intense_20_Emphasis", "Intense Emphasis"],
  SubtleReference: ["Subtle_20_Reference", "Subtle Reference"],
  IntenseReference: ["Intense_20_Reference", "Intense Reference"],
  BookTitle: ["Book_20_Title", "Book Title"],
};

const cm = (pt: number) => `${+((pt / 72) * 2.54).toFixed(4)}cm`;
const twCm = (tw: number) => `${+((tw / 1440) * 2.54).toFixed(4)}cm`;
const attrs = (o: Record<string, string | undefined>) =>
  Object.entries(o)
    .filter(([, v]) => v !== undefined)
    .map(([k, v]) => ` ${k}="${esc(v!)}"`)
    .join("");

/** Karakter özellikleri → style:text-properties öznitelikleri. */
export function textPropsAttrs(p: CharProps, fonts: Set<string>): Record<string, string | undefined> {
  const o: Record<string, string | undefined> = {};
  if (p.font) {
    fonts.add(p.font);
    o["style:font-name"] = p.font;
  }
  if (p.size) o["fo:font-size"] = `${+p.size.toFixed(2)}pt`;
  if (p.bold !== undefined) o["fo:font-weight"] = p.bold ? "bold" : "normal";
  if (p.italic !== undefined) o["fo:font-style"] = p.italic ? "italic" : "normal";
  if (p.underline) {
    const u = p.underline;
    o["style:text-underline-style"] = u === "dotted" ? "dotted" : u === "dashed" ? "dash" : u === "wave" ? "wave" : "solid";
    o["style:text-underline-width"] = u === "thick" ? "bold" : "auto";
    o["style:text-underline-color"] = p.underlineColor ?? "font-color";
    if (u === "double") o["style:text-underline-type"] = "double";
    if (u === "words") o["style:text-underline-mode"] = "skip-white-space";
  }
  if (p.strike || p.dstrike) {
    o["style:text-line-through-style"] = "solid";
    if (p.dstrike) o["style:text-line-through-type"] = "double";
  }
  if (p.sup) o["style:text-position"] = "super 58%";
  else if (p.sub) o["style:text-position"] = "sub 58%";
  if (p.color) o["fo:color"] = p.color;
  if (p.highlight || p.bg) o["fo:background-color"] = (p.highlight ?? p.bg)!;
  if (p.caps) o["fo:text-transform"] = "uppercase";
  if (p.smallCaps) o["fo:font-variant"] = "small-caps";
  if (p.hidden) o["text:display"] = "none";
  if (p.spacing) o["fo:letter-spacing"] = cm(p.spacing);
  if (p.shadow) o["fo:text-shadow"] = "1pt 1pt";
  if (p.outline) o["style:text-outline"] = "true";
  if (p.emboss) o["style:font-relief"] = "embossed";
  else if (p.imprint) o["style:font-relief"] = "engraved";
  return o;
}

/** Paragraf öznitelikleri → style:paragraph-properties öznitelikleri. */
export function paraPropsAttrs(a: Record<string, unknown>, brk: "page" | "column" | null): Record<string, string | undefined> {
  const o: Record<string, string | undefined> = {};
  const al = a.textAlign as string | null;
  if (al) o["fo:text-align"] = al === "left" ? "start" : al === "right" ? "end" : al;
  if (a.indentLeft != null) o["fo:margin-left"] = cm(a.indentLeft as number);
  if (a.indentRight != null) o["fo:margin-right"] = cm(a.indentRight as number);
  if (a.indentFirst != null) o["fo:text-indent"] = cm(a.indentFirst as number);
  if (a.spaceBefore != null) o["fo:margin-top"] = cm(a.spaceBefore as number);
  if (a.spaceAfter != null) o["fo:margin-bottom"] = cm(a.spaceAfter as number);
  if (a.lineHeight != null) {
    const rule = (a.lineRule as string | null) ?? "auto";
    if (rule === "auto") o["fo:line-height"] = `${Math.round((a.lineHeight as number) * 100)}%`;
    else if (rule === "exact") o["fo:line-height"] = `${a.lineHeight}pt`;
    else o["style:line-height-at-least"] = `${a.lineHeight}pt`;
  }
  if (a.shading) o["fo:background-color"] = a.shading as string;
  if (a.keepNext) o["fo:keep-with-next"] = "always";
  if (a.keepLines) o["fo:keep-together"] = "always";
  if (brk || a.pageBreakBefore) o["fo:break-before"] = brk ?? "page";
  const b = a.borders as Record<string, { style: string; size: number; color: string } | undefined> | null;
  if (b) {
    for (const side of ["top", "bottom", "left", "right"] as const) {
      const s = b[side];
      if (s) o[`fo:border-${side}`] = `${s.size}pt ${s.style === "double" ? "double" : s.style === "dotted" ? "dotted" : s.style === "dashed" ? "dashed" : "solid"} ${s.color}`;
    }
  }
  return o;
}

export interface ListLevelDef {
  ordered: boolean;
  style: string | null;
  start: number;
  multilevel: string | null;
}

const BULLETS: Record<string, string> = { disc: "●", circle: "○", square: "■", diamond: "❖", arrow: "➢", check: "✓", dash: "–" };
const LEVEL_BULLETS = ["●", "○", "■"];

/** Otomatik stil defteri (content.xml ve styles.xml için ayrı önekli). */
export class StyleBook {
  private keys = new Map<string, string>();
  private xml: string[] = [];
  private counters = new Map<string, number>();

  constructor(
    readonly prefix: string,
    readonly fonts: Set<string>,
  ) {}

  private add(key: string, base: string, make: (name: string) => string): string {
    const hit = this.keys.get(key);
    if (hit) return hit;
    const n = (this.counters.get(base) ?? 0) + 1;
    this.counters.set(base, n);
    const name = `${this.prefix}${base}${n}`;
    this.keys.set(key, name);
    this.xml.push(make(name));
    return name;
  }

  /** Paragraf otomatik stili (özellik yoksa üst stil adı döner). inner: paragraph-properties alt öğeleri (ör. style:drop-cap). */
  para(parent: string, pp: Record<string, string | undefined>, extra: Record<string, string | undefined> = {}, inner = ""): string {
    const p = attrs(pp);
    const e = attrs(extra);
    if (!p && !e && !inner) return parent;
    const props = inner ? `<style:paragraph-properties${p}>${inner}</style:paragraph-properties>` : p ? `<style:paragraph-properties${p}/>` : "";
    return this.add(`P|${parent}|${p}|${e}|${inner}`, "P", (n) => `<style:style style:name="${n}" style:family="paragraph" style:parent-style-name="${parent}"${e}>${props}</style:style>`);
  }

  /** Metin otomatik stili; özellik yoksa adlandırılmış stil ya da null. */
  text(p: CharProps, parent: string | null): string | null {
    const t = attrs(textPropsAttrs(p, this.fonts));
    if (!t) return parent;
    return this.add(`T|${parent}|${t}`, "T", (n) => `<style:style style:name="${n}" style:family="text"${parent ? ` style:parent-style-name="${parent}"` : ""}><style:text-properties${t}/></style:style>`);
  }

  /** Tarih/saat alanının veri stili (biçim resminden). */
  dataStyle(pic: string, time: boolean): string {
    return this.add(`N|${time ? "t" : "d"}|${pic}`, "N", (n) => pictureToOdf(n, pic, time));
  }

  raw(family: string, base: string, inner: string, extra = ""): string {
    return this.add(`${family}|${inner}|${extra}`, base, (n) => `<style:style style:name="${n}" style:family="${family}"${extra}>${inner}</style:style>`);
  }

  list(levels: ListLevelDef[]): string {
    const inner = levels
      .map((l, k) => {
        const pos = `<style:list-level-properties text:list-level-position-and-space-mode="label-alignment"><style:list-level-label-alignment text:label-followed-by="listtab" text:list-tab-stop-position="${twCm(720 * (k + 1))}" fo:text-indent="-0.635cm" fo:margin-left="${twCm(720 * (k + 1))}"/></style:list-level-properties>`;
        if (!l.ordered) return `<text:list-level-style-bullet text:level="${k + 1}" text:bullet-char="${esc(BULLETS[l.style ?? ""] ?? LEVEL_BULLETS[k % 3])}">${pos}</text:list-level-style-bullet>`;
        const fmt = l.style === "upper-roman" ? "I" : l.style === "lower-roman" ? "i" : l.style === "upper-alpha" ? "A" : l.style === "lower-alpha" || l.style === "lower-alpha-paren" ? "a" : "1";
        const suffix = l.style === "decimal-paren" || l.style === "lower-alpha-paren" ? ")" : ".";
        const display = l.multilevel === "outline" ? ` text:display-levels="${k + 1}"` : "";
        return `<text:list-level-style-number text:level="${k + 1}" style:num-suffix="${suffix}" style:num-format="${fmt}"${l.start !== 1 ? ` text:start-value="${l.start}"` : ""}${display}>${pos}</text:list-level-style-number>`;
      })
      .join("");
    return this.add(`L|${inner}`, "L", (n) => `<text:list-style style:name="${n}">${inner}</text:list-style>`);
  }

  toXml(): string {
    return this.xml.join("");
  }
}

/** Font bildirimleri. */
export function fontDecls(fonts: Set<string>): string {
  return `<office:font-face-decls>${[...fonts].map((f) => `<style:font-face style:name="${esc(f)}" svg:font-family="${esc(/\s/.test(f) ? `'${f}'` : f)}"/>`).join("")}</office:font-face-decls>`;
}

/** Bizim stilimizin ODF adlandırılmış stili (karakter + paragraf özellikleri). */
function namedPara(styles: Record<string, StyleDef>, id: string, fonts: Set<string>, parent: string | null, extra = ""): string {
  const [name, display] = ODF_PARA[id];
  const c = styleChar(styles, id);
  const sp = c.base;
  const tp = attrs(textPropsAttrs({ font: c.font, size: c.size, bold: c.bold || undefined, italic: c.italic || undefined, color: c.color, caps: c.caps || undefined }, fonts));
  const pp = attrs(paraPropsAttrs({ spaceBefore: sp.spaceBefore ?? 0, spaceAfter: sp.spaceAfter ?? 0, lineHeight: sp.line ?? null, textAlign: sp.align ?? null, keepNext: sp.keepNext, indentLeft: sp.indentLeft ?? null }, null));
  return `<style:style style:name="${name}" style:display-name="${display}" style:family="paragraph"${parent ? ` style:parent-style-name="${parent}"` : ""} style:class="text"${extra}><style:paragraph-properties${pp}/><style:text-properties${tp}/></style:style>`;
}

/** styles.xml'in office:styles bölümü. */
export function namedStyles(styles: Record<string, StyleDef>, fonts: Set<string>): string {
  const normal = styleChar(styles, "Normal");
  fonts.add(normal.font ?? "Aptos");
  let x = `<style:default-style style:family="paragraph"><style:paragraph-properties style:writing-mode="lr-tb"/><style:text-properties style:font-name="${esc(normal.font ?? "Aptos")}" fo:font-size="${normal.size}pt" fo:language="tr" fo:country="TR"/></style:default-style>`;
  x += `<style:default-style style:family="graphic"><style:graphic-properties svg:stroke-color="#3465a4" draw:fill-color="#729fcf"/></style:default-style>`;
  x += `<style:default-style style:family="table"><style:table-properties table:border-model="collapsing"/></style:default-style>`;
  x += namedPara(styles, "Normal", fonts, null);
  for (let i = 1; i <= 6; i++) x += namedPara(styles, `Heading${i}`, fonts, "Standard", ` style:next-style-name="Standard" style:default-outline-level="${i}"`);
  for (const id of ["Title", "Subtitle", "Quote", "IntenseQuote", "NoSpacing", "ListParagraph", "Caption"]) x += namedPara(styles, id, fonts, "Standard", ' style:next-style-name="Standard"');
  const small = (name: string, display: string, extra: string) => `<style:style style:name="${name}" style:display-name="${display}" style:family="paragraph" style:parent-style-name="Standard" style:class="extra">${extra}</style:style>`;
  x += small("Footnote", "Footnote", '<style:paragraph-properties fo:margin-left="0.6cm" fo:margin-top="0cm" fo:margin-bottom="0cm" fo:text-indent="-0.6cm"/><style:text-properties fo:font-size="10pt"/>');
  x += small("Endnote", "Endnote", '<style:paragraph-properties fo:margin-left="0.6cm" fo:margin-top="0cm" fo:margin-bottom="0cm" fo:text-indent="-0.6cm"/><style:text-properties fo:font-size="10pt"/>');
  x += small("Header", "Header", '<style:paragraph-properties fo:margin-top="0cm" fo:margin-bottom="0cm"/>');
  x += small("Footer", "Footer", '<style:paragraph-properties fo:margin-top="0cm" fo:margin-bottom="0cm"/>');
  x += small("Table_20_Contents", "Table Contents", '<style:paragraph-properties fo:margin-top="0cm" fo:margin-bottom="0cm"/>');
  x += small("Horizontal_20_Line", "Horizontal Line", '<style:paragraph-properties fo:margin-top="0cm" fo:margin-bottom="0.25cm" fo:border-bottom="0.75pt solid #808080" fo:padding="0cm"/><style:text-properties fo:font-size="6pt"/>');
  x += small("Contents_20_Heading", "Contents Heading", '<style:paragraph-properties fo:margin-top="0.4cm" fo:margin-bottom="0.2cm"/><style:text-properties fo:font-size="16pt" fo:font-weight="bold"/>');
  for (let i = 1; i <= 9; i++) x += small(`Contents_20_${i}`, `Contents ${i}`, `<style:paragraph-properties fo:margin-left="${+((i - 1) * 0.5).toFixed(2)}cm" fo:margin-top="0cm" fo:margin-bottom="0.1cm"><style:tab-stops><style:tab-stop style:position="16cm" style:type="right" style:leader-style="dotted" style:leader-text="."/></style:tab-stops></style:paragraph-properties>`);
  for (const id of Object.keys(ODF_CHAR)) {
    const [name, display] = ODF_CHAR[id];
    const p = styles[id]?.props ?? {};
    const tp = attrs(textPropsAttrs({ bold: p.bold || undefined, italic: p.italic || undefined, color: p.color ?? null, caps: p.caps || undefined, spacing: p.spacing ?? null, underline: p.underline ? "single" : null }, fonts));
    x += `<style:style style:name="${name}" style:display-name="${display}" style:family="text"><style:text-properties${tp}/></style:style>`;
  }
  x += '<style:style style:name="Internet_20_link" style:display-name="Internet link" style:family="text"><style:text-properties fo:color="#0563C1" style:text-underline-style="solid" style:text-underline-width="auto" style:text-underline-color="font-color"/></style:style>';
  x += '<style:style style:name="Visited_20_Internet_20_Link" style:display-name="Visited Internet Link" style:family="text"><style:text-properties fo:color="#954F72" style:text-underline-style="solid" style:text-underline-width="auto" style:text-underline-color="font-color"/></style:style>';
  x += '<style:style style:name="Footnote_20_anchor" style:display-name="Footnote anchor" style:family="text"><style:text-properties style:text-position="super 58%"/></style:style>';
  x += '<style:style style:name="Graphics" style:family="graphic"><style:graphic-properties text:anchor-type="as-char" svg:y="0cm" style:vertical-pos="top" style:vertical-rel="baseline"/></style:style>';
  x += '<text:outline-style style:name="Outline">';
  for (let i = 1; i <= 10; i++) x += `<text:outline-level-style text:level="${i}" style:num-format=""><style:list-level-properties text:list-level-position-and-space-mode="label-alignment"><style:list-level-label-alignment text:label-followed-by="listtab"/></style:list-level-properties></text:outline-level-style>`;
  x += "</text:outline-style>";
  x += '<text:notes-configuration text:note-class="footnote" text:citation-style-name="Footnote_20_anchor" style:num-format="1" text:start-value="0" text:footnotes-position="page" text:start-numbering-at="document"/>';
  x += '<text:notes-configuration text:note-class="endnote" text:citation-style-name="Footnote_20_anchor" style:num-format="i" text:start-value="0"/>';
  return `<office:styles>${x}</office:styles>`;
}

/** Sayfa düzeni: ODF'de üst kenar boşluğu üst bilgiye kadardır; gövde üst bilgi yüksekliği + aralıkla başlar. */
export function pageLayout(name: string, p: PageSetup, hasHeader: boolean, hasFooter: boolean, pageColor: string | null = null): string {
  const m = p.margin;
  const hf = (has: boolean, edge: number, body: number, key: "bottom" | "top") => {
    if (!has) return "";
    const gap = Math.max(0, body - edge);
    const spacing = Math.min(gap, 283);
    return `<style:header-footer-properties fo:min-height="${twCm(Math.max(0, gap - spacing))}" fo:margin-${key}="${twCm(spacing)}" fo:margin-left="0cm" fo:margin-right="0cm" style:dynamic-spacing="false"/>`;
  };
  const top = hasHeader ? m.header : m.top;
  const bottom = hasFooter ? m.footer : m.bottom;
  const cols = p.columns > 1 ? `<style:columns fo:column-count="${p.columns}" fo:column-gap="${twCm(p.columnSpace)}"/>` : "";
  // Sayfa rengi (Tasarım › Sayfa Rengi): ODF 1.2 özniteliği ve LibreOffice'in dolgu biçimi birlikte.
  const bg = pageColor ? ` fo:background-color="${esc(pageColor)}" draw:fill="solid" draw:fill-color="${esc(pageColor)}"` : "";
  return (
    `<style:page-layout style:name="${name}"><style:page-layout-properties fo:page-width="${twCm(p.width)}" fo:page-height="${twCm(p.height)}" style:print-orientation="${p.orientation}" ` +
    `fo:margin-top="${twCm(top)}" fo:margin-bottom="${twCm(bottom)}" fo:margin-left="${twCm(m.left + m.gutter)}" fo:margin-right="${twCm(m.right)}"${bg} style:writing-mode="lr-tb">${cols}</style:page-layout-properties>` +
    `<style:header-style>${hf(hasHeader, m.header, m.top, "bottom")}</style:header-style><style:footer-style>${hf(hasFooter, m.footer, m.bottom, "top")}</style:footer-style></style:page-layout>`
  );
}

/** meta.xml */
export function metaXml(s: DocSettings): string {
  const p = s.properties;
  const t = (tag: string, v: string) => (v ? `<${tag}>${esc(v)}</${tag}>` : "");
  const date = (v: string) => (v && Number.isFinite(new Date(v).getTime()) ? new Date(v).toISOString().replace(/\.\d+Z$/, "") : "");
  const keywords = p.keywords
    .split(/[,;]/)
    .map((k) => k.trim())
    .filter(Boolean)
    .map((k) => `<meta:keyword>${esc(k)}</meta:keyword>`)
    .join("");
  return `${XML_DECL}<office:document-meta ${ODF_NS} office:version="1.3"><office:meta><meta:generator>OfficeHere Belge</meta:generator>${t("dc:title", p.title)}${t("dc:subject", p.subject)}${keywords}${t("meta:initial-creator", p.author)}${t("dc:creator", p.author)}${t("meta:creation-date", date(p.created))}${t("dc:date", date(p.modified))}<dc:language>tr-TR</dc:language></office:meta></office:document-meta>`;
}

/** Üst/alt bilgi içeriği olan belge mi? */
export const hasContent = (d: JSONContent | null | undefined) => !!d?.content?.length && JSON.stringify(d.content).match(/"text":"[^"]*\S|pageField|"image"/) !== null;
