// styles.xml üretimi: belge stilleri, tablo stilleri ve özellik dönüşümleri.
import type { BorderSpec, StyleDef, StyleProps } from "../../model/styles";
import { STYLE_ORDER } from "../../model/styles";
import { TABLE_STYLES } from "../../editor/table";
import { NS, XML_HEAD, esc, hex6 } from "./xml";

/** Yerleşik İngilizce stil adları (Word yerelleştirir). */
export const BUILTIN_NAMES: Record<string, string> = {
  Normal: "Normal",
  NoSpacing: "No Spacing",
  Heading1: "heading 1",
  Heading2: "heading 2",
  Heading3: "heading 3",
  Heading4: "heading 4",
  Heading5: "heading 5",
  Heading6: "heading 6",
  Title: "Title",
  Subtitle: "Subtitle",
  SubtleEmphasis: "Subtle Emphasis",
  Emphasis: "Emphasis",
  IntenseEmphasis: "Intense Emphasis",
  Strong: "Strong",
  Quote: "Quote",
  IntenseQuote: "Intense Quote",
  SubtleReference: "Subtle Reference",
  IntenseReference: "Intense Reference",
  BookTitle: "Book Title",
  ListParagraph: "List Paragraph",
};

export const BORDER_VAL: Record<BorderSpec["style"], string> = { single: "single", double: "double", dotted: "dotted", dashed: "dashed", thick: "thick" };

export function borderXml(tag: string, b: BorderSpec): string {
  const sz = Math.max(2, Math.min(96, Math.round(b.size * 8)));
  return `<w:${tag} w:val="${BORDER_VAL[b.style] ?? "single"}" w:sz="${sz}" w:space="${Math.round(b.space ?? 1)}" w:color="${hex6(b.color) ?? "auto"}"/>`;
}

export function fontsXml(font: string): string {
  const f = esc(font);
  return `<w:rFonts w:ascii="${f}" w:hAnsi="${f}" w:eastAsia="${f}" w:cs="${f}"/>`;
}

/** Stil özelliklerinden pPr içeriği (şema sırasıyla). */
export function stylePPr(p: StyleProps, extra: { outline?: number } = {}): string {
  let x = "";
  if (p.keepNext) x += "<w:keepNext/><w:keepLines/>";
  if (p.borderTop || p.borderBottom) x += `<w:pBdr>${p.borderTop ? borderXml("top", p.borderTop) : ""}${p.borderBottom ? borderXml("bottom", p.borderBottom) : ""}</w:pBdr>`;
  if (p.shading) x += `<w:shd w:val="clear" w:color="auto" w:fill="${hex6(p.shading)}"/>`;
  const sp: string[] = [];
  if (p.spaceBefore !== undefined) sp.push(`w:before="${Math.round(p.spaceBefore * 20)}"`);
  if (p.spaceAfter !== undefined) sp.push(`w:after="${Math.round(p.spaceAfter * 20)}"`);
  if (p.line !== undefined) sp.push(`w:line="${Math.round(p.line * 240)}" w:lineRule="auto"`);
  if (sp.length) x += `<w:spacing ${sp.join(" ")}/>`;
  if (p.indentLeft !== undefined || p.indentRight !== undefined)
    x += `<w:ind${p.indentLeft !== undefined ? ` w:left="${Math.round(p.indentLeft * 20)}"` : ""}${p.indentRight !== undefined ? ` w:right="${Math.round(p.indentRight * 20)}"` : ""}/>`;
  if (p.align) x += `<w:jc w:val="${p.align === "justify" ? "both" : p.align}"/>`;
  if (extra.outline !== undefined) x += `<w:outlineLvl w:val="${extra.outline}"/>`;
  return x ? `<w:pPr>${x}</w:pPr>` : "";
}

export function styleRPr(p: StyleProps): string {
  let x = "";
  if (p.font) x += fontsXml(p.font);
  if (p.bold !== undefined) x += p.bold ? "<w:b/><w:bCs/>" : '<w:b w:val="0"/>';
  if (p.italic !== undefined) x += p.italic ? "<w:i/><w:iCs/>" : '<w:i w:val="0"/>';
  if (p.caps) x += "<w:caps/>";
  if (p.color) x += `<w:color w:val="${hex6(p.color)}"/>`;
  if (p.spacing) x += `<w:spacing w:val="${Math.round(p.spacing * 20)}"/>`;
  if (p.size) x += `<w:sz w:val="${Math.round(p.size * 2)}"/><w:szCs w:val="${Math.round(p.size * 2)}"/>`;
  if (p.underline) x += '<w:u w:val="single"/>';
  return x ? `<w:rPr>${x}</w:rPr>` : "";
}

function tableStyleXml(id: string): string {
  const s = TABLE_STYLES.find((x) => x.id === id)!;
  const sid = id === "grid" ? "TableGrid" : `OH${id.replace(/[^a-z0-9]/gi, "")}`;
  const name = id === "grid" ? "Table Grid" : `OH ${s.label}`;
  const c = s.noBorders ? null : hex6(s.border);
  const b = (tag: string) => (c ? `<w:${tag} w:val="single" w:sz="4" w:space="0" w:color="${c}"/>` : `<w:${tag} w:val="none" w:sz="0" w:space="0" w:color="auto"/>`);
  const borders = s.outerOnly ? `<w:tblBorders>${b("top")}${b("left")}${b("bottom")}${b("right")}</w:tblBorders>` : `<w:tblBorders>${b("top")}${b("left")}${b("bottom")}${b("right")}${b("insideH")}${b("insideV")}</w:tblBorders>`;
  let cond = "";
  if (s.header) cond += `<w:tblStylePr w:type="firstRow"><w:rPr><w:b/><w:bCs/>${s.headerText ? `<w:color w:val="${hex6(s.headerText)}"/>` : ""}</w:rPr><w:tcPr><w:shd w:val="clear" w:color="auto" w:fill="${hex6(s.header)}"/></w:tcPr></w:tblStylePr>`;
  cond += `<w:tblStylePr w:type="lastRow"><w:rPr><w:b/><w:bCs/></w:rPr></w:tblStylePr>`;
  cond += `<w:tblStylePr w:type="firstCol"><w:rPr><w:b/><w:bCs/></w:rPr></w:tblStylePr>`;
  if (s.band) cond += `<w:tblStylePr w:type="band1Horz"><w:tcPr><w:shd w:val="clear" w:color="auto" w:fill="${hex6(s.band)}"/></w:tcPr></w:tblStylePr>`;
  return `<w:style w:type="table" w:styleId="${sid}"><w:name w:val="${esc(name)}"/><w:basedOn w:val="TableNormal"/><w:uiPriority w:val="39"/><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:tblPr><w:tblStyleRowBandSize w:val="1"/><w:tblStyleColBandSize w:val="1"/>${borders}</w:tblPr>${cond}</w:style>`;
}

export function tableStyleId(id: string | null | undefined): string {
  if (!id || id === "grid") return "TableGrid";
  if (id === "none") return "TableNormal";
  return `OH${id.replace(/[^a-z0-9]/gi, "")}`;
}

export function tableStyleFromId(styleId: string | null): string | null {
  if (!styleId) return null;
  if (styleId === "TableGrid") return "grid";
  if (styleId === "TableNormal") return "none";
  const hit = TABLE_STYLES.find((s) => `OH${s.id.replace(/[^a-z0-9]/gi, "")}` === styleId);
  return hit?.id ?? null;
}

export function buildStylesXml(styles: Record<string, StyleDef>): string {
  const normal = styles.Normal.props;
  const parts: string[] = [];
  parts.push(
    `<w:docDefaults><w:rPrDefault><w:rPr>${fontsXml(normal.font ?? "Aptos")}<w:kern w:val="2"/><w:sz w:val="${Math.round((normal.size ?? 11) * 2)}"/><w:szCs w:val="${Math.round((normal.size ?? 11) * 2)}"/><w:lang w:val="tr-TR" w:eastAsia="en-US" w:bidi="ar-SA"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after="${Math.round((normal.spaceAfter ?? 8) * 20)}" w:line="${Math.round((normal.line ?? 1.08) * 240)}" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>`,
  );
  for (const id of STYLE_ORDER) {
    const s = styles[id];
    if (!s) continue;
    const name = BUILTIN_NAMES[id] ?? s.name;
    const head = `<w:name w:val="${esc(name)}"/>${s.basedOn ? `<w:basedOn w:val="${s.basedOn}"/>` : ""}${s.next ? `<w:next w:val="${s.next}"/>` : ""}<w:uiPriority w:val="${s.level ? 9 : id === "Normal" ? 0 : 20}"/><w:qFormat/>`;
    if (s.type === "character") {
      parts.push(`<w:style w:type="character" w:styleId="${id}">${head}${styleRPr(s.props)}</w:style>`);
    } else {
      const p = id === "Normal" ? { ...s.props, font: undefined, size: undefined, spaceAfter: undefined, line: undefined } : s.props;
      const ppr = stylePPr(p, s.level ? { outline: s.level - 1 } : {});
      const rpr = styleRPr(p);
      parts.push(`<w:style w:type="paragraph"${id === "Normal" ? ' w:default="1"' : ""} w:styleId="${id}">${head}${ppr}${rpr}</w:style>`);
    }
  }
  // Yardımcı stiller
  parts.push('<w:style w:type="character" w:default="1" w:styleId="DefaultParagraphFont"><w:name w:val="Default Paragraph Font"/><w:uiPriority w:val="1"/><w:semiHidden/><w:unhideWhenUsed/></w:style>');
  parts.push('<w:style w:type="character" w:styleId="Hyperlink"><w:name w:val="Hyperlink"/><w:basedOn w:val="DefaultParagraphFont"/><w:uiPriority w:val="99"/><w:unhideWhenUsed/><w:rPr><w:color w:val="467886"/><w:u w:val="single"/></w:rPr></w:style>');
  parts.push('<w:style w:type="paragraph" w:styleId="Header"><w:name w:val="header"/><w:basedOn w:val="Normal"/><w:uiPriority w:val="99"/><w:unhideWhenUsed/><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:style>');
  parts.push('<w:style w:type="paragraph" w:styleId="Footer"><w:name w:val="footer"/><w:basedOn w:val="Normal"/><w:uiPriority w:val="99"/><w:unhideWhenUsed/><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:style>');
  parts.push('<w:style w:type="paragraph" w:styleId="CommentText"><w:name w:val="annotation text"/><w:basedOn w:val="Normal"/><w:uiPriority w:val="99"/><w:unhideWhenUsed/><w:pPr><w:spacing w:line="240" w:lineRule="auto"/></w:pPr><w:rPr><w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr></w:style>');
  parts.push('<w:style w:type="character" w:styleId="CommentReference"><w:name w:val="annotation reference"/><w:basedOn w:val="DefaultParagraphFont"/><w:uiPriority w:val="99"/><w:semiHidden/><w:unhideWhenUsed/><w:rPr><w:sz w:val="16"/><w:szCs w:val="16"/></w:rPr></w:style>');
  for (const k of ["Footnote", "Endnote"]) {
    parts.push(`<w:style w:type="paragraph" w:styleId="${k}Text"><w:name w:val="${k.toLowerCase()} text"/><w:basedOn w:val="Normal"/><w:uiPriority w:val="99"/><w:unhideWhenUsed/><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:rPr><w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr></w:style>`);
    parts.push(`<w:style w:type="character" w:styleId="${k}Reference"><w:name w:val="${k.toLowerCase()} reference"/><w:basedOn w:val="DefaultParagraphFont"/><w:uiPriority w:val="99"/><w:unhideWhenUsed/><w:rPr><w:vertAlign w:val="superscript"/></w:rPr></w:style>`);
  }
  parts.push(`<w:style w:type="paragraph" w:styleId="TOCHeading"><w:name w:val="TOC Heading"/><w:basedOn w:val="Heading1"/><w:next w:val="Normal"/><w:uiPriority w:val="39"/><w:unhideWhenUsed/><w:qFormat/><w:pPr><w:outlineLvl w:val="9"/></w:pPr></w:style>`);
  for (let l = 1; l <= 3; l++)
    parts.push(`<w:style w:type="paragraph" w:styleId="TOC${l}"><w:name w:val="toc ${l}"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:uiPriority w:val="39"/><w:unhideWhenUsed/><w:pPr><w:spacing w:after="100"/>${l > 1 ? `<w:ind w:left="${(l - 1) * 220}"/>` : ""}</w:pPr></w:style>`);
  parts.push('<w:style w:type="table" w:default="1" w:styleId="TableNormal"><w:name w:val="Normal Table"/><w:uiPriority w:val="99"/><w:semiHidden/><w:unhideWhenUsed/><w:tblPr><w:tblInd w:w="0" w:type="dxa"/><w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:left w:w="108" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar></w:tblPr></w:style>');
  for (const t of TABLE_STYLES) if (t.id !== "none") parts.push(tableStyleXml(t.id));
  parts.push('<w:style w:type="numbering" w:default="1" w:styleId="NoList"><w:name w:val="No List"/><w:uiPriority w:val="99"/><w:semiHidden/><w:unhideWhenUsed/></w:style>');
  return `${XML_HEAD}<w:styles xmlns:w="${NS.w}" xmlns:r="${NS.r}">${parts.join("")}</w:styles>`;
}
