/** Tema, asıl slayt, düzen, not asıl sayfası, tablo stilleri ve paket parçaları. */
import type { LevelStyle, Presentation, SpShape, TextStyles, Theme } from "../model/types";
import { TABLE_STYLES, tableStyleById, type BorderSpec, type PartStyle } from "../render/tableStyles";
import { colorXml, pPrXml, type WriteCtx } from "./writeDrawing";
import { attrs, esc, NS, SECTION_EXT_URI } from "./xml";
import { sectionsOf } from "../model/sections";

export const NS_DECL = `xmlns:a="${NS.a}" xmlns:r="${NS.r}" xmlns:p="${NS.p}"`;

export function themeXml(t: Theme, name = t.name, masterEdited = false): string {
  const c = t.colors;
  const slot = (k: keyof Theme["colors"], tag: string) => `<a:${tag}><a:srgbClr val="${c[k].replace("#", "").toUpperCase()}"/></a:${tag}>`;
  const clr =
    `<a:clrScheme name="${esc(name)}">` +
    slot("dk1", "dk1") +
    slot("lt1", "lt1") +
    slot("dk2", "dk2") +
    slot("lt2", "lt2") +
    slot("accent1", "accent1") +
    slot("accent2", "accent2") +
    slot("accent3", "accent3") +
    slot("accent4", "accent4") +
    slot("accent5", "accent5") +
    slot("accent6", "accent6") +
    slot("hlink", "hlink") +
    slot("folHlink", "folHlink") +
    "</a:clrScheme>";
  const font = (f: string) => `<a:latin typeface="${esc(f)}" panose="020B0604020202020204"/><a:ea typeface=""/><a:cs typeface=""/>`;
  const fonts = `<a:fontScheme name="${esc(name)}"><a:majorFont>${font(t.fonts.major)}</a:majorFont><a:minorFont>${font(t.fonts.minor)}</a:minorFont></a:fontScheme>`;
  const ph = '<a:schemeClr val="phClr"/>';
  const fmt =
    `<a:fmtScheme name="${esc(name)}">` +
    `<a:fillStyleLst><a:solidFill>${ph}</a:solidFill><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:lumMod val="110000"/><a:satMod val="105000"/><a:tint val="67000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:lumMod val="105000"/><a:satMod val="109000"/><a:tint val="81000"/></a:schemeClr></a:gs></a:gsLst><a:lin ang="5400000" scaled="0"/></a:gradFill><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:satMod val="103000"/><a:lumMod val="102000"/><a:tint val="94000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:lumMod val="99000"/><a:satMod val="120000"/><a:shade val="78000"/></a:schemeClr></a:gs></a:gsLst><a:lin ang="5400000" scaled="0"/></a:gradFill></a:fillStyleLst>` +
    `<a:lnStyleLst><a:ln w="12700" cap="flat" cmpd="sng" algn="ctr"><a:solidFill>${ph}</a:solidFill><a:prstDash val="solid"/><a:miter lim="800000"/></a:ln><a:ln w="19050" cap="flat" cmpd="sng" algn="ctr"><a:solidFill>${ph}</a:solidFill><a:prstDash val="solid"/><a:miter lim="800000"/></a:ln><a:ln w="25400" cap="flat" cmpd="sng" algn="ctr"><a:solidFill>${ph}</a:solidFill><a:prstDash val="solid"/><a:miter lim="800000"/></a:ln></a:lnStyleLst>` +
    `<a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst><a:outerShdw blurRad="57150" dist="19050" dir="5400000" algn="ctr" rotWithShape="0"><a:srgbClr val="000000"><a:alpha val="63000"/></a:srgbClr></a:outerShdw></a:effectLst></a:effectStyle></a:effectStyleLst>` +
    `<a:bgFillStyleLst><a:solidFill>${ph}</a:solidFill><a:solidFill><a:schemeClr val="phClr"><a:tint val="95000"/><a:satMod val="170000"/></a:schemeClr></a:solidFill><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="93000"/><a:satMod val="150000"/><a:shade val="98000"/><a:lumMod val="102000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:shade val="63000"/><a:satMod val="120000"/></a:schemeClr></a:gs></a:gsLst><a:lin ang="5400000" scaled="0"/></a:gradFill></a:bgFillStyleLst>` +
    "</a:fmtScheme>";
  const ext = t.id ? `<a:extLst><a:ext uri="{0FF1CE00-5A5A-4E7E-9B0B-5448454D4531}"><oh:theme xmlns:oh="${NS.oh}"${attrs({ id: t.id, variant: t.variant ?? 0, edited: masterEdited ? "1" : undefined })}/></a:ext></a:extLst>` : "";
  return `<a:theme xmlns:a="${NS.a}" name="${esc(name)}"><a:themeElements>${clr}${fonts}${fmt}</a:themeElements><a:objectDefaults/><a:extraClrSchemeLst/>${ext}</a:theme>`;
}

function levelsXml(levels: LevelStyle[], ctx: WriteCtx, count = 9): string {
  let s = "";
  for (let i = 0; i < count; i++) {
    const l = levels[i] ?? levels[levels.length - 1];
    if (!l) break;
    s += pPrXml(`a:lvl${i + 1}pPr`, l.para, ctx, { defRPr: l.run });
  }
  return s;
}

export function txStylesXml(ts: TextStyles, ctx: WriteCtx): string {
  return `<p:txStyles><p:titleStyle>${levelsXml(ts.title, ctx, 1)}</p:titleStyle><p:bodyStyle>${levelsXml(ts.body, ctx)}</p:bodyStyle><p:otherStyle><a:defPPr><a:defRPr lang="tr-TR"/></a:defPPr>${levelsXml(ts.other, ctx)}</p:otherStyle></p:txStyles>`;
}

export function defaultTextStyleXml(ts: TextStyles, ctx: WriteCtx): string {
  return `<p:defaultTextStyle><a:defPPr><a:defRPr lang="tr-TR"/></a:defPPr>${levelsXml(ts.other, ctx)}</p:defaultTextStyle>`;
}

export const CLR_MAP = '<p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>';

/** Asıl/düzen yer tutucularına yazılacak örnek metin. */
export function samplePrompt(s: SpShape, isMaster: boolean): string | undefined {
  if (!s.ph) return undefined;
  if (s.text?.prompt && !isMaster) return s.text.prompt;
  switch (s.ph.type) {
    case "title":
    case "ctrTitle":
      return "Asıl başlık stilini düzenlemek için tıklayın";
    case "subTitle":
      return "Asıl alt başlık stilini düzenlemek için tıklayın";
    case "body":
    case "obj":
      return "Asıl metin stillerini düzenlemek için tıklayın";
    case "pic":
      return "Resim eklemek için simgeye tıklayın";
    default:
      return undefined;
  }
}

/* ------------------------------------------------------------------ */
/* Tablo stilleri                                                       */
/* ------------------------------------------------------------------ */

function bdr(tag: string, b: BorderSpec | null | undefined, ctx: WriteCtx): string {
  if (b === undefined) return "";
  if (b === null) return `<a:${tag}><a:ln><a:noFill/></a:ln></a:${tag}>`;
  return `<a:${tag}><a:ln w="${Math.round(b.width * 12700)}" cmpd="sng"><a:solidFill>${colorXml(b.color, ctx)}</a:solidFill></a:ln></a:${tag}>`;
}

function partXml(tag: string, p: PartStyle | undefined, ctx: WriteCtx): string {
  if (!p) return "";
  const txt =
    p.text || p.bold !== undefined
      ? `<a:tcTxStyle${attrs({ b: p.bold === undefined ? undefined : p.bold ? "on" : "off" })}><a:fontRef idx="minor"><a:prstClr val="black"/></a:fontRef>${colorXml(p.text ?? { scheme: "dk1" }, ctx)}</a:tcTxStyle>`
      : "";
  const borders = bdr("left", p.left, ctx) + bdr("right", p.right, ctx) + bdr("top", p.top, ctx) + bdr("bottom", p.bottom, ctx) + bdr("insideH", p.insideH, ctx) + bdr("insideV", p.insideV, ctx);
  const fill = p.fill === undefined ? "" : p.fill === null ? "<a:fill><a:noFill/></a:fill>" : `<a:fill><a:solidFill>${colorXml(p.fill, ctx)}</a:solidFill></a:fill>`;
  const tc = borders || fill ? `<a:tcStyle>${borders ? `<a:tcBdr>${borders}</a:tcBdr>` : ""}${fill}</a:tcStyle>` : "<a:tcStyle/>";
  return `<a:${tag}>${txt}${tc}</a:${tag}>`;
}

export function tableStylesXml(used: Set<string>, ctx: WriteCtx): string {
  const def = tableStyleById("medium2-accent1");
  const list = TABLE_STYLES.filter((s) => used.has(s.id) || s.id === def.id);
  const body = list
    .map(
      (s) =>
        `<a:tblStyle styleId="${s.guid}" styleName="${esc(s.name)}">${partXml("wholeTbl", s.wholeTbl, ctx)}${partXml("band1H", s.band1H, ctx)}${partXml("band1V", s.band1V, ctx)}${partXml("lastCol", s.lastCol, ctx)}${partXml("firstCol", s.firstCol, ctx)}${partXml("lastRow", s.lastRow, ctx)}${partXml("firstRow", s.firstRow, ctx)}</a:tblStyle>`,
    )
    .join("");
  return `<a:tblStyleLst xmlns:a="${NS.a}" def="${def.guid}">${body}</a:tblStyleLst>`;
}

/* ------------------------------------------------------------------ */
/* Not asıl sayfası ve not slaytları                                     */
/* ------------------------------------------------------------------ */

const NOTES_W = 6858000;
const NOTES_H = 9144000;

export function notesMasterXml(): string {
  const tree =
    '<p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>' +
    `<p:sp><p:nvSpPr><p:cNvPr id="2" name="Slayt Görüntüsü Yer Tutucusu 1"/><p:cNvSpPr><a:spLocks noGrp="1" noRot="1" noChangeAspect="1"/></p:cNvSpPr><p:nvPr><p:ph type="sldImg" idx="2"/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x="685800" y="1143000"/><a:ext cx="5486400" cy="3086100"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/><a:ln w="12700"><a:solidFill><a:prstClr val="black"/></a:solidFill></a:ln></p:spPr></p:sp>` +
    `<p:sp><p:nvSpPr><p:cNvPr id="3" name="Not Yer Tutucusu 2"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="body" sz="quarter" idx="3"/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x="685800" y="4400550"/><a:ext cx="5486400" cy="3600450"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr><p:txBody><a:bodyPr vert="horz" lIns="91440" tIns="45720" rIns="91440" bIns="45720" rtlCol="0"/><a:lstStyle/><a:p><a:pPr lvl="0"/><a:r><a:rPr lang="tr-TR"/><a:t>Asıl metin stillerini düzenlemek için tıklayın</a:t></a:r></a:p></p:txBody></p:sp>` +
    "</p:spTree>";
  const style = `<p:notesStyle><a:lvl1pPr marL="0" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:defRPr sz="1200" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl1pPr></p:notesStyle>`;
  return `<p:notesMaster ${NS_DECL}><p:cSld><p:bg><p:bgRef idx="1001"><a:schemeClr val="bg1"/></p:bgRef></p:bg>${tree}</p:cSld>${CLR_MAP}${style}</p:notesMaster>`;
}

export function notesSlideXml(text: string): string {
  const paras = text
    .split("\n")
    .map((l) => (l ? `<a:p><a:r><a:rPr lang="tr-TR" dirty="0"/><a:t>${esc(l)}</a:t></a:r></a:p>` : '<a:p><a:endParaRPr lang="tr-TR" dirty="0"/></a:p>'))
    .join("");
  const tree =
    '<p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>' +
    '<p:sp><p:nvSpPr><p:cNvPr id="2" name="Slayt Görüntüsü Yer Tutucusu 1"/><p:cNvSpPr><a:spLocks noGrp="1" noRot="1" noChangeAspect="1"/></p:cNvSpPr><p:nvPr><p:ph type="sldImg"/></p:nvPr></p:nvSpPr><p:spPr/></p:sp>' +
    `<p:sp><p:nvSpPr><p:cNvPr id="3" name="Not Yer Tutucusu 2"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="body" idx="1"/></p:nvPr></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/>${paras}</p:txBody></p:sp>` +
    "</p:spTree>";
  return `<p:notes ${NS_DECL}><p:cSld>${tree}</p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:notes>`;
}

export function presentationXml(pres: Presentation, masterRid: string, notesRid: string | null, slideRids: string[], ctx: WriteCtx): string {
  const sldIds = slideRids.map((rid, i) => `<p:sldId id="${256 + i}" r:id="${rid}"/>`).join("");
  const type = Math.abs(pres.width / pres.height - 4 / 3) < 0.01 ? ' type="screen4x3"' : "";
  return (
    `<p:presentation ${NS_DECL} saveSubsetFonts="1">` +
    `<p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="${masterRid}"/></p:sldMasterIdLst>` +
    (notesRid ? `<p:notesMasterIdLst><p:notesMasterId r:id="${notesRid}"/></p:notesMasterIdLst>` : "") +
    (sldIds ? `<p:sldIdLst>${sldIds}</p:sldIdLst>` : "") +
    `<p:sldSz cx="${Math.round(pres.width * 9525)}" cy="${Math.round(pres.height * 9525)}"${type}/>` +
    `<p:notesSz cx="${NOTES_W}" cy="${NOTES_H}"/>` +
    defaultTextStyleXml(pres.master.textStyles, ctx) +
    sectionsXml(pres) +
    "</p:presentation>"
  );
}

/** Dizgeden kararlı GUID ("{8-4-4-4-12}"). */
export function guidFrom(str: string): string {
  let h1 = 0x811c9dc5;
  let h2 = 0x01000193;
  const hex: string[] = [];
  for (let round = 0; round < 4; round++) {
    for (let i = 0; i < str.length; i++) {
      h1 = Math.imul(h1 ^ str.charCodeAt(i), 16777619) >>> 0;
      h2 = Math.imul(h2 + str.charCodeAt(i) + round, 2246822519) >>> 0;
    }
    hex.push(((h1 ^ h2 ^ (round * 0x9e3779b9)) >>> 0).toString(16).toUpperCase().padStart(8, "0"));
  }
  const x = hex.join("");
  return `{${x.slice(0, 8)}-${x.slice(8, 12)}-${x.slice(12, 16)}-${x.slice(16, 20)}-${x.slice(20, 32)}}`;
}

/** Bölümler (p14:sectionLst, PowerPoint 2010 uzantısı; sldId kimlikleri 256'dan başlar). */
function sectionsXml(pres: Presentation): string {
  const secs = sectionsOf(pres.slides);
  if (!secs.length) return "";
  const idx = new Map(pres.slides.map((s, i) => [s.id, i]));
  const body = secs
    .map((sec) => `<p14:section name="${esc(sec.name)}" id="${guidFrom(sec.id)}"><p14:sldIdLst>${sec.slideIds.map((id) => `<p14:sldId id="${256 + idx.get(id)!}"/>`).join("")}</p14:sldIdLst></p14:section>`)
    .join("");
  return `<p:extLst><p:ext uri="${SECTION_EXT_URI}"><p14:sectionLst xmlns:p14="${NS.p14}">${body}</p14:sectionLst></p:ext></p:extLst>`;
}

export function coreXml(pres: Presentation): string {
  const now = new Date().toISOString().replace(/\.\d+Z$/, "Z");
  const created = pres.props.created ? new Date(pres.props.created).toISOString().replace(/\.\d+Z$/, "Z") : now;
  return (
    '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">' +
    `<dc:title>${esc(pres.props.title ?? "")}</dc:title>` +
    (pres.props.subject ? `<dc:subject>${esc(pres.props.subject)}</dc:subject>` : "") +
    `<dc:creator>${esc(pres.props.creator ?? "")}</dc:creator>` +
    `<cp:lastModifiedBy>${esc(pres.props.creator ?? "")}</cp:lastModifiedBy>` +
    `<dcterms:created xsi:type="dcterms:W3CDTF">${created}</dcterms:created>` +
    `<dcterms:modified xsi:type="dcterms:W3CDTF">${now}</dcterms:modified>` +
    "</cp:coreProperties>"
  );
}

export function appXml(pres: Presentation): string {
  const fmt = Math.abs(pres.width / pres.height - 16 / 9) < 0.01 ? "Geniş ekran" : Math.abs(pres.width / pres.height - 4 / 3) < 0.01 ? "Ekran Gösterisi (4:3)" : "Özel";
  return `<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"><Application>OfficeHere Sunum</Application><PresentationFormat>${fmt}</PresentationFormat><Slides>${pres.slides.length}</Slides><Notes>${pres.slides.filter((s) => s.notes).length}</Notes><HiddenSlides>${pres.slides.filter((s) => s.hidden).length}</HiddenSlides><AppVersion>16.0000</AppVersion></Properties>`;
}

export const PRES_PROPS = `<p:presentationPr ${NS_DECL}/>`;
export const VIEW_PROPS = `<p:viewPr ${NS_DECL}><p:normalViewPr><p:restoredLeft sz="15620"/><p:restoredTop sz="94660"/></p:normalViewPr><p:gridSpacing cx="72008" cy="72008"/></p:viewPr>`;
