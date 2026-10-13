// DOCX yazıcı: ProseMirror JSON + belge ayarları → OOXML paketi (ECMA-376).
import JSZip from "jszip";
import { BIB_NS, MERGE_NS, mergeToXml, sourcesToXml, writeCustomXml } from "./customXml";
import { hfKeyOf, hfOwner } from "../../editor/sectionHf";
import { CT_CHART } from "./chartXml";
import type { JSONContent } from "@tiptap/core";
import type { BelgeDocument, CommentData, DocSettings, HfVariant, PageSetup, SectionKind } from "../../model/document";
import { normalizeDoc } from "../../model/document";
import { resolveStyles, type BorderSpec } from "../../model/styles";
import { NS, REL, XML_HEAD, cleanText, esc, hex6 } from "./xml";
import { borderXml, buildStylesXml, fontsXml } from "./styleXml";
import { levelDef } from "./numbering";

import { Ctx, DOC_NS, PartWriter, Rels, type NoteData } from "./partWriter";


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

/** Dipnot / son not parçası (ayraçlar + notlar). Zengin içerik paragraf paragraf yazılır. */
function notesXml(kind: "footnote" | "endnote", notes: NoteData[], ctx: Ctx, rels: Rels): string {
  const tag = kind === "footnote" ? "footnote" : "endnote";
  const root = kind === "footnote" ? "footnotes" : "endnotes";
  const style = kind === "footnote" ? "Footnote" : "Endnote";
  const sep = (id: number, type: string, el: string) => `<w:${tag} w:type="${type}" w:id="${id}"><w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:r><w:${el}/></w:r></w:p></w:${tag}>`;
  let x = `${XML_HEAD}<w:${root} ${DOC_NS}>${sep(-1, "separator", "separator")}${sep(0, "continuationSeparator", "continuationSeparator")}`;
  const ref = `<w:r><w:rPr><w:rStyle w:val="${style}Reference"/></w:rPr><w:${tag}Ref/></w:r><w:r><w:t xml:space="preserve"> </w:t></w:r>`;
  const pw = new PartWriter(ctx, rels, 9000);
  notes.forEach((n, i) => {
    x += `<w:${tag} w:id="${i + 1}">`;
    const paras: JSONContent[] = n.content?.length ? n.content : (n.text || "").split("\n").map((t) => ({ type: "paragraph", ...(t ? { content: [{ type: "text", text: t }] } : {}) }));
    paras.forEach((p, k) => {
      if (p.type === "paragraph" || p.type === "heading") {
        const body = pw.inline(p.content ?? []);
        x += `<w:p><w:pPr><w:pStyle w:val="${style}Text"/></w:pPr>${k === 0 ? ref : ""}${body}</w:p>`;
      } else x += pw.block(p);
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

interface HfRefs {
  header: Partial<Record<HfVariant, string>>;
  footer: Partial<Record<HfVariant, string>>;
}

const SECT_TYPE: Record<SectionKind, string> = { nextPage: "nextPage", continuous: "continuous", evenPage: "evenPage", oddPage: "oddPage" };

/** Bölüm özellikleri (üst/alt bilgiler belge genelidir; her bölüm aynı parçalara başvurur). */
function sectPr(s: DocSettings, p: PageSetup, kind: SectionKind, refs: HfRefs | null): string {
  const m = p.margin;
  let x = "<w:sectPr>";
  if (refs) {
    for (const v of ["default", "first", "even"] as const) {
      if (refs.header[v]) x += `<w:headerReference w:type="${v}" r:id="${refs.header[v]}"/>`;
      if (refs.footer[v]) x += `<w:footerReference w:type="${v}" r:id="${refs.footer[v]}"/>`;
    }
  }
  if (kind !== "nextPage") x += `<w:type w:val="${SECT_TYPE[kind]}"/>`;
  x += `<w:pgSz w:w="${p.width}" w:h="${p.height}"${p.orientation === "landscape" ? ' w:orient="landscape"' : ""}/>`;
  x += `<w:pgMar w:top="${m.top}" w:right="${m.right}" w:bottom="${m.bottom}" w:left="${m.left}" w:header="${m.header}" w:footer="${m.footer}" w:gutter="${m.gutter}"/>`;
  if (s.pageBorder) {
    const b: BorderSpec = { style: s.pageBorder.style, size: s.pageBorder.width, color: s.pageBorder.color, space: 24 };
    x += `<w:pgBorders w:offsetFrom="page">${borderXml("top", b)}${borderXml("left", b)}${borderXml("bottom", b)}${borderXml("right", b)}</w:pgBorders>`;
  }
  // Şema sırası: pgBorders → lnNumType → cols.
  const ln = p.lineNumbers;
  // Satır numaraları: OOXML'de başlangıç sıfır tabanlıdır (1 → yazılmaz).
  if (ln) x += `<w:lnNumType w:countBy="${Math.max(1, ln.countBy ?? 1)}"${ln.start && ln.start > 1 ? ` w:start="${ln.start - 1}"` : ""}${ln.distance ? ` w:distance="${ln.distance}"` : ""} w:restart="${ln.restart}"/>`;
  x += `<w:cols w:space="${p.columnSpace}"${p.columns > 1 ? ` w:num="${p.columns}"` : ""}${p.columns > 1 && p.columnLine ? ' w:sep="1"' : ""}/>`;
  if (p.titlePage) x += "<w:titlePg/>";
  x += `<w:docGrid w:linePitch="360"/></w:sectPr>`;
  return x;
}

/** Belgenin bölümleri (bölümsüz eski içerik tek bölüme sarılır). */
function sectionsOf(content: JSONContent): JSONContent[] {
  return normalizeDoc(content).content ?? [];
}

/** Bölümün sectPr'ının yazılacağı son paragraf (liste içindeyse son liste paragrafı). */
function lastParagraph(nodes: JSONContent[] | undefined): JSONContent | null {
  const last = nodes?.[nodes.length - 1];
  if (!last) return null;
  if (last.type === "paragraph" || last.type === "heading") return last;
  if (last.type === "bulletList" || last.type === "orderedList") {
    const li = last.content?.[last.content.length - 1];
    return lastParagraph(li?.content);
  }
  return null;
}

/** Düzenlemeyi kısıtla (w:documentProtection; ISO/IEC 29500 parola karması). */
function protectionXml(p: DocSettings["protection"]): string {
  if (!p) return "";
  const pw = p.hash && p.salt ? ` w:algorithmName="${esc(p.algorithm ?? "SHA-512")}" w:hashValue="${esc(p.hash)}" w:saltValue="${esc(p.salt)}" w:spinCount="${p.spinCount ?? 100000}"` : "";
  return `<w:documentProtection w:edit="${p.mode}" w:enforcement="${p.enforced ? 1 : 0}"${pw}/>`;
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

  // Üst / alt bilgi parçaları (varsayılan, ilk sayfa, çift sayfa). Filigran tüm üst bilgilere yazılır.
  const hdrNs = `<w:hdr ${DOC_NS}>`;
  const refs: HfRefs = { header: {}, footer: {} };
  const anyTitle = sectionsOf(doc.content).some((sec) => !!((sec.attrs?.page as PageSetup | null) ?? settings.page).titlePage);
  let hfNo = 0;
  const hfParts: { name: string; kind: "header" | "footer" }[] = [];
  /** Üst/alt bilgi parçası yazar; ilişki kimliğini döner. `own`: bölümün kendi (boş olsa da yazılır). */
  const writeHf = (which: "header" | "footer", v: HfVariant, json: JSONContent | null, own = false): string | undefined => {
    const has = hasHfContent(json);
    const needWm = which === "header" && !!settings.watermark && (v === "default" || (v === "first" && anyTitle) || (v === "even" && settings.oddEven));
    if (!own && !has && !needWm && !(v !== "default" && json)) return undefined;
    const rels = new Rels();
    const pw = new PartWriter(ctx, rels, contentWidthTw);
    hfNo++;
    const name = `${which}${hfNo}.xml`;
    let rid: string;
    if (which === "header") {
      let inner = has ? pw.blocks(json?.content) : "";
      inner = watermarkXml(settings, settings.page.width) + (inner || '<w:p><w:pPr><w:pStyle w:val="Header"/></w:pPr></w:p>');
      zip.file(`word/${name}`, `${XML_HEAD}${hdrNs}${inner}</w:hdr>`);
      rid = docRels.add(REL.header, name);
    } else {
      const inner = (has ? pw.blocks(json?.content) : "") || '<w:p><w:pPr><w:pStyle w:val="Footer"/></w:pPr></w:p>';
      zip.file(`word/${name}`, `${XML_HEAD}<w:ftr ${DOC_NS}>${inner}</w:ftr>`);
      rid = docRels.add(REL.footer, name);
    }
    if (rels.size) zip.file(`word/_rels/${name}.rels`, rels.xml());
    hfParts.push({ name, kind: which });
    return rid;
  };
  const variants: HfVariant[] = settings.oddEven ? ["default", "first", "even"] : ["default", "first"];
  for (const v of variants) {
    for (const w of ["header", "footer"] as const) {
      const rid = writeHf(w, v, settings[hfKeyOf(w, v)]);
      if (rid) refs[w][v] = rid;
    }
  }
  // Bağı kaldırılmış bölümlerin kendi üst/alt bilgileri: her bölüm etkin parçalara açıkça başvurur.
  const ids = sectionsOf(doc.content).map((sec) => (sec.attrs?.sid as string | null | undefined) ?? null);
  const ownRid = new Map<string, string | undefined>();
  const sectionRefs = ids.map((_, i): HfRefs => {
    const r: HfRefs = { header: {}, footer: {} };
    for (const v of variants) {
      for (const w of ["header", "footer"] as const) {
        const key = hfKeyOf(w, v);
        const owner = hfOwner(ids, settings, i, key);
        if (owner < 0) {
          if (refs[w][v]) r[w][v] = refs[w][v];
          continue;
        }
        const ck = `${owner}:${key}`;
        if (!ownRid.has(ck)) ownRid.set(ck, writeHf(w, v, settings.sectionHf?.[ids[owner]!]?.[key] ?? null, true));
        const rid = ownRid.get(ck);
        if (rid) r[w][v] = rid;
      }
    }
    return r;
  });

  // Gövde: her bölümün sectPr'ı son paragrafına, son bölümünkü gövdenin sonuna yazılır.
  const body = new PartWriter(ctx, docRels, contentWidthTw);
  const sections = sectionsOf(doc.content);
  let bodyXml = "";
  let finalSect = "";
  sections.forEach((sec, i) => {
    const page = (sec.attrs?.page as PageSetup | null) ?? settings.page;
    const kind = ((sec.attrs?.kind as SectionKind | undefined) ?? "nextPage") as SectionKind;
    // Başvurular her bölüme yazılır (aynı parçalar): devralmayı tam desteklemeyen okuyucular için.
    const xml = sectPr(settings, page, i === 0 ? "nextPage" : kind, sectionRefs[i] ?? refs);
    const blocks = sec.content ?? [];
    body.contentWidthTw = Math.round((page.width - page.margin.left - page.margin.right - page.margin.gutter - (page.columns > 1 ? (page.columns - 1) * page.columnSpace : 0)) / Math.max(1, page.columns));
    if (i === sections.length - 1) {
      finalSect = xml;
      body.sectTarget = null;
      bodyXml += body.blocks(blocks);
      return;
    }
    const target = lastParagraph(blocks);
    body.sectTarget = target;
    body.sectXml = xml;
    bodyXml += body.blocks(blocks);
    if (!target) bodyXml += `<w:p><w:pPr>${xml}</w:pPr></w:p>`;
    body.sectTarget = null;
  });
  // Kapanmamış açıklama aralıkları son paragrafın içinde kapanır (yeni boş paragraf eklenirse her
  // kaydet-aç turunda belgenin sonuna bir paragraf daha eklenirdi).
  let ends = "";
  for (const id of [...ctx.openComments]) {
    const num = ctx.commentNum(id);
    ends += `<w:commentRangeEnd w:id="${num}"/><w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr><w:commentReference w:id="${num}"/></w:r>`;
  }
  if (ends) bodyXml = bodyXml.endsWith("</w:p>") ? `${bodyXml.slice(0, -"</w:p>".length)}${ends}</w:p>` : `${bodyXml}<w:p>${ends}</w:p>`;
  ctx.openComments.clear();
  const lastIsTable = /<\/w:tbl>$/.test(bodyXml) || /<\/w:sdt>$/.test(bodyXml);
  if (!bodyXml || lastIsTable) bodyXml += "<w:p/>";

  const bg = settings.pageColor ? `<w:background w:color="${hex6(settings.pageColor)}"/>` : "";
  zip.file("word/document.xml", `${XML_HEAD}<w:document ${DOC_NS} mc:Ignorable="w14 wp14"><w:body>${bodyXml}${finalSect}</w:body></w:document>`.replace("<w:body>", `${bg}<w:body>`));

  const styles = resolveStyles(settings.styleSet, settings.theme, settings.fontTheme, settings.paragraphSpacing, settings.customStyles);
  zip.file("word/styles.xml", buildStylesXml(styles));
  zip.file("word/numbering.xml", numberingXml(ctx));
  zip.file(
    "word/settings.xml",
    `${XML_HEAD}<w:settings xmlns:w="${NS.w}" xmlns:r="${NS.r}"><w:zoom w:percent="100"/>${settings.pageColor ? "<w:displayBackgroundShape/>" : ""}${protectionXml(settings.protection)}<w:defaultTabStop w:val="708"/><w:hyphenationZone w:val="425"/><w:characterSpacingControl w:val="doNotCompress"/>${settings.oddEven ? "<w:evenAndOddHeaders/>" : ""}${ctx.footnotes.length ? '<w:footnotePr><w:footnote w:id="-1"/><w:footnote w:id="0"/></w:footnotePr>' : ""}${ctx.endnotes.length ? '<w:endnotePr><w:endnote w:id="-1"/><w:endnote w:id="0"/></w:endnotePr>' : ""}<w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat><w:themeFontLang w:val="tr-TR"/></w:settings>`,
  );
  if (ctx.usedComments.size) {
    zip.file("word/comments.xml", commentsXml(ctx, settings.comments));
    docRels.add(REL.comments, "comments.xml");
  }
  for (const [kind, list, part, rel] of [
    ["footnote", ctx.footnotes, "footnotes.xml", REL.footnotes],
    ["endnote", ctx.endnotes, "endnotes.xml", REL.endnotes],
  ] as const) {
    if (!list.length) continue;
    const rels = new Rels();
    zip.file(`word/${part}`, notesXml(kind, list, ctx, rels));
    if (rels.size) zip.file(`word/_rels/${part}.rels`, rels.xml());
    docRels.add(rel, part);
  }
  for (const m of ctx.media) zip.file(`word/media/${m.name}`, m.data);
  for (const ch of ctx.charts) zip.file(`word/charts/${ch.name}`, ch.xml);
  // Özel XML: kaynakça kaynakları (Kaynakları Yönet listesi) ve birleştirme veri kaynağı.
  const customParts: { xml: string; schema: string }[] = [];
  if (settings.sources?.length) customParts.push({ xml: sourcesToXml(settings.sources, settings.bibStyle), schema: BIB_NS });
  if (settings.mailMerge?.list.records.length) customParts.push({ xml: mergeToXml(settings.mailMerge), schema: MERGE_NS });
  const customOverrides = writeCustomXml(zip, docRels, customParts);
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
      hfParts.map((h) => ov(`word/${h.name}`, `wordprocessingml.${h.kind}+xml`)).join("") +
      (ctx.usedComments.size ? ov("word/comments.xml", "wordprocessingml.comments+xml") : "") +
      (ctx.footnotes.length ? ov("word/footnotes.xml", "wordprocessingml.footnotes+xml") : "") +
      (ctx.endnotes.length ? ov("word/endnotes.xml", "wordprocessingml.endnotes+xml") : "") +
      `<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>` +
      ov("docProps/app.xml", "extended-properties+xml") +
      ov("docProps/custom.xml", "custom-properties+xml") +
      customOverrides +
      ctx.charts.map((ch) => `<Override PartName="/word/charts/${ch.name}" ContentType="${CT_CHART}"/>`).join("") +
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
  const custom = { styleSet: settings.styleSet, theme: settings.theme, fontTheme: settings.fontTheme, paragraphSpacing: settings.paragraphSpacing, watermark: settings.watermark, customStyles: settings.customStyles, sources: settings.sources, bibStyle: settings.bibStyle };
  zip.file(
    "docProps/custom.xml",
    `${XML_HEAD}<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"><property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="2" name="OfficeHereBelge"><vt:lpwstr>${esc(JSON.stringify(custom))}</vt:lpwstr></property></Properties>`,
  );
  return zip.generateAsync({ type: "uint8array", compression: "DEFLATE", mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document" });
}
