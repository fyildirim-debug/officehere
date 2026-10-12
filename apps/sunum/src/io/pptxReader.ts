/**
 * PPTX okuyucu: slayt boyutu, sıra, asıl slayt/düzen kalıtımı, tema renkleri
 * ve yazı tipleri, şekiller, resimler, tablolar, gruplar, arka planlar,
 * notlar, gizli slaytlar, geçişler ve animasyonlar.
 */
import JSZip from "jszip";
import { uid } from "../model/factory";
import { DEFAULT_HF } from "../model/presentation";
import { findPlaceholder, placeholderKind } from "../model/textStyle";
import type { Fill, Layout, LayoutType, LevelStyle, Master, Placeholder, Presentation, SchemeColorName, Shape, Slide, SpShape, TextStyles, Theme } from "../model/types";
import { buildTextStyles, LAYOUT_NAMES } from "../themes/build";
import { getThemeDef } from "../themes/themes";
import { EXT_MIME } from "./pptxWriter";
import { DEFAULT_CLRMAP, readColor, readFill, readLevels, readShapeList, readTextBody, type ReadCtx } from "./readDrawing";
import { readAnimations, readTransition } from "./readTiming";
import { attr, fromEmu, kid, kidPath, kids, numAttr, parseRels, parseXml, relId, relsPathFor, REL, type Rel } from "./xml";

export class PptxError extends Error {}

function bytesToBase64(bytes: Uint8Array): string {
  let s = "";
  const chunk = 0x8000;
  for (let i = 0; i < bytes.length; i += chunk) s += String.fromCharCode(...bytes.subarray(i, i + chunk));
  return btoa(s);
}

const SCHEME_KEYS: SchemeColorName[] = ["dk1", "lt1", "dk2", "lt2", "accent1", "accent2", "accent3", "accent4", "accent5", "accent6", "hlink", "folHlink"];

function readTheme(xml: string | undefined): { theme: Theme; lnWidths: number[] } {
  const def: Theme = {
    name: "Tema",
    colors: { dk1: "#000000", lt1: "#FFFFFF", dk2: "#44546A", lt2: "#E7E6E6", accent1: "#4472C4", accent2: "#ED7D31", accent3: "#A5A5A5", accent4: "#FFC000", accent5: "#5B9BD5", accent6: "#70AD47", hlink: "#0563C1", folHlink: "#954F72" },
    fonts: { major: "Calibri Light", minor: "Calibri" },
  };
  if (!xml) return { theme: def, lnWidths: [0.75, 1.5, 2] };
  const doc = parseXml(xml);
  const root = doc.documentElement;
  const te = kid(root, "themeElements");
  const cs = kid(te, "clrScheme");
  const colors = { ...def.colors };
  for (const k of SCHEME_KEYS) {
    const el = kid(cs, k);
    const c = kids(el)[0];
    if (!c) continue;
    if (c.localName === "srgbClr") colors[k] = `#${(attr(c, "val") ?? "000000").toUpperCase()}`;
    else if (c.localName === "sysClr") colors[k] = `#${(attr(c, "lastClr") ?? (k.startsWith("lt") ? "FFFFFF" : "000000")).toUpperCase()}`;
  }
  const fs = kid(te, "fontScheme");
  const major = attr(kidPath(fs, "majorFont", "latin"), "typeface") || def.fonts.major;
  const minor = attr(kidPath(fs, "minorFont", "latin"), "typeface") || def.fonts.minor;
  const lnWidths = kids(kidPath(te, "fmtScheme", "lnStyleLst"), "ln").map((l) => (numAttr(l, "w") ?? 12700) / 12700);
  const theme: Theme = { name: attr(root, "name") ?? def.name, colors, fonts: { major, minor } };
  for (const e of kids(kid(root, "extLst"), "ext")) {
    const t = kid(e, "theme");
    if (t && attr(t, "id") && getThemeDef(attr(t, "id"))) {
      theme.id = attr(t, "id");
      theme.variant = numAttr(t, "variant") ?? 0;
    }
  }
  return { theme, lnWidths: lnWidths.length ? lnWidths : [0.75, 1.5, 2] };
}

function readClrMap(el: Element | undefined, base: Record<string, SchemeColorName> = DEFAULT_CLRMAP): Record<string, SchemeColorName> {
  if (!el) return base;
  const map = { ...base };
  for (const k of ["bg1", "tx1", "bg2", "tx2", "accent1", "accent2", "accent3", "accent4", "accent5", "accent6", "hlink", "folHlink"]) {
    const v = attr(el, k) as SchemeColorName | undefined;
    if (v) map[k] = v;
  }
  return map;
}

function readBg(cSld: Element | undefined, ctx: ReadCtx): Fill | undefined {
  const bg = kid(cSld, "bg");
  if (!bg) return undefined;
  const pr = kid(bg, "bgPr");
  if (pr) return readFill(pr, ctx);
  const ref = kid(bg, "bgRef");
  if (ref) {
    const c = readColor(ref, ctx);
    return c ? { type: "solid", color: c } : undefined;
  }
  return undefined;
}

function levelsToStyles(levels: Partial<LevelStyle>[] | undefined, fallback: LevelStyle[]): LevelStyle[] {
  return fallback.map((f, i) => {
    const l = levels?.[i];
    if (!l) return f;
    return { para: { ...f.para, ...strip(l.para ?? {}) }, run: { ...f.run, ...strip(l.run ?? {}) } };
  });
}

function strip<T extends object>(o: T): Partial<T> {
  const r: Partial<T> = {};
  for (const [k, v] of Object.entries(o)) if (v !== undefined) (r as Record<string, unknown>)[k] = v;
  return r;
}

/** Asıl/düzen yer tutucularının örnek metnini boşaltır (özel istem metni korunur). */
function cleanPlaceholderText(shapes: Shape[], tree: Element | undefined, isLayout: boolean, idMap: Map<number, string>): Shape[] {
  const inverse = new Map([...idMap.entries()].map(([k, v]) => [v, k]));
  const customIds = new Set<number>();
  if (tree) {
    for (const sp of kids(tree, "sp")) {
      const ph = kidPath(sp, "nvSpPr", "nvPr", "ph");
      if (ph && (attr(ph, "hasCustomPrompt") === "1" || attr(ph, "hasCustomPrompt") === "true")) {
        const id = numAttr(kidPath(sp, "nvSpPr", "cNvPr"), "id");
        if (id !== undefined) customIds.add(id);
      }
    }
  }
  return shapes.map((s) => {
    if (s.kind !== "sp" || !s.ph || !s.text) return s;
    if (s.ph.type === "dt" || s.ph.type === "ftr" || s.ph.type === "sldNum") return s;
    const plain = s.text.paragraphs.map((p) => p.runs.map((r) => r.text).join("")).join("\n");
    const nid = inverse.get(s.id);
    const custom = isLayout && !!plain && nid !== undefined && customIds.has(nid);
    return { ...s, text: { ...s.text, prompt: custom ? plain : undefined, paragraphs: [{ runs: [], level: 0, props: {} }] } };
  });
}

/** Düzey stillerini birleştirir (b, a'nın üzerine yazar). */
function mergeLevels(a: Partial<LevelStyle>[] | undefined, b: Partial<LevelStyle>[] | undefined): Partial<LevelStyle>[] | undefined {
  if (!a?.length) return b;
  if (!b?.length) return a;
  const n = Math.max(a.length, b.length);
  const out: Partial<LevelStyle>[] = [];
  for (let i = 0; i < n; i++) {
    const x = a[i] ?? {};
    const y = b[i] ?? {};
    out.push({ para: { ...(x.para ?? {}), ...(y.para ?? {}) }, run: { ...(x.run ?? {}), ...(y.run ?? {}) } } as Partial<LevelStyle>);
  }
  return out;
}

const LAYOUT_TYPES = new Set(["title", "obj", "secHead", "twoObj", "twoTxTwoObj", "titleOnly", "blank", "objTx", "picTx", "vertTx"]);

/** PPTX baytlarını sunu modeline okur. */
export async function readPptx(data: ArrayBuffer | Uint8Array): Promise<Presentation> {
  let zip: JSZip;
  try {
    zip = await JSZip.loadAsync(data);
  } catch {
    throw new PptxError("Dosya geçerli bir PPTX (ZIP) paketi değil.");
  }
  const text = async (p: string) => {
    const f = zip.file(p);
    return f ? f.async("string") : undefined;
  };

  // Ortam dosyaları
  const media = new Map<string, string>();
  const mediaFiles = Object.keys(zip.files).filter((p) => /\.(png|jpe?g|gif|bmp|svg|webp|tiff?|emf|wmf)$/i.test(p) && !zip.files[p].dir);
  await Promise.all(
    mediaFiles.map(async (p) => {
      const bytes = await zip.file(p)!.async("uint8array");
      const ext = p.split(".").pop()!.toLowerCase();
      media.set(p, `data:${EXT_MIME[ext] ?? "application/octet-stream"};base64,${bytesToBase64(bytes)}`);
    }),
  );

  const rootRels = parseRels(await text("_rels/.rels"), "");
  const presPath = [...rootRels.values()].find((r) => r.type === REL.officeDocument)?.target ?? "ppt/presentation.xml";
  const presXml = await text(presPath);
  if (!presXml) throw new PptxError("Sunu bölümü (presentation.xml) bulunamadı.");
  const presDoc = parseXml(presXml).documentElement;
  const presRels = parseRels(await text(relsPathFor(presPath)), presPath);
  const sz = kid(presDoc, "sldSz");
  const width = fromEmu(numAttr(sz, "cx"), 1280) || 1280;
  const height = fromEmu(numAttr(sz, "cy"), 720) || 720;

  // Asıl slayt
  const masterRid = relId(kid(kid(presDoc, "sldMasterIdLst"), "sldMasterId"));
  const masterPath = (masterRid && presRels.get(masterRid)?.target) || [...presRels.values()].find((r) => r.type === REL.slideMaster)?.target;
  if (!masterPath) throw new PptxError("Asıl slayt bulunamadı.");
  const masterDoc = parseXml((await text(masterPath)) ?? "<x/>").documentElement;
  const masterRels = parseRels(await text(relsPathFor(masterPath)), masterPath);
  const themePath = [...masterRels.values()].find((r) => r.type === REL.theme)?.target ?? [...presRels.values()].find((r) => r.type === REL.theme)?.target;
  const { theme, lnWidths } = readTheme(themePath ? await text(themePath) : undefined);
  const clrMap = readClrMap(kid(masterDoc, "clrMap"));

  const mkCtx = (rels: Map<string, Rel>, map = clrMap, inherit?: ReadCtx["inherit"]): ReadCtx => ({ rels, media, clrMap: map, idMap: new Map(), inherit, lnWidths });

  // Açıklama yazarları (yalnız gerektiğinde okunur).
  let authorsCache: Map<string, string> | undefined;
  const commentAuthors = async (): Promise<Map<string, string>> => {
    if (authorsCache) return authorsCache;
    authorsCache = new Map();
    const rel = [...presRels.values()].find((r) => r.type === REL.commentAuthors);
    const xml = rel ? await text(rel.target) : undefined;
    if (xml) for (const a of kids(parseXml(xml).documentElement, "cmAuthor")) authorsCache.set(attr(a, "id") ?? "", attr(a, "name") ?? "Yazar");
    return authorsCache;
  };

  // Asıl slaytlar: model tek asıl slayt taşır. İlki asıl slayt olur; diğerlerinin
  // arka planı ve süs şekilleri kendi düzenlerine aktarılır (tüm düzenler okunur).
  const masterIdEls = kids(kid(presDoc, "sldMasterIdLst"), "sldMasterId");
  const masterPaths = masterIdEls.map((e) => relId(e)).map((rid) => (rid ? presRels.get(rid)?.target : undefined)).filter(Boolean) as string[];
  if (!masterPaths.includes(masterPath)) masterPaths.unshift(masterPath);

  const layouts: Layout[] = [];
  const layoutByPath = new Map<string, Layout>();
  const layoutMasterShapes = new Map<Layout, Shape[]>();
  let master: Master | undefined;
  for (const [mi, mp] of masterPaths.entries()) {
    const mDoc = mi === 0 ? masterDoc : parseXml((await text(mp)) ?? "<x/>").documentElement;
    const mRels = mi === 0 ? masterRels : parseRels(await text(relsPathFor(mp)), mp);
    const mClr = mi === 0 ? clrMap : readClrMap(kid(mDoc, "clrMap"));
    const mctx = mkCtx(mRels, mClr);
    const mCSld = kid(mDoc, "cSld");
    const mTree = kid(mCSld, "spTree");
    const mShapes = cleanPlaceholderText(readShapeList(mTree, mctx), mTree, false, mctx.idMap);
    const mBg = readBg(mCSld, mctx) ?? { type: "solid", color: { scheme: "lt1" } };
    const tx = kid(mDoc, "txStyles");
    const base = buildTextStyles({ bg: { type: "none" } });
    const defaultTs = readLevels(kid(presDoc, "defaultTextStyle"), mctx);
    const otherLv = readLevels(kid(tx, "otherStyle"), mctx);
    const textStyles: TextStyles = {
      title: levelsToStyles(readLevels(kid(tx, "titleStyle"), mctx), base.title),
      body: levelsToStyles(readLevels(kid(tx, "bodyStyle"), mctx), base.body),
      other: levelsToStyles(otherLv ?? defaultTs, base.other),
    };
    if (mi === 0) master = { background: mBg, shapes: mShapes, textStyles };
    // İkincil asıl slaytın metin stilleri farklıysa düzen yer tutucularına taşınır.
    const ownStyles = mi > 0 && master && JSON.stringify(textStyles) !== JSON.stringify(master.textStyles) ? textStyles : undefined;
    const mDeco = mShapes.filter((x) => !(x.kind === "sp" && x.ph));

    // Düzenler
    const layoutRids = kids(kid(mDoc, "sldLayoutIdLst"), "sldLayoutId").map((e) => relId(e)).filter(Boolean) as string[];
    const layoutPaths = layoutRids.map((rid) => mRels.get(rid)?.target).filter(Boolean) as string[];
    for (const r of mRels.values()) if (r.type === REL.slideLayout && !layoutPaths.includes(r.target)) layoutPaths.push(r.target);
    for (const lp of layoutPaths) {
      if (layoutByPath.has(lp)) continue;
      const xml = await text(lp);
      if (!xml) continue;
      const doc = parseXml(xml).documentElement;
      const rels = parseRels(await text(relsPathFor(lp)), lp);
      const lctx = mkCtx(rels, readClrMap(kidPath(doc, "clrMapOvr", "overrideClrMapping"), mClr), (ph: Placeholder) => findPlaceholder(mShapes, { type: ph.type }));
      const cSld = kid(doc, "cSld");
      const tree = kid(cSld, "spTree");
      let shapes = cleanPlaceholderText(readShapeList(tree, lctx), tree, true, lctx.idMap);
      const t = attr(doc, "type") ?? "cust";
      let showMasterShapes = attr(doc, "showMasterSp") !== "0" && attr(doc, "showMasterSp") !== "false";
      let background = readBg(cSld, lctx);
      if (mi > 0) {
        // İkincil asıl slaytın görünümü düzene taşınır.
        background ??= mBg;
        if (showMasterShapes) shapes = [...mDeco.map((d) => ({ ...d, id: uid("s") })), ...shapes];
        showMasterShapes = false;
        const st = ownStyles;
        if (st) {
          shapes = shapes.map((sh) => {
            if (sh.kind !== "sp" || !sh.ph) return sh;
            const mph = findPlaceholder(mShapes, { type: sh.ph.type });
            const levels = mergeLevels(mergeLevels(st[placeholderKind(sh.ph)], mph?.text?.levels), sh.text?.levels);
            return { ...sh, text: { ...(sh.text ?? { paragraphs: [{ runs: [], level: 0, props: {} }] }), levels } };
          });
        }
      }
      const layout: Layout = {
        id: uid("L"),
        type: (LAYOUT_TYPES.has(t) ? t : "cust") as LayoutType,
        name: attr(cSld, "name") || LAYOUT_NAMES[t] || "Özel Düzen",
        shapes,
        background,
        showMasterShapes,
      };
      layouts.push(layout);
      layoutByPath.set(lp, layout);
      layoutMasterShapes.set(layout, mShapes);
    }
  }
  if (!master) throw new PptxError("Asıl slayt okunamadı.");
  const masterShapes = master.shapes;
  if (!layouts.length) {
    const l: Layout = { id: uid("L"), type: "blank", name: "Boş", shapes: [], showMasterShapes: true };
    layouts.push(l);
  }

  // Slaytlar
  const slides: Slide[] = [];
  const sldIds = kids(kid(presDoc, "sldIdLst"), "sldId");
  for (const sid of sldIds) {
    const rid = relId(sid);
    const sp = rid ? presRels.get(rid)?.target : undefined;
    if (!sp) continue;
    const xml = await text(sp);
    if (!xml) continue;
    const doc = parseXml(xml).documentElement;
    const rels = parseRels(await text(relsPathFor(sp)), sp);
    const lrel = [...rels.values()].find((r) => r.type === REL.slideLayout);
    const layout = (lrel && layoutByPath.get(lrel.target)) || layouts[0];
    const sctx = mkCtx(rels, readClrMap(kidPath(doc, "clrMapOvr", "overrideClrMapping"), clrMap), (ph: Placeholder) => findPlaceholder(layout.shapes, ph) ?? findPlaceholder(layoutMasterShapes.get(layout) ?? masterShapes, { type: ph.type }));
    const cSld = kid(doc, "cSld");
    const shapes = readShapeList(kid(cSld, "spTree"), sctx);
    const slide: Slide = {
      id: uid("sl"),
      layoutId: layout.id,
      shapes,
      background: readBg(cSld, sctx),
      hideMasterShapes: attr(doc, "showMasterSp") === "0" || attr(doc, "showMasterSp") === "false" || undefined,
      hidden: attr(doc, "show") === "0" || attr(doc, "show") === "false" || undefined,
      notes: "",
      transition: readTransition(doc),
      animations: readAnimations(doc, sctx.idMap),
    };
    const crel = [...rels.values()].find((r) => r.type === REL.comments);
    const cxml = crel ? await text(crel.target) : undefined;
    if (cxml) {
      const names = await commentAuthors();
      const list = kids(parseXml(cxml).documentElement, "cm").map((cm) => {
        const pos = kid(cm, "pos");
        const dt = Date.parse(attr(cm, "dt") ?? "");
        return {
          id: uid("c"),
          author: names.get(attr(cm, "authorId") ?? "") ?? "Yazar",
          text: kid(cm, "text")?.textContent ?? "",
          x: (numAttr(pos, "x") ?? 0) / 6,
          y: (numAttr(pos, "y") ?? 0) / 6,
          date: Number.isFinite(dt) ? dt : Date.now(),
        };
      });
      if (list.length) slide.comments = list;
    }
    const nrel = [...rels.values()].find((r) => r.type === REL.notesSlide);
    if (nrel) {
      const nxml = await text(nrel.target);
      if (nxml) {
        const ndoc = parseXml(nxml).documentElement;
        const nctx = mkCtx(parseRels(await text(relsPathFor(nrel.target)), nrel.target));
        for (const nsp of kids(kidPath(ndoc, "cSld", "spTree"), "sp")) {
          const ph = kidPath(nsp, "nvSpPr", "nvPr", "ph");
          if (attr(ph, "type") !== "body") continue;
          const body = readTextBody(kid(nsp, "txBody"), nctx);
          slide.notes = body ? body.paragraphs.map((p) => p.runs.map((r) => r.text.replace(/\v/g, "\n")).join("")).join("\n").replace(/\s+$/, "") : "";
          break;
        }
      }
    }
    slides.push(slide);
  }

  // Belge özellikleri
  const props: Presentation["props"] = {};
  const coreRel = [...rootRels.values()].find((r) => r.type === REL.coreProps);
  const core = coreRel ? await text(coreRel.target) : undefined;
  if (core) {
    const d = parseXml(core).documentElement;
    props.title = kid(d, "title")?.textContent || undefined;
    props.subject = kid(d, "subject")?.textContent || undefined;
    props.creator = kid(d, "creator")?.textContent || undefined;
    props.created = kid(d, "created")?.textContent || undefined;
    props.modified = kid(d, "modified")?.textContent || undefined;
  }

  // Üst bilgi / alt bilgi durumu slaytlardan çıkarılır.
  const hasPh = (t: string) => slides.some((s) => s.shapes.some((x) => x.kind === "sp" && x.ph?.type === t));
  const ftr = slides.flatMap((s) => s.shapes).find((x): x is SpShape => x.kind === "sp" && x.ph?.type === "ftr");
  const headerFooter = {
    ...DEFAULT_HF,
    dateTime: hasPh("dt"),
    slideNumber: hasPh("sldNum"),
    footer: hasPh("ftr"),
    footerText: ftr?.text ? ftr.text.paragraphs.map((p) => p.runs.map((r) => r.text).join("")).join(" ") : "",
  };

  return { width, height, theme, master, layouts, slides, headerFooter, props };
}
