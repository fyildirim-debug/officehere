/**
 * PPTX okuyucu: slayt boyutu, sıra, asıl slayt/düzen kalıtımı, tema renkleri
 * ve yazı tipleri, şekiller, resimler, tablolar, gruplar, arka planlar,
 * notlar, gizli slaytlar, geçişler ve animasyonlar.
 */
import { uid } from "../model/factory";
import { mapSlideLinks } from "../model/links";
import { normalizeSections } from "../model/sections";
import { DEFAULT_HF, DEFAULT_NOTES_HF } from "../model/presentation";
import { findPlaceholder, placeholderKind } from "../model/textStyle";
import type { ChartData, CustomShow, Fill, Layout, LayoutType, LevelStyle, Master, NotesHeaderFooter, Placeholder, Presentation, SchemeColorName, Shape, Slide, SpShape, TextStyles, Theme } from "../model/types";
import { buildTextStyles, LAYOUT_NAMES } from "../themes/build";
import { getThemeDef } from "../themes/themes";
import { EXT_MIME, MEDIA_MIME } from "./pptxWriter";
import { readChartXml } from "./readChart";
import { DEFAULT_CLRMAP, readColor, readEffects, readFill, readLevels, readShapeList, readTextBody, type ReadCtx } from "./readDrawing";
import { readAnimations, readTransition } from "./readTiming";
import { DEFAULT_PRINT_SETTINGS, notesOrientation, readPresProps } from "./presProps";
import { attr, boolAttr, descendants, fromEmu, kid, kidPath, kids, numAttr, parseRels, parseXml, relId, relsPathFor, REL, type Rel } from "./xml";
import { LoadWarnings } from "./loadWarnings";
import { openZip, PackageError, PartError, xmlHealth, type SafeZip } from "./zipSafe";

/** PPTX okuma hatası: Türkçe ileti (+ günlük için teknik ayrıntı). */
export class PptxError extends PackageError {}

/** Not asıl sayfasındaki p:hf ve üst/alt bilgi metinleri; p:hf yoksa tanımsız (varsayılan). */
function readNotesHf(xml: string): NotesHeaderFooter | undefined {
  let root: Element;
  try {
    root = parseXml(xml).documentElement;
  } catch {
    return undefined;
  }
  const hf = kid(root, "hf");
  if (!hf) return undefined;
  const on = (k: string) => boolAttr(hf, k) !== false;
  const phText = (type: string) => {
    const sp = descendants(root, "sp").find((e) => attr(descendants(e, "ph")[0], "type") === type);
    return sp ? descendants(sp, "p").map((p) => descendants(p, "t").map((t) => t.textContent ?? "").join("")).join(" ").trim() : "";
  };
  const dtSp = descendants(root, "sp").find((e) => attr(descendants(e, "ph")[0], "type") === "dt");
  const dateAuto = !dtSp || descendants(dtSp, "fld").some((f) => (attr(f, "type") ?? "").startsWith("datetime"));
  return {
    ...DEFAULT_NOTES_HF,
    pageNumber: on("sldNum"),
    header: on("hdr"),
    headerText: phText("hdr"),
    footer: on("ftr"),
    footerText: phText("ftr"),
    dateTime: on("dt"),
    dateAuto,
    dateFixed: dateAuto ? "" : phText("dt"),
  };
}

function bytesToBase64(bytes: Uint8Array): string {
  let s = "";
  const chunk = 0x8000;
  for (let i = 0; i < bytes.length; i += chunk) s += String.fromCharCode(...bytes.subarray(i, i + chunk));
  return btoa(s);
}

const SCHEME_KEYS: SchemeColorName[] = ["dk1", "lt1", "dk2", "lt2", "accent1", "accent2", "accent3", "accent4", "accent5", "accent6", "hlink", "folHlink"];

function readTheme(xml: string | undefined): { theme: Theme; lnWidths: number[]; edited?: boolean; effectEls: Element[] } {
  const def: Theme = {
    name: "Tema",
    colors: { dk1: "#000000", lt1: "#FFFFFF", dk2: "#44546A", lt2: "#E7E6E6", accent1: "#4472C4", accent2: "#ED7D31", accent3: "#A5A5A5", accent4: "#FFC000", accent5: "#5B9BD5", accent6: "#70AD47", hlink: "#0563C1", folHlink: "#954F72" },
    fonts: { major: "Calibri Light", minor: "Calibri" },
  };
  if (!xml) return { theme: def, lnWidths: [0.75, 1.5, 2], effectEls: [] };
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
  let edited: boolean | undefined;
  for (const e of kids(kid(root, "extLst"), "ext")) {
    const t = kid(e, "theme");
    if (t && attr(t, "id") && getThemeDef(attr(t, "id"))) {
      theme.id = attr(t, "id");
      theme.variant = numAttr(t, "variant") ?? 0;
      edited = attr(t, "edited") === "1" || undefined;
    }
  }
  // Efekt stilleri (effectRef idx 1..3) ana slayt renk eşlemesi bilindikten sonra okunur.
  const effectEls = kids(kidPath(te, "fmtScheme", "effectStyleLst"), "effectStyle");
  return { theme, lnWidths: lnWidths.length ? lnWidths : [0.75, 1.5, 2], edited, effectEls };
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

/** Bir sunuda okunacak en çok slayt (bozuk sldIdLst'e karşı). */
const MAX_SLIDES = 5000;
const SLIDE_CT = "application/vnd.openxmlformats-officedocument.presentationml.slide+xml";
const MEDIA_RE = /\.(png|jpe?g|gif|bmp|svg|webp|tiff?|emf|wmf|mp4|m4v|webm|ogv|mov|mp3|m4a|wav|ogg|oga|aac|flac)$/i;

export interface PptxReadOptions {
  /** Açılış uyarıları (eksik/bozuk parçalar); verilmezse yok sayılır. */
  warnings?: LoadWarnings;
}

type PartProblem = "missing" | "corrupt" | "tooLarge" | "partial";
interface Part {
  root?: Element;
  xml?: string;
  problem?: PartProblem;
  detail?: string;
}

/** Yol sonundaki sayıya göre sıralama ("slide10" "slide9"dan sonra). */
const byNumber = (a: string, b: string) => {
  const na = Number(/(\d+)\.xml$/i.exec(a)?.[1] ?? 0);
  const nb = Number(/(\d+)\.xml$/i.exec(b)?.[1] ?? 0);
  return na - nb || a.localeCompare(b);
};

/** PPTX baytlarını sunu modeline okur. Eksik/bozuk parçalar `opts.warnings`'e yazılır. */
export async function readPptx(data: ArrayBuffer | Uint8Array, opts: PptxReadOptions = {}): Promise<Presentation> {
  const W = opts.warnings ?? new LoadWarnings();
  let pkg: SafeZip;
  try {
    pkg = await openZip(data, "Dosya geçerli bir PPTX (ZIP) paketi değil.");
  } catch (e) {
    if (e instanceof PackageError) throw new PptxError(e.message, e.detail);
    throw e;
  }
  /** Metin parçası; bozuk ya da sınırı aşan parça atlanır (verilirse uyarı yazılır). */
  const text = async (p: string, label?: string): Promise<string | undefined> => {
    try {
      return await pkg.text(p);
    } catch (e) {
      if (!(e instanceof PartError)) throw e;
      if (label) W.add(label, `${p}: ${e.detail ?? e.message}`);
      else W.detail(`${p}: ${e.detail ?? e.message}`);
      return undefined;
    }
  };
  /** XML parçası + sağlık durumu (eksik, bozuk, çok büyük, kısmen okunabilir). */
  const part = async (p: string | undefined, roots: string | string[]): Promise<Part> => {
    if (!p) return { problem: "missing", detail: "ilişki hedefi yok" };
    let xml: string | undefined;
    try {
      xml = await pkg.text(p);
    } catch (e) {
      if (!(e instanceof PartError)) throw e;
      return { problem: e.kind, detail: `${p}: ${e.detail ?? e.message}` };
    }
    if (xml === undefined) return { problem: "missing", detail: `${p}: pakette yok` };
    if (!xml.trim()) return { problem: "corrupt", detail: `${p}: boş parça` };
    const doc = parseXml(xml);
    const health = xmlHealth(doc, roots);
    if (health === "bad") {
      const r = doc.documentElement?.localName;
      return { problem: "corrupt", detail: `${p}: ${!r || r === "parsererror" ? "XML ayrıştırılamadı" : `beklenmeyen kök öğe <${r}>`}` };
    }
    return health === "partial" ? { root: doc.documentElement, xml, problem: "partial", detail: `${p}: XML yalnız kısmen ayrıştırılabildi` } : { root: doc.documentElement, xml };
  };

  // İçerik türleri: bozuk ilişkilerde parçaların türünü doğrulamak için.
  const contentTypes = new Map<string, string>();
  const ctXml = await text("[Content_Types].xml");
  if (ctXml) for (const o of kids(parseXml(ctXml).documentElement, "Override")) contentTypes.set((attr(o, "PartName") ?? "").replace(/^\//, ""), attr(o, "ContentType") ?? "");
  else W.add("Paketin içerik türleri bölümü ([Content_Types].xml) eksik ya da bozuk; dosya onarılarak açıldı.");
  const isSlidePart = (p: string) => (contentTypes.has(p) ? contentTypes.get(p) === SLIDE_CT : /^ppt\/slides\/[^/]+\.xml$/i.test(p));
  const pkgParts = (re: RegExp) => pkg.paths.filter((p) => re.test(p)).sort(byNumber);
  let recoveredFromPackage = false;

  // Ortam dosyaları
  const media = new Map<string, string>();
  const badMedia = new Set<string>();
  const mediaFiles = pkg.paths.filter((p) => MEDIA_RE.test(p));
  await Promise.all(
    mediaFiles.map(async (p) => {
      try {
        const bytes = await pkg.bytes(p);
        if (!bytes) return;
        const ext = p.split(".").pop()!.toLowerCase();
        media.set(p, `data:${EXT_MIME[ext] ?? MEDIA_MIME[ext] ?? "application/octet-stream"};base64,${bytesToBase64(bytes)}`);
      } catch (e) {
        if (!(e instanceof PartError)) throw e;
        badMedia.add(p);
        W.detail(`${p}: ${e.detail ?? e.message}`);
      }
    }),
  );
  if (badMedia.size) W.add(`${badMedia.size} resim ya da ortam dosyası bozuk ya da çok büyük olduğu için yüklenemedi; yerlerinde boş çerçeve görünür.`);
  const missingMedia = new Set<string>();

  const rootRels = parseRels(await text("_rels/.rels"), "");
  let presPath = [...rootRels.values()].find((r) => r.type === REL.officeDocument)?.target ?? "ppt/presentation.xml";
  if (!pkg.has(presPath) && pkg.has("ppt/presentation.xml")) presPath = "ppt/presentation.xml";
  const presPart = await part(presPath, "presentation");
  if (!presPart.root) {
    if (presPart.problem === "missing") throw new PptxError("Sunu bölümü (presentation.xml) bulunamadı.", presPart.detail);
    throw new PptxError(presPart.problem === "tooLarge" ? "Sunu bölümü (presentation.xml) olağandışı büyük; dosya bozuk ya da zararlı olabilir." : "Sunu bölümü (presentation.xml) bozuk; dosya açılamadı.", presPart.detail);
  }
  if (presPart.problem === "partial") W.add("Sunu bölümü (presentation.xml) bozuk; yalnız okunabilen kısmı kullanıldı.", presPart.detail);
  const presDoc = presPart.root;
  const presRelsXml = await text(relsPathFor(presPath));
  if (presRelsXml === undefined) W.detail(`${relsPathFor(presPath)}: pakette yok ya da okunamadı`);
  const presRels = parseRels(presRelsXml, presPath);
  /** İlişki hedefi (pakette varsa); yoksa ilişkiler bozukken alışılmış parça yolu. */
  const presPartOf = (type: string, fallback: RegExp): string | undefined => {
    const t = [...presRels.values()].find((r) => r.type === type)?.target;
    if (t && pkg.has(t)) return t;
    const alt = pkgParts(fallback)[0];
    if (alt) W.detail(`${type.split("/").pop()}: ${t ?? "(ilişki yok)"} yerine ${alt} kullanıldı`);
    return alt ?? t;
  };
  const sz = kid(presDoc, "sldSz");
  const width = fromEmu(numAttr(sz, "cx"), 1280) || 1280;
  const height = fromEmu(numAttr(sz, "cy"), 720) || 720;

  // Asıl slayt (ilişkisi bozuksa paket içinden ilk okunabilen asıl slayt).
  const masterRid = relId(kid(kid(presDoc, "sldMasterIdLst"), "sldMasterId"));
  const masterRelTarget = (masterRid && presRels.get(masterRid)?.target) || [...presRels.values()].find((r) => r.type === REL.slideMaster)?.target;
  // Sunu ilişkilerinin (sağlam olarak) başvurduğu asıl slaytlar: kurtarmada başka asıl slaydın yerine geçmez.
  const MASTER_RE = /^ppt\/slideMasters\/[^/]+\.xml$/i;
  const referencedMasters = new Set([...presRels.values()].filter((r) => r.type === REL.slideMaster && pkg.has(r.target)).map((r) => r.target));
  let masterPath = masterRelTarget;
  let masterPart = await part(masterPath, "sldMaster");
  let masterRecovered = false;
  if (!masterPart.root) {
    for (const cand of pkgParts(MASTER_RE)) {
      if (cand === masterPath || referencedMasters.has(cand)) continue;
      const p = await part(cand, "sldMaster");
      if (p.root) {
        W.detail(`asıl slayt ${masterPath ?? "(ilişki yok)"} yerine ${cand} kullanıldı`);
        masterPath = cand;
        masterPart = p;
        recoveredFromPackage = masterRecovered = true;
        break;
      }
    }
  }
  if (!masterPart.root) W.add("Asıl slayt okunamadı; varsayılan asıl slayt kullanıldı.", masterPart.detail);
  else if (masterPart.problem === "partial") W.add("Asıl slayt bozuk; yalnız okunabilen kısmı kullanıldı.", masterPart.detail);
  const masterDoc = masterPart.root ?? parseXml("<x/>").documentElement;
  const masterRelsXml = masterPath && masterPart.root ? await text(relsPathFor(masterPath)) : undefined;
  const masterRels = masterPath && masterPart.root ? parseRels(masterRelsXml, masterPath) : new Map<string, Rel>();
  // Asıl slayt ya da ilişkileri okunamazsa düzenler paket içinden bulunur (slaytlar düzenlerine bağlı kalsın).
  const recoverLayouts = !masterPart.root || masterRelsXml === undefined;
  if (masterPart.root && masterRelsXml === undefined) W.add("Asıl slaydın ilişki bölümü eksik ya da bozuk; düzenler paket içinden bulundu, asıl slayttaki resimler kaybolmuş olabilir.", `${relsPathFor(masterPath!)}: pakette yok ya da okunamadı`);
  /** Başka asıl sayfaların (diğer asıl slaytlar, not/dinleyici notu asılları) temaları: yedek tema olarak seçilmez. */
  const foreignThemes = async (): Promise<Set<string>> => {
    const out = new Set<string>();
    for (const p of pkgParts(/^ppt\/(slideMasters|notesMasters|handoutMasters)\/[^/]+\.xml$/i)) {
      if (p === masterPath) continue;
      for (const r of parseRels(await text(relsPathFor(p)), p).values()) if (r.type === REL.theme) out.add(r.target);
    }
    return out;
  };

  // Tema (ilişkisi bozuksa paket içindeki ilk okunabilen tema).
  let themePath = [...masterRels.values()].find((r) => r.type === REL.theme)?.target ?? [...presRels.values()].find((r) => r.type === REL.theme)?.target;
  let themePart = await part(themePath, "theme");
  if (!themePart.root) {
    const hadRel = !!themePath;
    const foreign = await foreignThemes();
    for (const cand of pkgParts(/^ppt\/theme\/[^/]+\.xml$/i)) {
      if (cand === themePath || foreign.has(cand)) continue;
      const p = await part(cand, "theme");
      if (p.root) {
        W.detail(`tema ${themePath ?? "(ilişki yok)"} yerine ${cand} kullanıldı`);
        themePath = cand;
        themePart = p;
        if (hadRel) recoveredFromPackage = true;
        break;
      }
    }
    if (!themePart.root && (hadRel || masterPart.root === undefined)) W.add("Tema okunamadı; varsayılan tema renkleri ve yazı tipleri kullanıldı.", themePart.detail);
  }
  const { theme, lnWidths, edited: masterEdited, effectEls } = readTheme(themePart.root ? themePart.xml : undefined);
  const clrMap = readClrMap(kid(masterDoc, "clrMap"));
  const effectStyles = effectEls.map((e) => readEffects(e, { rels: new Map(), media: new Map(), clrMap, idMap: new Map(), lnWidths }));

  // Grafik parçaları (önbellekli verileriyle) önceden okunur.
  const charts = new Map<string, ChartData>();
  let badCharts = 0;
  for (const cp of pkg.paths.filter((p) => /^ppt\/charts\/[^/]+\.xml$/i.test(p))) {
    const cx = await text(cp);
    if (!cx) {
      badCharts++;
      continue;
    }
    try {
      const data = readChartXml(cx, { rels: new Map(), media, clrMap, idMap: new Map(), lnWidths });
      if (data) charts.set(cp, data);
    } catch (e) {
      // Okunamayan grafik yedek resmiyle (varsa) gösterilir.
      badCharts++;
      W.detail(`${cp}: ${(e as Error)?.message ?? e}`);
    }
  }
  if (badCharts) W.add("Bazı grafikler okunamadı; varsa yedek resimleri gösteriliyor.");

  const mkCtx = (rels: Map<string, Rel>, map = clrMap, inherit?: ReadCtx["inherit"]): ReadCtx => ({ rels, media, clrMap: map, idMap: new Map(), inherit, lnWidths, charts, effectStyles, missingMedia });

  // Açıklama yazarları (yalnız gerektiğinde okunur).
  let authorsCache: Map<string, string> | undefined;
  const commentAuthors = async (): Promise<Map<string, string>> => {
    if (authorsCache) return authorsCache;
    authorsCache = new Map();
    const path = presPartOf(REL.commentAuthors, /^ppt\/commentAuthors\.xml$/i);
    const xml = path ? await text(path) : undefined;
    if (xml) for (const a of kids(parseXml(xml).documentElement, "cmAuthor")) authorsCache.set(attr(a, "id") ?? "", attr(a, "name") ?? "Yazar");
    return authorsCache;
  };

  // Asıl slaytlar: model tek asıl slayt taşır. İlki asıl slayt olur; diğerlerinin
  // arka planı ve süs şekilleri kendi düzenlerine aktarılır (tüm düzenler okunur).
  const masterIdEls = kids(kid(presDoc, "sldMasterIdLst"), "sldMasterId");
  const masterPaths = masterIdEls.map((e) => relId(e)).map((rid) => (rid ? presRels.get(rid)?.target : undefined)).filter((p, i, all): p is string => !!p && p !== masterPath && p !== masterRelTarget && all.indexOf(p) === i);
  // İlişkiler bozuksa başvurusuz diğer asıl slaytlar da (düzenleriyle) okunur.
  if (masterRecovered) for (const cand of pkgParts(MASTER_RE)) if (cand !== masterPath && !referencedMasters.has(cand) && !masterPaths.includes(cand)) masterPaths.push(cand);
  if (masterPath && masterPart.root) masterPaths.unshift(masterPath);
  else masterPaths.unshift("");
  // Asıl slayt ya da ilişkileri okunamadıysa: diğer asıl slaytların kullanmadığı düzenler ilk asıl slayda bağlanır.
  const orphanLayouts: string[] = [];
  if (recoverLayouts) {
    const taken = new Set<string>();
    for (const mp of masterPaths.slice(1)) for (const r of parseRels(await text(relsPathFor(mp)), mp).values()) if (r.type === REL.slideLayout) taken.add(r.target);
    orphanLayouts.push(...pkgParts(/^ppt\/slideLayouts\/[^/]+\.xml$/i).filter((p) => !taken.has(p)));
  }

  const layouts: Layout[] = [];
  const layoutByPath = new Map<string, Layout>();
  const layoutMasterShapes = new Map<Layout, Shape[]>();
  let master: Master | undefined;
  let badLayouts = 0;
  let badMasters = 0;
  for (const [mi, mp] of masterPaths.entries()) {
    let mDoc = masterDoc;
    let mRels = masterRels;
    if (mi > 0) {
      const p = await part(mp, "sldMaster");
      if (!p.root) {
        badMasters++;
        W.detail(p.detail ?? mp);
        continue;
      }
      mDoc = p.root;
      mRels = parseRels(await text(relsPathFor(mp)), mp);
    }
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
    // İkincil asıl slaydın teması farklıysa düzenleri o temayla çizilir (Seçili Slaytlara Uygula).
    let designTheme: Theme | undefined;
    if (mi > 0) {
      const tp = [...mRels.values()].find((r) => r.type === REL.theme)?.target;
      const t2 = tp && tp !== themePath ? readTheme(await text(tp)).theme : undefined;
      if (t2 && (JSON.stringify(t2.colors) !== JSON.stringify(theme.colors) || JSON.stringify(t2.fonts) !== JSON.stringify(theme.fonts))) designTheme = t2;
    }

    // Düzenler
    const layoutRids = kids(kid(mDoc, "sldLayoutIdLst"), "sldLayoutId").map((e) => relId(e)).filter(Boolean) as string[];
    const layoutPaths = layoutRids.map((rid) => mRels.get(rid)?.target).filter(Boolean) as string[];
    for (const r of mRels.values()) if (r.type === REL.slideLayout && !layoutPaths.includes(r.target)) layoutPaths.push(r.target);
    if (mi === 0) for (const lp of orphanLayouts) if (!layoutPaths.includes(lp)) layoutPaths.push(lp);
    for (const lp of layoutPaths) {
      if (layoutByPath.has(lp)) continue;
      const lpart = await part(lp, "sldLayout");
      if (!lpart.root) {
        badLayouts++;
        W.detail(lpart.detail ?? lp);
        continue;
      }
      if (lpart.problem === "partial") {
        badLayouts++;
        W.detail(lpart.detail ?? lp);
      }
      const doc = lpart.root;
      const lRelsXml = await text(relsPathFor(lp));
      if (lRelsXml === undefined && /\sr:(embed|id|link)="/.test(lpart.xml ?? "")) {
        badLayouts++;
        W.detail(`${relsPathFor(lp)}: pakette yok ya da okunamadı`);
      }
      const rels = parseRels(lRelsXml, lp);
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
        ...(designTheme ? { theme: designTheme } : {}),
      };
      layouts.push(layout);
      layoutByPath.set(lp, layout);
      layoutMasterShapes.set(layout, mShapes);
    }
  }
  if (badMasters) W.add("Bazı asıl slaytlar okunamadı; onları kullanan slaytlar başka bir düzene bağlandı.");
  if (badLayouts) W.add("Bazı slayt düzenleri okunamadı ya da bozuk; ilgili slaytlar başka bir düzene bağlanmış olabilir.");
  if (!master) throw new PptxError("Asıl slayt okunamadı.");
  const masterShapes = master.shapes;
  if (!layouts.length) {
    const l: Layout = { id: uid("L"), type: "blank", name: "Boş", shapes: [], showMasterShapes: true };
    layouts.push(l);
  }

  // Slaytlar: önce her sldId'nin parçası çözülür. İlişkisi bozuk (hedefi yok ya da
  // slayt olmayan bir parçaya işaret eden) girdiler, paket içinde hiçbir sağlam
  // girdinin kullanmadığı slayt parçalarıyla sırasıyla eşlenir (kurtarma).
  let sldIds = kids(kid(presDoc, "sldIdLst"), "sldId");
  if (sldIds.length > MAX_SLIDES) {
    W.add(`Sunuda ${sldIds.length.toLocaleString("tr-TR")} slayt girdisi var; yalnız ilk ${MAX_SLIDES.toLocaleString("tr-TR")} slayt açıldı.`);
    sldIds = sldIds.slice(0, MAX_SLIDES);
  }
  const relTarget = (sid: Element) => {
    const rid = relId(sid);
    return rid ? presRels.get(rid)?.target : undefined;
  };
  const used = new Set<string>();
  const slidePaths: (string | undefined)[] = sldIds.map((sid) => {
    const t = relTarget(sid);
    if (!t || used.has(t) || !pkg.has(t) || !isSlidePart(t)) return undefined;
    used.add(t);
    return t;
  });
  const failedIdx = slidePaths.flatMap((p, i) => (p ? [] : [i]));
  if (failedIdx.length) {
    const orphans = pkg.paths.filter((p) => isSlidePart(p) && !used.has(p) && /\.xml$/i.test(p) && !/\/_rels\//.test(p)).sort(byNumber);
    let k = 0;
    for (const i of failedIdx) {
      if (k >= orphans.length) break;
      W.detail(`${i + 1}. slayt: ilişki hedefi ${relTarget(sldIds[i]) ?? "(yok)"} yerine ${orphans[k]} kullanıldı`);
      slidePaths[i] = orphans[k++];
    }
    if (k) {
      recoveredFromPackage = true;
      W.add(`Sununun ilişki bölümü bozuk; ${k} slayt paket içinden bulunarak açıldı. Slayt sırasını denetleyin.`);
    }
  }
  const blankLayout = layouts.find((l) => l.type === "blank") ?? layouts[0];
  const blankSlide = (): Slide => ({ id: uid("sl"), layoutId: blankLayout.id, shapes: [], notes: "", animations: [] });

  const slides: Slide[] = [];
  const slideAt: (Slide | undefined)[] = [];
  let readable = 0;
  for (const [i, sp] of slidePaths.entries()) {
    const n = i + 1;
    const sp0 = sp ?? relTarget(sldIds[i]);
    const sPart = sp ? await part(sp, "sld") : ({ problem: "missing", detail: `${sp0 ?? "(ilişki yok)"}: slayt parçası yok` } as Part);
    if (!sPart.root) {
      W.slide(sPart.problem === "tooLarge" ? "tooLarge" : sPart.problem === "missing" ? "missing" : "corrupt", n, sPart.detail);
      slides.push(blankSlide());
      slideAt.push(undefined);
      continue;
    }
    if (sPart.problem === "partial") W.slide("partial", n, sPart.detail);
    try {
      const slide = await readSlide(sp!, sPart, n);
      slides.push(slide);
      slideAt.push(slide);
      readable++;
    } catch (e) {
      W.slide("corrupt", n, `${sp}: ${(e as Error)?.message ?? e}`);
      slides.push(blankSlide());
      slideAt.push(undefined);
    }
  }
  if (sldIds.length && !readable) throw new PptxError("Dosya bozuk: sunudaki slaytların hiçbiri okunamadı.", W.details.join("\n"));

  async function readSlide(sp: string, sPart: Part, n: number): Promise<Slide> {
    const doc = sPart.root!;
    const relsXml = await text(relsPathFor(sp));
    if (relsXml === undefined) W.slide("rels", n, `${relsPathFor(sp)}: pakette yok ya da okunamadı`);
    const rels = parseRels(relsXml, sp);
    const lrel = [...rels.values()].find((r) => r.type === REL.slideLayout);
    const layout = (lrel && layoutByPath.get(lrel.target)) || layouts[0];
    const sctx = mkCtx(rels, readClrMap(kidPath(doc, "clrMapOvr", "overrideClrMapping"), clrMap), (ph: Placeholder) => findPlaceholder(layout.shapes, ph) ?? findPlaceholder(layoutMasterShapes.get(layout) ?? masterShapes, { type: ph.type }));
    const cSld = kid(doc, "cSld");
    const shapes = readShapeList(kid(cSld, "spTree"), sctx);
    const slide: Slide = {
      id: uid("sl"),
      ...(attr(cSld, "name") ? { name: attr(cSld, "name") } : {}),
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
    if (crel) {
      const cPart = await part(crel.target, "cmLst");
      if (cPart.root) {
        const names = await commentAuthors();
        const list = kids(cPart.root, "cm").map((cm) => {
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
      if (cPart.problem) W.slide("comments", n, cPart.detail);
    }
    const nrel = [...rels.values()].find((r) => r.type === REL.notesSlide);
    if (nrel) {
      const nPart = await part(nrel.target, "notes");
      if (nPart.problem) W.slide("notes", n, nPart.detail);
      if (nPart.root) {
        const ndoc = nPart.root;
        const nctx = mkCtx(parseRels(await text(relsPathFor(nrel.target)), nrel.target));
        const nsps = kids(kidPath(ndoc, "cSld", "spTree"), "sp");
        const phType = (nsp: Element) => {
          const ph = kidPath(nsp, "nvSpPr", "nvPr", "ph");
          return ph ? (attr(ph, "type") ?? "body") : undefined;
        };
        // Not gövdesi yer tutucusu; yoksa (bazı yazıcılar) yer tutucu olmayan ilk metin kutusu.
        const nsp = nsps.find((nn) => phType(nn) === "body") ?? nsps.find((nn) => phType(nn) === undefined && !!kid(nn, "txBody"));
        if (nsp) {
          const body = readTextBody(kid(nsp, "txBody"), nctx);
          slide.notes = body ? body.paragraphs.map((p) => p.runs.map((r) => r.text.replace(/\v/g, "\n")).join("")).join("\n").replace(/\s+$/, "") : "";
        }
      }
    }
    return slide;
  }

  // Bölümler (p14:sectionLst): her bölümün ilk slaydına işaret konur.
  const sldIdNum = new Map<string, Slide>();
  sldIds.forEach((sid, i) => {
    const sl = slides[i];
    const num = attr(sid, "id");
    if (sl && num) sldIdNum.set(num, sl);
  });
  for (const ext of kids(kid(presDoc, "extLst"), "ext")) {
    const lst = kid(ext, "sectionLst");
    if (!lst) continue;
    for (const sec of kids(lst, "section")) {
      const first = kids(kid(sec, "sldIdLst"), "sldId")
        .map((e) => sldIdNum.get(attr(e, "id") ?? ""))
        .find((x): x is Slide => !!x);
      if (!first || first.section) continue;
      const k = slides.indexOf(first);
      if (k >= 0) slides[k] = { ...first, section: { id: uid("sec"), name: attr(sec, "name") ?? "Bölüm" } };
    }
  }
  const sectioned = normalizeSections(slides);
  if (sectioned !== slides) slides.splice(0, slides.length, ...sectioned);

  // Slayt atlama köprüleri (parça yolu) slayt kimliklerine çevrilir.
  const idByPath = new Map<string, string>();
  slidePaths.forEach((sp, i) => {
    if (sp && slideAt[i]) idByPath.set(sp, slides[i].id);
  });
  for (let k = 0; k < slides.length; k++) {
    slides[k] = mapSlideLinks(slides[k], (l) => (l.startsWith("#sldpath:") ? (idByPath.has(l.slice(9)) ? `#slide:${idByPath.get(l.slice(9))}` : undefined) : l));
  }

  // Özel slayt gösterileri
  const customShows: CustomShow[] = [];
  const customShowIds = new Map<string, string>();
  for (const cs of kids(kid(presDoc, "custShowLst"), "custShow")) {
    const ids = kids(kid(cs, "sldLst"), "sld")
      .map((e) => {
        const rid = relId(e);
        const sp = rid ? presRels.get(rid)?.target : undefined;
        return sp ? idByPath.get(sp) : undefined;
      })
      .filter((x): x is string => !!x);
    const show = { id: uid("cs"), name: attr(cs, "name") ?? `Özel Gösteri ${customShows.length + 1}`, slideIds: ids };
    if (attr(cs, "id") !== undefined) customShowIds.set(attr(cs, "id")!, show.id);
    customShows.push(show);
  }

  // Not sayfaları/dinleyici notları üst-alt bilgisi (not asıl sayfası p:hf ve yer tutucu metinleri).
  const nmRid = relId(kid(kid(presDoc, "notesMasterIdLst"), "notesMasterId"));
  const nmRel = nmRid && presRels.get(nmRid)?.target;
  const nmPath = nmRel && pkg.has(nmRel) ? nmRel : presPartOf(REL.notesMaster, /^ppt\/notesMasters\/[^/]+\.xml$/i);
  const nmXml = nmPath ? await text(nmPath) : undefined;
  const notesHeaderFooter = nmXml ? readNotesHf(nmXml) : undefined;

  // Gösteri ve yazdırma ayarları (presProps.xml); not sayfası yönü p:notesSz'den.
  const propsPath = presPartOf(REL.presProps, /^ppt\/presProps\.xml$/i);
  const presProps = readPresProps(propsPath ? await text(propsPath, "Gösteri ve yazdırma ayarları okunamadı; varsayılanlar kullanıldı.") : undefined, customShowIds);
  const notesSz = kid(presDoc, "notesSz");
  if (notesOrientation(numAttr(notesSz, "cx"), numAttr(notesSz, "cy")) === "landscape") presProps.printSettings = { ...(presProps.printSettings ?? DEFAULT_PRINT_SETTINGS), orientation: "landscape" };

  // Belge özellikleri
  const props: Presentation["props"] = {};
  const coreRel = [...rootRels.values()].find((r) => r.type === REL.coreProps)?.target;
  const corePath = coreRel && pkg.has(coreRel) ? coreRel : pkg.has("docProps/core.xml") ? "docProps/core.xml" : coreRel;
  const core = corePath ? await text(corePath, "Belge özellikleri (başlık, yazar) okunamadı.") : undefined;
  if (core) {
    const d = parseXml(core).documentElement;
    props.title = kid(d, "title")?.textContent || undefined;
    props.subject = kid(d, "subject")?.textContent || undefined;
    props.creator = kid(d, "creator")?.textContent || undefined;
    props.created = kid(d, "created")?.textContent || undefined;
    props.modified = kid(d, "modified")?.textContent || undefined;
  }

  // Pakette bulunamayan (bozuk olanlar dışında) resim/ortam başvuruları.
  const lost = [...missingMedia].filter((p) => !badMedia.has(p));
  if (lost.length) W.add(`${lost.length} resim ya da ortam dosyası pakette bulunamadı; yerlerinde boş çerçeve görünür.`, `eksik ortam: ${lost.slice(0, 20).join(", ")}`);
  if (recoveredFromPackage && !W.list().some((w) => w.startsWith("Sununun ilişki bölümü bozuk"))) W.add("Sununun ilişki bölümü bozuk; asıl slayt ve tema paket içinden bulunarak kullanıldı.");

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

  return { width, height, theme, master, layouts, slides, headerFooter, props, masterEdited, ...(customShows.length ? { customShows } : {}), ...presProps, ...(notesHeaderFooter ? { notesHeaderFooter } : {}) };
}
