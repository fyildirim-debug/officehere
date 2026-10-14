/**
 * Asıl slaytlar (slideMaster). Model tek asıl slayt taşır: ilki asıl slayt
 * olur; diğerlerinin arka planı, süs şekilleri, farklı metin stilleri ve
 * teması kendi düzenlerine aktarılır (tüm düzenler okunur). Asıl slayt ya da
 * ilişkileri bozuksa asıl slayt ve düzenler paket içinden kurtarılır.
 * Geçiş kancası (dalga 3): asıl slaydın extLst'i, p:hf ve tanınmayan öğeleri
 * burada (readMasterLook) toplanacak.
 */
import { uid } from "../../model/factory";
import type { LevelStyle, Master, SchemeColorName, TextStyles, Theme } from "../../model/types";
import { buildTextStyles } from "../../themes/build";
import { readLevels, readShapeList, type ReadCtx } from "../readDrawing";
import { kid, kids, parseRels, parseXml, relId, relsPathFor, REL, type Rel } from "../xml";
import { cleanPlaceholderText, readBg, readClrMap, type MkCtx } from "./common";
import { readLayout, type LayoutSet, type MasterLook } from "./layout";
import { PptxError, type Part, type PptxPackage } from "./package";
import type { PresPart } from "./presentation";
import { readTheme } from "./theme";

const MASTER_RE = /^ppt\/slideMasters\/[^/]+\.xml$/i;

/** Ana (ilk) asıl slayt. */
export interface MainMaster {
  /** Okunan parçanın yolu (kurtarıldıysa paket içinden bulunan). */
  path?: string;
  /** Sunu ilişkisinin gösterdiği yol. */
  relTarget?: string;
  part: Part;
  /** Okunamadıysa boş bir öğe (varsayılan asıl slayt). */
  doc: Element;
  rels: Map<string, Rel>;
  /** Sunu ilişkilerinin (sağlam olarak) başvurduğu asıl slaytlar: kurtarmada başka asıl slaydın yerine geçmez. */
  referenced: Set<string>;
  /** Asıl slayt paket içinden bulundu. */
  recovered: boolean;
  /** Asıl slayt ya da ilişkileri okunamadı: düzenler paket içinden bulunur (slaytlar düzenlerine bağlı kalsın). */
  recoverLayouts: boolean;
}

/** Ana asıl slayt (ilişkisi bozuksa paket içinden ilk okunabilen asıl slayt). */
export async function resolveMaster(pres: PresPart): Promise<MainMaster> {
  const { pk } = pres;
  const masterRid = relId(kid(kid(pres.doc, "sldMasterIdLst"), "sldMasterId"));
  const relTarget = (masterRid && pres.rels.get(masterRid)?.target) || [...pres.rels.values()].find((r) => r.type === REL.slideMaster)?.target;
  const referenced = new Set([...pres.rels.values()].filter((r) => r.type === REL.slideMaster && pk.zip.has(r.target)).map((r) => r.target));
  let path = relTarget;
  let part = await pk.part(path, "sldMaster");
  const found = part.root ? undefined : await recoverMaster(pk, path, referenced);
  if (found) ({ path, part } = found);
  if (!part.root) pk.W.add("Asıl slayt okunamadı; varsayılan asıl slayt kullanıldı.", part.detail);
  else if (part.problem === "partial") pk.W.add("Asıl slayt bozuk; yalnız okunabilen kısmı kullanıldı.", part.detail);
  const doc = part.root ?? parseXml("<x/>").documentElement;
  const relsXml = path && part.root ? await pk.text(relsPathFor(path)) : undefined;
  const rels = path && part.root ? parseRels(relsXml, path) : new Map<string, Rel>();
  if (part.root && relsXml === undefined) pk.W.add("Asıl slaydın ilişki bölümü eksik ya da bozuk; düzenler paket içinden bulundu, asıl slayttaki resimler kaybolmuş olabilir.", `${relsPathFor(path!)}: pakette yok ya da okunamadı`);
  return { path, relTarget, part, doc, rels, referenced, recovered: !!found, recoverLayouts: !part.root || relsXml === undefined };
}

async function recoverMaster(pk: PptxPackage, path: string | undefined, referenced: Set<string>): Promise<{ path: string; part: Part } | undefined> {
  for (const cand of pk.parts(MASTER_RE)) {
    if (cand === path || referenced.has(cand)) continue;
    const p = await pk.part(cand, "sldMaster");
    if (p.root) {
      pk.W.detail(`asıl slayt ${path ?? "(ilişki yok)"} yerine ${cand} kullanıldı`);
      return { path: cand, part: p };
    }
  }
  return undefined;
}

/** Asıl slayt okuma ortamı. */
export interface MasterEnv {
  pres: PresPart;
  main: MainMaster;
  mkCtx: MkCtx;
  /** Ana asıl slaydın renk eşlemesi. */
  clrMap: Record<string, SchemeColorName>;
  /** Sununun teması ve yolu (ikincil asıl slaydın teması bununla karşılaştırılır). */
  theme: Theme;
  themePath?: string;
}

/** Tüm asıl slaytları ve düzenlerini okur; ana asıl slayt okunamazsa PptxError. */
export async function readMasters(env: MasterEnv): Promise<{ master: Master; set: LayoutSet }> {
  const { pres, main } = env;
  const { pk } = pres;
  const masterPaths = masterPathList(pres, main);
  const orphans = main.recoverLayouts ? await orphanLayoutPaths(pk, masterPaths) : [];
  const set: LayoutSet = { layouts: [], byPath: new Map(), masterShapes: new Map(), bad: 0 };
  let master: Master | undefined;
  let badMasters = 0;
  for (const [mi, mp] of masterPaths.entries()) {
    const opened = mi === 0 ? { doc: main.doc, rels: main.rels } : await openSecondaryMaster(pk, mp);
    if (!opened) {
      badMasters++;
      continue;
    }
    const look = await readMasterLook(env, opened, master?.textStyles);
    if (mi === 0) master = { background: look.background, shapes: look.shapes, textStyles: look.textStyles };
    for (const lp of layoutPathsOf(opened.doc, opened.rels, mi === 0 ? orphans : [])) await readLayout(pk, env.mkCtx, lp, look, set);
  }
  if (badMasters) pk.W.add("Bazı asıl slaytlar okunamadı; onları kullanan slaytlar başka bir düzene bağlandı.");
  if (set.bad) pk.W.add("Bazı slayt düzenleri okunamadı ya da bozuk; ilgili slaytlar başka bir düzene bağlanmış olabilir.");
  if (!master) throw new PptxError("Asıl slayt okunamadı.");
  if (!set.layouts.length) set.layouts.push({ id: uid("L"), type: "blank", name: "Boş", shapes: [], showMasterShapes: true });
  return { master, set };
}

/** Okunacak asıl slaytlar: ilki ana asıl slayt (okunamadıysa ""), sonra sunudaki diğerleri. */
function masterPathList(pres: PresPart, main: MainMaster): string[] {
  const masterPaths = kids(kid(pres.doc, "sldMasterIdLst"), "sldMasterId")
    .map((e) => relId(e))
    .map((rid) => (rid ? pres.rels.get(rid)?.target : undefined))
    .filter((p, i, all): p is string => !!p && p !== main.path && p !== main.relTarget && all.indexOf(p) === i);
  // İlişkiler bozuksa başvurusuz diğer asıl slaytlar da (düzenleriyle) okunur.
  if (main.recovered) for (const cand of pres.pk.parts(MASTER_RE)) if (cand !== main.path && !main.referenced.has(cand) && !masterPaths.includes(cand)) masterPaths.push(cand);
  masterPaths.unshift(main.path && main.part.root ? main.path : "");
  return masterPaths;
}

/** Asıl slayt ya da ilişkileri okunamadıysa: diğer asıl slaytların kullanmadığı düzenler ilk asıl slayda bağlanır. */
async function orphanLayoutPaths(pk: PptxPackage, masterPaths: string[]): Promise<string[]> {
  const taken = new Set<string>();
  for (const mp of masterPaths.slice(1)) for (const r of parseRels(await pk.text(relsPathFor(mp)), mp).values()) if (r.type === REL.slideLayout) taken.add(r.target);
  return pk.parts(/^ppt\/slideLayouts\/[^/]+\.xml$/i).filter((p) => !taken.has(p));
}

async function openSecondaryMaster(pk: PptxPackage, mp: string): Promise<{ doc: Element; rels: Map<string, Rel> } | undefined> {
  const p = await pk.part(mp, "sldMaster");
  if (!p.root) {
    pk.W.detail(p.detail ?? mp);
    return undefined;
  }
  return { doc: p.root, rels: parseRels(await pk.text(relsPathFor(mp)), mp) };
}

/** Asıl slaydın görünümü; `mainStyles` verilirse (ikincil asıl slayt) farklı metin stilleri ve teması da. */
async function readMasterLook(env: MasterEnv, m: { doc: Element; rels: Map<string, Rel> }, mainStyles: TextStyles | undefined): Promise<MasterLook & { textStyles: TextStyles }> {
  const secondary = !!mainStyles;
  const clrMap = secondary ? readClrMap(kid(m.doc, "clrMap")) : env.clrMap;
  const mctx = env.mkCtx(m.rels, clrMap);
  const mCSld = kid(m.doc, "cSld");
  const mTree = kid(mCSld, "spTree");
  const shapes = cleanPlaceholderText(readShapeList(mTree, mctx), mTree, false, mctx.idMap);
  const background = readBg(mCSld, mctx) ?? { type: "solid", color: { scheme: "lt1" } };
  const textStyles = readTextStyles(env.pres.doc, kid(m.doc, "txStyles"), mctx);
  // İkincil asıl slaytın metin stilleri farklıysa düzen yer tutucularına taşınır.
  const ownStyles = secondary && JSON.stringify(textStyles) !== JSON.stringify(mainStyles) ? textStyles : undefined;
  const deco = shapes.filter((x) => !(x.kind === "sp" && x.ph));
  const designTheme = secondary ? await designThemeOf(env, m.rels) : undefined;
  return { secondary, clrMap, shapes, background, deco, ownStyles, designTheme, textStyles };
}

function readTextStyles(presDoc: Element, tx: Element | undefined, mctx: ReadCtx): TextStyles {
  const base = buildTextStyles({ bg: { type: "none" } });
  const defaultTs = readLevels(kid(presDoc, "defaultTextStyle"), mctx);
  const otherLv = readLevels(kid(tx, "otherStyle"), mctx);
  return {
    title: levelsToStyles(readLevels(kid(tx, "titleStyle"), mctx), base.title),
    body: levelsToStyles(readLevels(kid(tx, "bodyStyle"), mctx), base.body),
    other: levelsToStyles(otherLv ?? defaultTs, base.other),
  };
}

/** İkincil asıl slaydın teması farklıysa düzenleri o temayla çizilir (Seçili Slaytlara Uygula). */
async function designThemeOf(env: MasterEnv, mRels: Map<string, Rel>): Promise<Theme | undefined> {
  const tp = [...mRels.values()].find((r) => r.type === REL.theme)?.target;
  const t2 = tp && tp !== env.themePath ? readTheme(await env.pres.pk.text(tp)).theme : undefined;
  if (t2 && (JSON.stringify(t2.colors) !== JSON.stringify(env.theme.colors) || JSON.stringify(t2.fonts) !== JSON.stringify(env.theme.fonts))) return t2;
  return undefined;
}

/** Asıl slaydın düzenleri: sldLayoutIdLst sırası, sonra ilişkilerdeki diğerleri ve `extra`. */
function layoutPathsOf(mDoc: Element, mRels: Map<string, Rel>, extra: string[]): string[] {
  const layoutRids = kids(kid(mDoc, "sldLayoutIdLst"), "sldLayoutId").map((e) => relId(e)).filter(Boolean) as string[];
  const layoutPaths = layoutRids.map((rid) => mRels.get(rid)?.target).filter(Boolean) as string[];
  for (const r of mRels.values()) if (r.type === REL.slideLayout && !layoutPaths.includes(r.target)) layoutPaths.push(r.target);
  for (const lp of extra) if (!layoutPaths.includes(lp)) layoutPaths.push(lp);
  return layoutPaths;
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
