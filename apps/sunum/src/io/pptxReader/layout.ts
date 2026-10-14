/**
 * Slayt düzenleri (slideLayout): yer tutucular, süs şekilleri, arka plan.
 * İkincil asıl slayda bağlı düzenlere o asıl slaydın görünümü (arka plan, süs
 * şekilleri, farklı metin stilleri ve teması) taşınır; model tek asıl slayt taşır.
 * Geçiş kancası (dalga 3): düzenin extLst'i ve tanınmayan spTree öğeleri burada
 * (readShapeList çevresinde) toplanacak.
 */
import { uid } from "../../model/factory";
import { findPlaceholder, placeholderKind } from "../../model/textStyle";
import type { Fill, Layout, LayoutType, LevelStyle, Placeholder, SchemeColorName, Shape, TextStyles, Theme } from "../../model/types";
import { LAYOUT_NAMES } from "../../themes/build";
import { readShapeList } from "../readDrawing";
import { attr, kid, kidPath, parseRels, relsPathFor, type Rel } from "../xml";
import { cleanPlaceholderText, readBg, readClrMap, type MkCtx } from "./common";
import type { PptxPackage } from "./package";

const LAYOUT_TYPES = new Set(["title", "obj", "secHead", "twoObj", "twoTxTwoObj", "titleOnly", "blank", "objTx", "picTx", "vertTx"]);

/** Düzenin bağlı olduğu asıl slaydın düzene aktarılan görünümü. */
export interface MasterLook {
  /** İkincil asıl slayt: görünümü düzene taşınır. */
  secondary: boolean;
  clrMap: Record<string, SchemeColorName>;
  shapes: Shape[];
  background: Fill;
  /** Yer tutucu olmayan (süs) şekiller. */
  deco: Shape[];
  /** İkincil asıl slaydın ana asıl slayttan farklı metin stilleri. */
  ownStyles?: TextStyles;
  /** İkincil asıl slaydın ana temadan farklı teması. */
  designTheme?: Theme;
}

/** Okunan düzenler (tüm asıl slaytlarınki). */
export interface LayoutSet {
  layouts: Layout[];
  byPath: Map<string, Layout>;
  /** Düzen → bağlı olduğu asıl slaydın şekilleri (yer tutucu kalıtımı için). */
  masterShapes: Map<Layout, Shape[]>;
  /** Okunamayan ya da bozuk düzen sayısı. */
  bad: number;
}

interface LayoutLook {
  shapes: Shape[];
  background?: Fill;
  showMasterShapes: boolean;
}

/** `lp` düzenini okuyup `set`'e ekler (daha önce okunduysa atlar). */
export async function readLayout(pk: PptxPackage, mkCtx: MkCtx, lp: string, m: MasterLook, set: LayoutSet): Promise<void> {
  if (set.byPath.has(lp)) return;
  const opened = await openLayout(pk, lp, set);
  if (!opened) return;
  const { doc, rels } = opened;
  const lctx = mkCtx(rels, readClrMap(kidPath(doc, "clrMapOvr", "overrideClrMapping"), m.clrMap), (ph: Placeholder) => findPlaceholder(m.shapes, { type: ph.type }));
  const cSld = kid(doc, "cSld");
  const tree = kid(cSld, "spTree");
  let look: LayoutLook = {
    shapes: cleanPlaceholderText(readShapeList(tree, lctx), tree, true, lctx.idMap),
    showMasterShapes: attr(doc, "showMasterSp") !== "0" && attr(doc, "showMasterSp") !== "false",
    background: readBg(cSld, lctx),
  };
  if (m.secondary) look = adoptMasterLook(look, m);
  const t = attr(doc, "type") ?? "cust";
  const layout: Layout = {
    id: uid("L"),
    type: (LAYOUT_TYPES.has(t) ? t : "cust") as LayoutType,
    name: attr(cSld, "name") || LAYOUT_NAMES[t] || "Özel Düzen",
    shapes: look.shapes,
    background: look.background,
    showMasterShapes: look.showMasterShapes,
    ...(m.designTheme ? { theme: m.designTheme } : {}),
  };
  set.layouts.push(layout);
  set.byPath.set(lp, layout);
  set.masterShapes.set(layout, m.shapes);
}

/** Düzen parçası ve ilişkileri; okunamazsa tanımsız. Sorunlar `set.bad`'e sayılır. */
async function openLayout(pk: PptxPackage, lp: string, set: LayoutSet): Promise<{ doc: Element; rels: Map<string, Rel> } | undefined> {
  const lpart = await pk.part(lp, "sldLayout");
  if (!lpart.root) {
    set.bad++;
    pk.W.detail(lpart.detail ?? lp);
    return undefined;
  }
  if (lpart.problem === "partial") {
    set.bad++;
    pk.W.detail(lpart.detail ?? lp);
  }
  const lRelsXml = await pk.text(relsPathFor(lp));
  if (lRelsXml === undefined && /\sr:(embed|id|link)="/.test(lpart.xml ?? "")) {
    set.bad++;
    pk.W.detail(`${relsPathFor(lp)}: pakette yok ya da okunamadı`);
  }
  return { doc: lpart.root, rels: parseRels(lRelsXml, lp) };
}

/** İkincil asıl slaytın görünümü düzene taşınır (arka plan, süs şekilleri, metin stilleri). */
function adoptMasterLook(look: LayoutLook, m: MasterLook): LayoutLook {
  const background = look.background ?? m.background;
  let shapes = look.showMasterShapes ? [...m.deco.map((d) => ({ ...d, id: uid("s") })), ...look.shapes] : look.shapes;
  const st = m.ownStyles;
  if (st) {
    shapes = shapes.map((sh) => {
      if (sh.kind !== "sp" || !sh.ph) return sh;
      const mph = findPlaceholder(m.shapes, { type: sh.ph.type });
      const levels = mergeLevels(mergeLevels(st[placeholderKind(sh.ph)], mph?.text?.levels), sh.text?.levels);
      return { ...sh, text: { ...(sh.text ?? { paragraphs: [{ runs: [], level: 0, props: {} }] }), levels } };
    });
  }
  return { shapes, background, showMasterShapes: false };
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
