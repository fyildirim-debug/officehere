/**
 * Tema parçası: renk düzeni, yazı tipleri, çizgi kalınlıkları ve efekt
 * stilleri; tema ilişkisi bozuksa paket içinden kurtarma.
 */
import type { SchemeColorName, Theme } from "../../model/types";
import { getThemeDef } from "../../themes/themes";
import { attr, kid, kidPath, kids, numAttr, parseRels, parseXml, relsPathFor, REL } from "../xml";
import type { MainMaster } from "./master";
import type { Part, PptxPackage } from "./package";
import type { PresPart } from "./presentation";

const SCHEME_KEYS: SchemeColorName[] = ["dk1", "lt1", "dk2", "lt2", "accent1", "accent2", "accent3", "accent4", "accent5", "accent6", "hlink", "folHlink"];

export function readTheme(xml: string | undefined): { theme: Theme; lnWidths: number[]; edited?: boolean; effectEls: Element[] } {
  const def: Theme = {
    name: "Tema",
    colors: { dk1: "#000000", lt1: "#FFFFFF", dk2: "#44546A", lt2: "#E7E6E6", accent1: "#4472C4", accent2: "#ED7D31", accent3: "#A5A5A5", accent4: "#FFC000", accent5: "#5B9BD5", accent6: "#70AD47", hlink: "#0563C1", folHlink: "#954F72" },
    fonts: { major: "Calibri Light", minor: "Calibri" },
  };
  if (!xml) return { theme: def, lnWidths: [0.75, 1.5, 2], effectEls: [] };
  const doc = parseXml(xml);
  const root = doc.documentElement;
  const te = kid(root, "themeElements");
  const colors = readSchemeColors(kid(te, "clrScheme"), def.colors);
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

function readSchemeColors(cs: Element | undefined, base: Theme["colors"]): Theme["colors"] {
  const colors = { ...base };
  for (const k of SCHEME_KEYS) {
    const el = kid(cs, k);
    const c = kids(el)[0];
    if (!c) continue;
    if (c.localName === "srgbClr") colors[k] = `#${(attr(c, "val") ?? "000000").toUpperCase()}`; // yerel-ayar: ascii
    else if (c.localName === "sysClr") colors[k] = `#${(attr(c, "lastClr") ?? (k.startsWith("lt") ? "FFFFFF" : "000000")).toUpperCase()}`; // yerel-ayar: ascii
  }
  return colors;
}

/** Ana asıl slaydın teması (ilişkisi bozuksa paket içindeki ilk okunabilen tema). */
export async function resolveTheme(pres: PresPart, main: MainMaster): Promise<{ path?: string; part: Part; recovered: boolean }> {
  const { pk } = pres;
  const path = [...main.rels.values()].find((r) => r.type === REL.theme)?.target ?? [...pres.rels.values()].find((r) => r.type === REL.theme)?.target;
  const part = await pk.part(path, "theme");
  if (part.root) return { path, part, recovered: false };
  const hadRel = !!path;
  const found = await recoverTheme(pk, path, main.path);
  if (!found && (hadRel || main.part.root === undefined)) pk.W.add("Tema okunamadı; varsayılan tema renkleri ve yazı tipleri kullanıldı.", part.detail);
  return found ? { ...found, recovered: hadRel } : { path, part, recovered: false };
}

/** Paket içinde, başka asıl sayfalara ait olmayan ilk okunabilen tema. */
async function recoverTheme(pk: PptxPackage, path: string | undefined, masterPath: string | undefined): Promise<{ path: string; part: Part } | undefined> {
  const foreign = await foreignThemes(pk, masterPath);
  for (const cand of pk.parts(/^ppt\/theme\/[^/]+\.xml$/i)) {
    if (cand === path || foreign.has(cand)) continue;
    const p = await pk.part(cand, "theme");
    if (p.root) {
      pk.W.detail(`tema ${path ?? "(ilişki yok)"} yerine ${cand} kullanıldı`);
      return { path: cand, part: p };
    }
  }
  return undefined;
}

/** Başka asıl sayfaların (diğer asıl slaytlar, not/dinleyici notu asılları) temaları: yedek tema olarak seçilmez. */
async function foreignThemes(pk: PptxPackage, masterPath: string | undefined): Promise<Set<string>> {
  const out = new Set<string>();
  for (const p of pk.parts(/^ppt\/(slideMasters|notesMasters|handoutMasters)\/[^/]+\.xml$/i)) {
    if (p === masterPath) continue;
    for (const r of parseRels(await pk.text(relsPathFor(p)), p).values()) if (r.type === REL.theme) out.add(r.target);
  }
  return out;
}
