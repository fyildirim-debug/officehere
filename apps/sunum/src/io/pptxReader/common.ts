/**
 * Asıl slayt, düzen ve slayt okuyucularının ortak yardımcıları: renk eşlemesi
 * (clrMap), arka plan ve yer tutucu örnek metinlerinin boşaltılması.
 */
import type { Fill, SchemeColorName, Shape } from "../../model/types";
import { DEFAULT_CLRMAP, readColor, readFill, type ReadCtx } from "../readDrawing";
import { attr, kid, kidPath, kids, numAttr, type Rel } from "../xml";

/** Parçanın ilişkileriyle şekil okuma bağlamı (ortam, grafikler, tema çizgi/efekt stilleri ortak). */
export type MkCtx = (rels: Map<string, Rel>, map?: Record<string, SchemeColorName>, inherit?: ReadCtx["inherit"]) => ReadCtx;

export function readClrMap(el: Element | undefined, base: Record<string, SchemeColorName> = DEFAULT_CLRMAP): Record<string, SchemeColorName> {
  if (!el) return base;
  const map = { ...base };
  for (const k of ["bg1", "tx1", "bg2", "tx2", "accent1", "accent2", "accent3", "accent4", "accent5", "accent6", "hlink", "folHlink"]) {
    const v = attr(el, k) as SchemeColorName | undefined;
    if (v) map[k] = v;
  }
  return map;
}

export function readBg(cSld: Element | undefined, ctx: ReadCtx): Fill | undefined {
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

/** Asıl/düzen yer tutucularının örnek metnini boşaltır (özel istem metni korunur). */
export function cleanPlaceholderText(shapes: Shape[], tree: Element | undefined, isLayout: boolean, idMap: Map<number, string>): Shape[] {
  const inverse = new Map([...idMap.entries()].map(([k, v]) => [v, k]));
  const customIds = customPromptIds(tree);
  return shapes.map((s) => {
    if (s.kind !== "sp" || !s.ph || !s.text) return s;
    if (s.ph.type === "dt" || s.ph.type === "ftr" || s.ph.type === "sldNum") return s;
    const plain = s.text.paragraphs.map((p) => p.runs.map((r) => r.text).join("")).join("\n");
    const nid = inverse.get(s.id);
    const custom = isLayout && !!plain && nid !== undefined && customIds.has(nid);
    return { ...s, text: { ...s.text, prompt: custom ? plain : undefined, paragraphs: [{ runs: [], level: 0, props: {} }] } };
  });
}

/** Özel istem metni olan yer tutucuların cNvPr kimlikleri. */
function customPromptIds(tree: Element | undefined): Set<number> {
  const customIds = new Set<number>();
  if (!tree) return customIds;
  for (const sp of kids(tree, "sp")) {
    const ph = kidPath(sp, "nvSpPr", "nvPr", "ph");
    if (ph && (attr(ph, "hasCustomPrompt") === "1" || attr(ph, "hasCustomPrompt") === "true")) {
      const id = numAttr(kidPath(sp, "nvSpPr", "cNvPr"), "id");
      if (id !== undefined) customIds.add(id);
    }
  }
  return customIds;
}
