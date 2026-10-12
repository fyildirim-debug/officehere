/** DrawingML renk, dolgu, çizgi ve efekt okuyucuları. */
import type {
  ArrowType,
  Color,
  DashType,
  Effects,
  Fill,
  Line,
  Placeholder,
  SchemeColorName,
  SpShape,
} from "../model/types";
import { attr, fromEmu, kid, kidPath, kids, numAttr, type Rel } from "./xml";

export interface ReadCtx {
  rels: Map<string, Rel>;
  media: Map<string, string>;
  /** schemeClr adı → tema yuvası (clrMap). */
  clrMap: Record<string, SchemeColorName>;
  /** cNvPr kimliği → model kimliği. */
  idMap: Map<number, string>;
  /** Yer tutucu kalıtımı (düzen/asıl). */
  inherit?: (ph: Placeholder) => SpShape | undefined;
  /** Tema satır kalınlıkları (lnStyleLst, pt). */
  lnWidths: number[];
  phClr?: Color;
}

export const DEFAULT_CLRMAP: Record<string, SchemeColorName> = {
  bg1: "lt1",
  tx1: "dk1",
  bg2: "lt2",
  tx2: "dk2",
  accent1: "accent1",
  accent2: "accent2",
  accent3: "accent3",
  accent4: "accent4",
  accent5: "accent5",
  accent6: "accent6",
  hlink: "hlink",
  folHlink: "folHlink",
  dk1: "dk1",
  lt1: "lt1",
  dk2: "dk2",
  lt2: "lt2",
};

const PRESET_COLORS: Record<string, string> = {
  black: "#000000",
  white: "#FFFFFF",
  red: "#FF0000",
  green: "#008000",
  lime: "#00FF00",
  blue: "#0000FF",
  yellow: "#FFFF00",
  gray: "#808080",
  grey: "#808080",
  silver: "#C0C0C0",
  orange: "#FFA500",
  purple: "#800080",
  navy: "#000080",
  maroon: "#800000",
  teal: "#008080",
  cyan: "#00FFFF",
  magenta: "#FF00FF",
  ltGray: "#D3D3D3",
  dkGray: "#A9A9A9",
};

const COLOR_TAGS = new Set(["srgbClr", "schemeClr", "sysClr", "prstClr", "scrgbClr", "hslClr"]);

function hslToHex(h: number, s: number, l: number): string {
  const f = (n: number) => {
    const k = (n + h / 30) % 12;
    const a = s * Math.min(l, 1 - l);
    return Math.round(255 * (l - a * Math.max(-1, Math.min(k - 3, 9 - k, 1))));
  };
  return `#${[f(0), f(8), f(4)].map((v) => v.toString(16).padStart(2, "0")).join("")}`.toUpperCase();
}

/** Renk öğesini (srgbClr vb.) okur. */
export function readColorEl(el: Element, ctx: ReadCtx): Color | undefined {
  let c: Color | undefined;
  switch (el.localName) {
    case "srgbClr":
      c = { rgb: `#${(attr(el, "val") ?? "000000").toUpperCase()}` };
      break;
    case "schemeClr": {
      const v = attr(el, "val") ?? "tx1";
      if (v === "phClr") {
        c = ctx.phClr ? { ...ctx.phClr } : { scheme: "accent1" };
        break;
      }
      c = { scheme: ctx.clrMap[v] ?? DEFAULT_CLRMAP[v] ?? "dk1" };
      break;
    }
    case "sysClr":
      c = { rgb: `#${(attr(el, "lastClr") ?? (attr(el, "val") === "window" ? "FFFFFF" : "000000")).toUpperCase()}` };
      break;
    case "prstClr":
      c = { rgb: PRESET_COLORS[attr(el, "val") ?? "black"] ?? "#000000" };
      break;
    case "scrgbClr": {
      const t = (k: string) => Math.round(Math.max(0, Math.min(1, (numAttr(el, k) ?? 0) / 100000)) ** (1 / 2.2) * 255);
      c = { rgb: `#${[t("r"), t("g"), t("b")].map((v) => v.toString(16).padStart(2, "0")).join("")}`.toUpperCase() };
      break;
    }
    case "hslClr":
      c = { rgb: hslToHex((numAttr(el, "hue") ?? 0) / 60000, (numAttr(el, "sat") ?? 0) / 100000, (numAttr(el, "lum") ?? 0) / 100000) };
      break;
    default:
      return undefined;
  }
  for (const m of kids(el)) {
    const v = (numAttr(m, "val") ?? 0) / 100000;
    switch (m.localName) {
      case "lumMod":
        c.lumMod = v;
        break;
      case "lumOff":
        c.lumOff = v;
        break;
      case "tint":
        c.tint = v;
        break;
      case "shade":
        c.shade = v;
        break;
      case "alpha":
        c.alpha = v;
        break;
      default:
        break;
    }
  }
  return c;
}

/** Bir kapsayıcıdaki ilk renk öğesi. */
export function readColor(parent: Element | undefined, ctx: ReadCtx): Color | undefined {
  if (!parent) return undefined;
  for (const k of kids(parent)) if (COLOR_TAGS.has(k.localName)) return readColorEl(k, ctx);
  return undefined;
}

export function mediaFor(rid: string | undefined, ctx: ReadCtx): string | undefined {
  if (!rid) return undefined;
  const rel = ctx.rels.get(rid);
  if (!rel || rel.external) return undefined;
  return ctx.media.get(rel.target);
}

/** Dolgu öğesi (noFill/solidFill/…); bulunamazsa undefined. */
export function readFill(parent: Element | undefined, ctx: ReadCtx): Fill | undefined {
  if (!parent) return undefined;
  for (const el of kids(parent)) {
    switch (el.localName) {
      case "noFill":
        return { type: "none" };
      case "solidFill": {
        const c = readColor(el, ctx);
        return c ? { type: "solid", color: c } : { type: "none" };
      }
      case "gradFill": {
        const stops = kids(kid(el, "gsLst"), "gs")
          .map((gs) => ({ pos: (numAttr(gs, "pos") ?? 0) / 100000, color: readColor(gs, ctx) ?? { rgb: "#000000" } }))
          .sort((a, b) => a.pos - b.pos);
        if (!stops.length) return { type: "none" };
        const lin = kid(el, "lin");
        const path = kid(el, "path");
        return { type: "gradient", kind: path ? "radial" : "linear", angle: lin ? (numAttr(lin, "ang") ?? 0) / 60000 : 90, stops };
      }
      case "blipFill": {
        const src = mediaFor(attr(kid(el, "blip"), "embed"), ctx);
        if (!src) return { type: "none" };
        const amt = numAttr(kidPath(el, "blip", "alphaModFix"), "amt");
        return { type: "image", src, mode: kid(el, "tile") ? "tile" : "stretch", alpha: amt !== undefined ? amt / 100000 : undefined };
      }
      case "pattFill": {
        const c = readColor(kid(el, "fgClr"), ctx);
        return c ? { type: "solid", color: c } : { type: "none" };
      }
      case "grpFill":
        return { type: "none" };
      default:
        break;
    }
  }
  return undefined;
}

const DASHES = new Set(["solid", "sysDot", "sysDash", "dash", "dashDot", "lgDash", "lgDashDot", "lgDashDotDot"]);

export function readLine(ln: Element | undefined, ctx: ReadCtx, base?: Line): Line | undefined {
  if (!ln) return base;
  const w = numAttr(ln, "w");
  const f = readFill(ln, ctx);
  const dashV = attr(kid(ln, "prstDash"), "val");
  const head = attr(kid(ln, "headEnd"), "type") as ArrowType | undefined;
  const tail = attr(kid(ln, "tailEnd"), "type") as ArrowType | undefined;
  const line: Line = {
    fill: f ? (f.type === "solid" ? f : f.type === "gradient" ? { type: "solid", color: f.stops[0].color } : { type: "none" }) : (base?.fill ?? { type: "none" }),
    width: w !== undefined ? w / 12700 : (base?.width ?? 0.75),
    dash: dashV && DASHES.has(dashV) ? (dashV as DashType) : dashV ? "dash" : (base?.dash ?? "solid"),
    head: head ?? base?.head,
    tail: tail ?? base?.tail,
  };
  const cap = attr(ln, "cap");
  if (cap === "rnd" || cap === "sq" || cap === "flat") line.cap = cap;
  else if (base?.cap) line.cap = base.cap;
  if (line.head === "none") line.head = undefined;
  if (line.tail === "none") line.tail = undefined;
  return line;
}

export function readEffects(parent: Element | undefined, ctx: ReadCtx): Effects | undefined {
  const lst = kid(parent, "effectLst");
  if (!lst) return undefined;
  const fx: Effects = {};
  const sh = kid(lst, "outerShdw");
  if (sh) fx.shadow = { color: readColor(sh, ctx) ?? { rgb: "#000000", alpha: 0.4 }, blur: fromEmu(numAttr(sh, "blurRad")), dist: fromEmu(numAttr(sh, "dist")), dir: (numAttr(sh, "dir") ?? 0) / 60000 };
  const glow = kid(lst, "glow");
  if (glow) fx.glow = { color: readColor(glow, ctx) ?? { scheme: "accent1" }, radius: fromEmu(numAttr(glow, "rad")) };
  if (kid(lst, "reflection")) fx.reflection = true;
  const se = kid(lst, "softEdge");
  if (se) fx.softEdge = fromEmu(numAttr(se, "rad"));
  return fx.shadow || fx.glow || fx.reflection || fx.softEdge ? fx : undefined;
}

/* ------------------------------------------------------------------ */
/* Metin                                                                */
/* ------------------------------------------------------------------ */
