// ODT stil çözümleyici: yazı tipi bildirimleri, adlandırılmış ve otomatik stiller (üst stil zinciri),
// varsayılan stiller, liste stilleri, sayfa düzenleri ve asıl sayfalar (OpenDocument 1.3).
import type { CharProps, UnderlineKind } from "../common";
import { lengthPt, normColor } from "../common";
import { attr, kid, kids } from "../docx/xml";
import { odfToPicture } from "./dataStyles";

export interface ParaProps {
  align?: string;
  indentLeft?: number;
  indentRight?: number;
  indentFirst?: number;
  spaceBefore?: number;
  spaceAfter?: number;
  lineHeight?: number;
  lineRule?: "auto" | "exact" | "atLeast";
  shading?: string;
  keepNext?: boolean;
  keepLines?: boolean;
  breakBefore?: "page" | "column";
  breakAfter?: "page" | "column";
  borderBottom?: boolean;
  borderOther?: boolean;
  /** style:drop-cap (satır sayısı, uzaklık pt, harf stili). */
  dropCap?: { lines: number; distance: number; textStyle: string | null } | null;
}

export interface OdtStyle {
  name: string;
  display: string;
  family: string;
  parent: string | null;
  auto: boolean;
  listStyle: string | null;
  masterPage: string | null;
  text: CharProps & { sizePct?: number };
  para: ParaProps;
  /** Tablo / hücre / sütun özellikleri. */
  cell: { bg?: string; valign?: string; border?: boolean };
  table: { width?: number; align?: string; breakBefore?: boolean; marginLeft?: number };
  column: { width?: number; rel?: number };
  /** Satır yüksekliği (pt): style:row-height tam, style:min-row-height en az. */
  row: { height?: number; exact?: boolean };
  graphic: { wrap?: string };
}

export interface ListLevel {
  bullet: boolean;
  char: string;
  format: string;
  suffix: string;
  start: number;
  display: number;
}

export interface PageLayout {
  width?: number;
  height?: number;
  landscape?: boolean;
  margin: { top?: number; bottom?: number; left?: number; right?: number };
  columns?: number;
  columnGap?: number;
  header?: { minHeight: number; spacing: number };
  footer?: { minHeight: number; spacing: number };
  /** Sayfa rengi (fo:background-color ya da draw:fill-color). */
  background?: string | null;
}

export interface MasterPage {
  name: string;
  layout: string | null;
  header: Element | null;
  footer: Element | null;
  headerFirst: Element | null;
  footerFirst: Element | null;
  headerLeft: Element | null;
  footerLeft: Element | null;
}

/** "Heading_20_1" → "Heading 1" (ODF stil adı kodlaması). */
export const decodeName = (n: string) => n.replace(/_([0-9a-f]{2})_/gi, (_, h: string) => String.fromCharCode(parseInt(h, 16)));

const NS_FO = "urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0";

/** Önekli öznitelik okuma (yerel ad + isteğe bağlı ad alanı denetimi). */
function a(el: Element | null, name: string): string | null {
  if (!el) return null;
  return attr(el, name);
}

function underlineOf(tp: Element): { kind: UnderlineKind | null; color: string | null } | null {
  const style = a(tp, "text-underline-style");
  if (style === null) return null;
  if (style === "none") return { kind: null, color: null };
  const type = a(tp, "text-underline-type");
  const width = a(tp, "text-underline-width");
  const mode = a(tp, "text-underline-mode");
  const color = a(tp, "text-underline-color");
  let kind: UnderlineKind = "single";
  if (type === "double") kind = "double";
  else if (style === "wave") kind = "wave";
  else if (style === "dotted") kind = "dotted";
  else if (/dash/.test(style)) kind = "dashed";
  else if (width === "bold" || width === "thick") kind = "thick";
  else if (mode === "skip-white-space") kind = "words";
  return { kind, color: color && color !== "font-color" ? normColor(color) : null };
}

/** style:text-properties → karakter özellikleri. */
export function parseTextProps(tp: Element | null, fonts: Map<string, string>): CharProps & { sizePct?: number } {
  const p: CharProps & { sizePct?: number } = {};
  if (!tp) return p;
  const fn = a(tp, "font-name");
  const ff = tp.getAttributeNS(NS_FO, "font-family") ?? null;
  if (fn) p.font = fonts.get(fn) ?? decodeName(fn);
  else if (ff) p.font = ff.replace(/^['"]|['"]$/g, "").split(/['"]?,/)[0].trim();
  const fs = tp.getAttributeNS(NS_FO, "font-size");
  if (fs) {
    if (fs.endsWith("%")) p.sizePct = parseFloat(fs) / 100;
    else {
      const v = lengthPt(fs);
      if (v && v > 0 && v < 1700) p.size = v;
    }
  }
  const w = tp.getAttributeNS(NS_FO, "font-weight");
  if (w) p.bold = w === "bold" || (/^\d+$/.test(w) && Number(w) >= 600);
  const st = tp.getAttributeNS(NS_FO, "font-style");
  if (st) p.italic = st === "italic" || st === "oblique";
  const ul = underlineOf(tp);
  if (ul) {
    p.underline = ul.kind;
    p.underlineColor = ul.color;
  }
  const lt = a(tp, "text-line-through-style");
  if (lt) {
    const dbl = a(tp, "text-line-through-type") === "double";
    p.strike = lt !== "none" && !dbl;
    p.dstrike = lt !== "none" && dbl;
  }
  const pos = a(tp, "text-position");
  if (pos) {
    const first = pos.trim().split(/\s+/)[0];
    const v = first === "super" ? 1 : first === "sub" ? -1 : parseFloat(first);
    p.sup = v > 0;
    p.sub = v < 0;
  }
  const color = tp.getAttributeNS(NS_FO, "color");
  if (color) p.color = normColor(color);
  const bg = tp.getAttributeNS(NS_FO, "background-color");
  if (bg) p.highlight = bg === "transparent" ? null : normColor(bg);
  const tt = tp.getAttributeNS(NS_FO, "text-transform");
  if (tt) p.caps = tt === "uppercase";
  const fv = tp.getAttributeNS(NS_FO, "font-variant");
  if (fv) p.smallCaps = fv === "small-caps";
  const disp = a(tp, "display");
  if (disp) p.hidden = disp === "none";
  const ls = tp.getAttributeNS(NS_FO, "letter-spacing");
  // cm → pt dönüşümünün yuvarlama artığı atılır (OOXML gibi 1/20 pt duyarlılık).
  const lsPt = ls && ls !== "normal" ? lengthPt(ls) : null;
  if (ls) p.spacing = lsPt === null ? null : Math.round(lsPt * 20) / 20;
  const sh = tp.getAttributeNS(NS_FO, "text-shadow");
  if (sh) p.shadow = sh !== "none";
  const ol = a(tp, "text-outline");
  if (ol) p.outline = ol === "true";
  const rel = a(tp, "font-relief");
  if (rel) {
    p.emboss = rel === "embossed";
    p.imprint = rel === "engraved";
  }
  return p;
}

/** style:paragraph-properties → paragraf özellikleri. */
export function parseParaProps(pp: Element | null): ParaProps {
  const p: ParaProps = {};
  if (!pp) return p;
  const fo = (n: string) => pp.getAttributeNS(NS_FO, n);
  const al = fo("text-align");
  if (al) p.align = al === "start" || al === "left" ? "left" : al === "end" || al === "right" ? "right" : al === "center" ? "center" : al === "justify" ? "justify" : undefined;
  const all = lengthPt(fo("margin"));
  const len = (n: string) => lengthPt(fo(n)) ?? all ?? undefined;
  const ml = len("margin-left");
  if (ml !== undefined) p.indentLeft = ml;
  const mr = len("margin-right");
  if (mr !== undefined) p.indentRight = mr;
  const ti = lengthPt(fo("text-indent"));
  if (ti !== null) p.indentFirst = ti;
  const mt = len("margin-top");
  if (mt !== undefined) p.spaceBefore = mt;
  const mb = len("margin-bottom");
  if (mb !== undefined) p.spaceAfter = mb;
  const lh = fo("line-height");
  if (lh && lh !== "normal") {
    if (lh.endsWith("%")) {
      p.lineHeight = parseFloat(lh) / 100;
      p.lineRule = "auto";
    } else {
      const v = lengthPt(lh);
      if (v) {
        p.lineHeight = v;
        p.lineRule = "exact";
      }
    }
  } else if (lh === "normal") {
    p.lineHeight = 1;
    p.lineRule = "auto";
  }
  const least = lengthPt(a(pp, "line-height-at-least"));
  if (least) {
    p.lineHeight = least;
    p.lineRule = "atLeast";
  }
  const bg = fo("background-color");
  if (bg && bg !== "transparent") p.shading = normColor(bg) ?? undefined;
  if (fo("keep-with-next") === "always") p.keepNext = true;
  if (fo("keep-together") === "always") p.keepLines = true;
  const bb = fo("break-before");
  if (bb === "page" || bb === "column") p.breakBefore = bb;
  const ba = fo("break-after");
  if (ba === "page" || ba === "column") p.breakAfter = ba;
  const vis = (v: string | null) => !!v && v !== "none" && !/^0(\.0+)?(pt|cm|in|mm)?\s/.test(v);
  const dc = kid(pp, "drop-cap");
  if (dc) {
    const lines = Number(a(dc, "lines") ?? 1) || 1;
    p.dropCap = lines > 1 && Number(a(dc, "length") ?? 1) !== 0 ? { lines, distance: lengthPt(a(dc, "distance")) ?? 0, textStyle: a(dc, "style-name") } : null;
  }
  if (vis(fo("border-bottom"))) p.borderBottom = true;
  if (vis(fo("border")) || vis(fo("border-top")) || vis(fo("border-left")) || vis(fo("border-right"))) p.borderOther = true;
  return p;
}

export class OdtStyles {
  fonts = new Map<string, string>();
  named = new Map<string, OdtStyle>();
  /** Otomatik stiller kapsam başına (content.xml / styles.xml) ayrı tutulur. */
  autoContent = new Map<string, OdtStyle>();
  autoStyles = new Map<string, OdtStyle>();
  defaults = new Map<string, OdtStyle>();
  lists = new Map<string, ListLevel[]>();
  layouts = new Map<string, PageLayout>();
  /** Tarih/saat veri stilleri: ad → biçim resmi. */
  dataStyles = new Map<string, string>();
  masters = new Map<string, MasterPage>();
  masterOrder: string[] = [];

  /** Belge kökünden (content.xml ya da styles.xml) stilleri toplar. */
  load(root: Element, scope: "content" | "styles") {
    for (const sec of kids(root)) {
      if (sec.localName === "font-face-decls") {
        for (const f of kids(sec, "font-face")) {
          const name = a(f, "name");
          const fam = a(f, "font-family");
          if (name) this.fonts.set(name, (fam ?? name).replace(/^['"]|['"]$/g, ""));
        }
      } else if (sec.localName === "styles" || sec.localName === "automatic-styles") {
        const auto = sec.localName === "automatic-styles";
        const target = !auto ? this.named : scope === "content" ? this.autoContent : this.autoStyles;
        for (const s of kids(sec)) {
          if (s.localName === "style") {
            const st = this.parseStyle(s, auto);
            target.set(`${st.family}:${st.name}`, st);
          } else if (s.localName === "default-style") {
            const st = this.parseStyle(s, false);
            this.defaults.set(st.family, st);
          } else if (s.localName === "list-style") {
            const name = a(s, "name");
            if (name) this.lists.set(`${auto ? scope : "named"}:${name}`, parseListStyle(s));
          } else if (s.localName === "page-layout") {
            const name = a(s, "name");
            if (name) this.layouts.set(name, parseLayout(s));
          } else if (s.localName === "date-style" || s.localName === "time-style") {
            const name = a(s, "name");
            const pic = odfToPicture(s);
            if (name && pic) this.dataStyles.set(name, pic);
          }
        }
      } else if (sec.localName === "master-styles") {
        for (const m of kids(sec, "master-page")) {
          const name = a(m, "name");
          if (!name) continue;
          this.masterOrder.push(name);
          this.masters.set(name, {
            name,
            layout: a(m, "page-layout-name"),
            header: kid(m, "header"),
            footer: kid(m, "footer"),
            headerFirst: kid(m, "header-first"),
            footerFirst: kid(m, "footer-first"),
            headerLeft: kid(m, "header-left"),
            footerLeft: kid(m, "footer-left"),
          });
        }
      }
    }
  }

  private parseStyle(s: Element, auto: boolean): OdtStyle {
    const name = a(s, "name") ?? "";
    const tp = kid(s, "text-properties");
    const pp = kid(s, "paragraph-properties");
    const cp = kid(s, "table-cell-properties");
    const tbl = kid(s, "table-properties");
    const col = kid(s, "table-column-properties");
    const rp = kid(s, "table-row-properties");
    const gp = kid(s, "graphic-properties");
    const cell: OdtStyle["cell"] = {};
    if (cp) {
      const bg = cp.getAttributeNS(NS_FO, "background-color");
      if (bg && bg !== "transparent") cell.bg = normColor(bg) ?? undefined;
      const va = a(cp, "vertical-align");
      if (va === "middle" || va === "bottom") cell.valign = va;
      const bs = ["border", "border-top", "border-bottom", "border-left", "border-right"].map((n) => cp.getAttributeNS(NS_FO, n)).filter((v): v is string => !!v);
      if (bs.length) cell.border = bs.some((v) => v !== "none" && !/^0(\.0+)?(pt|cm|in|mm)?\s/.test(v));
    }
    const table: OdtStyle["table"] = {};
    if (tbl) {
      const w = lengthPt(a(tbl, "width"));
      if (w) table.width = w;
      table.align = a(tbl, "align") ?? undefined;
      const ml = lengthPt(tbl.getAttributeNS(NS_FO, "margin-left"));
      if (ml) table.marginLeft = ml;
      if (tbl.getAttributeNS(NS_FO, "break-before") === "page") table.breakBefore = true;
    }
    const column: OdtStyle["column"] = {};
    if (col) {
      const w = lengthPt(a(col, "column-width"));
      if (w) column.width = w;
      const rel = parseFloat(a(col, "rel-column-width") ?? "");
      if (Number.isFinite(rel)) column.rel = rel;
    }
    const row: OdtStyle["row"] = {};
    if (rp) {
      const exact = lengthPt(a(rp, "row-height"));
      const min = lengthPt(a(rp, "min-row-height"));
      if (exact) Object.assign(row, { height: exact, exact: true });
      else if (min) row.height = min;
    }
    return {
      name,
      display: a(s, "display-name") ?? decodeName(name),
      family: a(s, "family") ?? "paragraph",
      parent: a(s, "parent-style-name"),
      auto,
      listStyle: a(s, "list-style-name"),
      masterPage: Array.from(s.attributes).find((x) => x.localName === "master-page-name")?.value ?? null,
      text: parseTextProps(tp, this.fonts),
      para: parseParaProps(pp),
      cell,
      table,
      column,
      row,
      graphic: { wrap: a(gp, "wrap") ?? undefined },
    };
  }

  /** Stil arama: önce kapsamın otomatik stilleri, sonra adlandırılmış stiller. */
  get(name: string | null, family: string, scope: "content" | "styles"): OdtStyle | undefined {
    if (!name) return undefined;
    const key = `${family}:${name}`;
    return (scope === "content" ? this.autoContent : this.autoStyles).get(key) ?? this.named.get(key);
  }

  /** Stil zinciri (kökten yaprağa doğru), döngüye karşı korumalı. */
  chain(name: string | null, family: string, scope: "content" | "styles"): OdtStyle[] {
    const out: OdtStyle[] = [];
    let cur = this.get(name, family, scope);
    const seen = new Set<string>();
    while (cur && !seen.has(cur.name) && out.length < 20) {
      seen.add(cur.name);
      out.unshift(cur);
      cur = cur.parent ? this.named.get(`${family}:${cur.parent}`) : undefined;
    }
    return out;
  }

  /** Zincirin birleştirilmiş karakter özellikleri (yüzdelik boyut üst boyuta göre çözülür). */
  mergeText(base: CharProps, list: OdtStyle[]): CharProps {
    const out: CharProps = { ...base };
    for (const s of list) {
      const { sizePct, ...rest } = s.text;
      for (const [k, v] of Object.entries(rest)) if (v !== undefined) (out as Record<string, unknown>)[k] = v;
      if (sizePct) out.size = Math.round((out.size ?? 12) * sizePct * 100) / 100;
    }
    return out;
  }

  mergePara(list: OdtStyle[]): ParaProps {
    const out: ParaProps = { ...this.defaults.get("paragraph")?.para };
    for (const s of list) for (const [k, v] of Object.entries(s.para)) if (v !== undefined) (out as Record<string, unknown>)[k] = v;
    return out;
  }

  /** Varsayılan paragraf metin özellikleri (yazı tipi ve boyut). */
  defaultText(): CharProps {
    const d = this.defaults.get("paragraph")?.text ?? {};
    const out: CharProps = { size: 12 };
    for (const [k, v] of Object.entries(d)) if (v !== undefined && k !== "sizePct") (out as Record<string, unknown>)[k] = v;
    return out;
  }

  list(name: string | null, scope: "content" | "styles"): ListLevel[] | undefined {
    if (!name) return undefined;
    return this.lists.get(`${scope}:${name}`) ?? this.lists.get(`named:${name}`);
  }
}

function parseListStyle(s: Element): ListLevel[] {
  const levels: ListLevel[] = [];
  for (const l of kids(s)) {
    const level = Number(a(l, "level") ?? 0) - 1;
    if (level < 0 || level > 9) continue;
    if (l.localName === "list-level-style-bullet" || l.localName === "list-level-style-image") {
      levels[level] = { bullet: true, char: a(l, "bullet-char") ?? "•", format: "", suffix: "", start: 1, display: 1 };
    } else if (l.localName === "list-level-style-number") {
      levels[level] = { bullet: false, char: "", format: a(l, "num-format") ?? "1", suffix: a(l, "num-suffix") ?? "", start: Number(a(l, "start-value") ?? 1) || 1, display: Number(a(l, "display-levels") ?? 1) || 1 };
    }
  }
  return levels;
}

function parseLayout(s: Element): PageLayout {
  const pp = kid(s, "page-layout-properties");
  const fo = (el: Element | null, n: string) => lengthPt(el?.getAttributeNS(NS_FO, n) ?? null) ?? undefined;
  const all = fo(pp, "margin");
  const lay: PageLayout = {
    width: fo(pp, "page-width"),
    height: fo(pp, "page-height"),
    landscape: a(pp, "print-orientation") === "landscape",
    margin: { top: fo(pp, "margin-top") ?? all, bottom: fo(pp, "margin-bottom") ?? all, left: fo(pp, "margin-left") ?? all, right: fo(pp, "margin-right") ?? all },
  };
  const bg = pp?.getAttributeNS(NS_FO, "background-color") || (a(pp, "fill") === "solid" ? a(pp, "fill-color") : null);
  if (bg && bg !== "transparent") lay.background = normColor(bg);
  const cols = kid(pp, "columns");
  const cc = Number(cols?.getAttributeNS(NS_FO, "column-count") ?? 1);
  if (cc > 1) {
    lay.columns = Math.min(13, cc);
    lay.columnGap = fo(cols, "column-gap");
  }
  for (const k of ["header", "footer"] as const) {
    const hfp = kid(kid(s, `${k}-style`), "header-footer-properties");
    if (hfp) lay[k] = { minHeight: fo(hfp, "min-height") ?? fo(hfp, "height") ?? 0, spacing: (k === "header" ? fo(hfp, "margin-bottom") : fo(hfp, "margin-top")) ?? 0 };
  }
  return lay;
}
