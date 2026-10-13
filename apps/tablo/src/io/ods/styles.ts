/**
 * ODF hücre/satır/sütun/tablo stilleri ⇄ CellStyle. Kalıtım (parent-style-name)
 * ve varsayılan stil (style:default-style) çözülür; uzunluklar piksele çevrilir.
 */
import type { BorderSide, BorderStyle, CellStyle, HAlign, VAlign } from "../../model/styles";
import { dataStyleToCode } from "./numfmt";
import { esc, kids, type XmlEl } from "./xml";

/** "1.25in", "2cm", "12pt", "3mm", "10px" → piksel (96 dpi). */
export function lengthToPx(v: string | undefined): number | undefined {
  if (!v) return undefined;
  const m = /^(-?[\d.]+)\s*(in|cm|mm|pt|pc|px)?$/.exec(v.trim());
  if (!m) return undefined;
  const n = Number(m[1]);
  if (!Number.isFinite(n)) return undefined;
  const unit = m[2] ?? "px";
  const px = unit === "in" ? n * 96 : unit === "cm" ? (n * 96) / 2.54 : unit === "mm" ? (n * 96) / 25.4 : unit === "pt" ? (n * 96) / 72 : unit === "pc" ? n * 16 : n;
  return px;
}

function lengthToPt(v: string | undefined): number | undefined {
  const px = lengthToPx(v);
  return px === undefined ? undefined : (px * 72) / 96;
}

export function pxToIn(px: number): string {
  return `${(px / 96).toFixed(4)}in`;
}

function color(v: string | undefined): string | undefined {
  if (!v || v === "transparent") return undefined;
  const m = /^#([0-9a-fA-F]{6})$/.exec(v.trim());
  return m ? "#" + m[1].toUpperCase() : undefined;
}

function parseBorder(v: string | undefined): BorderSide | null | undefined {
  if (v === undefined) return undefined;
  const t = v.trim();
  if (!t || t === "none" || /\bhidden\b/.test(t)) return null;
  const parts = t.split(/\s+/);
  let width = 0.74;
  let kind = "solid";
  let col = "#000000";
  for (const p of parts) {
    if (/^#[0-9a-fA-F]{6}$/.test(p)) col = p.toUpperCase();
    else if (/^[\d.]+(in|cm|mm|pt|px)$/.test(p)) width = lengthToPt(p) ?? width;
    else if (/^(solid|dashed|dotted|double|double-thin|fine-dashed|dash-dot|dash-dot-dot|thin|thick|medium)$/.test(p)) kind = p;
  }
  let style: BorderStyle;
  const heavy = width >= 1.2;
  if (kind === "double" || kind === "double-thin") style = "double";
  else if (kind === "dashed" || kind === "fine-dashed") style = heavy ? "mediumDashed" : "dashed";
  else if (kind === "dotted") style = "dotted";
  else if (kind === "dash-dot") style = heavy ? "mediumDashDot" : "dashDot";
  else if (kind === "dash-dot-dot") style = "dashDotDot";
  else style = width < 0.5 ? "hair" : width < 1.2 ? "thin" : width < 2.2 ? "medium" : "thick";
  return { style, color: col };
}

/** Okuma bağlamı: stil öğeleri, veri stilleri ve yazı tipi bildirimleri. */
export class OdsStyleSheet {
  /** "table-cell:ce1" → öğe. */
  private styles = new Map<string, XmlEl>();
  private defaults = new Map<string, XmlEl>();
  dataStyles = new Map<string, XmlEl>();
  private fonts = new Map<string, string>();
  private cache = new Map<string, CellStyle>();

  /** office:styles / office:automatic-styles / office:font-face-decls içeriğini ekler. */
  add(container: XmlEl | undefined): void {
    for (const el of kids(container)) {
      if (el.name === "style:style") {
        const fam = el.attrs["style:family"];
        const name = el.attrs["style:name"];
        if (fam && name) this.styles.set(`${fam}:${name}`, el);
      } else if (el.name === "style:default-style") {
        const fam = el.attrs["style:family"];
        if (fam) this.defaults.set(fam, el);
      } else if (el.name.startsWith("number:") && el.attrs["style:name"]) {
        this.dataStyles.set(el.attrs["style:name"], el);
      } else if (el.name === "style:font-face") {
        const fam = (el.attrs["svg:font-family"] ?? el.attrs["style:name"] ?? "").replace(/^['"]|['"]$/g, "");
        if (el.attrs["style:name"] && fam) this.fonts.set(el.attrs["style:name"], fam);
      }
    }
  }

  get(family: string, name: string | undefined): XmlEl | undefined {
    return name ? this.styles.get(`${family}:${name}`) : undefined;
  }

  /** Kalıtım zinciri (kökten yaprağa). */
  private chain(family: string, name: string | undefined): XmlEl[] {
    const out: XmlEl[] = [];
    let cur = this.get(family, name);
    const seen = new Set<string>();
    while (cur && !seen.has(cur.attrs["style:name"]) && out.length < 16) {
      seen.add(cur.attrs["style:name"]);
      out.unshift(cur);
      cur = this.get(family, cur.attrs["style:parent-style-name"]);
    }
    const def = this.defaults.get(family);
    if (def) out.unshift(def);
    return out;
  }

  /** Bir aile için birleştirilmiş özellik öznitelikleri (ör. "style:table-column-properties"). */
  props(family: string, name: string | undefined, propEl: string): Record<string, string> {
    const out: Record<string, string> = {};
    for (const el of this.chain(family, name)) for (const p of kids(el, propEl)) Object.assign(out, p.attrs);
    return out;
  }

  /** Hücre stili (tam çözülmüş). Ad yoksa "Default" kullanılır. */
  cellStyle(name: string | undefined): CellStyle {
    const key = name ?? "";
    const hit = this.cache.get(key);
    if (hit) return hit;
    const chain = this.chain("table-cell", name ?? "Default");
    const cp: Record<string, string> = {};
    const pp: Record<string, string> = {};
    const tp: Record<string, string> = {};
    let dataStyle: string | undefined;
    for (const el of chain) {
      for (const p of kids(el, "style:table-cell-properties")) Object.assign(cp, p.attrs);
      for (const p of kids(el, "style:paragraph-properties")) Object.assign(pp, p.attrs);
      for (const p of kids(el, "style:text-properties")) Object.assign(tp, p.attrs);
      if (el.attrs["style:data-style-name"]) dataStyle = el.attrs["style:data-style-name"];
    }
    const st: CellStyle = {};
    const font = tp["style:font-name"] ? (this.fonts.get(tp["style:font-name"]) ?? tp["style:font-name"]) : tp["fo:font-family"]?.replace(/^['"]|['"]$/g, "");
    if (font) st.font = font;
    const size = tp["fo:font-size"];
    if (size && !size.endsWith("%")) {
      const pt = lengthToPt(size);
      if (pt) st.size = Math.round(pt * 2) / 2;
    }
    const w = tp["fo:font-weight"];
    if (w === "bold" || (w && Number(w) >= 600)) st.bold = true;
    if (tp["fo:font-style"] === "italic" || tp["fo:font-style"] === "oblique") st.italic = true;
    const ul = tp["style:text-underline-style"];
    if (ul && ul !== "none") st.underline = tp["style:text-underline-type"] === "double" ? "double" : "single";
    const lt = tp["style:text-line-through-style"];
    if (lt && lt !== "none") st.strike = true;
    const fc = color(tp["fo:color"]);
    if (fc && fc !== "#000000") st.color = fc;
    const bg = color(cp["fo:background-color"]);
    if (bg) st.fill = bg;
    const all = parseBorder(cp["fo:border"]);
    for (const [side, key] of [["top", "bTop"], ["bottom", "bBottom"], ["left", "bLeft"], ["right", "bRight"]] as const) {
      const b = parseBorder(cp[`fo:border-${side}`]);
      const v = b === undefined ? all : b;
      if (v) st[key] = v;
    }
    const va = cp["style:vertical-align"];
    if (va === "top" || va === "middle") st.vAlign = va as VAlign;
    if (cp["fo:wrap-option"] === "wrap") st.wrap = true;
    if (cp["style:shrink-to-fit"] === "true") st.shrink = true;
    if (cp["style:direction"] === "ttb") st.rotation = 255;
    else if (cp["style:rotation-angle"]) {
      const a = ((Number(cp["style:rotation-angle"].replace(/deg$/, "")) % 360) + 360) % 360;
      if (a > 0 && a <= 90) st.rotation = Math.round(a);
      else if (a >= 270) st.rotation = Math.round(a - 360);
    }
    const prot = cp["style:cell-protect"];
    if (prot === "none") st.locked = false;
    if (prot && /hidden/.test(prot)) st.hidden = true;
    if (cp["style:text-align-source"] !== "value-type") {
      const ta = pp["fo:text-align"];
      const h: HAlign | undefined = ta === "start" || ta === "left" ? "left" : ta === "center" ? "center" : ta === "end" || ta === "right" ? "right" : ta === "justify" ? "justify" : undefined;
      if (h) st.hAlign = h;
    }
    const ml = lengthToPt(pp["fo:margin-left"]);
    if (ml && ml >= 5) st.indent = Math.min(15, Math.round(ml / 10));
    if (dataStyle) {
      const code = dataStyleToCode(dataStyle, this.dataStyles);
      if (code) st.numFmt = code;
    }
    this.cache.set(key, st);
    return st;
  }
}

/* ------------------------------ Yazma ------------------------------ */

const BORDER_OUT: Record<BorderStyle, string> = {
  hair: "0.26pt solid",
  thin: "0.74pt solid",
  medium: "1.76pt solid",
  thick: "2.49pt solid",
  dashed: "0.74pt dashed",
  mediumDashed: "1.76pt dashed",
  dotted: "0.74pt dotted",
  double: "1.76pt double",
  dashDot: "0.74pt dash-dot",
  mediumDashDot: "1.76pt dash-dot",
  dashDotDot: "0.74pt dash-dot-dot",
  slantDashDot: "1.76pt dash-dot",
};

function borderOut(b: BorderSide | undefined): string {
  return b ? `${BORDER_OUT[b.style] ?? "0.74pt solid"} ${(b.color ?? "#000000").toLowerCase()}` : "none";
}

/** CellStyle → style:style öğesi (table-cell ailesi). */
export function cellStyleXml(name: string, s: CellStyle, dataStyle: string | undefined, fontName: (f: string) => string): string {
  const cp: string[] = [];
  if (s.fill) cp.push(`fo:background-color="${s.fill.toLowerCase()}"`);
  if (s.bTop || s.bBottom || s.bLeft || s.bRight) {
    cp.push(`fo:border-top="${borderOut(s.bTop)}"`, `fo:border-bottom="${borderOut(s.bBottom)}"`, `fo:border-left="${borderOut(s.bLeft)}"`, `fo:border-right="${borderOut(s.bRight)}"`);
  }
  if (s.vAlign) cp.push(`style:vertical-align="${s.vAlign === "top" ? "top" : s.vAlign === "bottom" ? "bottom" : "middle"}"`);
  if (s.wrap) cp.push('fo:wrap-option="wrap"');
  if (s.shrink) cp.push('style:shrink-to-fit="true"');
  if (s.rotation === 255) cp.push('style:direction="ttb"');
  else if (s.rotation) cp.push(`style:rotation-angle="${s.rotation > 0 ? s.rotation : 360 + s.rotation}"`);
  if (s.hAlign && s.hAlign !== "general") cp.push('style:text-align-source="fix"');
  if (s.locked === false || s.hidden) cp.push(`style:cell-protect="${s.hidden ? (s.locked === false ? "formula-hidden" : "hidden-and-protected") : "none"}"`);
  const pp: string[] = [];
  if (s.hAlign && s.hAlign !== "general") {
    const ta = s.hAlign === "left" ? "start" : s.hAlign === "right" ? "end" : s.hAlign === "justify" || s.hAlign === "distributed" ? "justify" : "center";
    pp.push(`fo:text-align="${ta}"`);
  }
  if (s.indent) pp.push(`fo:margin-left="${s.indent * 10}pt"`);
  const tp: string[] = [];
  if (s.font) tp.push(`style:font-name="${esc(fontName(s.font))}"`);
  if (s.size) tp.push(`fo:font-size="${s.size}pt"`, `style:font-size-asian="${s.size}pt"`, `style:font-size-complex="${s.size}pt"`);
  if (s.bold) tp.push('fo:font-weight="bold"', 'style:font-weight-asian="bold"', 'style:font-weight-complex="bold"');
  if (s.italic) tp.push('fo:font-style="italic"', 'style:font-style-asian="italic"', 'style:font-style-complex="italic"');
  if (s.underline) tp.push('style:text-underline-style="solid"', 'style:text-underline-width="auto"', 'style:text-underline-color="font-color"', ...(s.underline === "double" ? ['style:text-underline-type="double"'] : []));
  if (s.strike) tp.push('style:text-line-through-style="solid"', 'style:text-line-through-type="single"');
  if (s.color) tp.push(`fo:color="${s.color.toLowerCase()}"`);
  const ds = dataStyle ? ` style:data-style-name="${dataStyle}"` : "";
  return (
    `<style:style style:name="${name}" style:family="table-cell" style:parent-style-name="Default"${ds}>` +
    (cp.length ? `<style:table-cell-properties ${cp.join(" ")}/>` : "") +
    (pp.length ? `<style:paragraph-properties ${pp.join(" ")}/>` : "") +
    (tp.length ? `<style:text-properties ${tp.join(" ")}/>` : "") +
    "</style:style>"
  );
}
