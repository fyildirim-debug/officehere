/**
 * ODP yazıcısının stil toplayıcıları: otomatik stiller (içerik ve stil
 * bölümleri için ayrı önekle), ortak adlandırılmış stiller (gradyan, doku,
 * ok ucu, kesik çizgi), yazı tipi bildirimleri ve resim dosyaları.
 */
import { resolveHex } from "../../model/color";
import { resolveFont } from "../../model/textStyle";
import type { Fill, Line, ParaProps, RunProps, Theme } from "../../model/types";
import { dataUrlToBytes } from "../pptxWriter";
import { esc } from "../xml";
import { cm, odfAttrs, ptStr } from "./common";

export type Attrs = Record<string, string | number | undefined | null | false>;

export class AutoStyles {
  private seen = new Map<string, string>();
  private out: string[] = [];
  private n = new Map<string, number>();
  constructor(private prefix: string) {}

  /** Aynı içerikli stil bir kez yazılır; adı döner. */
  style(family: string, kind: string, body: string, parent?: string): string {
    const key = `${family}|${parent ?? ""}|${body}`;
    const hit = this.seen.get(key);
    if (hit) return hit;
    const k = (this.n.get(kind) ?? 0) + 1;
    this.n.set(kind, k);
    const name = `${this.prefix}${kind}${k}`;
    this.seen.set(key, name);
    this.out.push(`<style:style${odfAttrs({ "style:name": name, "style:family": family, "style:parent-style-name": parent })}>${body}</style:style>`);
    return name;
  }

  list(body: string): string {
    const key = `list|${body}`;
    const hit = this.seen.get(key);
    if (hit) return hit;
    const k = (this.n.get("L") ?? 0) + 1;
    this.n.set("L", k);
    const name = `${this.prefix}L${k}`;
    this.seen.set(key, name);
    this.out.push(`<text:list-style style:name="${name}">${body}</text:list-style>`);
    return name;
  }

  /** Ham öğe (sayfa düzeni vb.). */
  raw(xml: string) {
    this.out.push(xml);
  }

  xml(): string {
    return this.out.join("");
  }
}

/** Belge genelinde paylaşılan adlandırılmış stiller ve dosyalar. */
export class Shared {
  readonly fonts = new Set<string>();
  readonly named: string[] = [];
  private namedKeys = new Map<string, string>();
  readonly files: { path: string; bytes: Uint8Array; mime: string }[] = [];
  private pics = new Map<string, string>();
  constructor(readonly theme: Theme) {}

  private addNamed(key: string, make: (name: string) => string, prefix: string): string {
    const hit = this.namedKeys.get(key);
    if (hit) return hit;
    const name = `${prefix}${this.namedKeys.size + 1}`;
    this.namedKeys.set(key, name);
    this.named.push(make(name));
    return name;
  }

  gradient(fill: Extract<Fill, { type: "gradient" }>): string {
    const stops = [...fill.stops].sort((a, b) => a.pos - b.pos);
    const a = resolveHex(stops[0]?.color, this.theme, "#FFFFFF");
    const b = resolveHex(stops[stops.length - 1]?.color, this.theme, "#000000");
    const angle = (((90 - fill.angle) % 360) + 360) % 360;
    const style = fill.kind === "radial" ? "radial" : "linear";
    return this.addNamed(`g|${a}|${b}|${angle}|${style}`, (n) => `<draw:gradient${odfAttrs({ "draw:name": n, "draw:style": style, "draw:start-color": a.toLowerCase(), "draw:end-color": b.toLowerCase(), "draw:start-intensity": "100%", "draw:end-intensity": "100%", "draw:angle": `${Math.round(angle)}deg`, "draw:border": "0%", "draw:cx": style === "radial" ? "50%" : undefined, "draw:cy": style === "radial" ? "50%" : undefined })}/>`, "Gradyan");
  }

  fillImage(src: string): string | undefined {
    const path = this.picture(src);
    if (!path) return undefined;
    return this.addNamed(`i|${path}`, (n) => `<draw:fill-image${odfAttrs({ "draw:name": n, "xlink:href": path, "xlink:type": "simple", "xlink:show": "embed", "xlink:actuate": "onLoad" })}/>`, "Doku");
  }

  marker(): string {
    return this.addNamed("arrow", (n) => `<draw:marker draw:name="${n}" svg:viewBox="0 0 20 30" svg:d="M10 0l-10 30h20z"/>`, "Ok");
  }

  dash(kind: Line["dash"]): string {
    const dots = kind === "sysDot" ? { "draw:dots1": 1, "draw:dots1-length": "0.02cm", "draw:distance": "0.02cm" } : kind === "dashDot" || kind === "lgDashDot" || kind === "lgDashDotDot" ? { "draw:dots1": 1, "draw:dots1-length": "0.2cm", "draw:dots2": kind === "lgDashDotDot" ? 2 : 1, "draw:dots2-length": "0.02cm", "draw:distance": "0.1cm" } : { "draw:dots1": 1, "draw:dots1-length": kind === "lgDash" ? "0.4cm" : "0.2cm", "draw:distance": "0.1cm" };
    return this.addNamed(`d|${kind}`, (n) => `<draw:stroke-dash${odfAttrs({ "draw:name": n, "draw:style": "rect", ...dots })}/>`, "Kesik");
  }

  /** data: URL'li resmi Pictures/ altına koyar; paket yolunu döndürür. */
  picture(src: string): string | undefined {
    const hit = this.pics.get(src);
    if (hit) return hit;
    const d = dataUrlToBytes(src);
    if (!d) return undefined;
    const ext = d.mime === "image/jpeg" ? "jpg" : (d.mime.split("/")[1] ?? "png").replace("svg+xml", "svg").replace(/[^a-z0-9]/g, "") || "png";
    const path = `Pictures/resim${this.files.length + 1}.${ext}`;
    this.files.push({ path, bytes: d.bytes, mime: d.mime });
    this.pics.set(src, path);
    return path;
  }

  font(name: string): string {
    this.fonts.add(name);
    return name;
  }

  fontDecls(): string {
    return `<office:font-face-decls>${[...this.fonts].map((f) => `<style:font-face${odfAttrs({ "style:name": f, "svg:font-family": /\s/.test(f) ? `'${f}'` : f })}/>`).join("")}</office:font-face-decls>`;
  }

  color(c: Parameters<typeof resolveHex>[0], fallback = "#000000"): string {
    return resolveHex(c, this.theme, fallback).toLowerCase();
  }
}

/** Dolgu → grafik özellikleri. */
export function fillAttrs(fill: Fill | undefined, sh: Shared): Attrs {
  if (!fill || fill.type === "none") return { "draw:fill": "none" };
  if (fill.type === "solid") {
    const a = fill.color.alpha;
    return { "draw:fill": "solid", "draw:fill-color": sh.color(fill.color), "draw:opacity": a !== undefined && a < 1 ? `${Math.round(a * 100)}%` : undefined };
  }
  if (fill.type === "gradient") return { "draw:fill": "gradient", "draw:fill-gradient-name": sh.gradient(fill), "draw:fill-color": sh.color(fill.stops[0]?.color, "#ffffff") };
  const name = sh.fillImage(fill.src);
  if (!name) return { "draw:fill": "none" };
  return { "draw:fill": "bitmap", "draw:fill-image-name": name, "style:repeat": fill.mode === "tile" ? "repeat" : "stretch", "draw:opacity": fill.alpha !== undefined && fill.alpha < 1 ? `${Math.round(fill.alpha * 100)}%` : undefined };
}

/** Çizgi → grafik özellikleri. */
export function strokeAttrs(line: Line | undefined, sh: Shared): Attrs {
  if (!line || line.fill.type === "none") return { "draw:stroke": "none" };
  const a = line.fill.color.alpha;
  return {
    "draw:stroke": line.dash === "solid" ? "solid" : "dash",
    "draw:stroke-dash": line.dash === "solid" ? undefined : sh.dash(line.dash),
    "svg:stroke-color": sh.color(line.fill.color),
    "svg:stroke-width": cm((Math.max(0, line.width) * 96) / 72),
    "svg:stroke-opacity": a !== undefined && a < 1 ? `${Math.round(a * 100)}%` : undefined,
    "draw:marker-start": line.head && line.head !== "none" ? sh.marker() : undefined,
    "draw:marker-start-width": line.head && line.head !== "none" ? cm(Math.max(6, line.width * 4)) : undefined,
    "draw:marker-end": line.tail && line.tail !== "none" ? sh.marker() : undefined,
    "draw:marker-end-width": line.tail && line.tail !== "none" ? cm(Math.max(6, line.width * 4)) : undefined,
    "svg:stroke-linecap": line.cap === "rnd" ? "round" : line.cap === "sq" ? "square" : undefined,
  };
}

/** Çalıştırma biçimi → style:text-properties. */
export function textProps(r: RunProps, sh: Shared, extra: Attrs = {}): string {
  const font = r.font ? resolveFont(r.font, sh.theme) : undefined;
  const a: Attrs = {
    "fo:font-size": r.size ? ptStr(r.size) : undefined,
    "fo:font-weight": r.bold === undefined ? undefined : r.bold ? "bold" : "normal",
    "fo:font-style": r.italic === undefined ? undefined : r.italic ? "italic" : "normal",
    "style:text-underline-style": r.underline === undefined ? undefined : r.underline ? "solid" : "none",
    "style:text-underline-width": r.underline ? "auto" : undefined,
    "style:text-underline-color": r.underline ? "font-color" : undefined,
    "style:text-line-through-style": r.strike === undefined ? undefined : r.strike ? "solid" : "none",
    "fo:color": r.color ? sh.color(r.color) : undefined,
    "style:font-name": font ? sh.font(font) : undefined,
    "fo:font-family": font ? (/\s/.test(font) ? `'${font}'` : font) : undefined,
    "fo:text-transform": r.caps === "all" ? "uppercase" : r.caps === "none" ? "none" : undefined,
    "fo:font-variant": r.caps === "small" ? "small-caps" : r.caps === "none" ? "normal" : undefined,
    "style:text-position": r.baseline ? `${Math.round(r.baseline)}% 58%` : undefined,
    "fo:background-color": r.highlight ? sh.color(r.highlight) : undefined,
    "fo:text-shadow": r.shadow === undefined ? undefined : r.shadow ? "1pt 1pt" : "none",
    "fo:letter-spacing": r.spacing ? cm((r.spacing * 96) / 72) : undefined,
    ...extra,
  };
  const s = odfAttrs(a);
  return s ? `<style:text-properties${s}/>` : "";
}

const ALIGN: Record<string, string> = { l: "start", ctr: "center", r: "end", just: "justify" };

/** Paragraf biçimi → style:paragraph-properties (madde girintileri liste stiline gider). */
export function paraProps(p: ParaProps, inList: boolean): string {
  const a: Attrs = {
    "fo:text-align": p.align ? ALIGN[p.align] : undefined,
    "fo:margin-top": p.spaceBefore !== undefined ? cm((p.spaceBefore * 96) / 72) : undefined,
    "fo:margin-bottom": p.spaceAfter !== undefined ? cm((p.spaceAfter * 96) / 72) : undefined,
    "fo:line-height": p.lineSpacingPt ? cm((p.lineSpacingPt * 96) / 72) : p.lineSpacing ? `${Math.round(p.lineSpacing)}%` : undefined,
    "fo:margin-left": !inList && p.marL !== undefined ? cm(p.marL) : undefined,
    "fo:text-indent": !inList && p.indent !== undefined ? cm(p.indent) : undefined,
  };
  const s = odfAttrs(a);
  return s ? `<style:paragraph-properties${s}/>` : "";
}

export const xmlText = esc;
