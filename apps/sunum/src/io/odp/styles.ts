/**
 * ODF stil çözümleyici (okuma): ortak/otomatik stiller, üst stil zinciri,
 * varsayılan stiller, yazı tipi bildirimleri ve adlandırılmış çizim stilleri
 * (gradyan, doku resmi, kesik çizgi, ok ucu).
 */
import { attr, kids } from "../xml";
import { firstFamily } from "./common";

export type Props = Record<string, string>;

export interface Resolved {
  graphic: Props;
  paragraph: Props;
  text: Props;
  page: Props;
  /** Grafik özelliklerine gömülü ya da stil adıyla verilen liste stili. */
  listStyle?: Element;
}

const PROP_KIND: Record<string, keyof Omit<Resolved, "listStyle">> = {
  "graphic-properties": "graphic",
  "paragraph-properties": "paragraph",
  "text-properties": "text",
  "drawing-page-properties": "page",
  "table-cell-properties": "graphic",
  "table-column-properties": "graphic",
  "table-row-properties": "graphic",
};

function collectAttrs(el: Element, into: Props) {
  for (const a of Array.from(el.attributes)) into[a.localName] = a.value;
}

/** Bir stil öğesinin kendi özellikleri (üst stiller hariç). */
function ownProps(el: Element): Resolved {
  const r: Resolved = { graphic: {}, paragraph: {}, text: {}, page: {} };
  for (const c of kids(el)) {
    const kind = PROP_KIND[c.localName];
    if (kind) {
      collectAttrs(c, r[kind]);
      const ls = kids(c, "list-style")[0];
      if (ls) r.listStyle = ls;
    }
  }
  return r;
}

function merge(into: Resolved, from: Resolved) {
  Object.assign(into.graphic, from.graphic);
  Object.assign(into.paragraph, from.paragraph);
  Object.assign(into.text, from.text);
  Object.assign(into.page, from.page);
  if (from.listStyle) into.listStyle = from.listStyle;
}

export const EMPTY: Resolved = { graphic: {}, paragraph: {}, text: {}, page: {} };

export class StyleSheet {
  private styles = new Map<string, Element>();
  private defaults = new Map<string, Element>();
  private cache = new Map<string, Resolved>();
  /** text:list-style öğeleri (ada göre). */
  readonly lists = new Map<string, Element>();
  /** Adlandırılmış çizim stilleri: "gradient:Ad", "fill-image:Ad", "stroke-dash:Ad", "marker:Ad". */
  readonly named = new Map<string, Element>();
  /** style:font-face adı → aile adı. */
  readonly fonts = new Map<string, string>();
  /** Ortak (office:styles) stil adları; ana slayt stillerini ayırt etmek için. */
  readonly common = new Set<string>();

  /** Kaplar öncelik sırasıyla verilir (ilk bulunan kazanır). */
  constructor(containers: (Element | undefined)[], fontDecls: (Element | undefined)[], commonContainer?: Element) {
    for (const fd of fontDecls) {
      for (const f of kids(fd, "font-face")) {
        const name = attr(f, "name");
        const fam = firstFamily(attr(f, "font-family"));
        if (name && fam && !this.fonts.has(name)) this.fonts.set(name, fam);
      }
    }
    for (const c of containers) {
      if (!c) continue;
      for (const el of kids(c)) {
        const ln = el.localName;
        const name = attr(el, "name");
        if (ln === "style" && name) {
          const key = `${attr(el, "family") ?? ""}:${name}`;
          if (!this.styles.has(key)) this.styles.set(key, el);
          if (c === commonContainer) this.common.add(key);
        } else if (ln === "default-style") {
          const fam = attr(el, "family") ?? "";
          if (!this.defaults.has(fam)) this.defaults.set(fam, el);
        } else if (ln === "list-style" && name) {
          if (!this.lists.has(name)) this.lists.set(name, el);
        } else if ((ln === "gradient" || ln === "fill-image" || ln === "stroke-dash" || ln === "marker" || ln === "hatch" || ln === "opacity") && name) {
          const key = `${ln}:${name}`;
          if (!this.named.has(key)) this.named.set(key, el);
        }
      }
    }
  }

  has(family: string, name: string | undefined): boolean {
    return !!name && this.styles.has(`${family}:${name}`);
  }

  isCommon(family: string, name: string | undefined): boolean {
    return !!name && this.common.has(`${family}:${name}`);
  }

  parentOf(family: string, name: string): string | undefined {
    return attr(this.styles.get(`${family}:${name}`), "parent-style-name");
  }

  /**
   * Stilin zincirle çözülmüş özellikleri. `stopAtCommon`: yalnız otomatik
   * stillerin katkısı (ana slayt stilleri ayrı okunur). `withDefault`: ailenin
   * varsayılan stili zincirin başına eklenir.
   */
  resolve(family: string, name: string | undefined, opts: { stopAtCommon?: boolean; withDefault?: boolean } = {}): Resolved {
    const key = `${family}:${name ?? ""}:${opts.stopAtCommon ? 1 : 0}${opts.withDefault ? 1 : 0}`;
    const hit = this.cache.get(key);
    if (hit) return hit;
    const chain: Element[] = [];
    const seen = new Set<string>();
    let cur = name;
    while (cur && chain.length < 32) {
      const k = `${family}:${cur}`;
      if (seen.has(k)) break;
      seen.add(k);
      const el = this.styles.get(k);
      if (!el) break;
      if (opts.stopAtCommon && this.common.has(k)) break;
      chain.unshift(el);
      cur = attr(el, "parent-style-name");
    }
    const out: Resolved = { graphic: {}, paragraph: {}, text: {}, page: {} };
    if (opts.withDefault) {
      const d = this.defaults.get(family);
      if (d) merge(out, ownProps(d));
    }
    for (const el of chain) merge(out, ownProps(el));
    this.cache.set(key, out);
    return out;
  }

  /** Yazı tipi adını (style:font-name ya da fo:font-family) aile adına çevirir. */
  fontOf(text: Props): string | undefined {
    const n = text["font-name"];
    if (n) return this.fonts.get(n) ?? n;
    return firstFamily(text["font-family"]);
  }
}
