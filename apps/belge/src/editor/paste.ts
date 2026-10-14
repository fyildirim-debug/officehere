// Yapıştırma dönüştürücüleri: kelime işlemci (Office HTML) listeleri, gereksiz işaretler ve düz metin satırları.
import { Extension } from "@tiptap/core";
import { Fragment, Slice, type Node as PMNode, type ResolvedPos, type Schema } from "@tiptap/pm/model";
import { Plugin, PluginKey } from "@tiptap/pm/state";
import type { EditorView } from "@tiptap/pm/view";
import { syncDomSelection } from "./pagination";

interface ListPara {
  el: HTMLElement;
  list: string;
  level: number;
  ordered: boolean;
  style: string | null;
}

/** "mso-list:l0 level2 lfo1" → { list: "l0:lfo1", level: 2 } */
function msoList(el: HTMLElement): { list: string; level: number } | null {
  const st = el.getAttribute("style") ?? "";
  const m = /mso-list:\s*(l\d+)\s+level(\d+)\s+(lfo\d+)/i.exec(st);
  if (!m) return null;
  return { list: `${m[1]}:${m[3]}`, level: Math.max(1, Number(m[2])) };
}

/** Liste işaretinin metninden numaralandırma türü. */
function markerKind(marker: string): { ordered: boolean; style: string | null } {
  const t = marker.replace(/ /g, " ").trim();
  if (/^\d+[.)]$/.test(t)) return { ordered: true, style: t.endsWith(")") ? "decimal-paren" : null };
  if (/^[ivxlcdm]+[.)]$/.test(t)) return { ordered: true, style: "lower-roman" };
  if (/^[IVXLCDM]+[.)]$/.test(t)) return { ordered: true, style: "upper-roman" };
  if (/^[a-z][.)]$/.test(t)) return { ordered: true, style: t.endsWith(")") ? "lower-alpha-paren" : "lower-alpha" };
  if (/^[A-Z][.)]$/.test(t)) return { ordered: true, style: "upper-alpha" };
  if (t === "o") return { ordered: false, style: "circle" };
  if (t === "§" || t === "▪" || t === "■") return { ordered: false, style: "square" };
  if (t === "-" || t === "–") return { ordered: false, style: "dash" };
  if (t === "Ø" || t === "➢") return { ordered: false, style: "arrow" };
  if (t === "ü" || t === "✓") return { ordered: false, style: "check" };
  return { ordered: false, style: null };
}

/** Etkin içerik taşıyan öğeler: yapıştırılan HTML'den içerikleriyle birlikte atılır. */
const ACTIVE_TAGS = "script, iframe, frame, frameset, object, embed, applet, base, link, meta, noscript, template, portal";
/** URL taşıyan öznitelikler. */
const URL_ATTRS = new Set(["href", "src", "xlink:href", "action", "formaction", "data", "poster", "background", "srcset", "lowsrc", "dynsrc", "ping", "cite", "longdesc", "manifest", "codebase"]);

/** Betik çalıştırabilen adres: javascript:, vbscript:, (resim dışı) data:. Görünmez karakterler ve boşluklar yok sayılır. */
function unsafeUrl(name: string, value: string, tag: string): boolean {
  // eslint-disable-next-line no-control-regex
  const v = value.replace(/[\u0000-\u0020\u007f-\u00a0\u200b-\u200f\u2028\u2029\ufeff]+/g, "").toLowerCase();
  if (/^(javascript|vbscript|livescript|mocha):/.test(v) || /&#|\\u00/.test(v.slice(0, 12))) return true;
  if (v.startsWith("data:")) return !(tag === "img" && name === "src" && /^data:image\/(png|jpe?g|gif|bmp|webp)[;,]/.test(v));
  return false;
}

/**
 * Yapıştırılan (ya da açılan) HTML'i etkisiz bir belgede (DOMParser; betik çalışmaz, resim yüklenmez) ayrıştırıp
 * temizler: betik/çerçeve/nesne öğeleri, on* olay öznitelikleri, javascript:/vbscript: adresleri ve CSS ifadeleri
 * atılır; metin ve biçim korunur.
 */
export function sanitizePastedHtml(html: string): string {
  const doc = new DOMParser().parseFromString(html, "text/html");
  const root = doc.documentElement;
  root.querySelectorAll(ACTIVE_TAGS).forEach((e) => e.remove());
  for (const el of Array.from(root.querySelectorAll("*"))) {
    const tag = el.localName.toLowerCase();
    // SVG içindeki betik ve animasyonla öznitelik değiştirme (set/animate href) de etkisizleştirilir.
    if ((tag === "set" || tag === "animate" || tag === "animatetransform" || tag === "handler" || tag === "foreignobject") && el.namespaceURI !== "http://www.w3.org/1999/xhtml") {
      el.remove();
      continue;
    }
    for (const a of Array.from(el.attributes)) {
      const n = a.name.toLowerCase();
      if (n.startsWith("on") || n === "formaction" || n === "srcdoc") el.removeAttribute(a.name);
      else if (URL_ATTRS.has(n) && unsafeUrl(n, a.value, tag)) el.removeAttribute(a.name);
      else if (n === "style" && /expression\s*\(|javascript:|vbscript:|-moz-binding|behavior\s*:/i.test(a.value)) el.removeAttribute(a.name);
    }
  }
  // Gövde dışında kalan (head içindeki) stil dışı içerik zaten atıldı; yalnız gövde döner.
  return doc.body.innerHTML;
}

/** Office HTML'ini (mso sınıfları, koşullu yorumlar, o:p, mso-list paragrafları) sadeleştirir. */
export function cleanOfficeHtml(html: string): string {
  if (!/mso-|urn:schemas-microsoft-com|class="?Mso/i.test(html)) return html;
  // Koşullu bloklar: liste işaretleri ve VML yedekleri
  let src = html.replace(/<!\[if !supportLists\]>([\s\S]*?)<!\[endif\]>/gi, "$1");
  src = src.replace(/<!--\[if [^\]]*\]>[\s\S]*?<!\[endif\]-->/gi, "");
  src = src.replace(/<!\[if [^\]]*\]>|<!\[endif\]>/gi, "");
  const doc = new DOMParser().parseFromString(src, "text/html");
  const body = doc.body;
  body.querySelectorAll("style, meta, link, title, xml").forEach((e) => e.remove());
  // <o:p>&nbsp;</o:p> boş paragraf işaretidir.
  for (const op of Array.from(body.getElementsByTagName("o:p"))) op.remove();

  // Liste paragrafları
  const paras = Array.from(body.querySelectorAll<HTMLElement>("p, h1, h2, h3, h4, h5, h6"));
  const items: ListPara[] = [];
  for (const p of paras) {
    const info = msoList(p);
    if (!info) continue;
    const ign = p.querySelector<HTMLElement>('span[style*="mso-list:Ignore" i], span[style*="mso-list: Ignore" i]');
    const marker = ign?.textContent ?? "";
    // İşaret ve hizalama boşlukları atılır.
    (ign?.parentElement && ign.parentElement !== p && ign.parentElement.textContent === ign.textContent ? ign.parentElement : ign)?.remove();
    const kind = markerKind(marker);
    p.style.removeProperty("text-indent");
    p.style.removeProperty("margin-left");
    items.push({ el: p, list: info.list, level: info.level, ordered: kind.ordered, style: kind.style });
  }
  // Ardışık liste paragraflarını iç içe ul/ol yapısına çevir.
  let i = 0;
  while (i < items.length) {
    const start = items[i];
    const group: ListPara[] = [start];
    let j = i + 1;
    while (j < items.length && items[j].list === start.list && items[j].el.previousElementSibling === group[group.length - 1].el) group.push(items[j++]);
    const root = doc.createElement(start.ordered ? "ol" : "ul");
    start.el.before(root);
    const stack: { level: number; list: HTMLElement; lastLi: HTMLElement | null }[] = [{ level: start.level, list: root, lastLi: null }];
    if (start.style) root.setAttribute("data-list-style", start.style);
    for (const it of group) {
      while (stack.length > 1 && it.level < stack[stack.length - 1].level) stack.pop();
      let top = stack[stack.length - 1];
      if (it.level > top.level && top.lastLi) {
        const sub = doc.createElement(it.ordered ? "ol" : "ul");
        if (it.style) sub.setAttribute("data-list-style", it.style);
        top.lastLi.appendChild(sub);
        stack.push({ level: it.level, list: sub, lastLi: null });
        top = stack[stack.length - 1];
      }
      const li = doc.createElement("li");
      li.appendChild(it.el);
      top.list.appendChild(li);
      top.lastLi = li;
    }
    i = j;
  }
  // Boş paragraflardaki tek &nbsp; Word'ün boş paragraf işaretidir.
  for (const p of Array.from(body.querySelectorAll("p"))) {
    if (p.textContent === " " && !p.querySelector("img")) p.innerHTML = "";
  }
  // Office sınıfları ve mso- stil özellikleri
  body.querySelectorAll<HTMLElement>("[class]").forEach((e) => {
    if (/^Mso/i.test(e.className)) e.removeAttribute("class");
  });
  body.querySelectorAll<HTMLElement>("[style]").forEach((e) => {
    const kept = (e.getAttribute("style") ?? "")
      .split(";")
      .filter((d) => d.trim() && !/^\s*mso-/i.test(d))
      .join(";");
    if (kept) e.setAttribute("style", kept);
    else e.removeAttribute("style");
  });
  return body.innerHTML;
}

const BLOCK_SEL = "h1, h2, h3, h4, h5, h6, p, li, td, th, blockquote, div";
const INHERIT = ["color", "font-family", "font-size", "font-weight", "font-style"] as const;

/**
 * Kaynak biçimlendirmesini koru: blok öğelerdeki (ör. başlık rengi) yazı tipi özellikleri metin işaretlerine
 * aktarılabilsin diye içeriği saran bir span'e taşınır; <font> öğeleri span'e çevrilir.
 */
export function liftBlockStyles(html: string): string {
  if (!/style=|<font/i.test(html)) return html;
  const doc = new DOMParser().parseFromString(html, "text/html");
  for (const f of Array.from(doc.body.querySelectorAll("font"))) {
    const span = doc.createElement("span");
    const c = f.getAttribute("color");
    const face = f.getAttribute("face");
    if (c) span.style.color = c;
    if (face) span.style.fontFamily = face;
    span.append(...Array.from(f.childNodes));
    f.replaceWith(span);
  }
  for (const el of Array.from(doc.body.querySelectorAll<HTMLElement>(BLOCK_SEL))) {
    const decl: string[] = [];
    for (const k of INHERIT) {
      const v = el.style.getPropertyValue(k);
      if (!v || /inherit|initial|unset/.test(v)) continue;
      // Başlıklar zaten kalın gösterilir; yalnız yazı tipi kalınlığı farklıysa taşınır.
      if (k === "font-weight" && (/^h\d$/i.test(el.tagName) || !/bold|[6-9]00/.test(v))) continue;
      if (k === "font-style" && v !== "italic") continue;
      decl.push(`${k}:${v}`);
    }
    if (!decl.length || !el.firstChild || el.querySelector(BLOCK_SEL)) continue;
    const span = doc.createElement("span");
    span.setAttribute("style", decl.join(";"));
    span.append(...Array.from(el.childNodes));
    el.appendChild(span);
  }
  return doc.body.innerHTML;
}

/** Düz metin: her satır bir paragraf (boş satırlar korunur, sekmeler metinde kalır). */
export function textToSlice(schema: Schema, text: string): Slice {
  const lines = text.replace(/\r\n?/g, "\n").split("\n");
  if (lines.length > 1 && lines[lines.length - 1] === "") lines.pop();
  const paras = lines.map((l) => schema.nodes.paragraph.create(null, l ? schema.text(l) : undefined));
  return new Slice(Fragment.from(paras), 1, 1);
}

/**
 * Tek bölümün içinden kopyalanan içerik (ör. Ctrl+A ile tüm belge) bölüm düğümüyle gelir; olduğu gibi
 * yapıştırılırsa yeni bir bölüm (bölüm sonu + yeni sayfa) eklenirdi. Word'deki gibi yalnız içerik
 * yapıştırılır: bölüm sarmalı açılır ve kenardaki paragraflar imleçteki paragrafla birleşir.
 * Birden çok bölüme yayılan kopyalar bölüm sonlarını korur.
 */
export function unwrapSectionSlice(slice: Slice): Slice {
  const { content } = slice;
  if (content.childCount !== 1 || content.firstChild!.type.name !== "section") return slice;
  const inner = content.firstChild!.content;
  if (!inner.size) return slice;
  const openStart = slice.openStart > 0 ? slice.openStart - 1 : inner.firstChild!.isTextblock ? 1 : 0;
  const openEnd = slice.openEnd > 0 ? slice.openEnd - 1 : inner.lastChild!.isTextblock ? 1 : 0;
  return new Slice(inner, openStart, openEnd);
}

/**
 * Liste öğesine birden çok paragraf yapıştırılınca her paragraf ayrı liste öğesi olur (Word, LibreOffice ve
 * Google Dokümanlar'daki gibi); ilk paragraf imleçteki öğeyle, sonuncusu imleçten sonraki metinle birleşir.
 */
export function listifySlice(slice: Slice, $from: ResolvedPos): Slice {
  const { content } = slice;
  if (content.childCount < 2 || $from.depth < 2) return slice;
  const item = $from.node($from.depth - 1);
  if (item.type.name !== "listItem" || !$from.parent.isTextblock) return slice;
  let plain = true;
  content.forEach((n) => {
    if (n.type.name !== "paragraph") plain = false;
  });
  if (!plain) return slice;
  const items: PMNode[] = [];
  content.forEach((n) => items.push(item.type.create(null, n)));
  return new Slice(Fragment.from(items), 2, 2);
}

/** Uygulama içinde son kopyalanan/kesilen içerik (pano izni olmayan tarayıcıda şerit Yapıştır'ı için). */
let lastCopied: Slice | null = null;
export function lastInternalCopy(): Slice | null {
  return lastCopied;
}

/** Kopyalanan seçimi saklar (klavyeyle hemen önce değişen DOM seçimi önce okunur). */
function remember(view: EditorView) {
  syncDomSelection(view);
  if (!view.state.selection.empty) lastCopied = unwrapSectionSlice(view.state.selection.content());
}

export const PasteRules = Extension.create({
  name: "belgePaste",
  addProseMirrorPlugins() {
    return [
      new Plugin({
        key: new PluginKey("belgePaste"),
        props: {
          // Office HTML sadeleştirmesi ham koşullu yorumları gerektirir; ikisi de etkisiz belgede (DOMParser) çalışır.
          transformPastedHTML: (html) => liftBlockStyles(sanitizePastedHtml(cleanOfficeHtml(html))),
          clipboardTextParser: (text, _ctx, _plain, view) => textToSlice(view.state.schema, text),
          transformPasted: (slice, view) => listifySlice(unwrapSectionSlice(slice), view.state.selection.$from),
          handleDOMEvents: {
            copy: (view) => {
              remember(view);
              return false;
            },
            cut: (view) => {
              remember(view);
              return false;
            },
          },
        },
      }),
    ];
  },
});
