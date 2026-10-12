// Yapıştırma dönüştürücüleri: kelime işlemci (Office HTML) listeleri, gereksiz işaretler ve düz metin satırları.
import { Extension } from "@tiptap/core";
import { Fragment, Slice, type Schema } from "@tiptap/pm/model";
import { Plugin, PluginKey } from "@tiptap/pm/state";

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

/** Düz metin: her satır bir paragraf (boş satırlar korunur, sekmeler metinde kalır). */
export function textToSlice(schema: Schema, text: string): Slice {
  const lines = text.replace(/\r\n?/g, "\n").split("\n");
  if (lines.length > 1 && lines[lines.length - 1] === "") lines.pop();
  const paras = lines.map((l) => schema.nodes.paragraph.create(null, l ? schema.text(l) : undefined));
  return new Slice(Fragment.from(paras), 1, 1);
}

export const PasteRules = Extension.create({
  name: "belgePaste",
  addProseMirrorPlugins() {
    return [
      new Plugin({
        key: new PluginKey("belgePaste"),
        props: {
          transformPastedHTML: (html) => cleanOfficeHtml(html),
          clipboardTextParser: (text, _ctx, _plain, view) => textToSlice(view.state.schema, text),
        },
      }),
    ];
  },
});
