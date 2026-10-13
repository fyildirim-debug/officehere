/** Ana hat görünümü için başlık/gövde metni ile satır listesi arasında dönüşüm. */
import type { Paragraph, Presentation, Slide, SpShape, TextBody } from "./types";

const TITLE_TYPES = new Set(["title", "ctrTitle"]);
const BODY_TYPES = new Set(["body", "obj", "subTitle"]);

export function titleShape(s: Slide): SpShape | undefined {
  return s.shapes.find((x): x is SpShape => x.kind === "sp" && !!x.ph && TITLE_TYPES.has(x.ph.type));
}

export function bodyShape(s: Slide): SpShape | undefined {
  return s.shapes.find((x): x is SpShape => x.kind === "sp" && !!x.ph && BODY_TYPES.has(x.ph.type));
}

export const paraText = (p: Paragraph) => p.runs.map((r) => r.text).join("");

export function bodyToLines(t: TextBody | undefined): string {
  if (!t) return "";
  const lines = t.paragraphs.map((p) => "\t".repeat(p.level) + paraText(p));
  return lines.length === 1 && !lines[0] ? "" : lines.join("\n");
}

/** Satırları paragraflara çevirir; aynı sıradaki paragrafın biçimi korunur. */
export function linesToBody(text: string, old: TextBody | undefined): TextBody {
  const base: TextBody = old ?? { paragraphs: [] };
  const lines = text.split("\n");
  const paragraphs: Paragraph[] = lines.map((line, k) => {
    const m = /^\t*/.exec(line)![0].length;
    const body = line.slice(m);
    const prev = base.paragraphs[k];
    const runProps = prev?.runs.find((r) => !r.props.field)?.props ?? prev?.endProps ?? {};
    return { level: Math.min(8, m), props: prev?.props ?? {}, runs: body ? [{ text: body, props: { ...runProps } }] : [], endProps: prev?.endProps };
  });
  return { ...base, paragraphs };
}


const FOOTER_PH = new Set(["dt", "ftr", "sldNum"]);
const isEmpty = (t: TextBody | undefined) => !t || t.paragraphs.every((p) => !paraText(p).trim());
const withText = (s: SpShape, text: TextBody): SpShape => ({ ...s, text });

/**
 * Ana hat: başlıkta Tab — slayt bir düzey indirilir; başlığı önceki slaydın gövdesine
 * ilk düzey madde, gövde maddeleri bir alt düzey olarak eklenir ve slayt kaldırılır.
 * Önceki slaytta gövde yoksa ya da slaytta yer tutucu dışı içerik/not varsa yapılmaz.
 */
export function demoteSlide(pres: Presentation, index: number): Presentation | null {
  const prev = pres.slides[index - 1];
  const cur = pres.slides[index];
  if (!prev || !cur) return null;
  const pb = bodyShape(prev);
  if (!pb) return null;
  const t = titleShape(cur);
  const cb = bodyShape(cur);
  const removable = (s: Slide["shapes"][number]) => s.kind === "sp" && !!s.ph && (isEmpty(s.text) || FOOTER_PH.has(s.ph.type));
  if (cur.notes.trim() || cur.shapes.some((s) => s !== t && s !== cb && !removable(s))) return null;
  const base = pb.text ?? { paragraphs: [] };
  const kept = isEmpty(base) ? [] : base.paragraphs;
  const props = kept[kept.length - 1]?.runs[0]?.props ?? {};
  const title = t?.text?.paragraphs.map(paraText).join(" ").replace(/\v/g, " ").trim() ?? "";
  const added: Paragraph[] = [];
  if (title) added.push({ level: 0, props: {}, runs: [{ text: title, props: { ...props } }] });
  for (const p of cb?.text?.paragraphs ?? []) if (paraText(p).trim()) added.push({ ...p, level: Math.min(8, p.level + 1) });
  const paragraphs = [...kept, ...added];
  const body: TextBody = { ...base, paragraphs: paragraphs.length ? paragraphs : [{ runs: [], level: 0, props: {} }] };
  const slides = pres.slides.slice();
  slides[index - 1] = { ...prev, shapes: prev.shapes.map((s) => (s === pb ? withText(pb, body) : s)) };
  slides.splice(index, 1);
  return { ...pres, slides };
}

/**
 * Ana hat: gövdenin ilk düzeyindeki satırda Shift+Tab — satır yeni bir slaydın başlığı
 * olur; ardından gelen satırlar (düzeyleri korunarak) yeni slaydın gövdesine taşınır.
 */
export function promoteLine(pres: Presentation, index: number, para: number, make: () => Slide): Presentation | null {
  const cur = pres.slides[index];
  const cb = cur && bodyShape(cur);
  const ps = cb?.text?.paragraphs;
  if (!cb || !ps || !ps[para] || ps[para].level !== 0) return null;
  const ns = make();
  const nt = titleShape(ns);
  const nb = bodyShape(ns);
  if (!nt) return null;
  const empty = { runs: [], level: 0, props: {} };
  const head = ps.slice(0, para);
  const tail = ps.slice(para + 1);
  const newSlide: Slide = {
    ...ns,
    shapes: ns.shapes.map((s) => {
      if (s === nt) return withText(nt, { ...(nt.text ?? { paragraphs: [] }), paragraphs: [{ level: 0, props: {}, runs: [{ text: paraText(ps[para]), props: {} }] }] });
      if (s === nb && tail.length) return withText(nb, { ...(nb.text ?? { paragraphs: [] }), paragraphs: tail });
      return s;
    }),
  };
  const slides = pres.slides.slice();
  slides[index] = { ...cur, shapes: cur.shapes.map((s) => (s === cb ? withText(cb, { ...cb.text!, paragraphs: head.length ? head : [empty] }) : s)) };
  slides.splice(index + 1, 0, newSlide);
  return { ...pres, slides };
}
