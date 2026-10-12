/**
 * Metin biçimi kalıtımı: yerleşik varsayılan → asıl slayt metin stilleri
 * (başlık/gövde/diğer) → asıl yer tutucu → düzen yer tutucusu → şekil
 * düzey stili → paragraf/çalıştırma biçimi (ECMA-376 §19.3.1.x / §21.1.2).
 */
import type { Layout, LevelStyle, Master, Paragraph, ParaProps, Placeholder, Presentation, RunProps, Shape, Slide, SpShape, TextBody, Theme } from "./types";

export type TextKind = "title" | "body" | "other";

export function placeholderKind(ph: Placeholder | undefined): TextKind {
  if (!ph) return "other";
  if (ph.type === "title" || ph.type === "ctrTitle") return "title";
  if (ph.type === "body" || ph.type === "obj" || ph.type === "subTitle" || ph.type === "tbl" || ph.type === "pic") return "body";
  return "other";
}

function typeGroup(t: Placeholder["type"]): string {
  if (t === "title" || t === "ctrTitle") return "title";
  if (t === "body" || t === "obj" || t === "subTitle" || t === "tbl" || t === "pic") return "body";
  return t;
}

/** Düzende/asılda eşleşen yer tutucu şekli. */
export function findPlaceholder(shapes: Shape[], ph: Placeholder, strictIdx = false): SpShape | undefined {
  const sps = shapes.filter((s): s is SpShape => s.kind === "sp" && !!s.ph);
  if (ph.idx !== undefined) {
    const byIdx = sps.find((s) => s.ph!.idx === ph.idx && typeGroup(s.ph!.type) === typeGroup(ph.type));
    if (byIdx) return byIdx;
    const anyIdx = sps.find((s) => s.ph!.idx === ph.idx);
    if (anyIdx && ph.idx !== 0) return anyIdx;
    if (strictIdx) return undefined;
  }
  const exact = sps.find((s) => s.ph!.type === ph.type);
  if (exact) return exact;
  return sps.find((s) => typeGroup(s.ph!.type) === typeGroup(ph.type));
}

export function layoutOf(pres: Presentation, slide: Slide | undefined): Layout | undefined {
  if (!slide) return undefined;
  return pres.layouts.find((l) => l.id === slide.layoutId) ?? pres.layouts[0];
}

export interface TextContext {
  master: Master;
  layout?: Layout;
}

/** Yer tutucunun kalıtım zincirindeki şekiller (asıl → düzen). */
export function placeholderChain(shape: SpShape, ctx: TextContext): SpShape[] {
  if (!shape.ph) return [];
  const chain: SpShape[] = [];
  const m = findPlaceholder(ctx.master.shapes, { type: shape.ph.type });
  if (m && m !== shape) chain.push(m);
  if (ctx.layout) {
    const l = findPlaceholder(ctx.layout.shapes, shape.ph);
    if (l && l !== shape) chain.push(l);
  }
  return chain;
}

const BASE: LevelStyle = { para: { align: "l", lineSpacing: 100, marL: 0, indent: 0 }, run: { font: "+mn", size: 18, color: { scheme: "dk1" } } };

function mergeLevel(into: LevelStyle, src: Partial<LevelStyle> | undefined) {
  if (!src) return;
  if (src.para) Object.assign(into.para, stripUndef(src.para));
  if (src.run) Object.assign(into.run, stripUndef(src.run));
}

function stripUndef<T extends object>(o: T): Partial<T> {
  const r: Partial<T> = {};
  for (const k of Object.keys(o) as (keyof T)[]) if (o[k] !== undefined) r[k] = o[k];
  return r;
}

/** Bir şeklin metin düzeyleri için çözülmüş varsayılan stil üreticisi. */
export function levelResolver(shape: SpShape | { text?: TextBody; ph?: Placeholder; kind?: string }, ctx: TextContext, kindOverride?: TextKind): (lvl: number) => LevelStyle {
  const sp = shape as SpShape;
  const kind = kindOverride ?? placeholderKind(sp.ph);
  const chain = sp.kind === "sp" && sp.ph ? placeholderChain(sp, ctx) : [];
  const cache = new Map<number, LevelStyle>();
  return (lvl: number) => {
    const l = Math.max(0, Math.min(8, lvl));
    const hit = cache.get(l);
    if (hit) return hit;
    const st: LevelStyle = { para: { ...BASE.para }, run: { ...BASE.run } };
    const styles = ctx.master.textStyles[kind];
    mergeLevel(st, styles[l] ?? styles[styles.length - 1]);
    for (const c of chain) mergeLevel(st, c.text?.levels?.[l]);
    mergeLevel(st, sp.text?.levels?.[l]);
    cache.set(l, st);
    return st;
  };
}

/** Metin gövdesi özelliklerini (yaslama, iç boşluk, sığdırma) zincirden tamamlar. */
export function effectiveBodyProps(shape: SpShape, ctx: TextContext): Required<Pick<TextBody, "anchor" | "wrap" | "autofit" | "insets" | "vert">> {
  const chain = placeholderChain(shape, ctx);
  const pick = <K extends keyof TextBody>(k: K): TextBody[K] | undefined => {
    if (shape.text?.[k] !== undefined) return shape.text[k];
    for (let i = chain.length - 1; i >= 0; i--) if (chain[i].text?.[k] !== undefined) return chain[i].text![k];
    return undefined;
  };
  return {
    anchor: pick("anchor") ?? "t",
    wrap: pick("wrap") ?? true,
    autofit: pick("autofit") ?? "none",
    insets: pick("insets") ?? DEFAULT_INSETS,
    vert: pick("vert") ?? "horz",
  };
}

export const DEFAULT_INSETS = { l: 9.6, t: 4.8, r: 9.6, b: 4.8 };

export function resolveFont(font: string | undefined, theme: Theme): string {
  if (!font || font === "+mn" || font === "+mn-lt") return theme.fonts.minor;
  if (font === "+mj" || font === "+mj-lt") return theme.fonts.major;
  return font;
}

export function effectivePara(p: Paragraph, lvl: (n: number) => LevelStyle): ParaProps {
  return { ...lvl(p.level).para, ...stripUndef(p.props) };
}

export function effectiveRun(props: RunProps, p: Paragraph, lvl: (n: number) => LevelStyle): RunProps {
  return { ...lvl(p.level).run, ...stripUndef(props) };
}

/** Şekil yer tutucu istem metni. */
export function promptText(shape: SpShape, ctx: TextContext): string | undefined {
  if (!shape.ph) return undefined;
  if (shape.text?.prompt) return shape.text.prompt;
  // Yalnız düzendeki özel istem metni (custPrompt) slaytta gösterilir.
  const l = ctx.layout ? findPlaceholder(ctx.layout.shapes, shape.ph) : undefined;
  if (l?.text?.prompt) return l.text.prompt;
  return defaultPrompt(shape.ph.type);
}

export function defaultPrompt(t: Placeholder["type"]): string | undefined {
  switch (t) {
    case "title":
    case "ctrTitle":
      return "Başlık eklemek için tıklayın";
    case "subTitle":
      return "Alt başlık eklemek için tıklayın";
    case "body":
    case "obj":
      return "Metin eklemek için tıklayın";
    case "pic":
      return "Resim eklemek için simgeye tıklayın";
    case "tbl":
      return "Tablo eklemek için simgeye tıklayın";
    default:
      return undefined;
  }
}

export function isTextEmpty(body: TextBody | undefined): boolean {
  if (!body) return true;
  return body.paragraphs.every((p) => p.runs.every((r) => r.text === ""));
}

export function plainText(body: TextBody | undefined): string {
  if (!body) return "";
  return body.paragraphs.map((p) => p.runs.map((r) => r.text).join("")).join("\n");
}
