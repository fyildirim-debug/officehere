import type { Color, Fill, Line, Paragraph, PicShape, RunProps, Shape, SpShape, TableCell, TableShape, TextBody } from "./types";

let counter = 0;
/** Oturum içinde benzersiz kimlik. */
export function uid(prefix = "s"): string {
  counter += 1;
  return `${prefix}${Date.now().toString(36)}${counter.toString(36)}${Math.floor(Math.random() * 1296).toString(36)}`;
}

export const NO_FILL: Fill = { type: "none" };
export const NO_LINE: Line = { fill: { type: "none" }, width: 1, dash: "solid" };

export const solid = (color: Color): Fill => ({ type: "solid", color });
export const lineOf = (color: Color, width = 1, dash: Line["dash"] = "solid"): Line => ({ fill: { type: "solid", color }, width, dash });

/** Varsayılan şekil biçimi: Vurgu 1 dolgu, %50 koyu anahat (Office varsayılanına benzer). */
export const DEFAULT_SHAPE_FILL: Fill = { type: "solid", color: { scheme: "accent1" } };
export const DEFAULT_SHAPE_LINE: Line = { fill: { type: "solid", color: { scheme: "accent1", lumMod: 0.5 } }, width: 1, dash: "solid" };

export function para(text = "", props: RunProps = {}, level = 0): Paragraph {
  return { runs: text ? [{ text, props: { ...props } }] : [], level, props: {}, endProps: { ...props } };
}

/** Düz metinden (satır başına paragraf) metin gövdesi. "\t" ile başlayan satırlar alt düzeydir. */
export function textBody(text: string, opts: Partial<TextBody> = {}, props: RunProps = {}): TextBody {
  const lines = text.split("\n");
  return {
    paragraphs: lines.map((l) => {
      let level = 0;
      while (l.startsWith("\t")) {
        level += 1;
        l = l.slice(1);
      }
      return para(l, props, level);
    }),
    ...opts,
  };
}

export function emptyBody(opts: Partial<TextBody> = {}): TextBody {
  return { paragraphs: [para()], ...opts };
}

export function makeSp(geom: string, x: number, y: number, w: number, h: number, extra: Partial<SpShape> = {}): SpShape {
  return {
    kind: "sp",
    id: uid(),
    name: extra.name ?? "Şekil",
    x,
    y,
    w,
    h,
    rot: 0,
    geom,
    fill: NO_FILL,
    line: NO_LINE,
    ...extra,
  };
}

export function makeTextBox(x: number, y: number, w: number, h: number, text = "", autoWidth = false): SpShape {
  return makeSp("rect", x, y, w, h, {
    name: "Metin Kutusu",
    txBox: true,
    text: { ...textBody(text), wrap: !autoWidth, autofit: "shape", anchor: "t" },
  });
}

export function makePicture(src: string, x: number, y: number, w: number, h: number, name = "Resim"): PicShape {
  return { kind: "pic", id: uid(), name, x, y, w, h, rot: 0, src, line: NO_LINE };
}

export function emptyCell(): TableCell {
  return { text: emptyBody() };
}

export function makeTable(rows: number, cols: number, x: number, y: number, w: number, rowH = 40): TableShape {
  const cw = w / cols;
  return {
    kind: "table",
    id: uid(),
    name: "Tablo",
    x,
    y,
    w,
    h: rowH * rows,
    rot: 0,
    cols: Array.from({ length: cols }, () => cw),
    rows: Array.from({ length: rows }, () => ({ h: rowH, cells: Array.from({ length: cols }, emptyCell) })),
    style: { styleId: "medium2-accent1", firstRow: true, lastRow: false, firstCol: false, lastCol: false, bandRow: true, bandCol: false },
  };
}

/** Şekli (ve alt şekillerini) yeni kimliklerle kopyalar. */
export function cloneShape<T extends Shape>(s: T, idMap?: Map<string, string>): T {
  const copy = structuredClone(s) as T;
  const reid = (sh: Shape) => {
    const nid = uid();
    idMap?.set(sh.id, nid);
    sh.id = nid;
    if (sh.kind === "group") sh.children.forEach(reid);
  };
  reid(copy);
  return copy;
}
