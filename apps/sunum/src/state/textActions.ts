/**
 * Metin biçimi komutları: düzenleme modundaysa seçime, değilse seçili
 * şekillerin (ve tablo hücrelerinin) tüm metnine uygulanır.
 */
import { applyParaProps, applyRunPropsAll, changeCase, transformText, endPos, type CaseMode } from "../model/text";
import { effectiveBodyProps, effectivePara, effectiveRun, levelResolver, type TextContext } from "../model/textStyle";
import type { BulletDef, LevelStyle, Paragraph, ParaProps, RunProps, Shape, SpShape, TableShape, TextBody } from "../model/types";
import { cellFormat } from "../render/tableStyles";
import type { Actions } from "./actions";

export type RunPatch = (p: RunProps, para: Paragraph, eff: RunProps) => RunProps;

export interface TextInfo {
  run: RunProps;
  runs: RunProps[];
  para: ParaProps;
  paras: Paragraph[];
  body?: Pick<TextBody, "anchor" | "vert" | "autofit" | "wrap">;
  editing: boolean;
}

function ctxOf(a: Actions): TextContext {
  return { master: a.pres.master, layout: a.layout() };
}

function cellLvl(t: TableShape, r: number, c: number, base: (n: number) => LevelStyle) {
  const fmt = cellFormat(t, r, c);
  return (n: number): LevelStyle => {
    const b = base(n);
    return { para: b.para, run: { ...b.run, color: fmt.text ?? b.run.color, bold: fmt.bold || b.run.bold } };
  };
}

/** Şeklin içindeki metin gövdelerini (düzey çözücüsüyle) dolaşır ve dönüştürür. */
function mapBodies(s: Shape, ctx: TextContext, fn: (b: TextBody, lvl: (n: number) => LevelStyle) => TextBody, cells?: { r1: number; c1: number; r2: number; c2: number } | null): Shape {
  if (s.kind === "sp") {
    if (!s.text) return s;
    const lvl = levelResolver(s, ctx);
    return { ...s, text: fn(s.text, lvl) };
  }
  if (s.kind === "table") {
    const base = levelResolver({ kind: "sp" }, ctx, "other");
    return {
      ...s,
      rows: s.rows.map((row, r) => ({
        ...row,
        cells: row.cells.map((cell, c) => {
          if (cells && (r < Math.min(cells.r1, cells.r2) || r > Math.max(cells.r1, cells.r2) || c < Math.min(cells.c1, cells.c2) || c > Math.max(cells.c1, cells.c2))) return cell;
          return { ...cell, text: fn(cell.text, cellLvl(s, r, c, base)) };
        }),
      })),
    };
  }
  if (s.kind === "group") return { ...s, children: s.children.map((c) => mapBodies(c, ctx, fn)) };
  return s;
}

function collectBodies(s: Shape, ctx: TextContext, out: { body: TextBody; lvl: (n: number) => LevelStyle; shape: Shape }[]) {
  if (s.kind === "sp" && s.text) out.push({ body: s.text, lvl: levelResolver(s, ctx), shape: s });
  else if (s.kind === "table") {
    const base = levelResolver({ kind: "sp" }, ctx, "other");
    s.rows.forEach((row, r) => row.cells.forEach((cell, c) => out.push({ body: cell.text, lvl: cellLvl(s, r, c, base), shape: s })));
  } else if (s.kind === "group") s.children.forEach((c) => collectBodies(c, ctx, out));
}

export function applyRun(a: Actions, patch: RunPatch) {
  const h = a.textHandle;
  if (a.s.editing && h) {
    const lvlCache = new Map<Paragraph, RunProps>();
    void lvlCache;
    const effList = h.runProps();
    h.applyRun((p, para) => patch(p, para, { ...(effList[0] ?? {}), ...p }));
    a.store.set((s) => ({ textTick: s.textTick + 1 }));
    return;
  }
  const ctx = ctxOf(a);
  const cells = a.s.cellSel;
  a.updateShapes((s) => mapBodies(s, ctx, (b, lvl) => applyRunPropsAll(b, (p, para) => patch(p, para, effectiveRun(p, para, lvl))), s.kind === "table" ? cells : null));
}

export function applyPara(a: Actions, patch: (p: Paragraph) => Paragraph) {
  const h = a.textHandle;
  if (a.s.editing && h) {
    h.applyPara(patch);
    a.store.set((s) => ({ textTick: s.textTick + 1 }));
    return;
  }
  const ctx = ctxOf(a);
  a.updateShapes((s) => mapBodies(s, ctx, (b) => applyParaProps(b, 0, b.paragraphs.length - 1, patch), s.kind === "table" ? a.s.cellSel : null));
}

export function setParaProps(a: Actions, props: Partial<ParaProps>) {
  applyPara(a, (p) => ({ ...p, props: { ...p.props, ...props } }));
}

/** Numaralandırma için asılı girinti (514350 EMU ≈ 54 px). */
const NUM_INDENT = 54;

/**
 * Madde işareti/numaralandırma uygular. Numaralandırmada metnin numaradan
 * sonra başlaması için girinti genişletilir; madde işaretine dönülünce
 * numaralandırmanın koyduğu girinti kaldırılır (asıl slayt değeri geçerli olur).
 */
export function setListStyle(a: Actions, bullet: BulletDef, noInheritedIndent = false) {
  applyPara(a, (p) => {
    const props: ParaProps = { ...p.props, bullet };
    if (bullet.type === "auto") {
      props.marL = NUM_INDENT + p.level * 48;
      props.indent = -NUM_INDENT;
    } else if (bullet.type === "char" && noInheritedIndent) {
      props.marL = 24 + p.level * 48;
      props.indent = -24;
    } else if (p.props.indent === -NUM_INDENT) {
      delete props.marL;
      delete props.indent;
    }
    return { ...p, props };
  });
}

export function changeLevel(a: Actions, delta: number) {
  const h = a.textHandle;
  if (a.s.editing && h) {
    h.applyPara((p) => p, delta);
    return;
  }
  // Yer tutucularda 5 düzey (0–4), diğer metinlerde 9 düzey.
  const ctx = ctxOf(a);
  a.updateShapes((s) => {
    const max = s.kind === "sp" && s.ph ? 4 : 8;
    return mapBodies(s, ctx, (b) => applyParaProps(b, 0, b.paragraphs.length - 1, (p) => ({ ...p, level: Math.max(0, Math.min(max, p.level + delta)) })), s.kind === "table" ? a.s.cellSel : null);
  });
}

export function setBodyProps(a: Actions, patch: Partial<TextBody>) {
  a.updateSelectedDeep((s) => {
    if (s.kind === "sp") return { ...s, text: { ...(s.text ?? { paragraphs: [{ runs: [], level: 0, props: {} }] }), ...patch } };
    if (s.kind === "table" && patch.anchor) return { ...s, rows: s.rows.map((r) => ({ ...r, cells: r.cells.map((c) => ({ ...c, anchor: patch.anchor })) })) };
    return s;
  });
}

export function applyCase(a: Actions, mode: CaseMode) {
  const h = a.textHandle;
  if (a.s.editing && h) {
    h.transform((t, c) => changeCase(t, mode, c.paraStart));
    return;
  }
  const ctx = ctxOf(a);
  a.updateShapes((s) => mapBodies(s, ctx, (b) => transformText(b, { p: 0, o: 0 }, endPos(b), (t, c) => changeCase(t, mode, c.paraStart))));
}

/** Sunu yazı tipi boyutu listesi (Yazı Tipi Boyutu kutusu ve Ctrl+Shift+> / < adımları). */
export const SLIDE_FONT_SIZES = [8, 9, 10, 10.5, 11, 12, 14, 16, 18, 20, 24, 28, 32, 36, 40, 44, 48, 54, 60, 66, 72, 80, 88, 96];
const MAX_LISTED = SLIDE_FONT_SIZES[SLIDE_FONT_SIZES.length - 1];

export function stepFontSize(size: number, dir: 1 | -1): number {
  if (dir > 0) {
    const next = SLIDE_FONT_SIZES.find((v) => v > size + 0.01);
    return next ?? (Math.floor(size / 10) + 1) * 10;
  }
  if (size > MAX_LISTED) return Math.max(MAX_LISTED, Math.floor((size - 1) / 10) * 10);
  const prev = [...SLIDE_FONT_SIZES].reverse().find((v) => v < size - 0.01);
  return prev ?? Math.max(1, size - 1);
}

export function growFont(a: Actions, dir: 1 | -1) {
  applyRun(a, (p, _para, eff) => ({ ...p, size: stepFontSize(eff.size ?? 18, dir) }));
}

/** Şerit durumunu göstermek için seçimin metin bilgisi. */
export function textInfo(a: Actions): TextInfo | null {
  const s = a.s;
  const h = a.textHandle;
  const ctx = ctxOf(a);
  if (s.editing && h) {
    const runs = h.runProps();
    const paras = h.paragraphs();
    const shape = a.selectedShapes()[0];
    let lvl: ((n: number) => LevelStyle) | null = null;
    let body: TextInfo["body"];
    if (shape?.kind === "sp") {
      lvl = levelResolver(shape, ctx);
      body = effectiveBodyProps(shape, ctx);
    } else if (shape?.kind === "table" && s.editing.cell) {
      const [r, c] = s.editing.cell;
      lvl = cellLvl(shape, r, c, levelResolver({ kind: "sp" }, ctx, "other"));
      body = { anchor: shape.rows[r]?.cells[c]?.anchor ?? "t" };
    }
    return {
      run: runs[0] ?? {},
      runs,
      para: lvl && paras[0] ? effectivePara(paras[0], lvl) : paras[0]?.props ?? {},
      paras,
      body,
      editing: true,
    };
  }
  const shapes = a.selectedShapes();
  if (!shapes.length) return null;
  const list: { body: TextBody; lvl: (n: number) => LevelStyle; shape: Shape }[] = [];
  for (const sh of shapes) collectBodies(sh, ctx, list);
  if (!list.length) return null;
  const runs: RunProps[] = [];
  const paras: Paragraph[] = [];
  for (const { body, lvl } of list) {
    for (const p of body.paragraphs) {
      paras.push(p);
      if (!p.runs.length) runs.push(effectiveRun(p.endProps ?? {}, p, lvl));
      for (const r of p.runs) runs.push(effectiveRun(r.props, p, lvl));
    }
  }
  const first = list[0];
  const sp = first.shape.kind === "sp" ? (first.shape as SpShape) : undefined;
  return {
    run: runs[0] ?? {},
    runs,
    para: effectivePara(first.body.paragraphs[0] ?? { runs: [], level: 0, props: {} }, first.lvl),
    paras,
    body: sp ? effectiveBodyProps(sp, ctx) : { anchor: first.body.anchor ?? "t" },
    editing: false,
  };
}

/** Seçimde tüm değerler aynıysa o değer, değilse undefined. */
export function common<T>(list: RunProps[], key: keyof RunProps): T | undefined {
  if (!list.length) return undefined;
  const v = JSON.stringify(list[0][key]);
  return list.every((r) => JSON.stringify(r[key]) === v) ? (list[0][key] as T) : undefined;
}

export type { CaseMode };
