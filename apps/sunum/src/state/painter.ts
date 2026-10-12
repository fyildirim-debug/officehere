/** Biçim Boyacısı: şekil dolgu/anahat/efekt ve metin biçimini kopyalar. */
import { findShape } from "../model/ops";
import { applyRunPropsAll } from "../model/text";
import type { Shape } from "../model/types";
import type { Actions } from "./actions";

export function captureFormat(a: Actions, sticky: boolean) {
  const sh = a.selectedShapes()[0];
  if (!sh) return;
  let run = undefined;
  let para = undefined;
  if (a.s.editing && a.textHandle) {
    const rp = a.textHandle.runProps()[0];
    run = rp ? { ...rp } : undefined;
    para = a.textHandle.paragraphs()[0]?.props;
  } else if (sh.kind === "sp" && sh.text) {
    const p0 = sh.text.paragraphs.find((p) => p.runs.length) ?? sh.text.paragraphs[0];
    run = p0?.runs[0]?.props ?? p0?.endProps;
    para = p0?.props;
  }
  const shape: Partial<Shape> = sh.kind === "sp" ? { fill: sh.fill, line: sh.line, effects: sh.effects } : sh.kind === "pic" ? { line: sh.line, effects: sh.effects } : {};
  a.store.set({ painter: { shape, run, para, sticky } });
}

export function applyFormatPainter(a: Actions, targetId: string) {
  const p = a.s.painter;
  if (!p) return;
  const slide = a.slide();
  const target = slide ? findShape(slide.shapes, targetId) : undefined;
  if (!target) return;
  a.updateShapes((s) => {
    let out: Shape = s;
    if (p.shape && (s.kind === "sp" || s.kind === "pic")) {
      const src = p.shape as { line?: unknown; effects?: unknown };
      const patch = s.kind === "pic" ? { line: src.line ?? s.line, effects: src.effects } : p.shape;
      out = { ...s, ...(patch as object) } as Shape;
    }
    if (p.run && out.kind === "sp" && out.text) {
      const run = p.run;
      out = { ...out, text: applyRunPropsAll(out.text, () => ({ ...run })) };
      if (p.para) out = { ...out, text: { ...out.text!, paragraphs: out.text!.paragraphs.map((pp) => ({ ...pp, props: { ...p.para } })) } } as Shape;
    }
    return out;
  }, [targetId]);
  if (!p.sticky) a.store.set({ painter: null });
}
