/** Bul ve Değiştir: tüm slaytlardaki şekil ve tablo hücresi metinleri. */
import { deleteRange, findInBody, insertText, propsAt, type Pos } from "../model/text";
import type { Presentation, Shape, TextBody } from "../model/types";

export interface Match {
  slideIdx: number;
  shapeId: string;
  cell?: [number, number];
  from: Pos;
  to: Pos;
}

export interface FindOptions {
  matchCase: boolean;
  wholeWord: boolean;
}

function bodiesOf(s: Shape, out: { shapeId: string; cell?: [number, number]; body: TextBody }[]) {
  if (s.kind === "sp" && s.text) out.push({ shapeId: s.id, body: s.text });
  else if (s.kind === "table") s.rows.forEach((r, ri) => r.cells.forEach((c, ci) => out.push({ shapeId: s.id, cell: [ri, ci], body: c.text })));
  else if (s.kind === "group") s.children.forEach((c) => bodiesOf(c, out));
}

export function findAll(pres: Presentation, query: string, o: FindOptions): Match[] {
  const out: Match[] = [];
  pres.slides.forEach((sl, i) => {
    const list: { shapeId: string; cell?: [number, number]; body: TextBody }[] = [];
    sl.shapes.forEach((s) => bodiesOf(s, list));
    for (const b of list) for (const [from, to] of findInBody(b.body, query, o.matchCase, o.wholeWord)) out.push({ slideIdx: i, shapeId: b.shapeId, cell: b.cell, from, to });
  });
  return out;
}

function replaceInBody(body: TextBody, query: string, repl: string, o: FindOptions): { body: TextBody; count: number } {
  const matches = findInBody(body, query, o.matchCase, o.wholeWord);
  let b = body;
  for (let k = matches.length - 1; k >= 0; k--) {
    const [from, to] = matches[k];
    const props = propsAt(b, { p: from.p, o: from.o + 1 > to.o ? from.o : from.o + 1 });
    b = deleteRange(b, from, to);
    b = insertText(b, { anchor: from, focus: from }, repl, props).body;
  }
  return { body: b, count: matches.length };
}

function replaceShape(s: Shape, query: string, repl: string, o: FindOptions, counter: { n: number }): Shape {
  if (s.kind === "sp" && s.text) {
    const r = replaceInBody(s.text, query, repl, o);
    counter.n += r.count;
    return r.count ? { ...s, text: r.body } : s;
  }
  if (s.kind === "table") {
    let changed = false;
    const rows = s.rows.map((row) => ({
      ...row,
      cells: row.cells.map((c) => {
        const r = replaceInBody(c.text, query, repl, o);
        counter.n += r.count;
        if (r.count) changed = true;
        return r.count ? { ...c, text: r.body } : c;
      }),
    }));
    return changed ? { ...s, rows } : s;
  }
  if (s.kind === "group") return { ...s, children: s.children.map((c) => replaceShape(c, query, repl, o, counter)) };
  return s;
}

export function replaceAll(pres: Presentation, query: string, repl: string, o: FindOptions): { pres: Presentation; count: number } {
  const counter = { n: 0 };
  const slides = pres.slides.map((sl) => ({ ...sl, shapes: sl.shapes.map((s) => replaceShape(s, query, repl, o, counter)) }));
  return { pres: counter.n ? { ...pres, slides } : pres, count: counter.n };
}

/** Yazı Tiplerini Değiştir: tüm çalıştırmalarda yazı tipi adını değiştirir. */
export function replaceFonts(pres: Presentation, from: string, to: string): Presentation {
  const fixBody = (b: TextBody): TextBody => ({
    ...b,
    paragraphs: b.paragraphs.map((p) => ({ ...p, runs: p.runs.map((r) => (r.props.font === from ? { ...r, props: { ...r.props, font: to } } : r)), endProps: p.endProps?.font === from ? { ...p.endProps, font: to } : p.endProps })),
  });
  const fix = (s: Shape): Shape => {
    if (s.kind === "sp" && s.text) return { ...s, text: fixBody(s.text) };
    if (s.kind === "table") return { ...s, rows: s.rows.map((r) => ({ ...r, cells: r.cells.map((c) => ({ ...c, text: fixBody(c.text) })) })) };
    if (s.kind === "group") return { ...s, children: s.children.map(fix) };
    return s;
  };
  const theme = { ...pres.theme, fonts: { major: pres.theme.fonts.major === from ? to : pres.theme.fonts.major, minor: pres.theme.fonts.minor === from ? to : pres.theme.fonts.minor } };
  return { ...pres, theme, slides: pres.slides.map((sl) => ({ ...sl, shapes: sl.shapes.map(fix) })) };
}

export function usedFonts(pres: Presentation): string[] {
  const set = new Set<string>([pres.theme.fonts.major, pres.theme.fonts.minor]);
  const walk = (s: Shape) => {
    const bodies: TextBody[] = [];
    if (s.kind === "sp" && s.text) bodies.push(s.text);
    if (s.kind === "table") s.rows.forEach((r) => r.cells.forEach((c) => bodies.push(c.text)));
    if (s.kind === "group") s.children.forEach(walk);
    for (const b of bodies) for (const p of b.paragraphs) for (const r of p.runs) if (r.props.font && !r.props.font.startsWith("+")) set.add(r.props.font);
  };
  pres.slides.forEach((sl) => sl.shapes.forEach(walk));
  return [...set].sort((a, b) => a.localeCompare(b, "tr"));
}
