/**
 * Metin gövdesi düzenleme işlemleri (saf fonksiyonlar). Konum: paragraf
 * dizini + paragraf içi karakter uzaklığı. Satır sonu (Shift+Enter) "\v"
 * karakteriyle tutulur.
 */
import type { Paragraph, ParaProps, Run, RunProps, TextBody } from "./types";

export interface Pos {
  p: number;
  o: number;
}

export interface Sel {
  anchor: Pos;
  focus: Pos;
}

export const cmpPos = (a: Pos, b: Pos) => (a.p !== b.p ? a.p - b.p : a.o - b.o);
export const orderSel = (s: Sel): [Pos, Pos] => (cmpPos(s.anchor, s.focus) <= 0 ? [s.anchor, s.focus] : [s.focus, s.anchor]);
export const isCollapsed = (s: Sel) => cmpPos(s.anchor, s.focus) === 0;
export const caret = (p: number, o: number): Sel => ({ anchor: { p, o }, focus: { p, o } });

export function paraText(p: Paragraph): string {
  let s = "";
  for (const r of p.runs) s += r.text;
  return s;
}

export function paraLen(p: Paragraph): number {
  let n = 0;
  for (const r of p.runs) n += r.text.length;
  return n;
}

export function propsEqual(a: RunProps, b: RunProps): boolean {
  const ka = Object.keys(a).filter((k) => (a as Record<string, unknown>)[k] !== undefined);
  const kb = Object.keys(b).filter((k) => (b as Record<string, unknown>)[k] !== undefined);
  if (ka.length !== kb.length) return false;
  for (const k of ka) {
    if (JSON.stringify((a as Record<string, unknown>)[k]) !== JSON.stringify((b as Record<string, unknown>)[k])) return false;
  }
  return true;
}

export function normalizeRuns(runs: Run[]): Run[] {
  const out: Run[] = [];
  for (const r of runs) {
    if (!r.text) continue;
    const last = out[out.length - 1];
    if (last && !last.props.field && !r.props.field && propsEqual(last.props, r.props)) out[out.length - 1] = { text: last.text + r.text, props: last.props };
    else out.push(r);
  }
  return out;
}

/** [start, end) aralığındaki çalıştırmaları böler: önce, orta, sonra. */
function splitRuns(runs: Run[], start: number, end: number): [Run[], Run[], Run[]] {
  const a: Run[] = [];
  const m: Run[] = [];
  const b: Run[] = [];
  let pos = 0;
  for (const r of runs) {
    const rs = pos;
    const re = pos + r.text.length;
    pos = re;
    if (re <= start) a.push(r);
    else if (rs >= end) b.push(r);
    else {
      const s1 = Math.max(start, rs) - rs;
      const s2 = Math.min(end, re) - rs;
      if (s1 > 0) a.push({ text: r.text.slice(0, s1), props: r.props });
      m.push({ text: r.text.slice(s1, s2), props: r.props });
      if (s2 < r.text.length) b.push({ text: r.text.slice(s2), props: r.props });
    }
  }
  return [a, m, b];
}

/** İmleçteki (yazılacak metnin) biçimi. */
export function propsAt(body: TextBody, pos: Pos): RunProps {
  const p = body.paragraphs[pos.p];
  if (!p) return {};
  if (!p.runs.length) return p.endProps ?? {};
  let acc = 0;
  for (let i = 0; i < p.runs.length; i++) {
    const r = p.runs[i];
    acc += r.text.length;
    if (pos.o <= acc && (pos.o > acc - r.text.length || (pos.o === 0 && i === 0))) {
      // Köprünün sonunda yazılan metin köprüye katılmaz.
      return pos.o === acc ? stripLink(stripField(r.props)) : stripField(r.props);
    }
  }
  return stripLink(stripField(p.runs[p.runs.length - 1].props));
}

function stripLink(p: RunProps): RunProps {
  if (!p.link) return p;
  const { link: _l, ...rest } = p;
  void _l;
  return rest;
}

function stripField(p: RunProps): RunProps {
  if (!p.field) return p;
  const { field: _f, ...rest } = p;
  void _f;
  return rest;
}

/** Alan (slayt numarası vb.) çalıştırmalarını bölmemek için aralığı genişletir. */
function expandFields(p: Paragraph, s: number, e: number): [number, number] {
  let pos = 0;
  for (const r of p.runs) {
    const rs = pos;
    const re = pos + r.text.length;
    pos = re;
    if (!r.props.field) continue;
    if (s < re && e > rs) {
      s = Math.min(s, rs);
      e = Math.max(e, re);
    }
  }
  return [s, e];
}

function snapOutOfField(p: Paragraph, o: number): number {
  let pos = 0;
  for (const r of p.runs) {
    const rs = pos;
    const re = pos + r.text.length;
    pos = re;
    if (r.props.field && o > rs && o < re) return re;
  }
  return o;
}

export function deleteRange(body: TextBody, from: Pos, to: Pos): TextBody {
  if (cmpPos(from, to) > 0) [from, to] = [to, from];
  const paras = body.paragraphs.slice();
  if (from.p === to.p) {
    const p = paras[from.p];
    const [s, e] = expandFields(p, from.o, to.o);
    const [a, , b] = splitRuns(p.runs, s, e);
    const endProps = p.runs.length ? propsAt(body, { p: from.p, o: s }) : p.endProps;
    paras[from.p] = { ...p, runs: normalizeRuns([...a, ...b]), endProps: a.length + b.length ? p.endProps : endProps };
    return { ...body, paragraphs: paras };
  }
  const p1 = paras[from.p];
  const p2 = paras[to.p];
  // İlk paragrafın başından seçim başlangıcına kadarki kısım (splitRuns'un ORTA parçası) korunur;
  // son paragrafın seçim sonundan sonraki kısmı eklenir.
  const [, a] = splitRuns(p1.runs, 0, expandFields(p1, from.o, paraLen(p1))[0]);
  const [, , b] = splitRuns(p2.runs, 0, expandFields(p2, 0, to.o)[1]);
  const keepA = a.filter((r) => r.text);
  // Birleşen paragraf boş kalırsa yazılacak metnin biçimi: ilk paragrafın sonu ya da ikincinin başı.
  const endProps = keepA.length ? p1.endProps : p1.runs.length ? propsAt(body, from) : (p1.endProps ?? p2.endProps);
  const merged: Paragraph = { ...p1, runs: normalizeRuns([...keepA, ...b]), endProps };
  paras.splice(from.p, to.p - from.p + 1, merged);
  return { ...body, paragraphs: paras };
}

/** i. paragrafı bir sonrakiyle birleştirir (biçimli çalımlar korunur). */
export function mergeParagraphs(body: TextBody, i: number): TextBody {
  if (i < 0 || i >= body.paragraphs.length - 1) return body;
  return deleteRange(body, { p: i, o: paraLen(body.paragraphs[i]) }, { p: i + 1, o: 0 });
}

/** Backspace: seçimi ya da imleçten önceki karakteri siler. Girintili paragrafın başında önce düzeyi düşürür. */
export function removeBackward(body: TextBody, sel: Sel): { body: TextBody; caret: Pos } {
  const [a, z] = orderSel(sel);
  if (cmpPos(a, z) !== 0) return { body: deleteRange(body, a, z), caret: a };
  const p = body.paragraphs[a.p];
  if (a.o === 0 && p && p.level > 0) {
    const paragraphs = body.paragraphs.slice();
    paragraphs[a.p] = { ...p, level: p.level - 1 };
    return { body: { ...body, paragraphs }, caret: a };
  }
  const from = prevPos(body, a);
  if (cmpPos(from, a) === 0) return { body, caret: a };
  return { body: deleteRange(body, from, a), caret: from };
}

/** Delete: seçimi ya da imleçten sonraki karakteri siler (paragraf sonunda birleştirir). */
export function removeForward(body: TextBody, sel: Sel): { body: TextBody; caret: Pos } {
  const [a, z] = orderSel(sel);
  if (cmpPos(a, z) !== 0) return { body: deleteRange(body, a, z), caret: a };
  const to = nextPos(body, a);
  if (cmpPos(to, a) === 0) return { body, caret: a };
  return { body: deleteRange(body, a, to), caret: a };
}

/** Seçimi siler ve metni ekler; yeni imleç konumunu döndürür. */
export function insertText(body: TextBody, sel: Sel, text: string, props?: RunProps): { body: TextBody; caret: Pos } {
  let [from] = orderSel(sel);
  const insProps = props ?? propsAt(body, from);
  let b = isCollapsed(sel) ? body : deleteRange(body, sel.anchor, sel.focus);
  const p0 = b.paragraphs[from.p];
  from = { p: from.p, o: snapOutOfField(p0, from.o) };
  const lines = text.replace(/\r\n?/g, "\n").split("\n");
  const paras = b.paragraphs.slice();
  const p = paras[from.p];
  const [a, , rest] = splitRuns(p.runs, from.o, from.o);
  const first: Paragraph = { ...p, runs: normalizeRuns([...a, { text: lines[0], props: insProps }]) };
  if (lines.length === 1) {
    first.runs = normalizeRuns([...first.runs, ...rest]);
    if (!first.runs.length) first.endProps = insProps;
    paras[from.p] = first;
    b = { ...b, paragraphs: paras };
    return { body: b, caret: { p: from.p, o: from.o + lines[0].length } };
  }
  const newParas: Paragraph[] = [first];
  for (let i = 1; i < lines.length; i++) {
    const isLast = i === lines.length - 1;
    newParas.push({
      runs: normalizeRuns([{ text: lines[i], props: insProps }, ...(isLast ? rest : [])]),
      level: p.level,
      props: { ...p.props },
      endProps: insProps,
    });
  }
  paras.splice(from.p, 1, ...newParas);
  return { body: { ...b, paragraphs: paras }, caret: { p: from.p + lines.length - 1, o: lines[lines.length - 1].length } };
}

export function splitParagraph(body: TextBody, sel: Sel, props?: RunProps): { body: TextBody; caret: Pos } {
  const [from] = orderSel(sel);
  const b = isCollapsed(sel) ? body : deleteRange(body, sel.anchor, sel.focus);
  const paras = b.paragraphs.slice();
  const p = paras[from.p];
  const o = snapOutOfField(p, from.o);
  const endP = props ?? propsAt(b, { p: from.p, o });
  const [a, , rest] = splitRuns(p.runs, o, o);
  paras.splice(from.p, 1, { ...p, runs: a, endProps: endP }, { runs: rest, level: p.level, props: { ...p.props }, endProps: endP });
  return { body: { ...b, paragraphs: paras }, caret: { p: from.p + 1, o: 0 } };
}

/** Aralıktaki çalıştırmalara biçim uygular. */
export function applyRunProps(body: TextBody, from: Pos, to: Pos, patch: (p: RunProps, para: Paragraph) => RunProps): TextBody {
  if (cmpPos(from, to) > 0) [from, to] = [to, from];
  const paras = body.paragraphs.map((p, i) => {
    if (i < from.p || i > to.p) return p;
    const s = i === from.p ? from.o : 0;
    const e = i === to.p ? to.o : paraLen(p);
    const whole = s === 0 && e === paraLen(p);
    const [a, m, b] = splitRuns(p.runs, s, e);
    const mid = m.map((r) => ({ text: r.text, props: clean(patch(r.props, p)) }));
    const np: Paragraph = { ...p, runs: normalizeRuns([...a, ...mid, ...b]) };
    if (whole || !p.runs.length || (i < to.p && e === paraLen(p))) np.endProps = clean(patch(p.endProps ?? {}, p));
    return np;
  });
  return { ...body, paragraphs: paras };
}

function clean<T extends object>(p: T): T {
  const o = {} as T;
  for (const [k, v] of Object.entries(p)) if (v !== undefined) (o as Record<string, unknown>)[k] = v;
  return o;
}

export function applyRunPropsAll(body: TextBody, patch: (p: RunProps, para: Paragraph) => RunProps): TextBody {
  const last = body.paragraphs.length - 1;
  return applyRunProps(body, { p: 0, o: 0 }, { p: last, o: paraLen(body.paragraphs[last]) }, patch);
}

export function applyParaProps(body: TextBody, fromP: number, toP: number, patch: (p: Paragraph) => Paragraph): TextBody {
  if (fromP > toP) [fromP, toP] = [toP, fromP];
  return { ...body, paragraphs: body.paragraphs.map((p, i) => (i >= fromP && i <= toP ? patch(p) : p)) };
}

/**
 * fromP..toP paragraflarını bir paragraf yukarı (dir=-1) ya da aşağı (dir=1) taşır
 * (Alt+Shift+Yukarı/Aşağı). Sınırda değişiklik yoksa null.
 */
export function moveParagraphs(body: TextBody, fromP: number, toP: number, dir: -1 | 1): TextBody | null {
  if (fromP > toP) [fromP, toP] = [toP, fromP];
  const ps = body.paragraphs;
  if ((dir < 0 && fromP === 0) || (dir > 0 && toP >= ps.length - 1)) return null;
  const block = ps.slice(fromP, toP + 1);
  const rest = ps.slice(0, fromP).concat(ps.slice(toP + 1));
  const at = fromP + dir;
  return { ...body, paragraphs: [...rest.slice(0, at), ...block, ...rest.slice(at)] };
}

export function setParaProps(body: TextBody, fromP: number, toP: number, props: Partial<ParaProps>): TextBody {
  return applyParaProps(body, fromP, toP, (p) => ({ ...p, props: clean({ ...p.props, ...props }) as ParaProps }));
}

/** Aralıktaki tüm çalıştırma biçimleri (şerit durumunu göstermek için). */
export function runPropsInRange(body: TextBody, from: Pos, to: Pos): RunProps[] {
  if (cmpPos(from, to) > 0) [from, to] = [to, from];
  if (cmpPos(from, to) === 0) return [propsAt(body, from)];
  const out: RunProps[] = [];
  for (let i = from.p; i <= to.p; i++) {
    const p = body.paragraphs[i];
    const s = i === from.p ? from.o : 0;
    const e = i === to.p ? to.o : paraLen(p);
    const [, m] = splitRuns(p.runs, s, e);
    for (const r of m) out.push(r.props);
    if (!p.runs.length) out.push(p.endProps ?? {});
  }
  return out.length ? out : [propsAt(body, from)];
}

const WORD = /[\p{L}\p{N}_]/u;

export function wordRangeAt(body: TextBody, pos: Pos): [Pos, Pos] | null {
  const t = paraText(body.paragraphs[pos.p]);
  let s = pos.o;
  let e = pos.o;
  while (s > 0 && WORD.test(t[s - 1])) s--;
  while (e < t.length && WORD.test(t[e])) e++;
  if (s === e || s === pos.o || e === pos.o) return null;
  return [
    { p: pos.p, o: s },
    { p: pos.p, o: e },
  ];
}

export function wordBoundary(body: TextBody, pos: Pos, dir: -1 | 1): Pos {
  const t = paraText(body.paragraphs[pos.p]);
  let o = pos.o;
  if (dir < 0) {
    if (o === 0) return pos.p > 0 ? { p: pos.p - 1, o: paraLen(body.paragraphs[pos.p - 1]) } : pos;
    while (o > 0 && /\s/.test(t[o - 1])) o--;
    while (o > 0 && !/\s/.test(t[o - 1])) o--;
  } else {
    if (o >= t.length) return pos.p < body.paragraphs.length - 1 ? { p: pos.p + 1, o: 0 } : pos;
    while (o < t.length && !/\s/.test(t[o])) o++;
    while (o < t.length && /\s/.test(t[o])) o++;
  }
  return { p: pos.p, o };
}

export function prevPos(body: TextBody, pos: Pos): Pos {
  if (pos.o > 0) return { p: pos.p, o: pos.o - 1 };
  if (pos.p > 0) return { p: pos.p - 1, o: paraLen(body.paragraphs[pos.p - 1]) };
  return pos;
}

export function nextPos(body: TextBody, pos: Pos): Pos {
  if (pos.o < paraLen(body.paragraphs[pos.p])) return { p: pos.p, o: pos.o + 1 };
  if (pos.p < body.paragraphs.length - 1) return { p: pos.p + 1, o: 0 };
  return pos;
}

export function endPos(body: TextBody): Pos {
  const last = body.paragraphs.length - 1;
  return { p: last, o: paraLen(body.paragraphs[last]) };
}

export function selectAll(body: TextBody): Sel {
  return { anchor: { p: 0, o: 0 }, focus: endPos(body) };
}

/** Metni dönüştürür (büyük/küçük harf). */
export function transformText(body: TextBody, from: Pos, to: Pos, fn: (s: string, ctx: { paraStart: boolean }) => string): TextBody {
  if (cmpPos(from, to) > 0) [from, to] = [to, from];
  const paras = body.paragraphs.map((p, i) => {
    if (i < from.p || i > to.p) return p;
    const s = i === from.p ? from.o : 0;
    const e = i === to.p ? to.o : paraLen(p);
    const [a, m, b] = splitRuns(p.runs, s, e);
    let first = s === 0;
    const mid = m.map((r) => {
      const t = r.props.field ? r.text : fn(r.text, { paraStart: first });
      first = false;
      return { text: t, props: r.props };
    });
    return { ...p, runs: normalizeRuns([...a, ...mid, ...b]) };
  });
  return { ...body, paragraphs: paras };
}

export type CaseMode = "sentence" | "lower" | "upper" | "title" | "toggle";

export function changeCase(s: string, mode: CaseMode, paraStart: boolean): string {
  const up = (x: string) => x.toLocaleUpperCase("tr-TR");
  const lo = (x: string) => x.toLocaleLowerCase("tr-TR");
  switch (mode) {
    case "upper":
      return up(s);
    case "lower":
      return lo(s);
    case "title":
      return lo(s).replace(/(^|[\s\-("'“])(\p{L})/gu, (_m, p1: string, p2: string) => p1 + up(p2));
    case "toggle":
      return Array.from(s)
        .map((ch) => (ch === up(ch) ? lo(ch) : up(ch)))
        .join("");
    case "sentence": {
      const l = lo(s);
      let out = "";
      let start = paraStart;
      for (const ch of l) {
        if (start && /\p{L}/u.test(ch)) {
          out += up(ch);
          start = false;
        } else {
          out += ch;
          if (/[.!?]/.test(ch)) start = true;
        }
      }
      return out;
    }
  }
}

/** Aralıktaki metni düz metin olarak döndürür. */
export function textInRange(body: TextBody, from: Pos, to: Pos): string {
  if (cmpPos(from, to) > 0) [from, to] = [to, from];
  const parts: string[] = [];
  for (let i = from.p; i <= to.p; i++) {
    const t = paraText(body.paragraphs[i]);
    parts.push(t.slice(i === from.p ? from.o : 0, i === to.p ? to.o : t.length));
  }
  return parts.join("\n");
}

/** Aralığın zengin kopyası (yapıştırma için). */
export function sliceBody(body: TextBody, from: Pos, to: Pos): Paragraph[] {
  if (cmpPos(from, to) > 0) [from, to] = [to, from];
  const out: Paragraph[] = [];
  for (let i = from.p; i <= to.p; i++) {
    const p = body.paragraphs[i];
    const s = i === from.p ? from.o : 0;
    const e = i === to.p ? to.o : paraLen(p);
    const [, m] = splitRuns(p.runs, s, e);
    out.push({ ...p, runs: m });
  }
  return out;
}

/** Zengin paragrafları imlece ekler. */
export function insertParagraphs(body: TextBody, sel: Sel, paras: Paragraph[]): { body: TextBody; caret: Pos } {
  if (!paras.length) return { body, caret: orderSel(sel)[0] };
  if (paras.length === 1) {
    let b = isCollapsed(sel) ? body : deleteRange(body, sel.anchor, sel.focus);
    const [from] = orderSel(sel);
    const p = b.paragraphs[from.p];
    const [a, , rest] = splitRuns(p.runs, from.o, from.o);
    const list = b.paragraphs.slice();
    list[from.p] = { ...p, runs: normalizeRuns([...a, ...paras[0].runs, ...rest]) };
    b = { ...b, paragraphs: list };
    const added = paras[0].runs.reduce((n, r) => n + r.text.length, 0);
    return { body: b, caret: { p: from.p, o: from.o + added } };
  }
  const split = splitParagraph(body, sel);
  const b = split.body;
  const list = b.paragraphs.slice();
  const i = split.caret.p - 1;
  const head = list[i];
  const tail = list[i + 1];
  const mids = paras.slice(1, -1).map((p) => ({ ...p, props: { ...p.props } }));
  const lastIns = paras[paras.length - 1];
  const newHead: Paragraph = { ...head, runs: normalizeRuns([...head.runs, ...paras[0].runs]) };
  const newTail: Paragraph = { ...tail, level: lastIns.level, props: lastIns.props, runs: normalizeRuns([...lastIns.runs, ...tail.runs]) };
  list.splice(i, 2, newHead, ...mids, newTail);
  const lastLen = lastIns.runs.reduce((n, r) => n + r.text.length, 0);
  return { body: { ...b, paragraphs: list }, caret: { p: i + paras.length - 1, o: lastLen } };
}

/** Bul: tüm eşleşmeler. */
export function findInBody(body: TextBody, query: string, matchCase: boolean, wholeWord: boolean): [Pos, Pos][] {
  const out: [Pos, Pos][] = [];
  if (!query) return out;
  const q = matchCase ? query : query.toLocaleLowerCase("tr-TR");
  body.paragraphs.forEach((p, i) => {
    const t = matchCase ? paraText(p) : paraText(p).toLocaleLowerCase("tr-TR");
    let idx = t.indexOf(q);
    while (idx >= 0) {
      const ok = !wholeWord || ((idx === 0 || !WORD.test(t[idx - 1])) && (idx + q.length >= t.length || !WORD.test(t[idx + q.length])));
      if (ok)
        out.push([
          { p: i, o: idx },
          { p: i, o: idx + q.length },
        ]);
      idx = t.indexOf(q, idx + Math.max(1, q.length));
    }
  });
  return out;
}
