/**
 * AST değerlendiricisi: işleçler (dizi yayınımıyla), başvuru çözümleme,
 * adlar, fonksiyon çağrıları ve skaler fonksiyonların diziye uygulanması.
 */
import type { BinaryOp, Node } from "./ast";
import { compareScalars, toNumber, toText } from "./coerce";
import type { EvalHost, FnContext } from "./context";
import type { LocaleInfo } from "./coerce";
import { getFunction, type FnDef } from "./functions/registry";
import { matrix, scalar } from "./functions/helpers";
import { MAX_COLS, MAX_ROWS } from "./references";
import { Errors, FormulaError, isArray, isError, isLambda, isMultiRef, isRef, makeArray, type RefValue, type Scalar, type Value } from "./values";

export interface EvalOptions {
  host: EvalHost;
  sheet: number;
  row: number;
  col: number;
  locale: LocaleInfo;
  onDynamicRef?: (ref: RefValue) => void;
}

export function createContext(opts: EvalOptions): FnContext {
  const ctx: FnContext = {
    host: opts.host,
    sheet: opts.sheet,
    row: opts.row,
    col: opts.col,
    locale: opts.locale,
    onDynamicRef: opts.onDynamicRef,
    evalNode: (node, scope) => evaluateNode(node, scope ? { ...ctx, scope } : ctx),
  };
  return ctx;
}

/** Formülü değerlendirir; sonuç başvuru/dizi olabilir. */
export function evaluateNode(node: Node, ctx: FnContext): Value {
  try {
    return evalInner(node, ctx);
  } catch (e) {
    if (e instanceof FormulaError) return e;
    if (e instanceof RangeError) return Errors.NUM;
    throw e;
  }
}

function resolveRef(node: Extract<Node, { type: "ref" }>, ctx: FnContext): Value {
  const t = node.ref;
  if (t.refError) return Errors.REF;
  let sheet = ctx.sheet;
  if (t.sheet !== undefined) {
    const id = ctx.host.sheetIdByName(t.sheet);
    if (id === undefined) return Errors.REF;
    sheet = id;
  }
  switch (t.kind) {
    case "cell":
      return { kind: "ref", sheet, r1: t.a.row, c1: t.a.col, r2: t.a.row, c2: t.a.col };
    case "range": {
      const b = t.b!;
      return { kind: "ref", sheet, r1: Math.min(t.a.row, b.row), c1: Math.min(t.a.col, b.col), r2: Math.max(t.a.row, b.row), c2: Math.max(t.a.col, b.col) };
    }
    case "cols": {
      const b = t.b!;
      return { kind: "ref", sheet, r1: 0, r2: MAX_ROWS - 1, c1: Math.min(t.a.col, b.col), c2: Math.max(t.a.col, b.col) };
    }
    case "rows": {
      const b = t.b!;
      return { kind: "ref", sheet, c1: 0, c2: MAX_COLS - 1, r1: Math.min(t.a.row, b.row), r2: Math.max(t.a.row, b.row) };
    }
  }
}

function evalInner(node: Node, ctx: FnContext): Value {
  switch (node.type) {
    case "num":
      return node.value;
    case "str":
      return node.value;
    case "bool":
      return node.value;
    case "err":
      return errorOf(node.code);
    case "missing":
      return null;
    case "ref":
      return resolveRef(node, ctx);
    case "name": {
      const key = node.name.toLocaleUpperCase("tr-TR");
      if (ctx.scope) {
        const v = ctx.scope.get(key) ?? ctx.scope.get(node.name.toUpperCase());
        if (v !== undefined) return v;
      }
      const def = ctx.host.getDefinedName(node.name, ctx.sheet);
      if (!def) return Errors.NAME;
      return evalInner(def, ctx);
    }
    case "array":
      return makeArray(node.rows.map((r) => r.map((c) => toScalarConst(evalInner(c, ctx)))));
    case "unary": {
      const v = evalInner(node.arg, ctx);
      if (node.op === "+") return v;
      return liftUnary(v, ctx, (s) => {
        const n = toNumber(s, ctx.locale);
        return isError(n) ? n : n === 0 ? 0 : -n;
      });
    }
    case "postfix": {
      const v = evalInner(node.arg, ctx);
      return liftUnary(v, ctx, (s) => {
        const n = toNumber(s, ctx.locale);
        return isError(n) ? n : n / 100;
      });
    }
    case "binary": {
      const l = evalInner(node.left, ctx);
      if (node.op === ":") return rangeOp(l, evalInner(node.right, ctx));
      const r = evalInner(node.right, ctx);
      return binary(node.op, l, r, ctx);
    }
    case "union": {
      const areas: RefValue[] = [];
      for (const item of node.items) {
        const v = evalInner(item, ctx);
        if (isRef(v)) areas.push(v);
        else if (isMultiRef(v)) areas.push(...v.areas);
        else return Errors.VALUE;
      }
      return { kind: "refs", areas };
    }
    case "func":
      return callFunction(node.name, node.args, ctx);
  }
}

function errorOf(code: string): FormulaError {
  for (const e of Object.values(Errors)) if (e.code === code) return e;
  return new FormulaError(code as FormulaError["code"]);
}

function toScalarConst(v: Value): Scalar {
  if (v === null || typeof v !== "object" || isError(v)) return v as Scalar;
  return Errors.VALUE;
}

function rangeOp(l: Value, r: Value): Value {
  if (isError(l)) return l;
  if (isError(r)) return r;
  if (!isRef(l) || !isRef(r)) return Errors.VALUE;
  if (l.sheet !== r.sheet) return Errors.VALUE;
  return { kind: "ref", sheet: l.sheet, r1: Math.min(l.r1, r.r1), c1: Math.min(l.c1, r.c1), r2: Math.max(l.r2, r.r2), c2: Math.max(l.c2, r.c2) };
}

/** İşlenen: tek değer ya da 2B dizi. */
function operand(v: Value, ctx: FnContext): Scalar | Scalar[][] {
  if (isRef(v)) {
    if (v.r1 === v.r2 && v.c1 === v.c2) return ctx.host.getCellValue(v.sheet, v.r1, v.c1);
    return matrix(v, ctx);
  }
  if (isArray(v)) {
    if (v.data.length === 1 && v.data[0].length === 1) return v.data[0][0];
    return v.data;
  }
  if (isMultiRef(v) || isLambda(v)) return Errors.VALUE;
  return v as Scalar;
}

/** İki dizi/skaler için Excel yayınımı. */
export function broadcast(a: Scalar | Scalar[][], b: Scalar | Scalar[][], f: (x: Scalar, y: Scalar) => Scalar): Value {
  const am = Array.isArray(a);
  const bm = Array.isArray(b);
  if (!am && !bm) return f(a as Scalar, b as Scalar);
  const A = am ? (a as Scalar[][]) : [[a as Scalar]];
  const B = bm ? (b as Scalar[][]) : [[b as Scalar]];
  const ar = A.length;
  const ac = A[0]?.length ?? 0;
  const br = B.length;
  const bc = B[0]?.length ?? 0;
  if (ar === 0 || br === 0) return Errors.VALUE;
  const rows = ar === 1 ? br : br === 1 ? ar : Math.max(ar, br);
  const cols = ac === 1 ? bc : bc === 1 ? ac : Math.max(ac, bc);
  const out: Scalar[][] = [];
  for (let i = 0; i < rows; i++) {
    const row: Scalar[] = new Array(cols);
    for (let j = 0; j < cols; j++) {
      const x = ar === 1 ? A[0][ac === 1 ? 0 : j] : A[i]?.[ac === 1 ? 0 : j];
      const y = br === 1 ? B[0][bc === 1 ? 0 : j] : B[i]?.[bc === 1 ? 0 : j];
      row[j] = x === undefined || y === undefined ? Errors.NA : f(x, y);
    }
    out.push(row);
  }
  return makeArray(out);
}

/** Çok işlenenli yayınım (EĞER gibi). */
export function broadcastMany(ops: (Scalar | Scalar[][])[], f: (xs: Scalar[]) => Scalar): Value {
  const mats = ops.map((o) => (Array.isArray(o) ? o : [[o]]));
  if (!ops.some((o) => Array.isArray(o))) return f(ops as Scalar[]);
  let rows = 1;
  let cols = 1;
  for (const m of mats) {
    const r = m.length;
    const c = m[0]?.length ?? 0;
    rows = rows === 1 ? r : r === 1 ? rows : Math.max(rows, r);
    cols = cols === 1 ? c : c === 1 ? cols : Math.max(cols, c);
  }
  const out: Scalar[][] = [];
  for (let i = 0; i < rows; i++) {
    const row: Scalar[] = new Array(cols);
    for (let j = 0; j < cols; j++) {
      const xs: Scalar[] = [];
      let bad = false;
      for (const m of mats) {
        const x = m[m.length === 1 ? 0 : i]?.[(m[0]?.length ?? 0) === 1 ? 0 : j];
        if (x === undefined) bad = true;
        xs.push(x ?? null);
      }
      row[j] = bad ? Errors.NA : f(xs);
    }
    out.push(row);
  }
  return makeArray(out);
}

function liftUnary(v: Value, ctx: FnContext, f: (s: Scalar) => Scalar): Value {
  const o = operand(v, ctx);
  if (Array.isArray(o)) return makeArray(o.map((r) => r.map(f)));
  return f(o);
}

export function scalarBinary(op: BinaryOp, a: Scalar, b: Scalar, locale: LocaleInfo): Scalar {
  if (op === "&") {
    if (isError(a)) return a;
    if (isError(b)) return b;
    const x = toText(a, locale);
    const y = toText(b, locale);
    if (isError(x)) return x;
    if (isError(y)) return y;
    return x + y;
  }
  if (op === "=" || op === "<>" || op === "<" || op === ">" || op === "<=" || op === ">=") {
    if (isError(a)) return a;
    if (isError(b)) return b;
    const c = compareScalars(a, b, locale);
    switch (op) {
      case "=":
        return c === 0;
      case "<>":
        return c !== 0;
      case "<":
        return c < 0;
      case ">":
        return c > 0;
      case "<=":
        return c <= 0;
      default:
        return c >= 0;
    }
  }
  if (isError(a)) return a;
  if (isError(b)) return b;
  const x = toNumber(a, locale);
  if (isError(x)) return x;
  const y = toNumber(b, locale);
  if (isError(y)) return y;
  let r: number;
  switch (op) {
    case "+":
      r = x + y;
      break;
    case "-":
      r = x - y;
      break;
    case "*":
      r = x * y;
      break;
    case "/":
      if (y === 0) return Errors.DIV0;
      r = x / y;
      break;
    case "^":
      if (x === 0 && y === 0) return Errors.NUM;
      if (x === 0 && y < 0) return Errors.DIV0;
      r = Math.pow(x, y);
      if (Number.isNaN(r) && x < 0) {
        // Tek paydalı kesirli üsler (ör. (-8)^(1/3))
        const inv = 1 / y;
        if (Math.abs(inv - Math.round(inv)) < 1e-10 && Math.round(inv) % 2 !== 0) r = -Math.pow(-x, y);
      }
      break;
    default:
      return Errors.VALUE;
  }
  if (!Number.isFinite(r)) return Errors.NUM;
  return r;
}

function binary(op: BinaryOp, l: Value, r: Value, ctx: FnContext): Value {
  const a = operand(l, ctx);
  const b = operand(r, ctx);
  return broadcast(a, b, (x, y) => scalarBinary(op, x, y, ctx.locale));
}

/* ------------------------------------------------------------------ */
/* Fonksiyon çağrısı                                                    */
/* ------------------------------------------------------------------ */

function callFunction(name: string, argNodes: Node[], ctx: FnContext): Value {
  const def = getFunction(name);
  if (!def) {
    // LAMBDA değişkenine bağlı ad çağrısı (LET içinde)
    const key = name.toLocaleUpperCase("tr-TR");
    const lam = ctx.scope?.get(key) ?? ctx.scope?.get(name.toUpperCase());
    if (lam && isLambda(lam)) {
      const args = argNodes.map((n) => evalInner(n, ctx));
      return invokeLambda(lam, args, ctx);
    }
    return Errors.NAME;
  }
  if (argNodes.length < def.min || argNodes.length > def.max) return Errors.VALUE;
  if (def.lazy) {
    const thunks = argNodes.map((n) => () => evalInner(n, ctx));
    return def.lazy(thunks, ctx, argNodes);
  }
  const args: Value[] = new Array(argNodes.length);
  for (let i = 0; i < argNodes.length; i++) args[i] = evalInner(argNodes[i], ctx);
  return applyFunction(def, args, ctx);
}

export function invokeLambda(lam: { params: string[]; body: unknown; scope: Map<string, Value> | undefined }, args: Value[], ctx: FnContext): Value {
  if (args.length !== lam.params.length) return Errors.VALUE;
  const scope = new Map(lam.scope ?? ctx.scope ?? []);
  lam.params.forEach((p, i) => scope.set(p, args[i]));
  return evalInner(lam.body as Node, { ...ctx, scope });
}

export function applyFunction(def: FnDef, args: Value[], ctx: FnContext): Value {
  const impl = def.impl!;
  if (def.lift) {
    const positions = def.lift === true ? args.map((_, i) => i) : def.lift;
    let lifted: { idx: number; data: Scalar[][] }[] | null = null;
    const base = args.slice();
    for (const i of positions) {
      if (i >= base.length) continue;
      const v = base[i];
      if (isRef(v)) {
        if (v.r1 === v.r2 && v.c1 === v.c2) base[i] = ctx.host.getCellValue(v.sheet, v.r1, v.c1);
        else (lifted ??= []).push({ idx: i, data: matrix(v, ctx) });
      } else if (isArray(v)) {
        if (v.data.length === 1 && v.data[0].length === 1) base[i] = v.data[0][0];
        else (lifted ??= []).push({ idx: i, data: v.data });
      }
    }
    if (lifted) {
      let rows = 1;
      let cols = 1;
      for (const l of lifted) {
        const lr = l.data.length;
        const lc = l.data[0]?.length ?? 0;
        rows = rows === 1 ? lr : lr === 1 ? rows : Math.max(rows, lr);
        cols = cols === 1 ? lc : lc === 1 ? cols : Math.max(cols, lc);
      }
      const out: Scalar[][] = [];
      for (let i = 0; i < rows; i++) {
        const row: Scalar[] = [];
        for (let j = 0; j < cols; j++) {
          const a = base.slice();
          let bad = false;
          for (const l of lifted) {
            const lr = l.data.length;
            const lc = l.data[0]?.length ?? 0;
            const x = l.data[lr === 1 ? 0 : i]?.[lc === 1 ? 0 : j];
            if (x === undefined) bad = true;
            a[l.idx] = x ?? null;
          }
          row.push(bad ? Errors.NA : safeScalar(impl, a, ctx));
        }
        out.push(row);
      }
      return makeArray(out);
    }
    return safeCall(impl, base, ctx);
  }
  return safeCall(impl, args, ctx);
}

function safeCall(impl: (a: Value[], c: FnContext) => Value, args: Value[], ctx: FnContext): Value {
  try {
    return impl(args, ctx);
  } catch (e) {
    if (e instanceof FormulaError) return e;
    if (e instanceof RangeError) return Errors.NUM;
    throw e;
  }
}

function safeScalar(impl: (a: Value[], c: FnContext) => Value, args: Value[], ctx: FnContext): Scalar {
  const v = safeCall(impl, args, ctx);
  return scalar(v, ctx);
}
