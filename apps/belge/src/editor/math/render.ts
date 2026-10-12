// Denklem ağacının çıktıları: MathML (ekranda/HTML'de gösterim) ve OMML (DOCX, ECMA-376 §22.1);
// OMML → ağaç (DOCX okuma).
import { FUNCS, type MNode } from "./ast";

const escX = (s: string) => s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");

const LARGE_OPS = new Set(["∑", "∏", "∐", "⋃", "⋂"]);

/** Ağaç → MathML (math kökü olmadan). */
export function toMathML(n: MNode): string {
  switch (n.t) {
    case "row":
      return `<mrow>${n.c.map(toMathML).join("")}</mrow>`;
    case "id":
      return `<mi>${escX(n.v)}</mi>`;
    case "num":
      return `<mn>${escX(n.v)}</mn>`;
    case "op":
      return n.v === " " ? `<mspace width="0.25em"/>` : `<mo>${escX(n.v)}</mo>`;
    case "text":
      return `<mtext>${escX(n.v)}</mtext>`;
    case "frac":
      return `<mfrac>${toMathML(n.n)}${toMathML(n.d)}</mfrac>`;
    case "sup":
      return `<msup>${toMathML(n.b)}${toMathML(n.s)}</msup>`;
    case "sub":
      return `<msub>${toMathML(n.b)}${toMathML(n.s)}</msub>`;
    case "subsup":
      return `<msubsup>${toMathML(n.b)}${toMathML(n.sub)}${toMathML(n.sup)}</msubsup>`;
    case "sqrt":
      return n.deg ? `<mroot>${toMathML(n.e)}${toMathML(n.deg)}</mroot>` : `<msqrt>${toMathML(n.e)}</msqrt>`;
    case "nary": {
      const op = `<mo largeop="true">${escX(n.op)}</mo>`;
      const under = LARGE_OPS.has(n.op);
      const tag = n.sub && n.sup ? (under ? "munderover" : "msubsup") : n.sub ? (under ? "munder" : "msub") : n.sup ? (under ? "mover" : "msup") : "";
      const lim = tag ? `<${tag}>${op}${n.sub ? toMathML(n.sub) : ""}${n.sup ? toMathML(n.sup) : ""}</${tag}>` : op;
      return `<mrow>${lim}${toMathML(n.e)}</mrow>`;
    }
    case "fence":
      return `<mrow>${n.open ? `<mo fence="true">${escX(n.open)}</mo>` : ""}${toMathML(n.e)}${n.close ? `<mo fence="true">${escX(n.close)}</mo>` : ""}</mrow>`;
    case "func": {
      const name = `<mi mathvariant="normal">${escX(n.name)}</mi>`;
      return `<mrow>${n.lim ? `<munder>${name}${toMathML(n.lim)}</munder>` : name}<mo>&#x2061;</mo>${toMathML(n.e)}</mrow>`;
    }
    case "accent":
      return `<mover accent="true">${toMathML(n.e)}<mo>${escX(n.ch === "̅" ? "‾" : n.ch === "⃗" ? "→" : n.ch === "̂" ? "^" : n.ch === "̃" ? "~" : n.ch === "̇" ? "˙" : n.ch === "̈" ? "¨" : n.ch)}</mo></mover>`;
    case "matrix":
      return `<mtable>${n.rows.map((r) => `<mtr>${r.map((c) => `<mtd>${toMathML(c)}</mtd>`).join("")}</mtr>`).join("")}</mtable>`;
  }
}

export function mathElement(n: MNode, display: boolean): string {
  return `<math xmlns="http://www.w3.org/1998/Math/MathML"${display ? ' display="block"' : ""}>${toMathML(n)}</math>`;
}

/* ------------------------------------------------------------------ */
/* OMML                                                                */
/* ------------------------------------------------------------------ */

export const MATH_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math";

// Matematik koşusu (yazı tipi belirtilmez: düzenleyici varsayılan matematik yazı tipini kullanır).
const run = (t: string, plain = false) => `<m:r>${plain ? "<m:rPr><m:sty m:val=\"p\"/></m:rPr>" : ""}<m:t xml:space="preserve">${escX(t)}</m:t></m:r>`;
const e = (n: MNode) => `<m:e>${toOmml(n)}</m:e>`;

/** Ağaç → OMML içeriği (m:oMath içine). Ardışık harf/sayı/işleçler tek koşuda birleşir. */
export function toOmml(n: MNode): string {
  switch (n.t) {
    case "row": {
      let out = "";
      let buf = "";
      const flush = () => {
        if (buf) out += run(buf);
        buf = "";
      };
      for (const c of n.c) {
        if (c.t === "id" || c.t === "num" || c.t === "op") buf += c.v;
        else {
          flush();
          out += toOmml(c);
        }
      }
      flush();
      return out;
    }
    case "id":
    case "num":
    case "op":
      return run(n.v);
    case "text":
      return `<m:r><m:rPr><m:nor/></m:rPr><m:t xml:space="preserve">${escX(n.v)}</m:t></m:r>`;
    case "frac":
      return `<m:f><m:num>${toOmml(n.n)}</m:num><m:den>${toOmml(n.d)}</m:den></m:f>`;
    case "sup":
      return `<m:sSup>${e(n.b)}<m:sup>${toOmml(n.s)}</m:sup></m:sSup>`;
    case "sub":
      return `<m:sSub>${e(n.b)}<m:sub>${toOmml(n.s)}</m:sub></m:sSub>`;
    case "subsup":
      return `<m:sSubSup>${e(n.b)}<m:sub>${toOmml(n.sub)}</m:sub><m:sup>${toOmml(n.sup)}</m:sup></m:sSubSup>`;
    case "sqrt":
      return n.deg ? `<m:rad><m:deg>${toOmml(n.deg)}</m:deg>${e(n.e)}</m:rad>` : `<m:rad><m:radPr><m:degHide m:val="1"/></m:radPr><m:deg/>${e(n.e)}</m:rad>`;
    case "nary": {
      const pr = `<m:naryPr><m:chr m:val="${escX(n.op)}"/><m:limLoc m:val="${LARGE_OPS.has(n.op) ? "undOvr" : "subSup"}"/>${n.sub ? "" : '<m:subHide m:val="1"/>'}${n.sup ? "" : '<m:supHide m:val="1"/>'}</m:naryPr>`;
      return `<m:nary>${pr}<m:sub>${n.sub ? toOmml(n.sub) : ""}</m:sub><m:sup>${n.sup ? toOmml(n.sup) : ""}</m:sup>${e(n.e)}</m:nary>`;
    }
    case "fence": {
      const pr = n.open === "(" && n.close === ")" ? "" : `<m:dPr><m:begChr m:val="${escX(n.open)}"/><m:endChr m:val="${escX(n.close)}"/></m:dPr>`;
      return `<m:d>${pr}${e(n.e)}</m:d>`;
    }
    case "func": {
      const name = run(n.name, true);
      const fName = n.lim ? `<m:limLow><m:e>${name}</m:e><m:lim>${toOmml(n.lim)}</m:lim></m:limLow>` : name;
      return `<m:func><m:fName>${fName}</m:fName>${e(n.e)}</m:func>`;
    }
    case "accent":
      if (n.ch === "̅") return `<m:bar><m:barPr><m:pos m:val="top"/></m:barPr>${e(n.e)}</m:bar>`;
      return `<m:acc><m:accPr><m:chr m:val="${escX(n.ch)}"/></m:accPr>${e(n.e)}</m:acc>`;
    case "matrix":
      return `<m:m>${n.rows.map((r) => `<m:mr>${r.map((c) => e(c)).join("")}</m:mr>`).join("")}</m:m>`;
  }
}

/** Satır içi ya da (kendi paragrafında) görüntülü denklem. */
export function ommlOf(n: MNode, display: boolean): string {
  const m = `<m:oMath>${toOmml(n)}</m:oMath>`;
  return display ? `<m:oMathPara><m:oMathParaPr><m:jc m:val="centerGroup"/></m:oMathParaPr>${m}</m:oMathPara>` : m;
}

/* ------------------------------------------------------------------ */
/* OMML → ağaç                                                         */
/* ------------------------------------------------------------------ */

const kidsOf = (el: Element) => Array.from(el.children);
const kidOf = (el: Element, name: string) => kidsOf(el).find((c) => c.localName === name) ?? null;
const valOf = (el: Element | null) => el?.getAttributeNS(MATH_NS, "val") ?? el?.getAttribute("m:val") ?? el?.getAttribute("val") ?? null;
const flat = (c: MNode[]): MNode => (c.length === 1 ? c[0] : { t: "row", c });

/** Düz metni harf/sayı/işleç atomlarına ayırır. */
function atomsOf(text: string, plain: boolean): MNode[] {
  if (plain && FUNCS.has(text.trim())) return [{ t: "text", v: text }];
  const out: MNode[] = [];
  const re = /(\d+(?:[.,]\d+)*)|(\p{L})|(\s+)|(.)/gu;
  for (let m = re.exec(text); m; m = re.exec(text)) {
    if (m[1]) out.push({ t: "num", v: m[1] });
    else if (m[2]) out.push({ t: "id", v: m[2] });
    else if (m[3]) continue;
    else out.push({ t: "op", v: m[4] === "-" ? "−" : m[4] });
  }
  return out;
}

function arg(el: Element | null): MNode {
  return el ? flat(readChildren(el)) : { t: "row", c: [] };
}

function readChildren(el: Element): MNode[] {
  const out: MNode[] = [];
  for (const c of kidsOf(el)) {
    const n = readNode(c);
    if (Array.isArray(n)) out.push(...n);
    else if (n) out.push(n);
  }
  return out;
}

function readNode(c: Element): MNode | MNode[] | null {
  switch (c.localName) {
    case "r": {
      const t = Array.from(c.getElementsByTagNameNS(MATH_NS, "t")).map((x) => x.textContent ?? "").join("");
      const rPr = kidOf(c, "rPr");
      if (rPr && kidOf(rPr, "nor")) return { t: "text", v: t };
      return atomsOf(t, valOf(rPr ? kidOf(rPr, "sty") : null) === "p");
    }
    case "f":
      return { t: "frac", n: arg(kidOf(c, "num")), d: arg(kidOf(c, "den")) };
    case "sSup":
      return { t: "sup", b: arg(kidOf(c, "e")), s: arg(kidOf(c, "sup")) };
    case "sSub":
      return { t: "sub", b: arg(kidOf(c, "e")), s: arg(kidOf(c, "sub")) };
    case "sSubSup":
      return { t: "subsup", b: arg(kidOf(c, "e")), sub: arg(kidOf(c, "sub")), sup: arg(kidOf(c, "sup")) };
    case "rad": {
      const pr = kidOf(c, "radPr");
      const hide = pr && valOf(kidOf(pr, "degHide"));
      const deg = kidOf(c, "deg");
      const hasDeg = !(hide === "1" || hide === "on" || hide === "true") && deg && deg.children.length > 0;
      return { t: "sqrt", e: arg(kidOf(c, "e")), ...(hasDeg ? { deg: arg(deg) } : {}) };
    }
    case "nary": {
      const pr = kidOf(c, "naryPr");
      const op = valOf(pr ? kidOf(pr, "chr") : null) ?? "∫";
      const hidden = (n: string) => ["1", "on", "true"].includes(valOf(pr ? kidOf(pr, n) : null) ?? "");
      const sub = kidOf(c, "sub");
      const sup = kidOf(c, "sup");
      return { t: "nary", op, ...(sub && !hidden("subHide") && sub.children.length ? { sub: arg(sub) } : {}), ...(sup && !hidden("supHide") && sup.children.length ? { sup: arg(sup) } : {}), e: arg(kidOf(c, "e")) };
    }
    case "d": {
      const pr = kidOf(c, "dPr");
      const beg = pr ? kidOf(pr, "begChr") : null;
      const end = pr ? kidOf(pr, "endChr") : null;
      const items = kidsOf(c).filter((k) => k.localName === "e").map((k) => arg(k));
      const body: MNode[] = [];
      items.forEach((it, i) => {
        if (i) body.push({ t: "op", v: valOf(pr ? kidOf(pr, "sepChr") : null) ?? "|" });
        body.push(it);
      });
      return { t: "fence", open: beg ? (valOf(beg) ?? "") : "(", close: end ? (valOf(end) ?? "") : ")", e: flat(body) };
    }
    case "func": {
      const fn = kidOf(c, "fName");
      let name = "";
      let lim: MNode | undefined;
      const low = fn ? kidOf(fn, "limLow") : null;
      if (low) {
        name = (kidOf(low, "e")?.textContent ?? "").trim();
        lim = arg(kidOf(low, "lim"));
      } else name = (fn?.textContent ?? "").trim();
      return { t: "func", name: name || "f", e: arg(kidOf(c, "e")), ...(lim ? { lim } : {}) };
    }
    case "limLow": {
      const base = arg(kidOf(c, "e"));
      const name = base.t === "text" ? base.v.trim() : "";
      if (FUNCS.has(name)) return { t: "func", name, e: { t: "row", c: [] }, lim: arg(kidOf(c, "lim")) };
      return { t: "sub", b: base, s: arg(kidOf(c, "lim")) };
    }
    case "limUpp":
      return { t: "sup", b: arg(kidOf(c, "e")), s: arg(kidOf(c, "lim")) };
    case "acc": {
      const pr = kidOf(c, "accPr");
      return { t: "accent", ch: valOf(pr ? kidOf(pr, "chr") : null) ?? "̂", e: arg(kidOf(c, "e")) };
    }
    case "bar":
      return { t: "accent", ch: "̅", e: arg(kidOf(c, "e")) };
    case "m":
      return { t: "matrix", rows: kidsOf(c).filter((r) => r.localName === "mr").map((r) => kidsOf(r).filter((k) => k.localName === "e").map((k) => arg(k))) };
    case "eqArr":
      return { t: "matrix", rows: kidsOf(c).filter((k) => k.localName === "e").map((k) => [arg(k)]) };
    case "rPr":
    case "ctrlPr":
    case "argPr":
    case "oMathParaPr":
    case "fPr":
    case "sSupPr":
    case "sSubPr":
    case "sSubSupPr":
    case "radPr":
    case "naryPr":
    case "dPr":
    case "funcPr":
    case "accPr":
    case "barPr":
    case "mPr":
    case "limLowPr":
    case "limUppPr":
      return null;
    default:
      // box, borderBox, groupChr, sPre, phant… : içerik açılır.
      return readChildren(c);
  }
}

/** m:oMath öğesi → ağaç. */
export function fromOmml(oMath: Element): MNode {
  return flat(readChildren(oMath));
}
