// Denklem söz dizimi ağacı ve doğrusal (LaTeX benzeri) girişin ayrıştırıcısı/yazıcısı.
// Desteklenen: \frac{a}{b} ve a/b, ^ _, \sqrt[n]{x}, \sum \prod \int \oint (alt/üst sınır, gövde),
// ( ) [ ] \{ \} ve \left \right ayraçları, \sin \cos \log \lim… işlevleri, Yunan harfleri ve işleçler,
// \text{…}, \hat \bar \vec \dot \tilde vurguları, \begin{matrix|pmatrix|bmatrix} … \end{…}.

export type MNode =
  | { t: "row"; c: MNode[] }
  | { t: "id"; v: string }
  | { t: "num"; v: string }
  | { t: "op"; v: string }
  | { t: "text"; v: string }
  | { t: "frac"; n: MNode; d: MNode }
  | { t: "sup"; b: MNode; s: MNode }
  | { t: "sub"; b: MNode; s: MNode }
  | { t: "subsup"; b: MNode; sub: MNode; sup: MNode }
  | { t: "sqrt"; e: MNode; deg?: MNode }
  | { t: "nary"; op: string; sub?: MNode; sup?: MNode; e: MNode }
  | { t: "fence"; open: string; close: string; e: MNode }
  | { t: "func"; name: string; e: MNode; lim?: MNode }
  | { t: "accent"; ch: string; e: MNode }
  | { t: "matrix"; rows: MNode[][] };

export const GREEK: Record<string, string> = {
  alpha: "α", beta: "β", gamma: "γ", delta: "δ", epsilon: "ε", varepsilon: "ε", zeta: "ζ", eta: "η", theta: "θ", vartheta: "ϑ", iota: "ι", kappa: "κ", lambda: "λ", mu: "μ", nu: "ν", xi: "ξ",
  pi: "π", varpi: "ϖ", rho: "ρ", sigma: "σ", tau: "τ", upsilon: "υ", phi: "φ", varphi: "ϕ", chi: "χ", psi: "ψ", omega: "ω",
  Gamma: "Γ", Delta: "Δ", Theta: "Θ", Lambda: "Λ", Xi: "Ξ", Pi: "Π", Sigma: "Σ", Upsilon: "Υ", Phi: "Φ", Psi: "Ψ", Omega: "Ω",
};

export const SYMBOLS: Record<string, string> = {
  pm: "±", mp: "∓", times: "×", div: "÷", cdot: "⋅", ast: "∗", leq: "≤", le: "≤", geq: "≥", ge: "≥", neq: "≠", ne: "≠", approx: "≈", equiv: "≡", sim: "∼", propto: "∝",
  infty: "∞", partial: "∂", nabla: "∇", to: "→", rightarrow: "→", leftarrow: "←", Rightarrow: "⇒", Leftarrow: "⇐", Leftrightarrow: "⇔", leftrightarrow: "↔",
  in: "∈", notin: "∉", subset: "⊂", supset: "⊃", subseteq: "⊆", supseteq: "⊇", cup: "∪", cap: "∩", emptyset: "∅", forall: "∀", exists: "∃", neg: "¬", land: "∧", lor: "∨",
  cdots: "⋯", ldots: "…", dots: "…", degree: "°", circ: "∘", angle: "∠", perp: "⊥", parallel: "∥", prime: "′", hbar: "ℏ", ell: "ℓ", Re: "ℜ", Im: "ℑ",
};

export const NARY: Record<string, string> = { sum: "∑", prod: "∏", coprod: "∐", int: "∫", iint: "∬", iiint: "∭", oint: "∮", bigcup: "⋃", bigcap: "⋂" };
const NARY_CHARS = new Set(Object.values(NARY));
export const FUNCS = new Set(["sin", "cos", "tan", "cot", "sec", "csc", "arcsin", "arccos", "arctan", "sinh", "cosh", "tanh", "log", "ln", "lg", "exp", "lim", "max", "min", "sup", "inf", "det", "gcd", "deg", "dim", "ker", "arg"]);
export const ACCENTS: Record<string, string> = { hat: "̂", bar: "̅", overline: "̅", vec: "⃗", dot: "̇", ddot: "̈", tilde: "̃" };
const OPEN: Record<string, string> = { "(": ")", "[": "]", "{": "}", "|": "|", "⟨": "⟩" };

const row = (c: MNode[]): MNode => (c.length === 1 ? c[0] : { t: "row", c });
/** Dıştaki tek ayraç kaldırılır (kesir pay/paydası: (a+b)/c → a+b bölü c). */
const unwrap = (n: MNode): MNode => (n.t === "fence" && n.open === "(" && n.close === ")" ? n.e : n);

const BREAK_OPS = new Set(["=", "+", "-", "−", "<", ">", "≤", "≥", "≠", "≈", ",", "±", "→", "⇒"]);

class Parser {
  i = 0;
  constructor(private s: string) {}
  peek() {
    return this.s[this.i];
  }
  ws() {
    while (this.i < this.s.length && /\s/.test(this.s[this.i])) this.i++;
  }
  /** İfade dizisi: kapanış karakterine ya da sona kadar. */
  seq(stop: string[] = []): MNode[] {
    const out: MNode[] = [];
    for (;;) {
      this.ws();
      if (this.i >= this.s.length) break;
      if (stop.some((x) => this.s.startsWith(x, this.i))) break;
      const a = this.scripted();
      if (!a) break;
      // a/b doğrusal kesir
      this.ws();
      if (this.peek() === "/" && !stop.includes("/")) {
        this.i++;
        const d = this.scripted();
        out.push({ t: "frac", n: unwrap(a), d: d ? unwrap(d) : { t: "row", c: [] } });
        continue;
      }
      out.push(a);
    }
    return out;
  }
  /** Atom + isteğe bağlı ^/_ betikleri; büyük işleçte gövde de alınır. */
  scripted(): MNode | null {
    const base = this.atom();
    if (!base) return null;
    let sub: MNode | undefined;
    let sup: MNode | undefined;
    for (let k = 0; k < 2; k++) {
      this.ws();
      const c = this.peek();
      if (c === "_" && !sub) {
        this.i++;
        sub = this.arg();
      } else if (c === "^" && !sup) {
        this.i++;
        sup = this.arg();
      } else break;
    }
    if (base.t === "op" && NARY_CHARS.has(base.v)) {
      // Gövde: bir sonraki toplama/ilişki işlecine kadar.
      const body: MNode[] = [];
      for (;;) {
        this.ws();
        const save = this.i;
        if (this.i >= this.s.length || this.peek() === "}" || this.s.startsWith("\\right", this.i) || this.peek() === "&" || this.s.startsWith("\\\\", this.i)) break;
        const n = this.scripted();
        if (!n) break;
        if (n.t === "op" && BREAK_OPS.has(n.v)) {
          this.i = save;
          break;
        }
        body.push(n);
      }
      return { t: "nary", op: base.v, ...(sub ? { sub } : {}), ...(sup ? { sup } : {}), e: row(body) };
    }
    if (sub && sup) return { t: "subsup", b: base, sub, sup };
    if (sub) return { t: "sub", b: base, s: sub };
    if (sup) return { t: "sup", b: base, s: sup };
    return base;
  }
  /** Betik ya da komut bağımsız değişkeni: {…} ya da tek atom (tek karakter). */
  arg(): MNode {
    this.ws();
    if (this.peek() === "{") {
      this.i++;
      const c = this.seq(["}"]);
      if (this.peek() === "}") this.i++;
      return row(c);
    }
    const ch = this.s[this.i];
    if (ch && /[0-9]/.test(ch)) {
      this.i++;
      return { t: "num", v: ch };
    }
    return this.atom() ?? { t: "row", c: [] };
  }
  group(): MNode {
    this.ws();
    if (this.peek() !== "{") return this.arg();
    return this.arg();
  }
  atom(): MNode | null {
    this.ws();
    const s = this.s;
    const ch = s[this.i];
    if (ch === undefined) return null;
    if (ch === "{") {
      this.i++;
      const c = this.seq(["}"]);
      if (this.peek() === "}") this.i++;
      return row(c);
    }
    if (ch === "}") return null;
    if (/[0-9]/.test(ch)) {
      let j = this.i;
      while (j < s.length && /[0-9.,]/.test(s[j]) && !(/[.,]/.test(s[j]) && !/[0-9]/.test(s[j + 1] ?? ""))) j++;
      const v = s.slice(this.i, j);
      this.i = j;
      return { t: "num", v };
    }
    if (ch === "\\") return this.command();
    if (ch === "√") {
      this.i++;
      return { t: "sqrt", e: this.arg() };
    }
    if (OPEN[ch] && ch !== "|" && ch !== "{") {
      this.i++;
      const close = OPEN[ch];
      const c = this.seq([close]);
      if (this.peek() === close) this.i++;
      return { t: "fence", open: ch, close, e: row(c) };
    }
    if (/\p{L}/u.test(ch)) {
      this.i++;
      return { t: "id", v: ch };
    }
    this.i++;
    if (ch === "-") return { t: "op", v: "−" };
    if (ch === "*") return { t: "op", v: "∗" };
    if (ch === "'") return { t: "op", v: "′" };
    return { t: "op", v: ch };
  }
  name(): string {
    const m = /^[A-Za-z]+/.exec(this.s.slice(this.i));
    if (m) {
      this.i += m[0].length;
      return m[0];
    }
    const c = this.s[this.i] ?? "";
    this.i++;
    return c;
  }
  command(): MNode {
    this.i++; // \
    const n = this.name();
    if (n === "frac" || n === "dfrac" || n === "tfrac") {
      const a = this.group();
      const b = this.group();
      return { t: "frac", n: a, d: b };
    }
    if (n === "sqrt") {
      this.ws();
      let deg: MNode | undefined;
      if (this.peek() === "[") {
        this.i++;
        deg = row(this.seq(["]"]));
        if (this.peek() === "]") this.i++;
      }
      return { t: "sqrt", e: this.group(), ...(deg ? { deg } : {}) };
    }
    if (n === "text" || n === "mathrm" || n === "textrm") {
      this.ws();
      if (this.peek() === "{") {
        const end = this.s.indexOf("}", this.i);
        const v = this.s.slice(this.i + 1, end < 0 ? undefined : end);
        this.i = end < 0 ? this.s.length : end + 1;
        return { t: "text", v };
      }
      return { t: "text", v: "" };
    }
    if (n === "left") {
      this.ws();
      let open = this.s[this.i] === "\\" ? (this.i++, this.name()) : (this.s[this.i++] ?? "(");
      if (open === "{" ) open = "{";
      if (open === "lbrace") open = "{";
      if (open === "langle") open = "⟨";
      const c = this.seq(["\\right"]);
      let close = ")";
      if (this.s.startsWith("\\right", this.i)) {
        this.i += 6;
        this.ws();
        close = this.s[this.i] === "\\" ? (this.i++, this.name()) : (this.s[this.i++] ?? ")");
        if (close === "rbrace") close = "}";
        if (close === "rangle") close = "⟩";
      }
      return { t: "fence", open: open === "." ? "" : open, close: close === "." ? "" : close, e: row(c) };
    }
    if (n === "begin") {
      const env = this.group();
      const kind = env.t === "row" ? env.c.map((x) => ("v" in x ? x.v : "")).join("") : "v" in env ? env.v : "matrix";
      const rows: MNode[][] = [[]];
      for (;;) {
        const c = this.seq(["&", "\\\\", "\\end"]);
        rows[rows.length - 1].push(row(c));
        if (this.peek() === "&") this.i++;
        else if (this.s.startsWith("\\\\", this.i)) {
          this.i += 2;
          rows.push([]);
        } else break;
      }
      if (this.s.startsWith("\\end", this.i)) {
        this.i += 4;
        this.group();
      }
      const m: MNode = { t: "matrix", rows: rows.filter((r) => r.length) };
      if (kind === "pmatrix") return { t: "fence", open: "(", close: ")", e: m };
      if (kind === "bmatrix") return { t: "fence", open: "[", close: "]", e: m };
      if (kind === "vmatrix") return { t: "fence", open: "|", close: "|", e: m };
      return m;
    }
    if (ACCENTS[n]) return { t: "accent", ch: ACCENTS[n], e: this.group() };
    if (NARY[n]) return { t: "op", v: NARY[n] };
    if (FUNCS.has(n)) {
      // İşlev adı: bağımsız değişken sonraki atom (betikleri ile).
      this.ws();
      let sub: MNode | undefined;
      if (this.peek() === "_") {
        this.i++;
        sub = this.arg();
      }
      const e = this.scripted() ?? { t: "row", c: [] };
      return { t: "func", name: n, e, ...(sub ? { lim: sub } : {}) };
    }
    if (GREEK[n]) return { t: "id", v: GREEK[n] };
    if (SYMBOLS[n]) return { t: "op", v: SYMBOLS[n] };
    if (n === "{" || n === "}") return { t: "op", v: n };
    if (n === "," || n === ";" || n === " " || n === "quad") return { t: "op", v: " " };
    return { t: "text", v: n };
  }
}

/** Doğrusal giriş → ağaç. */
export function parseMath(src: string): MNode {
  return row(new Parser(src).seq());
}

const INV_GREEK = Object.fromEntries(Object.entries(GREEK).filter(([k]) => !k.startsWith("var")).map(([k, v]) => [v, k]));
const INV_SYM = Object.fromEntries(Object.entries(SYMBOLS).reverse().map(([k, v]) => [v, k]));
const INV_NARY = Object.fromEntries(Object.entries(NARY).map(([k, v]) => [v, k]));
const INV_ACC = Object.fromEntries(Object.entries(ACCENTS).reverse().map(([k, v]) => [v, k]));

const grp = (n: MNode) => {
  const s = toLinear(n);
  return s.length === 1 ? s : `{${s}}`;
};

/** Ağaç → doğrusal giriş (DOCX'ten okunan denklemin düzenlenebilir biçimi). */
export function toLinear(n: MNode): string {
  switch (n.t) {
    case "row": {
      let out = "";
      for (const c of n.c) {
        const s = toLinear(c);
        if (out && /[A-Za-z]$/.test(out) && /^[A-Za-z]/.test(s) && out.match(/\\[A-Za-z]+$/)) out += " ";
        out += s;
      }
      return out;
    }
    case "id":
      return INV_GREEK[n.v] ? `\\${INV_GREEK[n.v]} ` : n.v;
    case "num":
      return n.v;
    case "op":
      if (n.v === "−") return "-";
      if (n.v === "∗") return "*";
      if (n.v === "′") return "'";
      if (n.v === " ") return "\\,";
      if (n.v === "{" || n.v === "}") return `\\${n.v}`;
      if (INV_NARY[n.v]) return `\\${INV_NARY[n.v]} `;
      return INV_SYM[n.v] ? `\\${INV_SYM[n.v]} ` : n.v;
    case "text":
      return `\\text{${n.v}}`;
    case "frac":
      return `\\frac{${toLinear(n.n)}}{${toLinear(n.d)}}`;
    case "sup":
      return `${toLinear(n.b)}^${grp(n.s)}`;
    case "sub":
      return `${toLinear(n.b)}_${grp(n.s)}`;
    case "subsup":
      return `${toLinear(n.b)}_${grp(n.sub)}^${grp(n.sup)}`;
    case "sqrt":
      return n.deg ? `\\sqrt[${toLinear(n.deg)}]{${toLinear(n.e)}}` : `\\sqrt{${toLinear(n.e)}}`;
    case "nary": {
      const name = INV_NARY[n.op];
      return `${name ? `\\${name}` : n.op}${n.sub ? `_${grp(n.sub)}` : ""}${n.sup ? `^${grp(n.sup)}` : ""} ${toLinear(n.e)}`;
    }
    case "fence": {
      if (n.e.t === "matrix") {
        const env = n.open === "(" ? "pmatrix" : n.open === "[" ? "bmatrix" : n.open === "|" ? "vmatrix" : "";
        if (env) return `\\begin{${env}}${n.e.rows.map((r) => r.map(toLinear).join("&")).join("\\\\")}\\end{${env}}`;
      }
      if (n.open === "(" && n.close === ")") return `(${toLinear(n.e)})`;
      if (n.open === "[" && n.close === "]") return `[${toLinear(n.e)}]`;
      const o = n.open === "{" ? "\\{" : n.open === "⟨" ? "\\langle" : n.open || ".";
      const c = n.close === "}" ? "\\}" : n.close === "⟩" ? "\\rangle" : n.close || ".";
      return `\\left${o}${toLinear(n.e)}\\right${c}`;
    }
    case "func":
      return `\\${n.name}${n.lim ? `_${grp(n.lim)}` : ""} ${toLinear(n.e)}`;
    case "accent":
      return `\\${INV_ACC[n.ch] ?? "hat"}{${toLinear(n.e)}}`;
    case "matrix":
      return `\\begin{matrix}${n.rows.map((r) => r.map(toLinear).join("&")).join("\\\\")}\\end{matrix}`;
  }
}
