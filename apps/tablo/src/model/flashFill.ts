/**
 * Akıllı Doldurma (Ctrl+E): yazılmış örneklerden aynı satırdaki kaynak
 * sütunlara dayanan bir desen çıkarır ve sütunun geri kalanını doldurur.
 * Desen; ayırıcıyla bölünmüş parçalar, sütunun tamamı, baş harf, ilk/son
 * karakterler ve büyük/küçük harf dönüşümleriyle sabit metinlerin
 * birleşimidir (ör. "Ahmet Yılmaz" → "A. YILMAZ", "a@b.com" → "a").
 */

export type Transform = "none" | "upper" | "lower" | "title" | "initial";

export type Part =
  | { kind: "lit"; text: string }
  | { kind: "field"; col: number; delim: string | null; index: number; transform: Transform; take?: number; fromEnd?: boolean };

export type Program = Part[];

const DELIMS = [" ", ",", ";", "-", "_", "/", ".", "@", "|", ":"];

function title(s: string): string {
  return s
    .toLocaleLowerCase("tr-TR")
    .replace(/(^|[\s\-'’])(\p{L})/gu, (_m, a: string, b: string) => a + b.toLocaleUpperCase("tr-TR"));
}

function transform(s: string, t: Transform): string {
  switch (t) {
    case "upper":
      return s.toLocaleUpperCase("tr-TR");
    case "lower":
      return s.toLocaleLowerCase("tr-TR");
    case "title":
      return title(s);
    case "initial":
      return s.slice(0, 1).toLocaleUpperCase("tr-TR");
    default:
      return s;
  }
}

function tokens(v: string, delim: string): string[] {
  return v.split(delim).filter((x) => x !== "");
}

/** Bir alanın satırdaki değeri (yoksa null). */
export function fieldValue(p: Extract<Part, { kind: "field" }>, row: string[]): string | null {
  const v = row[p.col];
  if (v === undefined || v === "") return null;
  let base: string;
  if (p.delim === null) base = v;
  else {
    const t = tokens(v, p.delim);
    const i = p.index < 0 ? t.length + p.index : p.index;
    if (i < 0 || i >= t.length) return null;
    base = t[i];
  }
  if (p.take) base = p.fromEnd ? base.slice(-p.take) : base.slice(0, p.take);
  return transform(base, p.transform);
}

export function runProgram(prog: Program, row: string[]): string | null {
  let out = "";
  for (const p of prog) {
    if (p.kind === "lit") out += p.text;
    else {
      const v = fieldValue(p, row);
      if (v === null) return null;
      out += v;
    }
  }
  return out;
}

/** Satırdaki aday alanlar (değerleriyle). */
function candidateFields(row: string[]): { part: Extract<Part, { kind: "field" }>; value: string }[] {
  const out: { part: Extract<Part, { kind: "field" }>; value: string }[] = [];
  const transforms: Transform[] = ["none", "upper", "lower", "title", "initial"];
  row.forEach((v, col) => {
    if (!v) return;
    const bases: { delim: string | null; index: number }[] = [{ delim: null, index: 0 }];
    for (const d of DELIMS) {
      const t = tokens(v, d);
      if (t.length < 2) continue;
      for (let i = 0; i < t.length; i++) bases.push({ delim: d, index: i });
      bases.push({ delim: d, index: -1 });
    }
    for (const b of bases)
      for (const tr of transforms) {
        const part = { kind: "field" as const, col, delim: b.delim, index: b.index, transform: tr };
        const value = fieldValue(part, row);
        if (value) out.push({ part, value });
        // İlk/son n karakter (ör. telefon numarasının son 4 hanesi)
        if (tr === "none" && b.delim === null) {
          for (const n of [1, 2, 3, 4, 5]) {
            if (v.length <= n) continue;
            out.push({ part: { ...part, take: n }, value: v.slice(0, n) });
            out.push({ part: { ...part, take: n, fromEnd: true }, value: v.slice(-n) });
          }
        }
      }
  });
  return out;
}

/** Bir örneği üreten olası programlar (kısa ve alan ağırlıklı olanlar önce). */
function programsFor(example: string, row: string[], limit = 400): Program[] {
  const fields = candidateFields(row);
  const results: Program[] = [];
  const memo = new Map<number, Program[]>();
  const rec = (pos: number): Program[] => {
    if (pos === example.length) return [[]];
    const hit = memo.get(pos);
    if (hit) return hit;
    const outs: Program[] = [];
    const rest = example.slice(pos);
    // Alanlar (uzun eşleşmeler önce)
    const matches = fields.filter((f) => rest.startsWith(f.value)).sort((a, b) => b.value.length - a.value.length);
    for (const m of matches) {
      for (const tail of rec(pos + m.value.length)) {
        outs.push([m.part, ...tail]);
        if (outs.length > limit) break;
      }
      if (outs.length > limit) break;
    }
    // Sabit karakter
    if (outs.length <= limit) {
      for (const tail of rec(pos + 1)) {
        const head = tail[0];
        if (head && head.kind === "lit") outs.push([{ kind: "lit", text: example[pos] + head.text }, ...tail.slice(1)]);
        else outs.push([{ kind: "lit", text: example[pos] }, ...tail]);
        if (outs.length > limit) break;
      }
    }
    memo.set(pos, outs);
    return outs;
  };
  results.push(...rec(0));
  // Basitlik puanı: az sabit metin, az parça; ilk/son n karakter ve dönüşümler daha az tercih edilir
  const score = (p: Program) =>
    p.reduce((n, x) => n + (x.kind === "lit" ? x.text.length * 3 + 1 : 1 + (x.take ? 2 : 0) + (x.transform !== "none" ? 0.5 : 0) + (x.delim !== null ? 0.25 : 0)), 0);
  return results.filter((p) => p.some((x) => x.kind === "field")).sort((a, b) => score(a) - score(b));
}

/** Örneklerin tümüyle tutarlı ilk programı bulur. */
export function learnProgram(examples: { row: string[]; output: string }[]): Program | null {
  if (!examples.length) return null;
  const [first, ...rest] = examples;
  for (const prog of programsFor(first.output, first.row)) {
    if (rest.every((e) => runProgram(prog, e.row) === e.output)) return prog;
  }
  return null;
}
