// Karşılaştır (Gözden Geçir › Karşılaştır): özgün ve düzeltilmiş belge karşılaştırılır; sonuç düzeltilmiş
// belgenin yapısında, farklar izlenen değişiklik (ekleme/silme işaretleri) olarak yeni bir belgedir.
// Bloklar benzerliğe göre (en uzun ortak alt dizi) eşlenir; eşlenen paragraflarda sözcük düzeyinde fark alınır.
import type { JSONContent } from "@tiptap/core";

type Mark = NonNullable<JSONContent["marks"]>[number];

interface Token {
  /** Karşılaştırma anahtarı (sözcük, boşluk ya da satır içi düğümün JSON'u). */
  key: string;
  node: JSONContent;
}

const TEXTBLOCKS = new Set(["paragraph", "heading"]);
const CONTAINERS = new Set(["bulletList", "orderedList", "listItem", "table", "tableRow", "tableCell", "tableHeader", "blockquote", "textBox"]);

/** En uzun ortak alt dizi eşleşmeleri (i, j çiftleri). Ortak baş ve son ayrıca eşlenir (bellek için). */
export function lcsPairs<A, B>(a: A[], b: B[], eq: (x: A, y: B) => boolean): [number, number][] {
  let start = 0;
  while (start < a.length && start < b.length && eq(a[start], b[start])) start++;
  let ea = a.length;
  let eb = b.length;
  while (ea > start && eb > start && eq(a[ea - 1], b[eb - 1])) {
    ea--;
    eb--;
  }
  const head: [number, number][] = Array.from({ length: start }, (_, k) => [k, k]);
  const tail: [number, number][] = Array.from({ length: a.length - ea }, (_, k) => [ea + k, eb + k]);
  const n = ea - start;
  const m = eb - start;
  const mid: [number, number][] = [];
  if (n && m) {
    const dp: Uint32Array[] = Array.from({ length: n + 1 }, () => new Uint32Array(m + 1));
    for (let i = n - 1; i >= 0; i--) for (let j = m - 1; j >= 0; j--) dp[i][j] = eq(a[start + i], b[start + j]) ? dp[i + 1][j + 1] + 1 : Math.max(dp[i + 1][j], dp[i][j + 1]);
    let i = 0;
    let j = 0;
    while (i < n && j < m) {
      if (eq(a[start + i], b[start + j]) && dp[i][j] === dp[i + 1][j + 1] + 1) {
        mid.push([start + i, start + j]);
        i++;
        j++;
      } else if (dp[i + 1][j] >= dp[i][j + 1]) i++;
      else j++;
    }
  }
  return [...head, ...mid, ...tail];
}

const textOf = (n: JSONContent): string => (n.text ?? "") + (n.content ?? []).map(textOf).join(n.type && TEXTBLOCKS.has(n.type) ? "" : "\n");

/** İki metnin sözcük benzerliği (0–1). */
export function similarity(a: string, b: string): number {
  if (a === b) return 1;
  const wa = a.split(/\s+/).filter(Boolean);
  const wb = b.split(/\s+/).filter(Boolean);
  if (!wa.length || !wb.length) return 0;
  const common = lcsPairs(wa, wb, (x, y) => x === y).length;
  return (2 * common) / (wa.length + wb.length);
}

function tokens(block: JSONContent): Token[] {
  const out: Token[] = [];
  for (const c of block.content ?? []) {
    if (c.type === "text") {
      for (const part of (c.text ?? "").match(/\s+|[^\s]+/g) ?? []) out.push({ key: part, node: { ...c, text: part } });
    } else out.push({ key: JSON.stringify({ ...c, marks: undefined }), node: c });
  }
  return out;
}

const withMark = (n: JSONContent, m: Mark): JSONContent => ({ ...n, marks: [...(n.marks ?? []).filter((x) => x.type !== "insertion" && x.type !== "deletion"), m] });

/** Bitişik ve aynı işaretli metin parçalarını birleştirir. */
function mergeText(nodes: JSONContent[]): JSONContent[] {
  const out: JSONContent[] = [];
  for (const n of nodes) {
    const last = out[out.length - 1];
    if (last && last.type === "text" && n.type === "text" && JSON.stringify(last.marks ?? []) === JSON.stringify(n.marks ?? [])) last.text = (last.text ?? "") + (n.text ?? "");
    else out.push({ ...n });
  }
  return out;
}

/** Tüm satır içi içeriği işaretler (eklenen ya da silinen blok). */
export function markAll(block: JSONContent, m: Mark): JSONContent {
  if (block.type === "text" || (!block.content && block.type !== "paragraph" && block.type !== "heading" && !CONTAINERS.has(block.type ?? "") && block.type !== "section")) return withMark(block, m);
  return { ...block, ...(block.content ? { content: block.content.map((c) => markAll(c, m)) } : {}) };
}

/** Eşlenen iki metin bloğu: sözcük düzeyinde fark (biçim düzeltilmiş belgeden). */
function diffTextblock(o: JSONContent, r: JSONContent, ins: Mark, del: Mark): JSONContent {
  const a = tokens(o);
  const b = tokens(r);
  const pairs = lcsPairs(a, b, (x, y) => x.key === y.key);
  const out: JSONContent[] = [];
  let i = 0;
  let j = 0;
  const flush = (ti: number, tj: number) => {
    for (; i < ti; i++) out.push(withMark(a[i].node, del));
    for (; j < tj; j++) out.push(withMark(b[j].node, ins));
  };
  for (const [pi, pj] of pairs) {
    flush(pi, pj);
    out.push(b[pj].node);
    i = pi + 1;
    j = pj + 1;
  }
  flush(a.length, b.length);
  const content = mergeText(out);
  return { ...r, ...(content.length ? { content } : {}) };
}

function same(a: JSONContent, b: JSONContent) {
  return JSON.stringify(a) === JSON.stringify(b);
}

/** İki blok eşlenebilir mi (aynı tür ve yeterince benzer). */
function matchable(a: JSONContent, b: JSONContent): boolean {
  if (a.type !== b.type) return TEXTBLOCKS.has(a.type ?? "") && TEXTBLOCKS.has(b.type ?? "") && similarity(textOf(a), textOf(b)) >= 0.5;
  if (same(a, b)) return true;
  const ta = textOf(a);
  const tb = textOf(b);
  if (!ta.trim() && !tb.trim()) return true;
  return similarity(ta, tb) >= 0.4;
}

/** Blok dizilerini karşılaştırır. */
function diffBlocks(orig: JSONContent[], rev: JSONContent[], ins: Mark, del: Mark): JSONContent[] {
  const pairs = lcsPairs(orig, rev, matchable);
  const out: JSONContent[] = [];
  let i = 0;
  let j = 0;
  const flush = (ti: number, tj: number) => {
    for (; i < ti; i++) out.push(markAll(orig[i], del));
    for (; j < tj; j++) out.push(markAll(rev[j], ins));
  };
  for (const [pi, pj] of pairs) {
    flush(pi, pj);
    out.push(diffBlock(orig[pi], rev[pj], ins, del));
    i = pi + 1;
    j = pj + 1;
  }
  flush(orig.length, rev.length);
  return out;
}

function diffBlock(o: JSONContent, r: JSONContent, ins: Mark, del: Mark): JSONContent {
  if (same(o, r)) return r;
  if (TEXTBLOCKS.has(o.type ?? "") && TEXTBLOCKS.has(r.type ?? "")) return diffTextblock(o, r, ins, del);
  if (o.type === r.type && CONTAINERS.has(o.type ?? "")) {
    // Tablo/liste: aynı yapıdaysa çocuklar eşlenerek karşılaştırılır.
    if (o.type === "table" || o.type === "tableRow") {
      const oc = o.content ?? [];
      const rc = r.content ?? [];
      if (oc.length === rc.length) return { ...r, content: rc.map((c, k) => diffBlock(oc[k], c, ins, del)) };
    } else if (o.type === "tableCell" || o.type === "tableHeader" || o.type === "listItem" || o.type === "blockquote" || o.type === "textBox") {
      return { ...r, content: diffBlocks(o.content ?? [], r.content ?? [], ins, del) };
    } else return { ...r, content: diffBlocks(o.content ?? [], r.content ?? [], ins, del) };
  }
  // Farklı yapı: özgün silinmiş, düzeltilmiş eklenmiş gösterilir.
  return r.type === "section" ? r : markAll(r, ins);
}

/**
 * Özgün ve düzeltilmiş belge içeriğini karşılaştırır. Sonuç düzeltilmiş belgenin bölüm yapısındadır;
 * silinen bloklar, ardından gelen düzeltilmiş bloğun bölümüne yerleşir.
 */
export function compareDocuments(orig: JSONContent, rev: JSONContent, author: string, date = new Date().toISOString()): JSONContent {
  const ins: Mark = { type: "insertion", attrs: { author, date } };
  const del: Mark = { type: "deletion", attrs: { author, date } };
  const flat = (d: JSONContent) => (d.content ?? []).flatMap((sec, si) => (sec.type === "section" ? (sec.content ?? []).map((b) => ({ b, si })) : [{ b: sec, si: 0 }]));
  const o = flat(orig);
  const r = flat(rev);
  const sections = (rev.content ?? []).filter((s) => s.type === "section");
  const outSecs: JSONContent[][] = sections.map(() => []);
  if (!outSecs.length) outSecs.push([]);
  const pairs = lcsPairs(
    o.map((x) => x.b),
    r.map((x) => x.b),
    matchable,
  );
  let i = 0;
  let j = 0;
  const put = (si: number, n: JSONContent) => outSecs[Math.min(si, outSecs.length - 1)].push(n);
  const flush = (ti: number, tj: number) => {
    const target = tj < r.length ? r[tj].si : outSecs.length - 1;
    for (; i < ti; i++) put(target, markAll(o[i].b, del));
    for (; j < tj; j++) put(r[j].si, markAll(r[j].b, ins));
  };
  for (const [pi, pj] of pairs) {
    flush(pi, pj);
    put(r[pj].si, diffBlock(o[pi].b, r[pj].b, ins, del));
    i = pi + 1;
    j = pj + 1;
  }
  flush(o.length, r.length);
  return {
    type: "doc",
    content: outSecs.map((c, k) => ({ type: "section", attrs: sections[k]?.attrs ?? { kind: "nextPage", page: null }, content: c.length ? c : [{ type: "paragraph" }] })),
  };
}
