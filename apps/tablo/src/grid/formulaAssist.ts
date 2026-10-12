/** Formül yazarken otomatik tamamlama ve bağımsız değişken ipucu yardımcıları. */
import { FUNCTIONS, TR_SYNTAX, argNames, resolveFunctionName, tokenize, type FnDef } from "@officehere/formula";

export interface AssistInfo {
  /** Tamamlanacak ad öneki ve konumu. */
  prefix?: { start: number; end: number; text: string };
  /** Tablo1[… içinde: tablo adı ve sütun öneki. */
  column?: { table: string; start: number; end: number; text: string };
  /** İmlecin içinde bulunduğu fonksiyon ve bağımsız değişken sırası. */
  call?: { def: FnDef; argIndex: number };
}

export function analyzeFormula(text: string, caret: number): AssistInfo {
  if (!text.startsWith("=")) return {};
  const body = text.slice(1, caret);
  const toks = tokenize(body, TR_SYNTAX);
  const info: AssistInfo = {};
  // Tablo1[Tu… : yapısal başvuruda sütun adı tamamlanır
  const col = /([\p{L}_\\][\p{L}\p{N}_.\\]*)?\[(?:@)?([^[\]]*)$/u.exec(body);
  if (col) {
    const textStart = body.length - col[2].length;
    info.column = { table: col[1] ?? "", start: textStart + 1, end: caret, text: col[2] };
  }
  // Önek: son belirteç bir ad ve imleçte bitiyor
  const last = toks[toks.length - 1];
  if (last && (last.type === "name" || last.type === "bool" || (last.type === "ref" && /^[A-Za-zÇĞİÖŞÜçğıöşü]+\d*$/.test(last.text))) && last.pos + last.text.length === body.length) {
    const prev = toks.slice(0, -1).filter((t) => t.type !== "ws").pop();
    if (!prev || prev.type === "op" || prev.type === "lparen" || prev.type === "sep") {
      info.prefix = { start: last.pos + 1, end: caret, text: last.text };
    }
  }
  // Fonksiyon yığını
  const stack: { name: string; arg: number }[] = [];
  let pendingFunc: string | null = null;
  for (const t of toks) {
    if (t.type === "func") pendingFunc = t.value as string;
    else if (t.type === "lparen") {
      stack.push({ name: pendingFunc ?? "", arg: 0 });
      pendingFunc = null;
    } else if (t.type === "rparen") stack.pop();
    else if (t.type === "sep" && stack.length) stack[stack.length - 1].arg++;
    else if (t.type !== "ws") pendingFunc = null;
  }
  for (let i = stack.length - 1; i >= 0; i--) {
    if (!stack[i].name) continue;
    const canonical = resolveFunctionName(stack[i].name, true);
    const def = canonical ? FUNCTIONS.get(canonical) : undefined;
    if (def) {
      info.call = { def, argIndex: stack[i].arg };
      break;
    }
  }
  return info;
}

/** Otomatik tamamlama önerisi: işlev, tablo adı, tanımlı ad ya da tablo sütunu. */
export interface AssistItem {
  key: string;
  kind: "fn" | "table" | "name" | "column" | "item";
  label: string;
  /** Eklenecek metin ve imlecin eklenen metne göre konumu. */
  insert: string;
  desc: string;
}

/** Öneke uyan işlevler, tablolar ve tanımlı adlar. */
export function suggestItems(prefix: string, tables: { name: string; columns: string[] }[], names: string[], limit = 14): AssistItem[] {
  const p = prefix.toLocaleUpperCase("tr-TR");
  const out: AssistItem[] = [];
  for (const t of tables) if (t.name.toLocaleUpperCase("tr-TR").startsWith(p)) out.push({ key: `t:${t.name}`, kind: "table", label: t.name, insert: t.name, desc: "Tablo" });
  for (const n of names) if (n.toLocaleUpperCase("tr-TR").startsWith(p)) out.push({ key: `n:${n}`, kind: "name", label: n, insert: n, desc: "Tanımlı ad" });
  for (const d of suggestFunctions(prefix, limit)) out.push({ key: `f:${d.name}`, kind: "fn", label: d.tr, insert: `${d.tr}(`, desc: d.desc });
  return out.slice(0, limit);
}

const SPECIAL_ITEMS: [string, string][] = [
  ["#Tümü", "Tablonun tamamı: başlıklar, veriler ve toplam satırı"],
  ["#Veri", "Tablonun veri satırları"],
  ["#Başlıklar", "Sütun başlıkları"],
  ["#Toplamlar", "Toplam satırı"],
  ["@", "Bu satır"],
];

/** Tablo1[… içindeki sütun adları ve özel öğeler. */
export function suggestColumns(table: { name: string; columns: string[] } | undefined, prefix: string): AssistItem[] {
  if (!table) return [];
  const p = prefix.replace(/^\[/, "").toLocaleUpperCase("tr-TR");
  const out: AssistItem[] = [];
  for (const c of table.columns) if (c.toLocaleUpperCase("tr-TR").startsWith(p)) out.push({ key: `c:${c}`, kind: "column", label: c, insert: /^[\p{L}\p{N}_.][\p{L}\p{N}_. ]*$/u.test(c) && !c.endsWith(" ") ? `${c}]` : `[${c.replace(/(['[\]#])/g, "'$1")}]]`, desc: `${table.name} tablosunun sütunu` });
  for (const [it, desc] of SPECIAL_ITEMS) if (it.toLocaleUpperCase("tr-TR").startsWith(p) && it !== "@") out.push({ key: `i:${it}`, kind: "item", label: it, insert: `${it}]`, desc });
  return out;
}

/** Öneke uyan Türkçe fonksiyon adları. */
export function suggestFunctions(prefix: string, limit = 12): FnDef[] {
  const p = prefix.toLocaleUpperCase("tr-TR");
  const pe = prefix.toUpperCase();
  const out: FnDef[] = [];
  for (const d of FUNCTIONS.values()) {
    if (d.tr.startsWith(p) || (pe.length > 1 && d.name.startsWith(pe) && !d.tr.startsWith(p))) out.push(d);
  }
  out.sort((a, b) => a.tr.localeCompare(b.tr, "tr"));
  return out.slice(0, limit);
}

/** İpucu metni parçaları: [ad, bağımsız değişkenler (etkin olan işaretli)]. */
export function hintParts(def: FnDef, argIndex: number): { name: string; args: { text: string; active: boolean }[] } {
  const names = argNames(def);
  const args = names.map((a) => ({ text: a, active: false }));
  if (args.length) {
    // "..." ile biten imzalarda fazladan değişkenler son öğeye kadar
    let idx = argIndex;
    const lastIsRepeat = names[names.length - 1] === "...";
    if (idx >= args.length) idx = lastIsRepeat ? Math.max(0, args.length - 2) : args.length - 1;
    if (args[idx]) args[idx].active = true;
  }
  return { name: def.tr, args };
}
