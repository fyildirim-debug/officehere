/** Formül yazarken otomatik tamamlama ve bağımsız değişken ipucu yardımcıları. */
import { FUNCTIONS, TR_SYNTAX, argNames, resolveFunctionName, tokenize, type FnDef } from "@officehere/formula";

export interface AssistInfo {
  /** Tamamlanacak ad öneki ve konumu. */
  prefix?: { start: number; end: number; text: string };
  /** İmlecin içinde bulunduğu fonksiyon ve bağımsız değişken sırası. */
  call?: { def: FnDef; argIndex: number };
}

export function analyzeFormula(text: string, caret: number): AssistInfo {
  if (!text.startsWith("=")) return {};
  const body = text.slice(1, caret);
  const toks = tokenize(body, TR_SYNTAX);
  const info: AssistInfo = {};
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
