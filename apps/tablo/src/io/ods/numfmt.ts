/**
 * Sayı biçimi kodu ("#,##0.00", "dd.mm.yyyy", "0%;[Red]-0%") ⇄ ODF veri
 * stilleri (number:number-style, number:date-style …, OpenDocument 1.3 §16.29).
 * Yapabildiği kadar çevirir; bilinmeyen öğeler atlanır.
 */
import { esc, kids, textOf, type XmlEl } from "./xml";

/* ------------------------------ Okuma ------------------------------ */

const NAMED_COLORS: Record<string, string> = { "#000000": "Black", "#ffffff": "White", "#ff0000": "Red", "#00ff00": "Green", "#0000ff": "Blue", "#ffff00": "Yellow", "#ff00ff": "Magenta", "#00ffff": "Cyan" };

function quoteLit(t: string): string {
  if (!t) return "";
  // Tek karakterlik yaygın ayırıcılar tırnaksız kalır
  if (/^[ .,:/\-()]+$/.test(t)) return t.replace(/,/g, "\\,");
  return `"${t.replace(/"/g, '\\"')}"`;
}

function numberPart(el: XmlEl): string {
  const a = el.attrs;
  // Ondalık basamağı belirtilmemiş sayı "Genel" biçimdir
  if (a["number:decimal-places"] === undefined && a["number:grouping"] !== "true") return "General";
  const dec = Math.max(0, Math.min(30, Number(a["number:decimal-places"] ?? 0) || 0));
  const minDec = Math.max(0, Math.min(dec, Number(a["number:min-decimal-places"] ?? dec) || 0));
  const minInt = Math.max(0, Math.min(30, Number(a["number:min-integer-digits"] ?? 1) || 0));
  let int = "0".repeat(minInt) || "#";
  if (a["number:grouping"] === "true") {
    // Sağdan minInt basamak "0", kalan "#"; son üç basamaktan önce binlik ayırıcı
    const digits = "#".repeat(Math.max(0, 4 - minInt)) + "0".repeat(minInt);
    int = digits.slice(0, -3) + "," + digits.slice(-3);
  }
  const decs = "0".repeat(minDec) + "#".repeat(dec - minDec);
  return int + (decs ? "." + decs : "");
}

/** Tek bir veri stilinin bölüm kodu. */
function sectionCode(el: XmlEl, kind: string): string {
  let out = "";
  const color = kids(el, "style:text-properties")[0]?.attrs["fo:color"]?.toLowerCase();
  if (color && NAMED_COLORS[color]) out += `[${NAMED_COLORS[color]}]`;
  for (const c of kids(el)) {
    const a = c.attrs;
    const long = a["number:style"] === "long";
    switch (c.name) {
      case "number:number":
        out += numberPart(c);
        break;
      case "number:scientific-number": {
        const dec = Math.min(30, Number(a["number:decimal-places"] ?? 2) || 0);
        const ex = Math.max(1, Math.min(4, Number(a["number:min-exponent-digits"] ?? 2) || 2));
        out += `0${dec ? "." + "0".repeat(dec) : ""}E+${"0".repeat(ex)}`;
        break;
      }
      case "number:fraction": {
        const den = a["number:denominator-value"];
        const dd = Math.max(1, Math.min(5, Number(a["number:min-denominator-digits"] ?? 1) || 1));
        const nd = Math.max(1, Math.min(5, Number(a["number:min-numerator-digits"] ?? 1) || 1));
        const minInt = Number(a["number:min-integer-digits"] ?? 0) || 0;
        out += `${minInt > 0 ? "0" : "#"} ${"?".repeat(nd)}/${den ?? "?".repeat(dd)}`;
        break;
      }
      case "number:text": {
        const t = textOf(c);
        // Yüzde stilindeki "%" ölçekleme işaretidir, tırnaklanmaz
        out += kind === "percentage" && t.trim() === "%" ? t : quoteLit(t);
        break;
      }
      case "number:currency-symbol": {
        const sym = textOf(c).trim();
        if (sym) out += `"${sym.replace(/"/g, "")}"`;
        break;
      }
      case "number:text-content":
        out += "@";
        break;
      case "number:fill-character":
        out += "*" + (textOf(c)[0] ?? " ");
        break;
      case "number:day":
        out += long ? "dd" : "d";
        break;
      case "number:month":
        out += a["number:textual"] === "true" ? (long ? "mmmm" : "mmm") : long ? "mm" : "m";
        break;
      case "number:year":
        out += long ? "yyyy" : "yy";
        break;
      case "number:day-of-week":
        out += long ? "dddd" : "ddd";
        break;
      case "number:hours": {
        const h = long ? "hh" : "h";
        out += el.attrs["number:truncate-on-overflow"] === "false" ? `[${h}]` : h;
        break;
      }
      case "number:minutes":
        out += long ? "mm" : "m";
        break;
      case "number:seconds": {
        const dec = Math.min(6, Number(a["number:decimal-places"] ?? 0) || 0);
        out += (long ? "ss" : "s") + (dec ? "." + "0".repeat(dec) : "");
        break;
      }
      case "number:am-pm":
        out += "AM/PM";
        break;
      case "number:boolean":
        out += "General";
        break;
    }
  }
  return out;
}

const STD_COND: Record<string, true> = { ">=0": true, ">0": true, "<0": true, "=0": true };

/** Veri stili (ad → öğe) tablosundan biçim kodu üretir. "General" eşdeğerinde undefined. */
export function dataStyleToCode(name: string, styles: Map<string, XmlEl>, depth = 0): string | undefined {
  const el = styles.get(name);
  if (!el || depth > 4) return undefined;
  const kind = el.name.replace(/^number:/, "").replace(/-style$/, "");
  if (kind === "boolean") return undefined;
  const main = sectionCode(el, kind);
  const sections: string[] = [];
  for (const m of kids(el, "style:map")) {
    const target = m.attrs["style:apply-style-name"];
    const tEl = target ? styles.get(target) : undefined;
    if (!tEl) continue;
    const cond = (m.attrs["style:condition"] ?? "").replace(/^value\(\)/, "").replace(/\s+/g, "");
    const code = sectionCode(tEl, tEl.name.replace(/^number:/, "").replace(/-style$/, ""));
    sections.push(STD_COND[cond] ? code : `[${cond}]${code}`);
  }
  const all = [...sections, main].join(";");
  if (!all || all === "General") return undefined;
  return all;
}

/* ------------------------------ Yazma ------------------------------ */

function splitSections(code: string): string[] {
  const out: string[] = [];
  let cur = "";
  let q = false;
  let br = false;
  for (let i = 0; i < code.length; i++) {
    const c = code[i];
    if (c === "\\" && !q) {
      cur += c + (code[i + 1] ?? "");
      i++;
      continue;
    }
    if (c === '"') q = !q;
    else if (!q && c === "[") br = true;
    else if (!q && c === "]") br = false;
    if (c === ";" && !q && !br) {
      out.push(cur);
      cur = "";
    } else cur += c;
  }
  out.push(cur);
  return out;
}

type Tok = { t: "lit"; v: string } | { t: "code"; v: string } | { t: "color"; v: string } | { t: "cond"; v: string };

function tokenizeSection(s: string): Tok[] {
  const out: Tok[] = [];
  let i = 0;
  const lit = (v: string) => {
    const last = out[out.length - 1];
    if (last?.t === "lit") last.v += v;
    else out.push({ t: "lit", v });
  };
  while (i < s.length) {
    const c = s[i];
    if (c === '"') {
      const e = s.indexOf('"', i + 1);
      lit(s.slice(i + 1, e < 0 ? s.length : e));
      i = e < 0 ? s.length : e + 1;
    } else if (c === "\\") {
      lit(s[i + 1] ?? "");
      i += 2;
    } else if (c === "_") {
      lit(" ");
      i += 2;
    } else if (c === "*") {
      i += 2;
    } else if (c === "[") {
      const e = s.indexOf("]", i);
      const inner = s.slice(i + 1, e < 0 ? s.length : e);
      i = e < 0 ? s.length : e + 1;
      if (/^\$/.test(inner)) {
        const sym = inner.slice(1).split("-")[0];
        if (sym) lit(sym);
      } else if (/^[<>=]/.test(inner)) out.push({ t: "cond", v: inner });
      else if (/^(h+|m+|s+)$/i.test(inner)) out.push({ t: "code", v: `[${inner.toLowerCase()}]` });
      else out.push({ t: "color", v: inner });
    } else if (/^(AM\/PM|A\/P)/i.test(s.slice(i))) {
      const m = /^(AM\/PM|A\/P)/i.exec(s.slice(i))!;
      out.push({ t: "code", v: "AM/PM" });
      i += m[0].length;
    } else if (/[yYmMdDhHsS]/.test(c)) {
      let j = i;
      while (j < s.length && s[j].toLowerCase() === c.toLowerCase()) j++;
      out.push({ t: "code", v: s.slice(i, j).toLowerCase() });
      i = j;
    } else if (/[0#?.,%Ee@/+]/.test(c)) {
      let j = i;
      while (j < s.length && /[0#?.,%Ee@/+\-]/.test(s[j]) && !(s[j] === "-" && !/[Ee]/.test(s[j - 1] ?? ""))) j++;
      out.push({ t: "code", v: s.slice(i, j) });
      i = j;
    } else {
      lit(c);
      i++;
    }
  }
  return out;
}

const COLORS: Record<string, string> = { black: "#000000", white: "#FFFFFF", red: "#FF0000", green: "#00FF00", blue: "#0000FF", yellow: "#FFFF00", magenta: "#FF00FF", cyan: "#00FFFF" };

function isDateSection(toks: Tok[]): boolean {
  return toks.some((t) => t.t === "code" && /^(y+|d+|h+|s+|\[h+\]|\[m+\]|\[s+\]|AM\/PM)$/.test(t.v));
}

/** Bir bölümü veri stili öğesine çevirir. */
function sectionXml(name: string, section: string, volatile: boolean): string {
  const toks = tokenizeSection(section);
  const colorTok = toks.find((t) => t.t === "color");
  const color = colorTok ? COLORS[colorTok.v.toLowerCase()] : undefined;
  const textProps = color ? `<style:text-properties fo:color="${color}"/>` : "";
  const vol = volatile ? ' style:volatile="true"' : "";
  const litXml = (v: string) => (v ? `<number:text>${esc(v)}</number:text>` : "");
  if (isDateSection(toks)) {
    const date = toks.some((t) => t.t === "code" && /^(y+|d+)$/.test(t.v)) || toks.some((t) => t.t === "code" && /^m{3,}$/.test(t.v));
    const elapsed = toks.some((t) => t.t === "code" && /^\[/.test(t.v));
    let body = "";
    // "m" dakika mı ay mı: saatten sonra ya da saniyeden önceyse dakika
    const codes = toks.map((t, i) => ({ t, i })).filter((x) => x.t.t === "code");
    for (const { t, i } of toks.map((t, i) => ({ t, i }))) {
      if (t.t === "lit") body += litXml(t.v);
      if (t.t !== "code") continue;
      const v = t.v.replace(/^\[|\]$/g, "");
      if (/^y+$/.test(v)) body += `<number:year${v.length > 2 ? ' number:style="long"' : ""}/>`;
      else if (/^d+$/.test(v)) body += v.length >= 3 ? `<number:day-of-week${v.length > 3 ? ' number:style="long"' : ""}/>` : `<number:day${v.length > 1 ? ' number:style="long"' : ""}/>`;
      else if (/^h+$/.test(v)) body += `<number:hours${v.length > 1 ? ' number:style="long"' : ""}/>`;
      else if (/^s+(\.0+)?$/.test(v) || /^s+$/.test(v)) body += `<number:seconds${v.length > 1 ? ' number:style="long"' : ""}/>`;
      else if (v === "AM/PM") body += "<number:am-pm/>";
      else if (/^m+$/.test(v)) {
        const k = codes.findIndex((x) => x.i === i);
        const prev = codes[k - 1]?.t.v ?? "";
        const next = codes[k + 1]?.t.v ?? "";
        const minute = v.length <= 2 && (/h/.test(prev) || /^s/.test(next) || /^\[s/.test(next));
        if (minute) body += `<number:minutes${v.length > 1 ? ' number:style="long"' : ""}/>`;
        else body += `<number:month${v.length >= 3 ? ' number:textual="true"' : ""}${v.length === 2 || v.length === 4 ? ' number:style="long"' : ""}/>`;
      } else if (/^\.0+$/.test(v) && /<number:seconds[^>]*\/>$/.test(body)) {
        // Saniye kesri: son saniye öğesine ondalık basamak
        body = body.replace(/<number:seconds([^>]*)\/>$/, `<number:seconds$1 number:decimal-places="${v.length - 1}"/>`);
      } else body += litXml(v);
    }
    const tag = date ? "number:date-style" : "number:time-style";
    const trunc = elapsed ? ' number:truncate-on-overflow="false"' : "";
    return `<${tag} style:name="${name}"${vol}${trunc}>${textProps}${body}</${tag}>`;
  }
  const codeTok = toks.filter((t) => t.t === "code").map((t) => t.v).join("");
  if (codeTok.includes("@")) {
    let body = "";
    for (const t of toks) {
      if (t.t === "lit") body += litXml(t.v);
      else if (t.t === "code" && t.v.includes("@")) body += "<number:text-content/>";
    }
    return `<number:text-style style:name="${name}"${vol}>${textProps}${body}</number:text-style>`;
  }
  const percent = codeTok.includes("%");
  let tag = percent ? "number:percentage-style" : "number:number-style";
  let body = "";
  let numberDone = false;
  for (const t of toks) {
    if (t.t === "lit") {
      body += litXml(t.v);
      continue;
    }
    if (t.t !== "code") continue;
    if (numberDone) {
      if (t.v.includes("%")) body += litXml("%");
      continue;
    }
    numberDone = true;
    const v = t.v.replace(/%/g, "");
    const exp = /[Ee][+-]/.exec(v);
    if (exp) {
      const mant = v.slice(0, exp.index);
      const dec = (mant.split(".")[1] ?? "").length;
      const ex = v.slice(exp.index + 2).length || 2;
      body += `<number:scientific-number number:decimal-places="${dec}" number:min-decimal-places="${dec}" number:min-integer-digits="1" number:min-exponent-digits="${ex}"/>`;
    } else if (v.includes("/")) {
      const [left, den] = v.split("/");
      const num = left.trim().split(/\s+/).pop() ?? "?";
      const fixed = /^\d+$/.test(den ?? "") ? ` number:denominator-value="${den}"` : "";
      body += `<number:fraction number:min-integer-digits="0" number:min-numerator-digits="${num.length}" number:min-denominator-digits="${(den ?? "?").length}"${fixed}/>`;
    } else {
      const [int, dec = ""] = v.split(".");
      const minInt = (int.match(/0/g) ?? []).length;
      const minDec = (dec.match(/0/g) ?? []).length;
      const allDec = (dec.match(/[0#?]/g) ?? []).length;
      const grouping = int.includes(",") ? ' number:grouping="true"' : "";
      body += `<number:number number:decimal-places="${allDec}" number:min-decimal-places="${minDec}" number:min-integer-digits="${minInt}"${grouping}/>`;
    }
    if (t.v.includes("%")) body += litXml("%");
  }
  if (!numberDone) tag = "number:number-style";
  return `<${tag} style:name="${name}"${vol}>${textProps}${body}</${tag}>`;
}

/**
 * Biçim kodunu ODF veri stillerine çevirir: ana stil "name", ek bölümler
 * "nameP0", "nameP1"… olarak yazılır (style:map ile bağlanır).
 */
export function codeToDataStyles(name: string, code: string): string {
  const secs = splitSections(code).slice(0, 4);
  if (secs.length === 1) return sectionXml(name, secs[0], false);
  const conds = secs.length === 2 ? [">=0"] : secs.length === 3 ? [">0", "<0"] : [">0", "<0", "=0"];
  let out = "";
  const maps: string[] = [];
  conds.forEach((c, i) => {
    const custom = /^\[([<>=][^\]]*)\]/.exec(secs[i]);
    const cond = custom ? custom[1] : c;
    out += sectionXml(`${name}P${i}`, secs[i].replace(/^\[[<>=][^\]]*\]/, ""), true);
    maps.push(`<style:map style:condition="value()${esc(cond)}" style:apply-style-name="${name}P${i}"/>`);
  });
  const main = sectionXml(name, secs[secs.length - 1], false);
  // Haritalar ana stilin sonuna eklenir
  return out + main.replace(/(<\/number:[a-z-]+>)$/, maps.join("") + "$1");
}
