/** Türkçe biçim kodu dönüşümü (kanonik ↔ yerel). */
import { COLOR_EN, COLOR_TR, parseSection, splitSections } from "./numfmt-parse";

/* ------------------------------------------------------------------ */
/* Türkçe biçim kodu dönüşümü                                          */
/* ------------------------------------------------------------------ */

/** Kod içindeki tırnaksız/parantezsiz kısımları dönüştürür. */
function mapUnquoted(code: string, fn: (seg: string) => string, bracket?: (inner: string) => string): string {
  let out = "";
  let seg = "";
  for (let i = 0; i < code.length; i++) {
    const c = code[i];
    if (c === '"') {
      out += fn(seg);
      seg = "";
      const j = code.indexOf('"', i + 1);
      const end = j < 0 ? code.length : j + 1;
      out += code.slice(i, end);
      i = end - 1;
    } else if (c === "[") {
      out += fn(seg);
      seg = "";
      const j = code.indexOf("]", i);
      const end = j < 0 ? code.length : j + 1;
      const inner = code.slice(i + 1, end - 1);
      out += "[" + (bracket ? bracket(inner) : inner) + "]";
      i = end - 1;
    } else if (c === "\\" || c === "_" || c === "*") {
      out += fn(seg);
      seg = "";
      out += code.slice(i, i + 2);
      i++;
    } else seg += c;
  }
  out += fn(seg);
  return out;
}

/** Kanonik kod → Türkçe gösterim ("#,##0.00" → "#.##0,00", "dd.mm.yyyy" → "gg.aa.yyyy"). */
export function formatToLocal(code: string): string {
  if (code === "General") return "Genel";
  return splitSections(code)
    .map((section) => {
      const sec = parseSection(section);
      const isDate = sec.kind === "date";
      return mapUnquoted(
        section,
        (seg) => {
          let r = seg.replace(/General/gi, "Genel");
          if (isDate) {
            r = r.replace(/AM\/PM/gi, "ÖÖ/ÖS").replace(/A\/P/gi, "Ö/S");
            // Dakika/ay ayrımı için parse edilmiş belirteçleri kullanmak yerine basit sıra kuralı
            let out = "";
            let lastWasH = false;
            for (let i = 0; i < r.length; ) {
              const c = r[i].toLowerCase();
              if ("ymdhs".includes(c)) {
                let j = i;
                while (j < r.length && r[j].toLowerCase() === c) j++;
                const run = j - i;
                let rest = r.slice(j);
                const nextIsS = /^[^a-z]*s/i.test(rest);
                let rep: string;
                if (c === "y") rep = "y".repeat(run);
                else if (c === "d") rep = "g".repeat(run);
                else if (c === "h") rep = "s".repeat(run);
                else if (c === "s") rep = "n".repeat(run);
                else rep = (lastWasH || nextIsS) && run <= 2 ? "d".repeat(run) : "a".repeat(run);
                lastWasH = c === "h";
                out += rep;
                i = j;
                rest = "";
                void rest;
              } else {
                out += r[i];
                i++;
              }
            }
            r = out;
            r = r.replace(/\./g, "\u0001").replace(/,/g, ".").replace(/\u0001/g, ".");
          } else {
            r = r.replace(/\./g, "\u0001").replace(/,/g, ".").replace(/\u0001/g, ",");
          }
          return r;
        },
        (inner) => {
          const key = Object.keys(COLOR_TR).find((k) => k.toLowerCase() === inner.toLowerCase());
          if (key) return COLOR_TR[key];
          if (/^(h+|m+|s+)$/i.test(inner)) return inner.toLowerCase().replace(/h/g, "s").replace(/m/g, "d").replace(/s(?=s*$)/g, (x) => x);
          return inner.replace(/^(<=|>=|<>|<|>|=)(\s*-?)(\d+)\.(\d+)$/, "$1$2$3,$4");
        },
      );
    })
    .join(";");
}

/** Kodun Türkçe yazımla mı girildiğini tahmin eder. */
export function looksLocalFormat(code: string): boolean {
  let letters = "";
  mapUnquoted(code, (seg) => {
    letters += seg.replace(/genel|general|am\/pm|a\/p|öö\/ös|ö\/s|e[+-]/gi, "");
    return seg;
  });
  const l = letters.toLocaleLowerCase("tr-TR");
  if (/[mh]/.test(l)) return false;
  if (/[gan]/.test(l)) return true;
  if (/s/.test(l) && /d/.test(l)) return true;
  if (/[dsy]/.test(l)) return false;
  // Yalnız sayı biçimi: ayırıcıların sırasına bak
  const code2 = code.replace(/"[^"]*"|\[[^\]]*\]|\\./g, "");
  const iDot = code2.indexOf(".");
  const iCom = code2.indexOf(",");
  if (iDot >= 0 && iCom >= 0) return iDot < iCom;
  if (iDot >= 0) return /\.(##0|###)/.test(code2);
  if (iCom >= 0) return !/,(##0|###)/.test(code2) && !/[0#],+(\s|$|;|%)/.test(code2);
  return true;
}

/** Türkçe kod → kanonik ("#.##0,00" → "#,##0.00", "gg.aa.yyyy" → "dd.mm.yyyy"). */
export function formatFromLocal(code: string): string {
  const t = code.trim();
  if (/^genel$/i.test(t)) return "General";
  return splitSections(t)
    .map((section) => {
      const lower = section.toLocaleLowerCase("tr-TR");
      const unq = lower.replace(/"[^"]*"|\[[^\]]*\]/g, "").replace(/öö\/ös|ö\/s|genel/g, "");
      const isDate = /[gaysdn]/.test(unq.replace(/e[+-]/g, ""));
      return mapUnquoted(
        section,
        (seg) => {
          let r = seg.replace(/Genel/gi, "General");
          if (isDate) {
            r = r.replace(/ÖÖ\/ÖS/gi, "AM/PM").replace(/Ö\/S/gi, "A/P");
            let out = "";
            for (let i = 0; i < r.length; ) {
              const c = r[i].toLocaleLowerCase("tr-TR");
              if ("gaysdn".includes(c)) {
                let j = i;
                while (j < r.length && r[j].toLocaleLowerCase("tr-TR") === c) j++;
                const run = j - i;
                const map: Record<string, string> = { g: "d", a: "m", y: "y", s: "h", d: "m", n: "s" };
                out += map[c].repeat(run);
                i = j;
              } else {
                out += r[i];
                i++;
              }
            }
            r = out;
          } else {
            r = r.replace(/\./g, "\u0001").replace(/,/g, ".").replace(/\u0001/g, ",");
          }
          return r;
        },
        (inner) => {
          const en = COLOR_EN[inner.toLocaleLowerCase("tr-TR")];
          if (en) return en;
          if (/^(s+|d+|n+)$/i.test(inner)) return inner.toLowerCase().replace(/s/g, "h").replace(/d/g, "m").replace(/n/g, "s");
          return inner.replace(/^(<=|>=|<>|<|>|=)(\s*-?)(\d+),(\d+)$/, "$1$2$3.$4");
        },
      );
    })
    .join(";");
}
