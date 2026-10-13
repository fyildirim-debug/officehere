// RTF okuyucu durumu: karakter ve paragraf biçimi denetim sözcükleri, atlanan hedefler, sabitler.
import type { UnderlineKind } from "../common";

export interface Chr {
  b: boolean;
  i: boolean;
  ul: UnderlineKind | null;
  strike: boolean;
  dstrike: boolean;
  sup: boolean;
  sub: boolean;
  fs: number | null;
  cf: number;
  hl: number;
  cb: number;
  f: number | null;
  caps: boolean;
  scaps: boolean;
  v: boolean;
  expnd: number;
  uc: number;
  link: string | null;
}

export interface Par {
  align: string | null;
  li: number | null;
  ri: number | null;
  fi: number | null;
  sb: number | null;
  sa: number | null;
  sl: number | null;
  slmult: boolean;
  keepn: boolean;
  keep: boolean;
  pagebb: boolean;
  intbl: boolean;
  ls: number | null;
  ilvl: number;
  s: number;
  outline: number | null;
  shade: number | null;
  brdrb: boolean;
  pn: { ordered: boolean; level: number } | null;
}

export const newChr = (): Chr => ({ b: false, i: false, ul: null, strike: false, dstrike: false, sup: false, sub: false, fs: null, cf: 0, hl: 0, cb: 0, f: null, caps: false, scaps: false, v: false, expnd: 0, uc: 1, link: null });
export const newPar = (): Par => ({ align: null, li: null, ri: null, fi: null, sb: null, sa: null, sl: null, slmult: false, keepn: false, keep: false, pagebb: false, intbl: false, ls: null, ilvl: 0, s: 0, outline: null, shade: null, brdrb: false, pn: null });

/** İçeriği okunmadan atlanan hedefler (yıldızsız olanlar dahil). */
export const SKIP_DESTS = new Set([
  "author", "buptim", "comment", "company", "creatim", "doccomm", "docvar", "hlinkbase", "keywords", "operator", "printim", "revtim", "subject", "title",
  "xe", "tc", "txe", "rxe", "private", "template", "ftnsep", "ftnsepc", "ftncn", "aftnsep", "aftnsepc", "aftncn", "nonshppict", "filetbl", "revtbl",
  "rsidtbl", "listpicture", "mmathPr", "themedata", "colorschememapping", "datastore", "latentstyles", "generator", "xmlnstbl", "fchars", "lchars",
  "pgptbl", "protusertbl", "bkmkend", "atnid", "atnauthor", "annotation", "atndate", "atnref", "atntime", "atnparent", "listtext", "pntext", "pntxta",
  "pntxtb", "objdata", "objclass", "objname", "shpinst", "sp", "sn", "sv", "background", "userprops", "category", "manager", "falt", "panose",
  "fname", "fontemb", "fontfile", "stylerestrictions", "ud", "upr", "wgrffmtfilter", "pgdsctbl", "oldcprops", "oldpprops", "oldtprops", "oldsprops",
  "defchp", "defpap", "pnseclvl", "picprop", "blipuid", "macpict", "pmmetafile",
]);

/** Yazıcının bağlantılara verdiği renk (okurken biçim sayılmaz). */
export const LINK_COLOR = "#0563C1";

/** \highlightN renk tablosu yoksa kullanılan standart vurgu renkleri. */
export const HL_COLORS = ["", "#000000", "#0000FF", "#00FFFF", "#00FF00", "#FF00FF", "#FF0000", "#FFFF00", "#FFFFFF", "#000080", "#008080", "#008000", "#800080", "#800000", "#808000", "#808080", "#C0C0C0"];

/** \levelnfc → liste türü. */
export const NFC_STYLE: Record<number, { ordered: boolean; style: string | null }> = {
  0: { ordered: true, style: null },
  1: { ordered: true, style: "upper-roman" },
  2: { ordered: true, style: "lower-roman" },
  3: { ordered: true, style: "upper-alpha" },
  4: { ordered: true, style: "lower-alpha" },
  23: { ordered: false, style: null },
};

const UL: Record<string, UnderlineKind> = {
  ul: "single",
  uldb: "double",
  ulth: "thick",
  ulthd: "thick",
  ulthdash: "thick",
  uld: "dotted",
  uldash: "dashed",
  uldashd: "dashed",
  uldashdd: "dashed",
  ulldash: "dashed",
  ulwave: "wave",
  ululdbwave: "wave",
  ulhwave: "wave",
  ulw: "words",
};

/** Karakter biçimi sözcüğü; işlendiyse true. */
export function charWord(c: Chr, w: string, param: number | null): boolean {
  const on = param === null || param !== 0;
  if (UL[w]) {
    c.ul = on ? UL[w] : null;
    return true;
  }
  switch (w) {
    case "plain":
      Object.assign(c, newChr(), { uc: c.uc, link: c.link });
      return true;
    case "ulnone":
      c.ul = null;
      return true;
    case "b":
    case "i":
    case "strike":
    case "caps":
    case "scaps":
    case "v":
      c[w] = on;
      return true;
    case "striked":
      c.dstrike = on;
      return true;
    case "super":
    case "sub":
      c.sup = w === "super" && on;
      c.sub = w === "sub" && on;
      return true;
    case "nosupersub":
      c.sup = c.sub = false;
      return true;
    case "up":
      c.sup = (param ?? 6) > 0;
      return true;
    case "dn":
      c.sub = (param ?? 6) > 0;
      return true;
    case "fs":
      c.fs = param && param > 0 && param < 3300 ? param : null;
      return true;
    case "f":
      c.f = param;
      return true;
    case "cf":
      c.cf = param ?? 0;
      return true;
    case "highlight":
      c.hl = param ?? 0;
      return true;
    case "cb":
    case "chcbpat":
      c.cb = param ?? 0;
      return true;
    case "expnd":
      c.expnd = (param ?? 0) * 5;
      return true;
    case "expndtw":
      c.expnd = param ?? 0;
      return true;
    case "uc":
      c.uc = Math.max(0, param ?? 1);
      return true;
  }
  return false;
}

const ALIGN: Record<string, string> = { ql: "left", qc: "center", qr: "right", qj: "justify", qd: "justify" };

/** Paragraf biçimi sözcüğü; işlendiyse true. */
export function paraWord(pa: Par, w: string, param: number | null): boolean {
  const on = param === null || param !== 0;
  if (ALIGN[w]) {
    pa.align = ALIGN[w];
    return true;
  }
  switch (w) {
    case "pard":
      Object.assign(pa, newPar());
      return true;
    case "li":
    case "lin":
      pa.li = param ?? 0;
      return true;
    case "ri":
    case "rin":
      pa.ri = param ?? 0;
      return true;
    case "fi":
    case "sb":
    case "sa":
      pa[w] = param ?? 0;
      return true;
    case "sl":
      pa.sl = param;
      return true;
    case "slmult":
    case "keepn":
    case "keep":
    case "pagebb":
      pa[w] = on;
      return true;
    case "intbl":
      pa.intbl = true;
      return true;
    case "itap":
      pa.intbl = (param ?? 0) > 0;
      return true;
    case "ls":
      pa.ls = param;
      return true;
    case "ilvl":
      pa.ilvl = Math.max(0, Math.min(8, param ?? 0));
      return true;
    case "s":
      pa.s = param ?? 0;
      return true;
    case "outlinelevel":
      pa.outline = param;
      return true;
    case "cbpat":
      pa.shade = param;
      return true;
    case "brdrb":
      pa.brdrb = true;
      return true;
  }
  return false;
}

/** \info alanı → belge özellikleri. */
export function applyInfo(p: { title: string; author: string; subject: string; keywords: string; created: string; modified: string }, key: string, v: string) {
  if (key === "title" || key === "author" || key === "subject" || key === "keywords") p[key] = v;
  else if (key === "creatim" || key === "revtim") {
    const n = (k: string) => Number(new RegExp(`${k}(-?\\d+)`).exec(v)?.[1] ?? 0);
    const yr = n("yr");
    if (yr > 1900 && yr < 3000) {
      const d = new Date(Date.UTC(yr, Math.max(0, n("mo") - 1), Math.max(1, n("dy")), n("hr"), n("min")));
      if (key === "creatim") p.created = d.toISOString();
      else p.modified = d.toISOString();
    }
  }
}

/** Stil adından başlık düzeyi ("heading 2", "Başlık 2"); ad yoksa anahat düzeyi. */
export function headingLevelOf(styleName: string | undefined, outline: number | null): number | null {
  const name = (styleName ?? "").toLowerCase().trim();
  const m = /^(?:heading|başlık|baslik|überschrift|titre)\s*(\d)$/.exec(name);
  if (m) return Math.min(6, Math.max(1, Number(m[1])));
  if (outline !== null && outline >= 0 && outline < 6 && !/^(normal|default|standard|body)/.test(name)) return outline + 1;
  return null;
}

/** Stil adından bizim paragraf stilimiz. */
export function styleIdByName(styleName: string | undefined): string | null {
  const name = (styleName ?? "").toLowerCase().trim();
  if (name === "title" || name === "konu başlığı") return "Title";
  if (name === "subtitle" || name === "alt konu başlığı") return "Subtitle";
  if (name === "quote" || name === "quotations" || name === "alıntı") return "Quote";
  if (name === "intense quote" || name === "belirgin alıntı") return "IntenseQuote";
  if (name === "no spacing" || name === "aralık yok") return "NoSpacing";
  return null;
}
