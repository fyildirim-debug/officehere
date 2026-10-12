/**
 * Yapısal başvurular (tablo başvuruları): Tablo1[Tutar], Tablo1[@Tutar],
 * Tablo1[[#Tümü];[Tutar]], [@Adet] … Kanonik (dosya) biçim İngilizcedir:
 * Table1[[#This Row],[Tutar]]; Türkçe arayüz @ kısaltmasını ve Türkçe özel
 * öğe adlarını (#Tümü, #Veri, #Başlıklar, #Toplamlar, #Bu Satır) kullanır.
 */

export type StructItem = "#All" | "#Data" | "#Headers" | "#Totals" | "#This Row";

export interface StructRef {
  /** Tablo adı; yoksa formülün bulunduğu tablo. */
  table?: string;
  items: StructItem[];
  col1?: string;
  col2?: string;
}

/** Hesap motorunun bildiği tablo. */
export interface TableInfo {
  name: string;
  sheet: number;
  r1: number;
  c1: number;
  r2: number;
  c2: number;
  headerRow: boolean;
  totalRow: boolean;
  columns: string[];
}

const ITEM_NAMES: Record<string, StructItem> = {
  "#ALL": "#All",
  "#DATA": "#Data",
  "#HEADERS": "#Headers",
  "#TOTALS": "#Totals",
  "#THIS ROW": "#This Row",
  "#TÜMÜ": "#All",
  "#VERİ": "#Data",
  "#BAŞLIKLAR": "#Headers",
  "#TOPLAMLAR": "#Totals",
  "#BU SATIR": "#This Row",
};

const TR_ITEM: Record<StructItem, string> = { "#All": "#Tümü", "#Data": "#Veri", "#Headers": "#Başlıklar", "#Totals": "#Toplamlar", "#This Row": "#Bu Satır" };

function itemOf(s: string): StructItem | undefined {
  return ITEM_NAMES[s.trim().replace(/\s+/g, " ").toLocaleUpperCase("tr-TR")] ?? ITEM_NAMES[s.trim().replace(/\s+/g, " ").toUpperCase()];
}

function unescapeCol(s: string): string {
  return s.replace(/'(.)/g, "$1");
}

function escapeCol(s: string): string {
  return s.replace(/(['[\]#])/g, "'$1");
}

/** Köşeli parantez içindeki konumdan eşleşen kapanışı bulur (' kaçışlarıyla). Bulunamazsa -1. */
export function matchBracket(src: string, open: number): number {
  let depth = 0;
  for (let i = open; i < src.length; i++) {
    const ch = src[i];
    if (ch === "'") {
      i++;
      continue;
    }
    if (ch === "[") depth++;
    else if (ch === "]") {
      depth--;
      if (depth === 0) return i;
    }
  }
  return -1;
}

/** Dış köşeli parantezlerin içini ayrıştırır. */
export function parseStructInner(inner: string): Omit<StructRef, "table"> | null {
  const s = inner.trim();
  if (s === "") return { items: ["#Data"] };
  if (s.startsWith("@")) {
    const rest = s.slice(1).trim();
    if (rest === "") return { items: ["#This Row"] };
    if (!rest.startsWith("[")) return { items: ["#This Row"], col1: unescapeCol(rest) };
    const cols = parseColumnSpec(rest);
    return cols ? { items: ["#This Row"], ...cols } : null;
  }
  if (s.startsWith("#")) {
    const it = itemOf(s);
    return it ? { items: [it] } : null;
  }
  if (!s.startsWith("[")) return { items: [], col1: unescapeCol(s) };
  // Liste: [#Öğe],[#Öğe],[Sütun] ya da [Sütun1]:[Sütun2]
  const parts: string[] = [];
  let i = 0;
  while (i < s.length) {
    while (i < s.length && /[\s,;]/.test(s[i])) i++;
    if (i >= s.length) break;
    if (s[i] !== "[") return null;
    const end = matchBracket(s, i);
    if (end < 0) return null;
    let part = s.slice(i, end + 1);
    i = end + 1;
    // Sütun aralığı: [A]:[B]
    let j = i;
    while (j < s.length && s[j] === " ") j++;
    if (s[j] === ":") {
      j++;
      while (j < s.length && s[j] === " ") j++;
      if (s[j] !== "[") return null;
      const e2 = matchBracket(s, j);
      if (e2 < 0) return null;
      part += ":" + s.slice(j, e2 + 1);
      i = e2 + 1;
    }
    parts.push(part);
  }
  const out: Omit<StructRef, "table"> = { items: [] };
  for (const p of parts) {
    const body = p.slice(1, p.indexOf("]") >= 0 ? -1 : undefined);
    if (p.startsWith("[#")) {
      const it = itemOf(p.slice(1, -1));
      if (!it) return null;
      out.items.push(it);
      continue;
    }
    if (out.col1 !== undefined) return null;
    const cols = parseColumnSpec(p);
    if (!cols) return null;
    Object.assign(out, cols);
    void body;
  }
  return out;
}

function parseColumnSpec(s: string): { col1: string; col2?: string } | null {
  const e1 = matchBracket(s, 0);
  if (e1 < 0) return null;
  const col1 = unescapeCol(s.slice(1, e1));
  const rest = s.slice(e1 + 1).trim();
  if (!rest) return { col1 };
  if (!rest.startsWith(":")) return null;
  const r2 = rest.slice(1).trim();
  const e2 = matchBracket(r2, 0);
  if (e2 < 0 || r2.slice(e2 + 1).trim()) return null;
  return { col1, col2: unescapeCol(r2.slice(1, e2)) };
}

function colPart(c: string): string {
  return `[${escapeCol(c)}]`;
}

/** Sütun adı köşeli parantezsiz yazılabilir mi (Tablo1[Tutar])? */
function simpleCol(c: string): boolean {
  return /^[\p{L}\p{N}_.][\p{L}\p{N}_. ]*$/u.test(c) && !c.endsWith(" ");
}

/** Yapısal başvuruyu metne çevirir: "en" kanonik (dosya) biçimi, "tr" Türkçe arayüz biçimi. */
export function renderStructRef(r: StructRef, locale: "tr" | "en" = "en"): string {
  const t = r.table ?? "";
  const sep = locale === "tr" ? ";" : ",";
  const item = (i: StructItem) => `[${locale === "tr" ? TR_ITEM[i] : i}]`;
  const cols = r.col1 !== undefined ? (r.col2 !== undefined ? `${colPart(r.col1)}:${colPart(r.col2)}` : colPart(r.col1)) : "";
  const thisRowOnly = r.items.length === 1 && r.items[0] === "#This Row";
  if (locale === "tr" && thisRowOnly) {
    if (r.col1 === undefined) return `${t}[@]`;
    if (r.col2 === undefined && simpleCol(r.col1)) return `${t}[@${r.col1}]`;
    return `${t}[@${cols}]`;
  }
  if (!r.items.length) {
    if (r.col1 === undefined) return `${t}[]`;
    if (r.col2 === undefined && simpleCol(r.col1)) return `${t}[${r.col1}]`;
    return `${t}[${cols}]`;
  }
  if (r.col1 === undefined && r.items.length === 1) return `${t}${item(r.items[0])}`;
  const parts = r.items.map(item);
  if (cols) parts.push(cols);
  return `${t}[${parts.join(sep)}]`;
}

/** Başvurunun gösterdiği hücre dikdörtgeni; hata durumunda null ("#BAŞV!" / "#DEĞER!"). */
export function resolveStructRect(
  r: StructRef,
  t: TableInfo,
  row: number,
): { sheet: number; r1: number; c1: number; r2: number; c2: number } | { error: "#REF!" | "#VALUE!" } {
  const dataStart = t.r1 + (t.headerRow ? 1 : 0);
  const dataEnd = t.r2 - (t.totalRow ? 1 : 0);
  let c1 = t.c1;
  let c2 = t.c2;
  if (r.col1 !== undefined) {
    const find = (name: string) => t.columns.findIndex((c) => c.toLocaleUpperCase("tr-TR") === name.toLocaleUpperCase("tr-TR"));
    const i1 = find(r.col1);
    const i2 = r.col2 !== undefined ? find(r.col2) : i1;
    if (i1 < 0 || i2 < 0) return { error: "#REF!" };
    c1 = t.c1 + Math.min(i1, i2);
    c2 = t.c1 + Math.max(i1, i2);
  }
  const items = r.items.length ? r.items : (["#Data"] as StructItem[]);
  if (items.includes("#This Row")) {
    if (row < dataStart || row > dataEnd) return { error: "#VALUE!" };
    return { sheet: t.sheet, r1: row, c1, r2: row, c2 };
  }
  let r1 = Infinity;
  let r2 = -Infinity;
  const span = (a: number, b: number) => {
    if (b < a) return;
    r1 = Math.min(r1, a);
    r2 = Math.max(r2, b);
  };
  for (const it of items) {
    if (it === "#All") span(t.r1, t.r2);
    else if (it === "#Data") span(dataStart, dataEnd);
    else if (it === "#Headers") {
      if (!t.headerRow) return { error: "#REF!" };
      span(t.r1, t.r1);
    } else if (it === "#Totals") {
      if (!t.totalRow) return { error: "#REF!" };
      span(t.r2, t.r2);
    }
  }
  if (!Number.isFinite(r1)) return { error: "#REF!" };
  return { sheet: t.sheet, r1, c1, r2, c2 };
}

/** Formüldeki tablo adını değiştirir (Tablo Adı yeniden adlandırma). */
export function renameTableInRef(r: StructRef, oldName: string, newName: string): StructRef {
  if (r.table && r.table.toLocaleUpperCase("tr-TR") === oldName.toLocaleUpperCase("tr-TR")) return { ...r, table: newName };
  return r;
}
