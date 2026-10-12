// Alıntılar ve Kaynakça: kaynak kayıtları ve APA / IEEE biçimlendirmesi.

export type SourceType = "book" | "article" | "web" | "report";

export interface Source {
  /** Etiket (ör. "Yil2020"); alıntı düğümleri bu kimlikle kaynağa bağlanır. */
  tag: string;
  type: SourceType;
  /** Yazarlar "Soyadı, Ad; Soyadı, Ad" biçiminde. */
  author: string;
  title: string;
  year: string;
  /** Kitap: yayınevi; makale: dergi adı; web: site adı. */
  publisher: string;
  city?: string;
  volume?: string;
  issue?: string;
  pages?: string;
  url?: string;
  /** Web sitesine erişim tarihi. */
  accessed?: string;
}

export const SOURCE_TYPE_LABEL: Record<SourceType, string> = { book: "Kitap", article: "Dergi Makalesi", web: "Web Sitesi", report: "Rapor" };

export type BibStyle = "apa" | "ieee";

export const BIB_STYLE_LABEL: Record<BibStyle, string> = { apa: "APA", ieee: "IEEE" };

interface Person {
  last: string;
  first: string;
}

export function parseAuthors(s: string): Person[] {
  return s
    .split(/;|\s+ve\s+|\s+and\s+/i)
    .map((a) => a.trim())
    .filter(Boolean)
    .map((a) => {
      if (a.includes(",")) {
        const [last, first] = a.split(",", 2).map((x) => x.trim());
        return { last, first: first ?? "" };
      }
      const parts = a.split(/\s+/);
      return { last: parts.pop() ?? a, first: parts.join(" ") };
    });
}

const initials = (first: string) =>
  first
    .split(/[\s-]+/)
    .filter(Boolean)
    .map((w) => `${w[0]!.toLocaleUpperCase("tr-TR")}.`)
    .join(" ");

/** Metin içi alıntı: APA "(Soyadı, Yıl)", IEEE "[n]". */
export function citationText(src: Source | undefined, style: BibStyle, index: number, pages?: string): string {
  if (!src) return style === "ieee" ? "[?]" : "(Kaynak yok)";
  if (style === "ieee") return `[${index + 1}${pages ? `, s. ${pages}` : ""}]`;
  const ps = parseAuthors(src.author);
  const who = !ps.length ? src.title : ps.length === 1 ? ps[0].last : ps.length === 2 ? `${ps[0].last} ve ${ps[1].last}` : `${ps[0].last} vd.`;
  return `(${who}, ${src.year || "t.y."}${pages ? `, s. ${pages}` : ""})`;
}

/** Kaynakça girdisinin parçaları (italik bölüm ayrı: dergi/kitap adı). */
export interface BibEntry {
  before: string;
  italic: string;
  after: string;
}

function apaAuthors(ps: Person[]): string {
  const f = ps.map((p) => (p.first ? `${p.last}, ${initials(p.first)}` : p.last));
  if (f.length <= 1) return f[0] ?? "";
  return `${f.slice(0, -1).join(", ")} ve ${f[f.length - 1]}`;
}

function ieeeAuthors(ps: Person[]): string {
  const f = ps.map((p) => (p.first ? `${initials(p.first)} ${p.last}` : p.last));
  if (f.length <= 2) return f.join(" ve ");
  return `${f.slice(0, -1).join(", ")} ve ${f[f.length - 1]}`;
}

/** Kaynakça girdisi (APA 7 / IEEE temel biçim). */
export function bibEntry(s: Source, style: BibStyle, index: number): BibEntry {
  const ps = parseAuthors(s.author);
  if (style === "ieee") {
    const lead = `[${index + 1}] ${ieeeAuthors(ps)}${ps.length ? ", " : ""}`;
    switch (s.type) {
      case "article":
        return { before: `${lead}“${s.title},” `, italic: s.publisher, after: `${s.volume ? `, c. ${s.volume}` : ""}${s.issue ? `, sy. ${s.issue}` : ""}${s.pages ? `, ss. ${s.pages}` : ""}, ${s.year}.` };
      case "web":
        return { before: `${lead}“${s.title},” `, italic: s.publisher, after: `${s.url ? `. [Çevrimiçi]. Erişim: ${s.url}` : ""}${s.accessed ? ` (Erişim tarihi: ${s.accessed})` : ""}.` };
      default:
        return { before: lead, italic: s.title, after: `. ${s.city ? `${s.city}: ` : ""}${s.publisher}, ${s.year}.` };
    }
  }
  const who = ps.length ? apaAuthors(ps) : s.title;
  const year = ` (${s.year || "t.y."}). `;
  switch (s.type) {
    case "article":
      return { before: `${who}${year}${s.title}. `, italic: `${s.publisher}${s.volume ? `, ${s.volume}` : ""}`, after: `${s.issue ? `(${s.issue})` : ""}${s.pages ? `, ${s.pages}` : ""}.${s.url ? ` ${s.url}` : ""}` };
    case "web":
      return { before: `${who}${year}`, italic: s.title, after: `. ${s.publisher}${s.publisher ? ". " : ""}${s.url ?? ""}` };
    default:
      return { before: `${who}${year}`, italic: s.title, after: `. ${s.publisher}.` };
  }
}

/** Kaynakça sırası: APA yazar soyadına göre, IEEE alıntı sırasına göre. */
export function orderSources(sources: Source[], style: BibStyle, citedOrder: string[]): Source[] {
  if (style === "ieee") {
    const rank = (t: string) => {
      const i = citedOrder.indexOf(t);
      return i < 0 ? 1e6 : i;
    };
    return sources.slice().sort((a, b) => rank(a.tag) - rank(b.tag));
  }
  const key = (s: Source) => (parseAuthors(s.author)[0]?.last ?? s.title).toLocaleLowerCase("tr-TR");
  return sources.slice().sort((a, b) => key(a).localeCompare(key(b), "tr"));
}

/** Kaynak için benzersiz etiket (Soyadı + Yıl). */
export function sourceTag(s: Pick<Source, "author" | "title" | "year">, taken: string[]): string {
  const last = parseAuthors(s.author)[0]?.last ?? s.title.split(/\s+/)[0] ?? "Kaynak";
  const base = `${last.replace(/[^\p{L}\d]/gu, "").slice(0, 12) || "Kaynak"}${s.year.replace(/\D/g, "").slice(0, 4)}`;
  let tag = base;
  for (let i = 2; taken.includes(tag); i++) tag = `${base}_${i}`;
  return tag;
}
