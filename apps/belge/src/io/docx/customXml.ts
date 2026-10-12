// Özel XML parçaları (customXml/itemN.xml): kaynakça kaynakları (b:Sources, ECMA-376 §22.6) ve adres
// mektup birleştirme veri kaynağı (alıcı listesi; OfficeHere ad alanı). Diğer düzenleyiciler parçaları korur.
import type JSZip from "jszip";
import type { BibStyle, Source, SourceType } from "../../model/sources";
import type { MailMergeData } from "../../model/document";
import { NS, XML_HEAD, cleanText, descendants, esc, kid, parseXml } from "./xml";
import type { Rels } from "./partWriter";

export const BIB_NS = "http://schemas.openxmlformats.org/officeDocument/2006/bibliography";
export const MERGE_NS = "urn:officehere:belge:mailmerge";
const DS_NS = "http://schemas.openxmlformats.org/officeDocument/2006/customXml";
const REL_CUSTOM_XML = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml";
const REL_CUSTOM_XML_PROPS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXmlProps";

const TYPE_OUT: Record<SourceType, string> = { book: "Book", article: "JournalArticle", web: "InternetSite", report: "Report" };
const TYPE_IN: Record<string, SourceType> = { Book: "book", BookSection: "book", JournalArticle: "article", ArticleInAPeriodical: "article", ConferenceProceedings: "article", InternetSite: "web", DocumentFromInternetSite: "web", Report: "report" };

const el = (name: string, v: string | undefined) => (v ? `<b:${name}>${esc(cleanText(v))}</b:${name}>` : "");

/** Kararlı GUID (etiketten): aynı kaynak her kayıtta aynı kimliği alır. */
function guidOf(s: string): string {
  let h1 = 0x811c9dc5;
  let h2 = 0x01000193;
  for (let i = 0; i < s.length; i++) {
    h1 = Math.imul(h1 ^ s.charCodeAt(i), 16777619) >>> 0;
    h2 = Math.imul(h2 + s.charCodeAt(i), 2246822519) >>> 0;
  }
  const hex = (n: number) => n.toString(16).toUpperCase().padStart(8, "0");
  const a = hex(h1);
  const b = hex(h2);
  const c = hex((h1 ^ h2) >>> 0);
  const d = hex(Math.imul(h1, 31) >>> 0);
  return `{${a}-${b.slice(0, 4)}-4${b.slice(5, 8)}-8${c.slice(1, 4)}-${c.slice(4)}${d}}`;
}

function personsXml(author: string): string {
  const people = author
    .split(";")
    .map((a) => a.trim())
    .filter(Boolean)
    .map((a) => {
      const [last, first] = a.includes(",") ? a.split(",", 2).map((x) => x.trim()) : [a.split(/\s+/).pop() ?? a, a.split(/\s+/).slice(0, -1).join(" ")];
      return `<b:Person><b:Last>${esc(last)}</b:Last>${first ? `<b:First>${esc(first)}</b:First>` : ""}</b:Person>`;
    });
  return people.length ? `<b:Author><b:Author><b:NameList>${people.join("")}</b:NameList></b:Author></b:Author>` : "";
}

/** Kaynaklar → b:Sources. */
export function sourcesToXml(sources: Source[], style: BibStyle | undefined): string {
  const items = sources.map((s) => {
    const accessed = s.accessed ? /^(\d{1,2})[./-](\d{1,2})[./-](\d{4})$/.exec(s.accessed.trim()) : null;
    return (
      `<b:Source><b:Tag>${esc(s.tag)}</b:Tag><b:SourceType>${TYPE_OUT[s.type] ?? "Book"}</b:SourceType><b:Guid>${guidOf(s.tag)}</b:Guid>${personsXml(s.author)}` +
      el("Title", s.title) +
      el("Year", s.year) +
      (s.type === "article" ? el("JournalName", s.publisher) : s.type === "web" ? el("InternetSiteTitle", s.publisher) : el("Publisher", s.publisher)) +
      el("City", s.city) +
      el("Volume", s.volume) +
      el("Issue", s.issue) +
      el("Pages", s.pages) +
      el("URL", s.url) +
      (accessed ? `<b:YearAccessed>${accessed[3]}</b:YearAccessed><b:MonthAccessed>${Number(accessed[2])}</b:MonthAccessed><b:DayAccessed>${Number(accessed[1])}</b:DayAccessed>` : "") +
      `</b:Source>`
    );
  });
  return `${XML_HEAD}<b:Sources xmlns:b="${BIB_NS}" xmlns="${BIB_NS}" StyleName="${style === "ieee" ? "IEEE" : "APA"}">${items.join("")}</b:Sources>`;
}

/** b:Sources → kaynaklar. */
export function sourcesFromXml(doc: Document): { sources: Source[]; style?: BibStyle } {
  const root = doc.documentElement;
  const text = (e: Element, n: string) => kid(e, n)?.textContent?.trim() || undefined;
  const sources: Source[] = [];
  for (const s of descendants(root, "Source")) {
    if (s.parentElement !== root) continue;
    const tag = text(s, "Tag");
    if (!tag) continue;
    const persons = descendants(kid(s, "Author") ?? s, "Person").map((p) => {
      const last = text(p, "Last") ?? "";
      const first = [text(p, "First"), text(p, "Middle")].filter(Boolean).join(" ");
      return first ? `${last}, ${first}` : last;
    });
    const corp = descendants(kid(s, "Author") ?? s, "Corporate")[0]?.textContent?.trim();
    const type = TYPE_IN[text(s, "SourceType") ?? ""] ?? "book";
    const y = text(s, "YearAccessed");
    const m = text(s, "MonthAccessed");
    const d = text(s, "DayAccessed");
    sources.push({
      tag,
      type,
      author: persons.length ? persons.join("; ") : (corp ?? ""),
      title: text(s, "Title") ?? "",
      year: text(s, "Year") ?? "",
      publisher: (type === "article" ? (text(s, "JournalName") ?? text(s, "PeriodicalTitle")) : type === "web" ? (text(s, "InternetSiteTitle") ?? text(s, "ProductionCompany")) : text(s, "Publisher")) ?? "",
      city: text(s, "City"),
      volume: text(s, "Volume"),
      issue: text(s, "Issue"),
      pages: text(s, "Pages"),
      url: text(s, "URL"),
      accessed: y && m && d ? `${d.padStart(2, "0")}.${m.padStart(2, "0")}.${y}` : undefined,
    });
  }
  const sn = (root.getAttribute("StyleName") ?? root.getAttribute("SelectedStyle") ?? "").toLowerCase();
  return { sources, style: sn.includes("ieee") ? "ieee" : sn.includes("apa") ? "apa" : undefined };
}

/** Birleştirme veri kaynağı → OfficeHere özel XML'i. */
export function mergeToXml(m: MailMergeData): string {
  const f = m.list.fields.map((x) => `<oh:field>${esc(cleanText(x))}</oh:field>`).join("");
  const r = m.list.records
    .map((rec, i) => `<oh:record${m.excluded.includes(i) ? ' excluded="1"' : ""}>${m.list.fields.map((x) => `<oh:v>${esc(cleanText(rec[x] ?? ""))}</oh:v>`).join("")}</oh:record>`)
    .join("");
  return `${XML_HEAD}<oh:mailMerge xmlns:oh="${MERGE_NS}" source="${esc(m.list.name)}" type="formLetters">${f}${r}</oh:mailMerge>`;
}

export function mergeFromXml(doc: Document): MailMergeData | null {
  const root = doc.documentElement;
  if (root.localName !== "mailMerge" || root.namespaceURI !== MERGE_NS) return null;
  const fields = descendants(root, "field").map((e) => e.textContent ?? "");
  if (!fields.length) return null;
  const excluded: number[] = [];
  const records = descendants(root, "record").map((r, i) => {
    if (r.getAttribute("excluded") === "1") excluded.push(i);
    const vals = descendants(r, "v").map((e) => e.textContent ?? "");
    return Object.fromEntries(fields.map((f, k) => [f, vals[k] ?? ""]));
  });
  return { list: { name: root.getAttribute("source") ?? "Alıcılar", fields, records }, excluded };
}

/** Özel XML parçalarını pakete yazar; içerik türü geçersiz kılmalarını döner. */
export function writeCustomXml(zip: JSZip, docRels: Rels, parts: { xml: string; schema: string }[]): string {
  let overrides = "";
  parts.forEach((p, i) => {
    const n = i + 1;
    zip.file(`customXml/item${n}.xml`, p.xml);
    zip.file(
      `customXml/itemProps${n}.xml`,
      `${XML_HEAD}<ds:datastoreItem ds:itemID="${guidOf(`officehere-item-${n}-${p.schema}`)}" xmlns:ds="${DS_NS}"><ds:schemaRefs><ds:schemaRef ds:uri="${p.schema}"/></ds:schemaRefs></ds:datastoreItem>`,
    );
    zip.file(`customXml/_rels/item${n}.xml.rels`, `${XML_HEAD}<Relationships xmlns="${NS.rel}"><Relationship Id="rId1" Type="${REL_CUSTOM_XML_PROPS}" Target="itemProps${n}.xml"/></Relationships>`);
    docRels.add(REL_CUSTOM_XML, `../customXml/item${n}.xml`);
    overrides += `<Override PartName="/customXml/itemProps${n}.xml" ContentType="application/vnd.openxmlformats-officedocument.customXmlProperties+xml"/>`;
  });
  return overrides;
}

/** Paketteki özel XML parçalarından kaynaklar ve birleştirme veri kaynağı. */
export async function readCustomXml(zip: JSZip): Promise<{ sources?: Source[]; bibStyle?: BibStyle; mailMerge?: MailMergeData }> {
  const out: { sources?: Source[]; bibStyle?: BibStyle; mailMerge?: MailMergeData } = {};
  const names = Object.keys(zip.files).filter((n) => /^customXml\/item\d+\.xml$/i.test(n));
  for (const name of names) {
    let doc: Document;
    try {
      doc = parseXml(await zip.file(name)!.async("string"));
    } catch {
      continue;
    }
    const root = doc.documentElement;
    if (root?.localName === "Sources" && root.namespaceURI === BIB_NS) {
      const r = sourcesFromXml(doc);
      if (r.sources.length) out.sources = r.sources;
      if (r.style) out.bibStyle = r.style;
    } else if (root?.localName === "mailMerge") {
      const m = mergeFromXml(doc);
      if (m) out.mailMerge = m;
    }
  }
  return out;
}
