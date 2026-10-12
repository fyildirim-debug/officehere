/**
 * OOXML parçalarının üst düzey öğelerini sıra tablosuna göre yöneten küçük
 * ağaç modeli. Çalışma sayfası ve çalışma kitabı XML'ine öğe eklerken dize
 * araması yerine bu model kullanılır: öğe eklenir/değiştirilir/kaldırılır ve
 * yazılırken ECMA-376 Bölüm 1 şemasındaki sıraya (CT_Worksheet §18.3.1.99,
 * CT_Workbook §18.2.27) göre dizilir. Bilinmeyen öğeler (ör. mc:AlternateContent)
 * kendinden önceki bilinen öğenin hemen arkasında kalır.
 */

/** CT_Worksheet alt öğe sırası. */
export const WORKSHEET_ORDER = [
  "sheetPr", "dimension", "sheetViews", "sheetFormatPr", "cols", "sheetData", "sheetCalcPr", "sheetProtection",
  "protectedRanges", "scenarios", "autoFilter", "sortState", "dataConsolidate", "customSheetViews", "mergeCells",
  "phoneticPr", "conditionalFormatting", "dataValidations", "hyperlinks", "printOptions", "pageMargins", "pageSetup",
  "headerFooter", "rowBreaks", "colBreaks", "customProperties", "cellWatches", "ignoredErrors", "smartTags", "drawing",
  "legacyDrawing", "legacyDrawingHF", "drawingHF", "picture", "oleObjects", "controls", "webPublishItems", "tableParts",
  "extLst",
] as const;

/** CT_Workbook alt öğe sırası. */
export const WORKBOOK_ORDER = [
  "fileVersion", "fileSharing", "workbookPr", "workbookProtection", "bookViews", "sheets", "functionGroups",
  "externalReferences", "definedNames", "calcPr", "oleSize", "customWorkbookViews", "pivotCaches", "smartTagPr",
  "smartTagTypes", "webPublishing", "fileRecoveryPr", "webPublishObjects", "extLst",
] as const;

/**
 * Çalışma sayfası extLst içindeki bilinen uzantıların alışılmış sırası
 * (koşullu biçimler, doğrulamalar, mini grafikler, dilimleyiciler…).
 */
export const WORKSHEET_EXT_ORDER = [
  "{78C0D931-6437-407d-A8EE-F0AAD7539E65}", // x14:conditionalFormattings
  "{CCE6A557-97BC-4b89-ADB6-D9C93CAAB3DF}", // x14:dataValidations
  "{05C60535-1F16-4fd2-B633-F4F36F0B64E0}", // x14:sparklineGroups
  "{A8765BA9-456A-4dab-B4F3-ACF838C121DE}", // x14:slicerList
  "{3A4CF648-6AED-40f4-86FF-DC5316D8AED3}", // x15:slicerList
];

export interface XmlChild {
  /** Önek olmadan yerel ad (ör. "sheetData"). */
  name: string;
  xml: string;
}

/** Öğenin yerel adı ("x:sheetData" → "sheetData"). */
export function localName(qname: string): string {
  const i = qname.indexOf(":");
  return i >= 0 ? qname.slice(i + 1) : qname;
}

const NAME_RE = /^<\/?([A-Za-z_][\w.\-]*(?::[A-Za-z_][\w.\-]*)?)/;

/** Bir işaretleme belirtecinin bitişi (tırnak içindeki ">" atlanır). */
function tagEnd(xml: string, start: number): number {
  if (xml.startsWith("<!--", start)) {
    const e = xml.indexOf("-->", start + 4);
    return e < 0 ? xml.length : e + 3;
  }
  if (xml.startsWith("<![CDATA[", start)) {
    const e = xml.indexOf("]]>", start + 9);
    return e < 0 ? xml.length : e + 3;
  }
  if (xml.startsWith("<?", start)) {
    const e = xml.indexOf("?>", start + 2);
    return e < 0 ? xml.length : e + 2;
  }
  let q = "";
  for (let i = start + 1; i < xml.length; i++) {
    const ch = xml[i];
    if (q) {
      if (ch === q) q = "";
    } else if (ch === '"' || ch === "'") q = ch;
    else if (ch === ">") return i + 1;
  }
  return xml.length;
}

interface Parsed {
  prolog: string;
  open: string;
  rootName: string;
  children: XmlChild[];
  close: string;
  epilog: string;
}

/** Kök öğeyi ve doğrudan çocuklarını ayrıştırır. */
function parseRoot(xml: string): Parsed {
  let i = 0;
  let prolog = "";
  // Kök öncesi: XML bildirimi, yorumlar, DOCTYPE
  for (;;) {
    const lt = xml.indexOf("<", i);
    if (lt < 0) throw new Error("XML kök öğesi bulunamadı");
    if (xml.startsWith("<?", lt) || xml.startsWith("<!", lt)) {
      const e = tagEnd(xml, lt);
      prolog += xml.slice(i, e);
      i = e;
      continue;
    }
    prolog += xml.slice(i, lt);
    i = lt;
    break;
  }
  const openEnd = tagEnd(xml, i);
  const open = xml.slice(i, openEnd);
  const rootName = NAME_RE.exec(open)?.[1] ?? "";
  if (open.endsWith("/>")) return { prolog, open: open.slice(0, -2) + ">", rootName, children: [], close: `</${rootName}>`, epilog: xml.slice(openEnd) };
  const children: XmlChild[] = [];
  let depth = 0;
  let childStart = -1;
  let pos = openEnd;
  for (;;) {
    const lt = xml.indexOf("<", pos);
    if (lt < 0) throw new Error(`Kapanmamış öğe: ${rootName}`);
    const end = tagEnd(xml, lt);
    pos = end;
    const c1 = xml[lt + 1];
    // Yorum, işleme yönergesi, CDATA
    if (c1 === "!" || c1 === "?") continue;
    if (c1 === "/") {
      if (depth === 0) {
        return { prolog, open, rootName, children, close: xml.slice(lt, end), epilog: xml.slice(end) };
      }
      depth--;
      if (depth === 0) {
        const raw = xml.slice(childStart, end);
        children.push({ name: localName(NAME_RE.exec(raw)?.[1] ?? ""), xml: raw });
      }
      continue;
    }
    const selfClose = xml[end - 2] === "/";
    if (depth === 0) childStart = lt;
    if (selfClose) {
      if (depth === 0) {
        const tok = xml.slice(lt, end);
        children.push({ name: localName(NAME_RE.exec(tok)?.[1] ?? ""), xml: tok });
      }
    } else depth++;
  }
}

/** Sıra tablosuna göre dizilen öğe kümesi (kök öğe ve doğrudan çocukları). */
export class XmlPart {
  prolog: string;
  open: string;
  readonly rootName: string;
  children: XmlChild[];
  private close: string;
  private epilog: string;
  private order: readonly string[];

  constructor(xml: string, order: readonly string[] = []) {
    const p = parseRoot(xml);
    this.prolog = p.prolog;
    this.open = p.open;
    this.rootName = p.rootName;
    this.children = p.children;
    this.close = p.close;
    this.epilog = p.epilog;
    this.order = order;
  }

  /** Verilen adlı ilk öğe. */
  get(name: string): string | undefined {
    return this.children.find((c) => c.name === name)?.xml;
  }

  all(name: string): string[] {
    return this.children.filter((c) => c.name === name).map((c) => c.xml);
  }

  has(name: string): boolean {
    return this.children.some((c) => c.name === name);
  }

  /** Öğeyi yerinde değiştirir; yoksa ekler. null kaldırır. Aynı addan birden çok varsa hepsinin yerine tek öğe kalır. */
  set(name: string, xml: string | null): void {
    const first = this.children.findIndex((c) => c.name === name);
    if (first < 0) {
      if (xml !== null) this.children.push({ name, xml });
      return;
    }
    const out: XmlChild[] = [];
    this.children.forEach((c, j) => {
      if (c.name !== name) out.push(c);
      else if (j === first && xml !== null) out.push({ name, xml });
    });
    this.children = out;
  }

  /** Aynı addan yeni bir öğe ekler (ör. birden çok conditionalFormatting). */
  add(name: string, xml: string): void {
    this.children.push({ name, xml });
  }

  /** Her öğeyi dönüştürür (null dönerse kaldırılır). */
  map(name: string, fn: (xml: string) => string | null): void {
    const out: XmlChild[] = [];
    for (const c of this.children) {
      if (c.name !== name) {
        out.push(c);
        continue;
      }
      const x = fn(c.xml);
      if (x !== null) out.push({ name, xml: x });
    }
    this.children = out;
  }

  /** Kök öğede ad alanı bildirimi yoksa ekler. */
  ensureNamespace(prefix: string, uri: string): void {
    if (new RegExp(`\\sxmlns:${prefix}=`).test(this.open)) return;
    this.open = this.open.replace(/^<([^\s>/]+)/, `<$1 xmlns:${prefix}="${uri}"`);
  }

  /** Şemadaki sırayla üst düzey öğe adları (yazılacak sırada). */
  orderedChildren(): XmlChild[] {
    const idx = new Map<string, number>();
    this.order.forEach((n, i) => idx.set(n, i * 2));
    let prev = -1;
    const keyed = this.children.map((c, i) => {
      const k = idx.get(c.name);
      if (k !== undefined) prev = k;
      return { c, i, key: k ?? prev + 1 };
    });
    keyed.sort((a, b) => a.key - b.key || a.i - b.i);
    return keyed.map((k) => k.c);
  }

  toString(): string {
    return `${this.prolog}${this.open}${this.orderedChildren().map((c) => c.xml).join("")}${this.close}${this.epilog}`;
  }
}

/** Öğenin doğrudan çocukları (ör. extLst içindeki ext'ler). */
export function childElements(xml: string): XmlChild[] {
  return parseRoot(xml).children;
}

/** Öğenin açılış etiketi ve çocuklarıyla yeniden kurulması. */
export function rebuildElement(xml: string, children: string[]): string {
  const p = parseRoot(xml);
  return `${p.open}${children.join("")}${p.close}`;
}

/** Bir öğenin özniteliği. */
export function attrOf(tag: string, name: string): string | undefined {
  const open = tag.slice(0, tagEnd(tag, 0));
  return new RegExp(`\\s${name}="([^"]*)"`).exec(open)?.[1];
}

/** Üst düzey öğe adlarının şema sırasına uyup uymadığı (doğrulama için). İlk aykırı çift döner. */
export function orderViolation(names: string[], order: readonly string[]): [string, string] | null {
  let last = -1;
  let lastName = "";
  for (const n of names) {
    const i = order.indexOf(n);
    if (i < 0) continue;
    if (i < last) return [lastName, n];
    last = i;
    lastName = n;
  }
  return null;
}

/** Bir parçanın üst düzey öğe adları (yazıldığı sırayla). */
export function topLevelNames(xml: string): string[] {
  return parseRoot(xml).children.map((c) => c.name);
}

/** extLst içine verilen uri'li ext öğesini ekler/değiştirir; ext'ler bilinen sıraya göre dizilir. */
export function upsertExt(part: XmlPart, uri: string, extXml: string | null, extOrder: readonly string[] = WORKSHEET_EXT_ORDER): void {
  const cur = part.get("extLst");
  let exts = cur ? childElements(cur).map((c) => c.xml) : [];
  exts = exts.filter((e) => attrOf(e, "uri")?.toUpperCase() !== uri.toUpperCase());
  if (extXml) exts.push(extXml);
  if (!exts.length) {
    part.set("extLst", null);
    return;
  }
  const rank = (e: string) => {
    const u = attrOf(e, "uri")?.toUpperCase() ?? "";
    const i = extOrder.findIndex((x) => x.toUpperCase() === u);
    return i < 0 ? extOrder.length : i;
  };
  exts = exts.map((e, i) => ({ e, i })).sort((a, b) => rank(a.e) - rank(b.e) || a.i - b.i).map((x) => x.e);
  part.set("extLst", `<extLst>${exts.join("")}</extLst>`);
}

/** extLst'teki verilen uri'li ext öğesi. */
export function getExt(part: XmlPart, uri: string): string | undefined {
  const cur = part.get("extLst");
  if (!cur) return undefined;
  return childElements(cur).find((c) => attrOf(c.xml, "uri")?.toUpperCase() === uri.toUpperCase())?.xml;
}
