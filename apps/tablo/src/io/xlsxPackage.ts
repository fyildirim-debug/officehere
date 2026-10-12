/**
 * XLSX paketini okumadan önce sağlamlaştırır:
 * - Mutlak ilişki hedeflerini (Target="/xl/...", openpyxl vb. üretir) kaynak
 *   parçaya göre göreli yola çevirir (ECMA-376 Bölüm 2: ikisi de geçerlidir,
 *   ancak bazı okuyucular yalnız göreli yolu çözer).
 * - Okunamayan isteğe bağlı parçaları (açıklamalar, çizimler, tablolar…)
 *   ilişkilerden çıkarıp kalanının açılabilmesini sağlar.
 */
import JSZip from "jszip";

/** İsteğe bağlı parça türleri: ilişki türü sonekine göre, kullanıcıya gösterilecek adla. */
export const OPTIONAL_PARTS: { kind: string; label: string; types: string[] }[] = [
  { kind: "comments", label: "açıklamalar (notlar)", types: ["/comments", "/vmlDrawing", "/threadedComment", "/person"] },
  { kind: "drawing", label: "resimler ve çizimler", types: ["/drawing"] },
  { kind: "table", label: "tablolar", types: ["/table"] },
  { kind: "pivot", label: "özet tablolar", types: ["/pivotTable", "/pivotCacheDefinition"] },
  { kind: "other", label: "diğer nesneler", types: ["/ctrlProp", "/oleObject", "/control", "/printerSettings", "/customProperty", "/slicer", "/timeline", "/webExtension"] },
];

function sourceOfRels(relsPath: string): string {
  // "xl/worksheets/_rels/sheet1.xml.rels" → "xl/worksheets/sheet1.xml"; "_rels/.rels" → ""
  const m = /^(.*?)_rels\/(.*)\.rels$/.exec(relsPath);
  if (!m) return "";
  return m[1] + m[2];
}

/** Hedefi kaynak parçanın klasörüne göre göreli yola çevirir. */
export function relativeTarget(source: string, absTarget: string): string {
  const target = absTarget.replace(/^\/+/, "");
  const from = source.split("/").slice(0, -1);
  const to = target.split("/");
  let i = 0;
  while (i < from.length && i < to.length - 1 && from[i] === to[i]) i++;
  const up = from.length - i;
  return [...Array<string>(up).fill(".."), ...to.slice(i)].join("/");
}

function relAttr(tag: string, name: string): string | undefined {
  return new RegExp(`\\b${name}="([^"]*)"`).exec(tag)?.[1];
}

/** .rels içindeki mutlak hedefleri göreli yapar. */
export function normalizeRelsXml(relsPath: string, xml: string): string {
  const source = sourceOfRels(relsPath);
  return xml.replace(/<Relationship\b[^>]*>/g, (tag) => {
    if (relAttr(tag, "TargetMode") === "External") return tag;
    const target = relAttr(tag, "Target");
    if (!target || !target.startsWith("/")) return tag;
    const rel = source ? relativeTarget(source, target) : target.replace(/^\/+/, "");
    return tag.replace(`Target="${target}"`, `Target="${rel}"`);
  });
}

/** Bilinen ad alanlarının alışılmış önekleri (önekli öğe adı bekleyen okuyucular için). */
const NS_PREFIX: Record<string, string> = {
  "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing": "xdr",
  "http://schemas.openxmlformats.org/drawingml/2006/main": "a",
  "http://schemas.openxmlformats.org/drawingml/2006/chart": "c",
  "http://schemas.openxmlformats.org/drawingml/2006/picture": "pic",
};

/**
 * Varsayılan ad alanıyla (öneksiz) yazılmış öğelere alışılmış öneki ekler:
 * <wsDr xmlns="…spreadsheetDrawing"><from> → <xdr:wsDr xmlns:xdr="…"><xdr:from>.
 */
export function prefixDefaultNamespaces(xml: string): string {
  const stack: (string | undefined)[] = [];
  let changed = false;
  const out = xml.replace(/<(\/?)([A-Za-z_][\w.-]*(?::[A-Za-z_][\w.-]*)?)((?:\s[^<>]*?)?)(\/?)>/g, (m, close: string, name: string, attrs: string, selfClose: string) => {
    if (close) {
      const prefix = stack.pop();
      if (name.includes(":") || !prefix) return m;
      changed = true;
      return `</${prefix}:${name}>`;
    }
    const declared = /\sxmlns="([^"]*)"/.exec(attrs)?.[1];
    const parentPrefix = stack.length ? stack[stack.length - 1] : undefined;
    let prefix: string | undefined;
    let newAttrs = attrs;
    if (declared !== undefined) {
      prefix = NS_PREFIX[declared];
      if (prefix) {
        // xmlns="…" → xmlns:önek="…" (önek zaten tanımlıysa yalnız kaldır)
        newAttrs = new RegExp(`\\sxmlns:${prefix}="`).test(attrs) ? attrs.replace(/\sxmlns="[^"]*"/, "") : attrs.replace(/\sxmlns="/, ` xmlns:${prefix}="`);
      }
    } else prefix = parentPrefix;
    const own = name.includes(":") ? undefined : prefix;
    if (!selfClose) stack.push(declared !== undefined ? prefix : parentPrefix);
    if (!own) return newAttrs === attrs ? m : `<${name}${newAttrs}${selfClose}>`;
    changed = true;
    return `<${own}:${name}${newAttrs}${selfClose}>`;
  });
  return changed ? out : xml;
}

const REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/";

/** Okuyucunun beklediği parça adları (tür → ad deseni ve yeni ad kalıbı). */
const PART_NAMES: { type: string; ok: RegExp; make: (n: number) => string }[] = [
  { type: "worksheet", ok: /^xl\/worksheets\/sheet\d+\.xml$/, make: (n) => `xl/worksheets/sheet${n}.xml` },
  { type: "comments", ok: /^xl\/comments\d+\.xml$/, make: (n) => `xl/comments${n}.xml` },
  { type: "drawing", ok: /^xl\/drawings\/[a-zA-Z0-9]+\.xml$/, make: (n) => `xl/drawings/drawing${n}.xml` },
  { type: "table", ok: /^xl\/tables\/table\d+\.xml$/, make: (n) => `xl/tables/table${n}.xml` },
];

async function renamePart(zip: JSZip, from: string, to: string): Promise<void> {
  const f = zip.file(from);
  if (!f) return;
  zip.file(to, await f.async("uint8array"));
  zip.remove(from);
  const rels = (p: string) => `${p.slice(0, p.lastIndexOf("/") + 1)}_rels/${p.slice(p.lastIndexOf("/") + 1)}.rels`;
  const oldRels = zip.file(rels(from));
  if (oldRels) {
    // Parçanın kendi ilişkileri: göreli hedefler yeni konuma göre yeniden yazılır
    const xml = await oldRels.async("string");
    const fixed = xml.replace(/<Relationship\b[^>]*>/g, (tag) => {
      const t = relAttr(tag, "Target");
      if (!t || relAttr(tag, "TargetMode") === "External") return tag;
      const abs = resolvePath(from, t);
      return tag.replace(`Target="${t}"`, `Target="${relativeTarget(to, abs)}"`);
    });
    zip.file(rels(to), fixed);
    zip.remove(rels(from));
  }
  // Bu parçayı gösteren tüm ilişkiler
  for (const name of Object.keys(zip.files)) {
    if (!name.endsWith(".rels")) continue;
    const src = sourceOfRels(name);
    const xml = await zip.file(name)!.async("string");
    const next = xml.replace(/<Relationship\b[^>]*>/g, (tag) => {
      const t = relAttr(tag, "Target");
      if (!t || relAttr(tag, "TargetMode") === "External" || resolvePath(src, t) !== from) return tag;
      return tag.replace(`Target="${t}"`, `Target="${src ? relativeTarget(src, to) : to}"`);
    });
    if (next !== xml) zip.file(name, next);
  }
  const ct = zip.file("[Content_Types].xml");
  if (ct) {
    const xml = await ct.async("string");
    zip.file("[Content_Types].xml", xml.split(`PartName="/${from}"`).join(`PartName="/${to}"`));
  }
}

/** Kaynak parçaya göre göreli ya da mutlak hedefi paket yoluna çevirir. */
export function resolvePath(source: string, target: string): string {
  if (target.startsWith("/")) return target.slice(1);
  const parts = source.split("/").slice(0, -1);
  for (const seg of target.split("/")) {
    if (seg === "..") parts.pop();
    else if (seg !== "." && seg !== "") parts.push(seg);
  }
  return parts.join("/");
}

/**
 * Paketi normalleştirir: mutlak hedefler göreli olur, parçalar alışılmış
 * adlarına taşınır, öneksiz DrawingML öğelerine önek eklenir ve VML not
 * çizimleri (yalnız görünüm bilgisi taşır) okuma için ayrılır.
 * Değişiklik yoksa özgün veriyi döndürür.
 */
export async function normalizePackage(data: ArrayBuffer): Promise<ArrayBuffer> {
  const zip = await JSZip.loadAsync(data);
  let changed = false;
  for (const name of Object.keys(zip.files)) {
    if (!name.endsWith(".rels") || zip.files[name].dir) continue;
    const xml = await zip.file(name)!.async("string");
    const next = normalizeRelsXml(name, xml);
    if (next !== xml) {
      zip.file(name, next);
      changed = true;
    }
  }
  // Parça adları
  for (const spec of PART_NAMES) {
    const targets: string[] = [];
    for (const name of Object.keys(zip.files)) {
      if (!name.endsWith(".rels")) continue;
      const src = sourceOfRels(name);
      const xml = await zip.file(name)!.async("string");
      for (const m of xml.matchAll(/<Relationship\b[^>]*>/g)) {
        if (relAttr(m[0], "Type") !== REL_TYPE + spec.type || relAttr(m[0], "TargetMode") === "External") continue;
        const p = resolvePath(src, relAttr(m[0], "Target") ?? "");
        if (!targets.includes(p)) targets.push(p);
      }
    }
    const bad = targets.filter((p) => !spec.ok.test(p) && zip.file(p));
    if (!bad.length) continue;
    let n = 1;
    for (const p of bad) {
      while (zip.file(spec.make(n))) n++;
      await renamePart(zip, p, spec.make(n));
      changed = true;
    }
  }
  // Öneksiz çizim/grafik XML'leri
  for (const name of Object.keys(zip.files)) {
    if (!/^xl\/(drawings|charts)\/[^/]+\.xml$/.test(name)) continue;
    const xml = await zip.file(name)!.async("string");
    const next = prefixDefaultNamespaces(xml);
    if (next !== xml) {
      zip.file(name, next);
      changed = true;
    }
  }
  // VML not çizimleri: yalnız Excel biçiminde (v:/x: önekli) olanlar kalır; diğerleri ilişkiden çıkarılır
  for (const name of Object.keys(zip.files)) {
    if (!/^xl\/worksheets\/_rels\/[^/]+\.rels$/.test(name)) continue;
    const src = sourceOfRels(name);
    const xml = await zip.file(name)!.async("string");
    const dropped: string[] = [];
    const next = await replaceAsync(xml, /<Relationship\b[^>]*>/g, async (tag) => {
      if (relAttr(tag, "Type") !== REL_TYPE + "vmlDrawing") return tag;
      const p = resolvePath(src, relAttr(tag, "Target") ?? "");
      const vml = zip.file(p) ? await zip.file(p)!.async("string") : "";
      if (/<v:shape\b/.test(vml) && /<x:ClientData\b/.test(vml) && /^xl\/drawings\/vmlDrawing\d+\.vml$/.test(p)) return tag;
      const id = relAttr(tag, "Id");
      if (id) dropped.push(id);
      return "";
    });
    if (!dropped.length) continue;
    zip.file(name, next);
    const sf = zip.file(src);
    if (sf) {
      let sx = await sf.async("string");
      for (const id of dropped) sx = sx.replace(new RegExp(`<(\\w+:)?legacyDrawing\\b[^>]*r:id="${id}"[^>]*/>`, "g"), "");
      zip.file(src, sx);
    }
    changed = true;
  }
  if (!changed) return data;
  return zip.generateAsync({ type: "arraybuffer" });
}

async function replaceAsync(s: string, re: RegExp, fn: (m: string) => Promise<string>): Promise<string> {
  const parts: Promise<string>[] = [];
  let last = 0;
  for (const m of s.matchAll(re)) {
    parts.push(Promise.resolve(s.slice(last, m.index)));
    parts.push(fn(m[0]));
    last = m.index! + m[0].length;
  }
  parts.push(Promise.resolve(s.slice(last)));
  return (await Promise.all(parts)).join("");
}

/** Sayfa ilişkilerinden verilen türlerdeki parçaları çıkarır (ilgili XML öğeleriyle birlikte). */
export async function stripOptionalParts(data: ArrayBuffer, kinds: string[]): Promise<ArrayBuffer> {
  const zip = await JSZip.loadAsync(data);
  const types = OPTIONAL_PARTS.filter((p) => kinds.includes(p.kind)).flatMap((p) => p.types);
  const isDropped = (type: string | undefined) => !!type && types.some((t) => type.endsWith(t));
  for (const name of Object.keys(zip.files)) {
    if (!/^xl\/worksheets\/_rels\/[^/]+\.rels$/.test(name)) continue;
    const xml = await zip.file(name)!.async("string");
    const dropped = new Set<string>();
    const next = xml.replace(/<Relationship\b[^>]*>/g, (tag) => {
      if (!isDropped(relAttr(tag, "Type"))) return tag;
      const id = relAttr(tag, "Id");
      if (id) dropped.add(id);
      return "";
    });
    if (!dropped.size) continue;
    zip.file(name, next);
    const sheetPath = sourceOfRels(name);
    const sheetFile = zip.file(sheetPath);
    if (!sheetFile) continue;
    let sx = await sheetFile.async("string");
    // Kaldırılan ilişkilere başvuran öğeler (drawing, legacyDrawing, tableParts içindeki tablePart…)
    for (const id of dropped) {
      sx = sx.replace(new RegExp(`<(\\w+:)?(drawing|legacyDrawing|legacyDrawingHF|tablePart|control|oleObject)\\b[^>]*r:id="${id}"[^>]*/>`, "g"), "");
    }
    sx = sx.replace(/<tableParts\b[^>]*count="\d+"[^>]*>\s*<\/tableParts>/g, "").replace(/<tableParts\b[^>]*\/>/g, "");
    sx = sx.replace(/<tableParts\b([^>]*)count="\d+"([^>]*)>((?:\s*<tablePart\b[^>]*\/>)+)\s*<\/tableParts>/g, (_m, a: string, b: string, inner: string) => {
      const n = (inner.match(/<tablePart\b/g) ?? []).length;
      return `<tableParts${a}count="${n}"${b}>${inner}</tableParts>`;
    });
    zip.file(sheetPath, sx);
  }
  return zip.generateAsync({ type: "arraybuffer" });
}

/** Hücre notlarını (comments parçaları) doğrudan okur: düz <t> ya da biçimli <r><t> metinleri. */
export async function readCommentsFromData(data: ArrayBuffer): Promise<Map<string, { ref: string; text: string; author?: string }[]>> {
  const zip = await JSZip.loadAsync(data);
  const out = new Map<string, { ref: string; text: string; author?: string }[]>();
  const wbXml = await zip.file("xl/workbook.xml")?.async("string");
  const wbRels = await zip.file("xl/_rels/workbook.xml.rels")?.async("string");
  if (!wbXml || !wbRels) return out;
  const relTarget = new Map<string, string>();
  for (const m of wbRels.matchAll(/<Relationship\b[^>]*>/g)) {
    const id = relAttr(m[0], "Id");
    const t = relAttr(m[0], "Target");
    if (id && t) relTarget.set(id, resolvePath("xl/workbook.xml", t));
  }
  const unesc = (s: string) => s.replace(/&lt;/g, "<").replace(/&gt;/g, ">").replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&#(\d+);/g, (_m, n: string) => String.fromCharCode(Number(n))).replace(/&amp;/g, "&");
  for (const m of wbXml.matchAll(/<(?:\w+:)?sheet\b[^>]*>/g)) {
    const name = relAttr(m[0], "name");
    const rid = /\br:id="([^"]+)"/.exec(m[0])?.[1];
    const sheetPath = rid ? relTarget.get(rid) : undefined;
    if (!name || !sheetPath) continue;
    const rp = `${sheetPath.slice(0, sheetPath.lastIndexOf("/") + 1)}_rels/${sheetPath.slice(sheetPath.lastIndexOf("/") + 1)}.rels`;
    const rels = await zip.file(rp)?.async("string");
    if (!rels) continue;
    for (const r of rels.matchAll(/<Relationship\b[^>]*>/g)) {
      if (!relAttr(r[0], "Type")?.endsWith("/comments")) continue;
      const cx = await zip.file(resolvePath(sheetPath, relAttr(r[0], "Target") ?? ""))?.async("string");
      if (!cx) continue;
      const authors = [...(/<(?:\w+:)?authors>([\s\S]*?)<\/(?:\w+:)?authors>/.exec(cx)?.[1] ?? "").matchAll(/<(?:\w+:)?author>([^<]*)<\/(?:\w+:)?author>/g)].map((a) => unesc(a[1]));
      const list: { ref: string; text: string; author?: string }[] = [];
      for (const c of cx.matchAll(/<(?:\w+:)?comment\b([^>]*)>([\s\S]*?)<\/(?:\w+:)?comment>/g)) {
        const ref = relAttr(c[1], "ref");
        if (!ref) continue;
        const text = [...c[2].matchAll(/<(?:\w+:)?t(?:\s[^>]*)?>([^<]*)<\/(?:\w+:)?t>/g)].map((t) => unesc(t[1])).join("");
        const ai = Number(relAttr(c[1], "authorId") ?? -1);
        list.push({ ref, text, author: authors[ai] });
      }
      out.set(unesc(name), [...(out.get(unesc(name)) ?? []), ...list]);
    }
  }
  return out;
}

/** workbook.xml içindeki tanımlı adlar (sabit ve formül adları dahil). */
export async function readDefinedNamesFromData(data: ArrayBuffer): Promise<{ name: string; formula: string; localSheetId?: number }[]> {
  const zip = await JSZip.loadAsync(data);
  const xml = await zip.file("xl/workbook.xml")?.async("string");
  if (!xml) return [];
  const out: { name: string; formula: string; localSheetId?: number }[] = [];
  const unesc = (s: string) => s.replace(/&lt;/g, "<").replace(/&gt;/g, ">").replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&amp;/g, "&");
  for (const m of xml.matchAll(/<(?:\w+:)?definedName\b([^>]*)>([^<]*)<\/(?:\w+:)?definedName>/g)) {
    const name = relAttr(m[1], "name");
    if (!name || name.startsWith("_xlnm") || relAttr(m[1], "hidden") === "1") continue;
    const local = relAttr(m[1], "localSheetId");
    out.push({ name: unesc(name), formula: unesc(m[2]).trim(), localSheetId: local !== undefined ? Number(local) : undefined });
  }
  return out;
}
