// Geçiş katmanı: yazıcının ürettiği paket + açılan özgün paket → birleşik paket.
//
// Paket düzeyi (carryParts), tek ve genel kural:
//   - Açık ilişki: kimliği kaynak parçanın XML'inde geçer (r:id, r:embed, r:dm…). Sahibi modeldir;
//     paket katmanı dokunmaz (öğe düzeyi taşıyıcı `Carry` ilgilenir).
//   - Örtük ilişki: kimliği XML'de geçmez (tema, fontTable, glossary, vbaProject, customXml…).
//     Politika tablosuna (policy.ts) göre taşınır, düşürülür ya da yazıcıya bırakılır.
//   - Yazıcı aynı kaynak parçada aynı türü zaten ürettiyse taşınmaz; customXml kök ad alanına göre eşlenir.
//   - Kapanış: taşınan parçanın kendi .rels'indeki iç hedefler de taşınır (döngü korumalı); eşlenen
//     parçalara bağlanır, eşlemesiz yapısal parçaya başvuru atılır, hedefi olmayan ilişki atılır.
//   - Ad çakışmasında "oh_tN_" önekiyle yeniden adlandırılır; öğe düzeyinde önek her zaman kullanılır.
//   - [Content_Types].xml: kaynaktaki tür taşınır; paket her zaman içerik türleri ilk parça olacak
//     şekilde kurulur. Makrolu hedefte ana parça türü makrolu olur.
//   - docProps/custom.xml: kullanıcı özellikleri yazıcınınkilerle birleştirilir.
//   - Yalnız eşlenen kaynak parçalardan ulaşılanlar taşınır; öğe düzeyi kopyalardan yeni paketin
//     kökünden ulaşılamayanlar atılır (çöp toplama).
// Taşınan parçalar yeniden sıkıştırılmadan ham kopyalanır.

import {
  CONTENT_TYPES_PART,
  ContentTypes,
  decodeText,
  escapeXml,
  explicitRelIds,
  isReservedName,
  OpcPackage,
  parseAttrs,
  relativeTarget,
  relsPathOf,
  relsXml,
  rewriteRelIds,
  RelsBuilder,
  XML_DECL,
  type Relationship,
} from "./package";
import { classifyRel, DROP_TYPES, macroContentType, OWN_TYPES, REL_TYPES, relTypeName, SINGLETON_TYPES, STRUCTURAL_TYPES, type DropReason, type OpcApp, type PolicyOverrides } from "./policy";
import { packageSources, type SourceRegistry } from "./source";
import { ZipWriter } from "./zip";

/* ------------------------------------------------------------ rapor */

export interface CarriedPart {
  /** Kaynak paketteki ad (dış ilişkide adres). */
  from: string;
  /** Yeni paketteki ad (dış ilişkide adres). */
  to: string;
  /** Parçaya ulaştıran ilişkinin türü (öğe düzeyi kopyada boş olabilir). */
  type: string;
  kind: "implicit" | "closure" | "element" | "external";
}

export interface DroppedPart {
  /** Kaynaktaki parça (ya da çözülemeyen hedef). */
  part: string;
  type: string;
  reason: DropReason;
  /** İlişkinin kaynak parçası ("" = kök). */
  source: string;
}

export interface RelInfo {
  source: string;
  id: string;
  type: string;
  /** Hedef parça (dışsa adres). */
  part: string;
}

export interface CarryReport {
  carried: CarriedPart[];
  dropped: DroppedPart[];
  /** Yazıcının sahip olduğu örtük ilişkiler (taşınmadı). */
  owned: RelInfo[];
  /** Açık ilişkiler: modelin (öğe düzeyi taşıyıcının) işi. */
  explicit: RelInfo[];
  /** Yazıcının custom.xml'ine eklenen kullanıcı özelliği sayısı. */
  customProperties: number;
  /** Ana parça türü makrolu yapıldı mı? */
  macro: boolean;
  /** Kaynak kimliği kayıtta bulunamadı: paket düzeyi taşıma yapılmadı. */
  sourceMissing: boolean;
  /** Kökten ulaşılamadığı için atılan öğe düzeyi kopya sayısı. */
  collected: number;
  ms: number;
}

/* ------------------------------------------------------------ öğe düzeyi taşıyıcı */

interface Plan {
  pkg: OpcPackage;
  from: string;
  to: string;
  type: string;
  kind: "implicit" | "closure" | "element";
}

/** Bir şekil/öğe örneği için kopyalayıcı. */
export interface CarryInstance {
  /** Örneğin kaynak paketi (kayıtta yoksa undefined: her kopya null döner). */
  readonly source: OpcPackage | undefined;
  /**
   * Kaynak parçayı (paket içi mutlak ad) bu örnek için kopyalanacak olarak işaretler ve yeni adını
   * döner ("…/oh_tN_ad"). Aynı örnekte aynı parça bir kez kopyalanır. Parça yoksa null.
   */
  copy(part: string): string | null;
  /** Bu örnekte daha önce kopyalanan parçanın yeni adı. */
  mapped(part: string): string | undefined;
}

const OH_PREFIX = /^(?:oh_t\d+_)+/i;

function prefixed(name: string, n: number): string {
  const i = name.lastIndexOf("/");
  return `${name.slice(0, i + 1)}oh_t${n}_${name.slice(i + 1).replace(OH_PREFIX, "")}`;
}

/**
 * Öğe düzeyi taşıma planı (bir yazma turu için). Yazıcı, opak öğenin ilişkilerini yeniden kurarken
 * `instance(kaynak).copy(parça)` ile yeni adları alır; baytlar `carryParts({ carry })` sırasında
 * kaynak paketten ham kopyalanır, kapanışları ve içerik türleriyle birlikte.
 */
export class Carry {
  private n = 0;
  private list: Plan[] = [];
  private patchMap = new Map<string, { to: string; fns: ((xml: string) => string)[] }>();
  private readonly registry: SourceRegistry;

  constructor(registry: SourceRegistry = packageSources) {
    this.registry = registry;
  }

  /** Yeni bir örnek kopyalayıcı (her opak öğe örneği kendi kopyalarını alır; paylaşım ve kimlik çakışması olmaz). */
  instance(source: string | OpcPackage | null | undefined): CarryInstance {
    const pkg = typeof source === "string" ? this.registry.get(source) : (source ?? undefined);
    const k = ++this.n;
    const map = new Map<string, string>();
    return {
      source: pkg,
      copy: (part) => {
        const name = pkg?.canonicalName(part);
        if (!pkg || !name || isReservedName(name)) return null;
        const key = name.toLowerCase(); // yerel-ayar: ascii
        let to = map.get(key);
        if (!to) {
          to = prefixed(name, k);
          map.set(key, to);
          this.list.push({ pkg, from: name, to, type: "", kind: "element" });
        }
        return to;
      },
      mapped: (part) => map.get(part.toLowerCase()), // yerel-ayar: ascii
    };
  }

  /** Kopyalanan parçanın (yeni adıyla) metnini yazmadan önce dönüştürür (ör. SmartArt relId yaması). */
  patch(to: string, fn: (xml: string) => string): void {
    const key = to.toLowerCase(); // yerel-ayar: ascii
    const p = this.patchMap.get(key) ?? { to, fns: [] };
    p.fns.push(fn);
    this.patchMap.set(key, p);
  }

  /** Planlanan kopya sayısı. */
  get size(): number {
    return this.list.length;
  }

  /** @internal */
  plans(): readonly Plan[] {
    return this.list;
  }

  /** @internal */
  patchesFor(to: string): ((xml: string) => string)[] | undefined {
    return this.patchMap.get(to.toLowerCase())?.fns; // yerel-ayar: ascii
  }

  /** @internal Sonraki önek numarası (kapanış kopyaları için). */
  nextNumber(): number {
    return ++this.n;
  }
}

/** Opak öğenin başvurduğu ilişki (okurken saklanır). */
export interface RawRef {
  id: string;
  type: string;
  /** İç hedefte kaynak paketteki parça adı, dış hedefte adres. */
  target: string;
  external: boolean;
}

/** Opak XML parçacığının başvurduğu ilişkileri kaynak parçanın .rels'inden toplar. Bilinmeyen kimlik atlanır. */
export async function collectRefs(pkg: OpcPackage, part: string, xml: string, opts: { prefixes?: readonly string[] } = {}): Promise<RawRef[]> {
  const ids = explicitRelIds(xml, opts);
  if (!ids.size) return [];
  const byId = new Map((await pkg.rels(part)).map((r) => [r.id, r]));
  const out: RawRef[] = [];
  for (const id of ids) {
    const r = byId.get(id);
    if (!r) continue;
    out.push({ id, type: r.type, external: r.external, target: r.external ? r.target : r.part ? (pkg.canonicalName(r.part) ?? r.part) : r.target });
  }
  return out;
}

/**
 * Opak XML'i yeni pakete bağlar: iç hedefleri `inst` ile kopyalanacak olarak işaretler, her ilişkiyi
 * `addRel` ile yazıcının .rels'ine ekletir (yeni kimliği döner; iç hedef yeni paketteki mutlak addır) ve
 * XML'deki kimlikleri yeniden yazar. Kopyalanamayan (kaynakta olmayan) başvurular `missing`'de döner,
 * XML'de eski kimlikleriyle kalır.
 */
export function reattach(
  xml: string,
  refs: readonly RawRef[],
  inst: CarryInstance,
  addRel: (rel: { type: string; target: string; external: boolean }) => string,
  opts: { prefixes?: readonly string[] } = {},
): { xml: string; ids: Map<string, string>; missing: RawRef[] } {
  const ids = new Map<string, string>();
  const missing: RawRef[] = [];
  for (const ref of refs) {
    if (ids.has(ref.id)) continue;
    if (ref.external) {
      ids.set(ref.id, addRel({ type: ref.type, target: ref.target, external: true }));
      continue;
    }
    const to = inst.copy(ref.target);
    if (!to) missing.push(ref);
    else ids.set(ref.id, addRel({ type: ref.type, target: to, external: false }));
  }
  return { xml: rewriteRelIds(xml, ids, opts), ids, missing };
}

/* ------------------------------------------------------------ paket düzeyi taşıma */

export interface CarryPartsOptions {
  app: OpcApp;
  /** Özgün paket ya da kayıt kimliği. Kimlik kayıtta yoksa paket düzeyi taşıma yapılmaz (`report.sourceMissing`). */
  source: OpcPackage | string | null | undefined;
  /** Yazıcının çıktısı. ZIP değilse OpcError. */
  output: Uint8Array | ArrayBuffer | OpcPackage;
  /** Hedef biçim makrolu mu (docm/xlsm/pptm)? Değilse makro projesi düşürülür. */
  macro?: boolean;
  /**
   * Kök ve ana parça dışındaki kaynak ↔ çıktı parça eşlemesi (sayfa adına göre sayfalar, slayt kimliğine
   * göre slaytlar…). Eşlenen parçaların örtük ilişkileri de taşınır; kapanıştaki başvurular bunlara bağlanır.
   */
  partMap?: Iterable<readonly [string, string]> | Record<string, string>;
  /** Öğe düzeyi taşıma planı. */
  carry?: Carry;
  /** Politika tablosuna uygulama ekleri. */
  overrides?: PolicyOverrides;
  /** Yazıcının sahip olduğu özel özellik adları (kaynaktakiler taşınmaz). Varsayılan: "OfficeHere" ile başlayanlar. */
  ownedCustomProperty?: (name: string) => boolean;
  /** `source` kimlikse aranacak kayıt (varsayılan: packageSources). */
  registry?: SourceRegistry;
}

const CT_CUSTOM = "application/vnd.openxmlformats-officedocument.custom-properties+xml";
const OCTET = "application/octet-stream";
const defaultOwnedProperty = (name: string) => /^officehere/i.test(name);

function pairsOf(map: CarryPartsOptions["partMap"]): [string, string][] {
  if (!map) return [];
  if (Symbol.iterator in (map as object)) return [...(map as Iterable<readonly [string, string]>)].map(([a, b]) => [a, b]);
  return Object.entries(map as Record<string, string>);
}

/** XML'in kök öğesinin ad alanı (customXml eşlemesi için). */
export function rootNamespace(xml: string): string {
  const body = xml.replace(/<\?[\s\S]*?\?>|<!--[\s\S]*?-->|<!DOCTYPE[^>]*>/g, "");
  const m = /<([\w.-]+:)?[\w.-]+(\s[^>]*)?>/.exec(body);
  if (!m) return "";
  const prefix = m[1] ? m[1].slice(0, -1) : "";
  const attrs = parseAttrs(m[2] ?? "");
  return attrs.get(prefix ? `xmlns:${prefix}` : "xmlns") ?? "";
}

/**
 * Yazıcının çıktısına özgün paketten taşınması gereken parçaları ekler. Kaynakla ilgili hiçbir sorun
 * (bozuk .rels, olmayan hedef, okunamayan parça) hata fırlatmaz; rapora yazılır. Yalnız çıktı ZIP
 * değilse OpcError fırlatır.
 */
export async function carryParts(opts: CarryPartsOptions): Promise<{ bytes: Uint8Array; report: CarryReport }> {
  const t0 = performance.now();
  const report: CarryReport = { carried: [], dropped: [], owned: [], explicit: [], customProperties: 0, macro: false, sourceMissing: false, collected: 0, ms: 0 };
  const out = opts.output instanceof OpcPackage ? opts.output : await OpcPackage.open(opts.output);
  const registry = opts.registry ?? packageSources;
  const src = typeof opts.source === "string" ? registry.get(opts.source) : (opts.source ?? undefined);
  if (typeof opts.source === "string" && !src) report.sourceMissing = true;
  const app = opts.app;
  const carry = opts.carry;
  const ct = out.contentTypes.clone();
  const modified = new Map<string, { name: string; text: string }>(); // küçük harf ad → yeni içerik
  const builders = new Map<string, RelsBuilder>(); // küçük harf çıktı parçası → .rels kurucusu
  const taken = new Set(out.names.map((n) => n.toLowerCase())); // yerel-ayar: ascii
  const plans: Plan[] = [];
  const planIndex = new Map<OpcPackage, Map<string, Plan>>();
  let seq = 0;
  const nextN = () => (carry ? carry.nextNumber() : ++seq);

  const drop = (part: string, type: string, reason: DropReason, source: string) => report.dropped.push({ part, type, reason, source });
  const indexOf = (pkg: OpcPackage) => {
    let m = planIndex.get(pkg);
    if (!m) planIndex.set(pkg, (m = new Map()));
    return m;
  };
  const builder = async (part: string) => {
    const key = part.toLowerCase(); // yerel-ayar: ascii
    let b = builders.get(key);
    if (!b) builders.set(key, (b = new RelsBuilder(part, await out.rels(part))));
    return b;
  };
  /** Kaynak parçayı kopya planına ekler (varsa olanı döner); ad çakışırsa (ya da öğe düzeyindeyse) önek alır. */
  const plan = (pkg: OpcPackage, from: string, type: string, kind: Plan["kind"], forcePrefix: boolean): Plan => {
    const idx = indexOf(pkg);
    const hit = idx.get(from.toLowerCase()); // yerel-ayar: ascii
    if (hit) return hit;
    let to = from;
    if (forcePrefix || taken.has(to.toLowerCase())) { // yerel-ayar: ascii
      do to = prefixed(from, nextN());
      while (taken.has(to.toLowerCase())); // yerel-ayar: ascii
    }
    taken.add(to.toLowerCase()); // yerel-ayar: ascii
    const p: Plan = { pkg, from, to, type, kind };
    idx.set(from.toLowerCase(), p); // yerel-ayar: ascii
    plans.push(p);
    return p;
  };

  // 1) Öğe düzeyi planlar: adları yazıcının parçalarıyla çakışmamalı
  for (const p of carry?.plans() ?? []) {
    if (taken.has(p.to.toLowerCase())) { // yerel-ayar: ascii
      drop(p.from, p.type, "conflict", "");
      continue;
    }
    taken.add(p.to.toLowerCase()); // yerel-ayar: ascii
    const idx = indexOf(p.pkg);
    if (!idx.has(p.from.toLowerCase())) idx.set(p.from.toLowerCase(), p); // yerel-ayar: ascii
    plans.push(p);
  }

  // 2) Paket düzeyi: eşlenen parçaların örtük ilişkileri
  const partMap = new Map<string, string>();
  const pairs: [string, string][] = [];
  if (src) {
    pairs.push(["", ""]);
    if (src.mainPart && out.mainPart) pairs.push([src.mainPart, out.mainPart]);
    for (const [a, b] of pairsOf(opts.partMap)) {
      const ca = src.canonicalName(a);
      const cb = out.canonicalName(b);
      if (ca && cb && !pairs.some(([x]) => x.toLowerCase() === ca.toLowerCase())) pairs.push([ca, cb]); // yerel-ayar: ascii
    }
    for (const [a, b] of pairs) if (a) partMap.set(a.toLowerCase(), b); // yerel-ayar: ascii
  }
  let customRel: Relationship | undefined;
  for (const [a, b] of pairs) {
    const srcRels = await src!.rels(a);
    const candidates: Relationship[] = [];
    for (const r of srcRels) {
      const short = relTypeName(r.type);
      if (a === "" && short === "officeDocument") continue;
      const target = r.external ? r.target : (r.part ?? r.target);
      const d = classifyRel(app, r.type, { macro: opts.macro, overrides: opts.overrides });
      if (d.action === "own") report.owned.push({ source: a, id: r.id, type: r.type, part: target });
      else if (d.action === "drop") drop(target, r.type, d.reason ?? "policy", a);
      else candidates.push(r);
    }
    if (!candidates.length) continue;
    let ids: Set<string> | null = new Set();
    if (a) ids = await src!.explicitRelIds(a).catch(() => null);
    const outRels = [...(await builder(b)).relationships];
    for (const r of candidates) {
      const short = relTypeName(r.type);
      const target = r.external ? r.target : (r.part ?? r.target);
      // Okunamayan kaynak parçada hepsi açık sayılır (taşımamak güvenli olan)
      if (!ids || ids.has(r.id)) {
        report.explicit.push({ source: a, id: r.id, type: r.type, part: target });
        continue;
      }
      if (a === "" && short === "custom-properties") {
        customRel ??= r;
        continue;
      }
      const same = outRels.filter((o) => relTypeName(o.type) === short);
      if (same.length) {
        if (short !== "customXml" || r.external || !r.part) {
          report.owned.push({ source: a, id: r.id, type: r.type, part: target });
          continue;
        }
        const nsA = rootNamespace((await src!.text(r.part).catch(() => null)) ?? "");
        let hit = false;
        for (const o of same) if (o.part && rootNamespace((await out.text(o.part).catch(() => null)) ?? "") === nsA) hit = true;
        if (hit) {
          report.owned.push({ source: a, id: r.id, type: r.type, part: target });
          continue;
        }
      }
      if (r.external) {
        (await builder(b)).addExternal(r.type, r.target);
        report.carried.push({ from: r.target, to: r.target, type: r.type, kind: "external" });
        continue;
      }
      const name = r.part ? src!.canonicalName(r.part) : undefined;
      if (!name || isReservedName(name)) {
        drop(target, r.type, "missing", a);
        continue;
      }
      const p = plan(src!, name, r.type, "implicit", false);
      (await builder(b)).addPart(r.type, p.to);
    }
  }

  // 3) Kapanış: taşınan her parçanın kendi ilişkileri (döngü korumalı; planlar kuyruk gibi büyür)
  const planRels = new Map<Plan, Relationship[]>();
  for (let i = 0; i < plans.length; i++) {
    const p = plans[i];
    const list: Relationship[] = [];
    for (const r of await p.pkg.rels(p.from)) {
      const short = relTypeName(r.type);
      if (r.external) {
        list.push({ ...r });
        continue;
      }
      const name = r.part ? p.pkg.canonicalName(r.part) : undefined;
      if (!name || isReservedName(name)) {
        drop(r.part ?? r.target, r.type, "missing", p.from);
        continue;
      }
      const link = (to: string) => list.push({ id: r.id, type: r.type, target: relativeTarget(p.to, to), external: false, part: to });
      const mapped = p.pkg === src ? partMap.get(name.toLowerCase()) : undefined; // yerel-ayar: ascii
      if (mapped) {
        link(mapped);
        continue;
      }
      const known = indexOf(p.pkg).get(name.toLowerCase()); // yerel-ayar: ascii
      if (known) {
        link(known.to);
        continue;
      }
      if (DROP_TYPES.get(short) === "signature") {
        drop(name, r.type, "signature", p.from);
        continue;
      }
      if (STRUCTURAL_TYPES.has(short)) {
        drop(name, r.type, "structural", p.from);
        continue;
      }
      if (SINGLETON_TYPES.has(short) && OWN_TYPES[app].has(short) && out.has(name) && out.contentType(name) === p.pkg.contentType(name)) {
        link(out.canonicalName(name)!);
        continue;
      }
      link(plan(p.pkg, name, r.type, "closure", p.kind === "element").to);
    }
    planRels.set(p, list);
  }

  // 4) Çöp toplama: öğe düzeyi kopyalar yalnız yeni paketin kökünden ulaşılıyorsa yazılır
  let live = plans;
  if (plans.some((p) => p.kind === "element")) {
    const byTo = new Map(plans.map((p) => [p.to.toLowerCase(), p])); // yerel-ayar: ascii
    const relsOf = async (name: string): Promise<Relationship[]> => {
      const key = name.toLowerCase(); // yerel-ayar: ascii
      const b = builders.get(key);
      if (b) return [...b.relationships];
      const p = byTo.get(key);
      if (p) return planRels.get(p) ?? [];
      return name === "" || out.has(name) ? out.rels(name) : [];
    };
    const seen = new Set<string>([""]);
    let queue = [""];
    while (queue.length) {
      const lists = await Promise.all(queue.map(relsOf));
      queue = [];
      for (const r of lists.flat()) {
        if (r.external || !r.part) continue;
        const k = r.part.toLowerCase(); // yerel-ayar: ascii
        if (seen.has(k)) continue;
        seen.add(k);
        queue.push(r.part);
      }
    }
    live = plans.filter((p) => seen.has(p.to.toLowerCase())); // yerel-ayar: ascii
    report.collected = plans.length - live.length;
  }

  // 5) Özel belge özellikleri: kullanıcının özellikleri yazıcınınkilerle birleştirilir
  if (src && customRel?.part) {
    await mergeCustomProperties(src, out, customRel, {
      owned: opts.ownedCustomProperty ?? defaultOwnedProperty,
      taken,
      nextN,
      report,
      modified,
      ct,
      rootBuilder: () => builder(""),
    });
  }

  // 6) İçerik türleri ve makro
  for (const p of live) {
    const type = p.pkg.contentType(p.from) ?? OCTET;
    if (ct.get(p.to) === type) continue;
    const dot = p.to.lastIndexOf(".");
    const ext = dot > p.to.lastIndexOf("/") ? p.to.slice(dot + 1) : "";
    if (ext && !p.pkg.contentTypes.getOverride(p.from) && !ct.getDefault(ext)) ct.setDefault(ext, type);
    else ct.setOverride(p.to, type);
  }
  if (opts.macro && out.mainPart) {
    const cur = ct.get(out.mainPart);
    const next = cur ? macroContentType(cur) : undefined;
    if (cur && next && next !== cur) {
      ct.setOverride(out.mainPart, next);
      report.macro = true;
    }
  }

  // 7) Rapor ve paketleme ([Content_Types].xml ilk parça; çıktı ve taşınan parçalar ham kopya)
  for (const p of live) report.carried.push({ from: p.from, to: p.to, type: p.type, kind: p.kind });
  for (const b of builders.values()) {
    if (b.size === (await out.rels(b.source)).length) continue;
    const path = relsPathOf(b.source);
    modified.set(path.toLowerCase(), { name: out.canonicalName(path) ?? path, text: b.xml() }); // yerel-ayar: ascii
  }
  const ctEntry = out.entry(CONTENT_TYPES_PART);
  const ctChanged = ct.toXml() !== out.contentTypes.toXml();
  if (!live.length && !modified.size && !ctChanged && out.entries[0]?.name === CONTENT_TYPES_PART) {
    report.ms = performance.now() - t0;
    return { bytes: out.bytes, report };
  }
  const zw = new ZipWriter(ctEntry ? { time: ctEntry.time, date: ctEntry.date } : {});
  if (ctChanged || !ctEntry) zw.add(CONTENT_TYPES_PART, ct.toXml());
  else zw.addRaw(CONTENT_TYPES_PART, out.bytes, ctEntry);
  for (const e of out.entries) {
    if (e.name === CONTENT_TYPES_PART) continue;
    const m = modified.get(e.name.toLowerCase()); // yerel-ayar: ascii
    if (m) {
      zw.add(e.name, m.text);
      modified.delete(e.name.toLowerCase()); // yerel-ayar: ascii
    } else zw.addRaw(e.name, out.bytes, e);
  }
  for (const m of modified.values()) zw.add(m.name, m.text);
  for (const p of live) {
    const e = p.pkg.entry(p.from)!;
    const fns = carry?.patchesFor(p.to);
    if (fns?.length) {
      try {
        let text = decodeText((await p.pkg.data(p.from)) ?? new Uint8Array());
        for (const fn of fns) text = fn(text);
        zw.add(p.to, text);
      } catch {
        // açılamayan parça yamasız (ham) yazılır
        zw.addRaw(p.to, p.pkg.bytes, e);
      }
    } else zw.addRaw(p.to, p.pkg.bytes, e);
    const rels = planRels.get(p);
    if (rels?.length) zw.add(relsPathOf(p.to), relsXml(rels));
  }
  const bytes = await zw.finish();
  report.ms = performance.now() - t0;
  return { bytes, report };
}

/* ------------------------------------------------------------ özel özellikler */

const PROP_RE = /<(?:[\w.-]+:)?property\b[^>]*?(?:\/>|>[\s\S]*?<\/(?:[\w.-]+:)?property\s*>)/g;
const ROOT_RE = /<((?:[\w.-]+:)?Properties)\b([^>]*?)(\/?)>/;

function propName(p: string): string {
  return parseAttrs(/^<[^>]*>/.exec(p)?.[0] ?? "").get("name") ?? "";
}

/** custom.xml birleştirme: yazıcının özellikleri + kaynakta olup yazıcıda olmayan kullanıcı özellikleri. */
async function mergeCustomProperties(
  src: OpcPackage,
  out: OpcPackage,
  rel: Relationship,
  ctx: {
    owned: (name: string) => boolean;
    taken: Set<string>;
    nextN: () => number;
    report: CarryReport;
    modified: Map<string, { name: string; text: string }>;
    ct: ContentTypes;
    rootBuilder: () => Promise<RelsBuilder>;
  },
): Promise<void> {
  const name = src.canonicalName(rel.part!);
  if (!name) {
    ctx.report.dropped.push({ part: rel.part!, type: rel.type, reason: "missing", source: "" });
    return;
  }
  const srcXml = await src.text(name).catch(() => null);
  const srcRoot = srcXml ? ROOT_RE.exec(srcXml) : null;
  if (!srcXml || !srcRoot) {
    ctx.report.dropped.push({ part: name, type: rel.type, reason: "unreadable", source: "" });
    return;
  }
  const outRel = (await out.rels("")).find((r) => relTypeName(r.type) === "custom-properties" && r.part && out.has(r.part));
  const outName = outRel ? out.canonicalName(outRel.part!)! : null;
  const outXml = outName ? await out.text(outName).catch(() => null) : null;
  const outRoot = outXml ? ROOT_RE.exec(outXml) : null;
  const outProps = outXml && outRoot ? [...outXml.matchAll(PROP_RE)].map((m) => m[0]) : [];
  const have = new Set(outProps.map((p) => propName(p).toLowerCase())); // yerel-ayar: ascii
  const add = [...srcXml.matchAll(PROP_RE)].map((m) => m[0]).filter((p) => {
    const n = propName(p);
    return n && !ctx.owned(n) && !have.has(n.toLowerCase()); // yerel-ayar: ascii
  });
  if (!add.length) return;
  const all = [...outProps, ...add].map((p, i) => p.replace(/^(<[^>]*?\bpid\s*=\s*)(?:"[^"]*"|'[^']*')/, `$1"${i + 2}"`));
  // Kök: yazıcınınki (yoksa kaynağınki); kaynağın ad alanı bildirimleri eksikse eklenir
  let open = outRoot ? `<${outRoot[1]}${outRoot[2]}>` : `<${srcRoot[1]}${srcRoot[2]}>`;
  const qname = outRoot ? outRoot[1] : srcRoot[1];
  const decl = parseAttrs(open);
  for (const [k, v] of parseAttrs(srcRoot[2])) if (k.startsWith("xmlns") && !decl.has(k)) open = open.replace(/>$/, ` ${k}="${escapeXml(v)}">`);
  const xml = `${XML_DECL}${open}${all.join("")}</${qname}>`;
  if (outName) ctx.modified.set(outName.toLowerCase(), { name: outName, text: xml }); // yerel-ayar: ascii
  else {
    let to = name;
    if (ctx.taken.has(to.toLowerCase())) { // yerel-ayar: ascii
      do to = prefixed(name, ctx.nextN());
      while (ctx.taken.has(to.toLowerCase())); // yerel-ayar: ascii
    }
    ctx.taken.add(to.toLowerCase()); // yerel-ayar: ascii
    ctx.modified.set(to.toLowerCase(), { name: to, text: xml }); // yerel-ayar: ascii
    (await ctx.rootBuilder()).addPart(REL_TYPES.customProperties, to);
    ctx.ct.setOverride(to, CT_CUSTOM);
    ctx.report.carried.push({ from: name, to, type: rel.type, kind: "implicit" });
  }
  ctx.report.customProperties = add.length;
}
