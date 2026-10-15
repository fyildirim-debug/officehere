/**
 * Geçiş katmanı (K§1.2.4 Sunum): modelin tanımadığı öğeler özgün XML'leriyle taşınır.
 *
 * Okurken: SmartArt, OLE nesnesi, mürekkep ve tanınmayan çizim öğeleri yedek görünümleriyle
 * (SmartArt → dsp çiziminden şekil grubu, OLE/mürekkep → yedek resim, bilinmeyen → yer tutucu resim)
 * modele girer; şeklin `raw` alanında özgün XML, ilişkileri ve görünüm imzası saklanır. Slaydın
 * tanınmayan öğeleri (custDataLst, ActiveX denetimleri, extLst) `Slide.raw`'da tutulur.
 *
 * Kaydederken: görünüm imzası değişmemişse özgün XML yazılır (konum/boyut ve cNvPr kimliği güncel,
 * ilişkiler yeni pakette yeniden kurulur, başvurulan parçalar kapanışlarıyla kaynak paketten
 * kopyalanır: core `Carry`). Görünüm düzenlendiyse yedek görünüm yazılır (SmartArt anlamı kaybolur,
 * görünüm kalır) ve Türkçe uyarı verilir. Kaynak paket bellekte kimlikle tutulur; model bayt taşımaz.
 */
import { attrOf, Carry, childElements, explicitRelIds, packageSources, PRESENTATION_ORDER, relativeTarget, relTypeName, rewriteRelIds, sha256Hex, XmlPart, type CarryInstance, type OpcPackage } from "@officehere/core/opc";
import type { PicShape, RawKind, RawRef, RawXml, Shape, Slide, SlideRaw, Theme } from "../model/types";
import { esc, RelsBuilder, SECTION_EXT_URI, type Rel } from "./xml";

/* ------------------------------------------------------------ kaynak paket kaydı */

/**
 * Bellekte tutulan en çok kaynak paket sayısı. Otomatik kayıt doğrulaması gibi tek kullanımlık
 * okumalar da kayıt açar; en uzun süredir kullanılmayan bırakılır (açık sunu her kayıtta kullanılır).
 */
const KAYNAK_SINIRI = 12;
const kullanim = new Map<string, true>();

function dokun(id: string): void {
  kullanim.delete(id);
  kullanim.set(id, true);
}

/**
 * Açılan paketi kaydeder ve kimliğini döner. Kimlik özgün baytların SHA-256'sıdır: oturumlar arası
 * benzersizdir (kurtarılan taslaktaki eski kimlik başka bir dosyanın kaynağına bağlanmaz) ve kalıcı
 * kaynak deposunun (IndexedDB) anahtarıyla aynıdır. Hesaplanamazsa kayıt sayaçlı kimlik verir.
 */
export async function kaynakKaydet(pkg: OpcPackage, name?: string): Promise<{ id: string; yeni: boolean }> {
  const sha = await sha256Hex(pkg.bytes).catch(() => undefined);
  // Aynı dosya yeniden açıldıysa var olan kayıt kullanılır (açık sunuların kaynağı değişmez).
  if (sha && packageSources.has(`kaynak-${sha}`)) {
    dokun(`kaynak-${sha}`);
    return { id: `kaynak-${sha}`, yeni: false };
  }
  const id = packageSources.register(pkg, { id: sha && `kaynak-${sha}`, name });
  dokun(id);
  while (kullanim.size > KAYNAK_SINIRI) {
    const eski = kullanim.keys().next().value as string;
    kullanim.delete(eski);
    packageSources.release(eski);
  }
  return { id, yeni: true };
}

/** Kaydı bırakır (okuma başarısız olduysa). */
export function kaynakBirak(id: string | undefined): void {
  if (!id) return;
  kullanim.delete(id);
  packageSources.release(id);
}

/** Kayıtlı kaynak paket (kullanıldı olarak işaretlenir); yoksa tanımsız. */
export function kaynakPaket(id: string | undefined): OpcPackage | undefined {
  const pkg = id ? packageSources.get(id) : undefined;
  if (pkg) dokun(id!);
  return pkg;
}

/* ------------------------------------------------------------ okurken */

const XMLNS = "http://www.w3.org/2000/xmlns/";

/** Açılış etiketinin bitişi (tırnak içindeki ">" atlanır). */
function openTagEnd(xml: string, start = 0): number {
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

/**
 * Öğenin XML'i; atalarda bildirilen bütün önekli ad alanları kök etikete eklenir (mc:Choice
 * Requires="p14" gibi değerle anılan önekler de kapsamda kalsın).
 */
export function rawOuterXml(el: Element): string {
  // Bildirimler kopyanın kökünde olsun: seri hâle getirici alt öğelerde yinelemesin.
  const copy = el.cloneNode(true) as Element;
  for (let p = el.parentElement; p; p = p.parentElement) {
    for (const a of Array.from(p.attributes)) {
      if (a.prefix !== "xmlns" || copy.hasAttribute(a.name)) continue;
      copy.setAttributeNS(XMLNS, a.name, a.value);
    }
  }
  return new XMLSerializer().serializeToString(copy);
}

/** XML'de geçen ilişki kimliklerinin parça ilişkilerindeki karşılıkları (bilinmeyen kimlik atlanır). */
export function rawRefs(xml: string, rels: Map<string, Rel>): RawRef[] {
  const out: RawRef[] = [];
  for (const id of explicitRelIds(xml)) {
    const r = rels.get(id);
    if (r) out.push({ id, type: r.type, target: r.target, external: r.external });
  }
  return out;
}

const ETIKET: Record<RawKind, string> = { smartart: "SmartArt", ole: "Gömülü nesne", murekkep: "Mürekkep", bilinmeyen: "Desteklenmeyen öğe" };
export const RAW_NOTE = "Desteklenmeyen öğe; kaydederken korunur";

function base64Utf8(s: string): string {
  const b = new TextEncoder().encode(s);
  let t = "";
  for (let i = 0; i < b.length; i += 0x8000) t += String.fromCharCode(...b.subarray(i, i + 0x8000));
  return btoa(t);
}

/** Yer tutucu resim (SVG): öğe türü ve "kaydederken korunur" notu; çerçeveyi her oranda doldurur. */
export function placeholderSrc(kind: RawKind): string {
  const svg =
    `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 320 120">` +
    `<rect x="-2000" y="-2000" width="4320" height="4120" fill="#EEF1F5"/>` +
    `<rect x="4" y="4" width="312" height="112" rx="6" fill="none" stroke="#8A94A6" stroke-width="2" stroke-dasharray="8 6"/>` +
    `<text x="160" y="54" text-anchor="middle" font-family="Carlito, Calibri, Arial, sans-serif" font-size="20" font-weight="bold" fill="#3D4757">${esc(ETIKET[kind])}</text>` +
    `<text x="160" y="84" text-anchor="middle" font-family="Carlito, Calibri, Arial, sans-serif" font-size="15" fill="#3D4757">${esc(RAW_NOTE)}</text></svg>`;
  return `data:image/svg+xml;base64,${base64Utf8(svg)}`;
}

/** Yer tutucu resim şekli (konum ve adıyla). */
export function placeholderPic(kind: RawKind, base: { id: string; name: string; descr?: string; hidden?: boolean; x: number; y: number; w: number; h: number }): PicShape {
  return { kind: "pic", ...base, name: base.name || ETIKET[kind], rot: 0, src: placeholderSrc(kind), line: { fill: { type: "none" }, width: 1, dash: "solid" } };
}

/* ------------------------------------------------------------ görünüm imzası */

function fnv(s: string, seed: number): string {
  let h = seed >>> 0;
  for (let i = 0; i < s.length; i++) h = Math.imul(h ^ s.charCodeAt(i), 16777619) >>> 0;
  return h.toString(16).padStart(8, "0");
}
const hash = (s: string) => fnv(s, 0x811c9dc5) + fnv(s, 0x01000193);

/** Anahtarları sıralı JSON (nesne kurulum sırası imzayı etkilemesin). */
function stable(v: unknown): string {
  if (v === null || typeof v !== "object") return JSON.stringify(v) ?? "null";
  if (Array.isArray(v)) return `[${v.map((x) => (x === undefined ? "null" : stable(x))).join(",")}]`;
  const o = v as Record<string, unknown>;
  return `{${Object.keys(o)
    .filter((k) => o[k] !== undefined)
    .sort()
    .map((k) => `${JSON.stringify(k)}:${stable(o[k])}`)
    .join(",")}}`;
}

/** Görünüme dahil olmayan (cNvPr'de güncellenen ya da yalnız oturum içi hesaplanan) alanlar. */
const SIG_SKIP = new Set(["id", "name", "descr", "hidden", "raw", "srcPart", "fontScale", "lnSpcReduction", "minH"]);

/**
 * Görünüm imzası: şekil ve alt şekillerinin bütün özellikleri; alt şekillerin konum/boyutu
 * şeklin kutusuna göre oranlanır (taşıma ve ölçekleme imzayı değiştirmez).
 */
export function rawSig(s: Shape): string {
  const W = s.w || 1;
  const H = s.h || 1;
  const geo: number[] = [];
  const walk = (v: unknown, holder: boolean): unknown => {
    if (Array.isArray(v)) return v.map((x) => walk(x, false));
    if (!v || typeof v !== "object") return typeof v === "string" && v.length > 200 ? hash(v) : v;
    const o = v as Record<string, unknown>;
    const isShape = typeof o.kind === "string" && typeof o.x === "number" && typeof o.w === "number";
    const out: Record<string, unknown> = {};
    for (const [k, x] of Object.entries(o)) {
      if (SIG_SKIP.has(k)) continue;
      if (isShape && (k === "x" || k === "y" || k === "w" || k === "h")) continue;
      if (isShape && k === "cols" && Array.isArray(x)) {
        for (const c of x) geo.push((c as number) / W);
        continue;
      }
      if (k === "h" && typeof x === "number" && "cells" in o) {
        geo.push(x / H);
        continue;
      }
      out[k] = walk(x, false);
    }
    if (isShape && !holder) geo.push(((o.x as number) - s.x) / W, ((o.y as number) - s.y) / H, (o.w as number) / W, (o.h as number) / H);
    return out;
  };
  const body = stable(walk(s, true));
  return `${hash(body)}|${geo.map((g) => (Math.round(g * 10000) / 10000).toString()).join(",")}`;
}

/** Temanın modeldeki imzası (ad, renkler, yazı tipleri). */
export function themeSig(t: Theme): string {
  return hash(stable({ name: t.name, colors: t.colors, fonts: t.fonts, id: t.id, variant: t.variant }));
}

/** Şeklin görünümü okunduğu andakiyle aynı mı (oranlarda küçük kayan nokta farkı yok sayılır)? */
export function rawSigMatches(s: Shape): boolean {
  const sig = s.raw?.sig;
  if (!sig) return false;
  const [h1, g1 = ""] = rawSig(s).split("|");
  const [h2, g2 = ""] = sig.split("|");
  if (h1 !== h2) return false;
  const a = g1 ? g1.split(",").map(Number) : [];
  const b = g2 ? g2.split(",").map(Number) : [];
  return a.length === b.length && a.every((v, i) => Math.abs(v - b[i]) < 2e-3);
}

/* ------------------------------------------------------------ yazarken */

/** Yazıcının şekil bağlamından gereken kısım (writeDrawing.WriteCtx ile uyumlu). */
export interface RawCtx {
  rels: { add(type: string, target: string, external?: boolean): string };
  part: string;
  nextId: () => number;
  idMap: Map<string, number>;
}

/** Yeni pakette karşılığı olan (taşınmayan) yapısal parçalar. */
const STRUCTURAL = new Set(["slide", "slideLayout", "slideMaster", "notesSlide", "notesMaster", "handoutMaster", "officeDocument"]);

export const RAW_WARN = {
  smartart: "Düzenlenen SmartArt grafikleri şekil grubu olarak kaydedildi; görünüm korundu, SmartArt yapısı kaldırıldı.",
  other: "Düzenlenen gömülü nesneler ve desteklenmeyen öğeler yedek görünümleriyle (resim/şekil) kaydedildi; özgün öğe kaldırıldı.",
  source: "Desteklenmeyen öğelerin özgün kaynağı bulunamadığı için yedek görünümleriyle (şekil grubu/resim) kaydedildi.",
  odp: "SmartArt, gömülü nesneler ve desteklenmeyen öğeler .odp biçiminde yedek görünümleriyle (şekil grubu/resim) kaydedildi.",
} as const;

const GUID_RE = /\{[0-9A-Fa-f]{8}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{12}\}/g;

function newGuid(): string {
  const c = globalThis.crypto;
  const u = c?.randomUUID ? c.randomUUID() : "xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx".replace(/[xy]/g, (ch) => ((Math.random() * 16) | (ch === "y" ? 8 : 0)).toString(16).slice(-1));
  return `{${u.toUpperCase()}}`; // yerel-ayar: ascii
}

/**
 * SmartArt kopyasının model kimliklerini (GUID) yeniden üretir. Veri ve çizim parçası aynı eşlemeyi
 * paylaşır; yalnız parçada `modelId` olarak geçen GUID'ler (ve onlara yapılan başvurular) değişir,
 * uzantı adresleri (`uri="{…}"`) dokunulmaz.
 */
export function regenerateModelIds(memo: Map<string, string>): (xml: string) => string {
  return (xml) => {
    const ids = new Set([...xml.matchAll(/\bmodelId="(\{[0-9A-Fa-f-]{36}\})"/g)].map((m) => m[1].toUpperCase())); // yerel-ayar: ascii
    if (!ids.size) return xml;
    return xml.replace(GUID_RE, (g) => {
      const k = g.toUpperCase(); // yerel-ayar: ascii
      if (!ids.has(k)) return g;
      let n = memo.get(k);
      if (!n) memo.set(k, (n = newGuid()));
      return n;
    });
  };
}

const XFRM_RE = /(<(?:[\w.-]+:)?xfrm\b[^>]*>\s*<(?:[\w.-]+:)?off\s+)x="(-?\d+)"\s+y="(-?\d+)"(\s*\/>\s*<(?:[\w.-]+:)?ext\s+)cx="(\d+)"\s+cy="(\d+)"/g;
const CNVPR_RE = /<(?:[\w.-]+:)?cNvPr\b[^>]*?>/g;
const EMU = 9525;

/** Şeklin kutusu özgün XML'deki ilk xfrm'den farklı boyuttaysa ölçek oranları; aynıysa null. */
function rawScale(xml: string, s: Shape): { sx: number; sy: number } | null {
  const m = new RegExp(XFRM_RE.source).exec(xml);
  if (!m) return null;
  const sx = Number(m[5]) ? Math.round(s.w * EMU) / Number(m[5]) : 1;
  const sy = Number(m[6]) ? Math.round(s.h * EMU) / Number(m[6]) : 1;
  return Math.abs(sx - 1) < 1e-4 && Math.abs(sy - 1) < 1e-4 ? null : { sx, sy };
}

const DRAWING_XFRM_RE = /(<(?:[\w.-]+:)?(?:xfrm|txXfrm)\b[^>]*>\s*<(?:[\w.-]+:)?off\s+)x="(-?\d+)"\s+y="(-?\d+)"(\s*\/>\s*<(?:[\w.-]+:)?ext\s+)cx="(\d+)"\s+cy="(\d+)"/g;

/** Çizim parçasındaki konum/boyutları (çerçeve köşesine göre) ölçekler; grup ölçeklemesiyle aynı sonuç. */
function scaleXfrms(sx: number, sy: number): (xml: string) => string {
  return (xml) =>
    xml.replace(DRAWING_XFRM_RE, (_m, a: string, x: string, y: string, b: string, cx: string, cy: string) => `${a}x="${Math.round(Number(x) * sx)}" y="${Math.round(Number(y) * sy)}"${b}cx="${Math.round(Number(cx) * sx)}" cy="${Math.round(Number(cy) * sy)}"`);
}

/** Açılış etiketinde özniteliği ayarlar (tanımsız → kaldırır). */
function setAttr(tag: string, name: string, value: string | undefined): string {
  const re = new RegExp(`\\s${name}="[^"]*"`);
  if (value === undefined) return tag.replace(re, "");
  if (re.test(tag)) return tag.replace(re, ` ${name}="${esc(value)}"`);
  return tag.replace(/(\/?>)$/, ` ${name}="${esc(value)}"$1`);
}

/** Özgün XML'in konumunu/boyutunu ve cNvPr özniteliklerini modele göre günceller. */
function placeXml(xml: string, s: Shape, id: number): string {
  const first = new RegExp(XFRM_RE.source).exec(xml);
  if (first) {
    const [ox, oy, ow, oh] = [Number(first[2]), Number(first[3]), Number(first[5]), Number(first[6])];
    const nx = Math.round(s.x * EMU);
    const ny = Math.round(s.y * EMU);
    const sx = ow ? Math.round(s.w * EMU) / ow : 1;
    const sy = oh ? Math.round(s.h * EMU) / oh : 1;
    xml = xml.replace(XFRM_RE, (_m, a: string, x: string, y: string, b: string, cx: string, cy: string) => {
      const X = nx + Math.round((Number(x) - ox) * sx);
      const Y = ny + Math.round((Number(y) - oy) * sy);
      return `${a}x="${X}" y="${Y}"${b}cx="${Math.max(0, Math.round(Number(cx) * sx))}" cy="${Math.max(0, Math.round(Number(cy) * sy))}"`;
    });
  }
  const firstTag = xml.match(CNVPR_RE)?.[0];
  const oid = firstTag ? attrOf(firstTag, "id") : undefined;
  if (oid === undefined) return xml;
  const oname = attrOf(firstTag!, "name") ?? "";
  const odescr = attrOf(firstTag!, "descr");
  const ohidden = attrOf(firstTag!, "hidden") === "1" || attrOf(firstTag!, "hidden") === "true";
  return xml.replace(CNVPR_RE, (tag) => {
    if (attrOf(tag, "id") !== oid) return tag;
    let t = setAttr(tag, "id", String(id));
    if (s.name && s.name !== oname) t = setAttr(t, "name", s.name);
    if ((s.descr ?? undefined) !== (odescr || undefined)) t = setAttr(t, "descr", s.descr || undefined);
    if (!!s.hidden !== ohidden) t = setAttr(t, "hidden", s.hidden ? "1" : undefined);
    return t;
  });
}

/**
 * Bir yazma turunun geçiş planı: ham şekillerin özgün XML'i, slayt düzeyi öğeler ve sunu bölümü
 * yaması. Kopyalanan parçaların baytları `carryParts({ carry })` sırasında yazılır.
 */
export class RawWriter {
  readonly carry = new Carry(packageSources);
  /** Kaydetme uyarıları (Türkçe, yinelenmez). */
  readonly warnings = new Set<string>();
  /** Kaynak slayt (`kimlik|parça`) → yeni slayt parçası: ham XML'deki slayt başvuruları için. */
  readonly slideParts = new Map<string, string>();
  private dataSeen = new Set<string>();
  private slideSeen = new Set<string>();
  private shapeSeen = new Set<string>();

  /** İç ilişki hedefini yeni pakette kurar: yapısal parça eşlenir, diğerleri kopyalanır. Olmazsa null. */
  private target(r: RawRef, src: string, inst: CarryInstance): string | null {
    if (r.external) return r.target;
    if (STRUCTURAL.has(relTypeName(r.type))) return relTypeName(r.type) === "slide" ? (this.slideParts.get(`${src}|${r.target}`) ?? null) : null;
    return inst.copy(r.target);
  }

  private relink(xml: string, refs: RawRef[], src: string, inst: CarryInstance, ctx: RawCtx): { xml: string; ids: Map<string, string> } | null {
    const ids = new Map<string, string>();
    for (const r of refs) {
      if (ids.has(r.id)) continue;
      const to = this.target(r, src, inst);
      if (to === null) return null;
      ids.set(r.id, ctx.rels.add(r.type, r.external ? to : relativeTarget(ctx.part, to), r.external));
    }
    return { xml: rewriteRelIds(xml, ids), ids };
  }

  /** Ham şeklin özgün XML'i; yazılamazsa (görünüm değişti, kaynak yok) null ve uyarı: yedek görünüm yazılır. */
  shape(s: Shape, ctx: RawCtx): string | null {
    const raw = s.raw;
    if (!raw) return null;
    if (!rawSigMatches(s)) {
      this.warnings.add(raw.kind === "smartart" ? RAW_WARN.smartart : RAW_WARN.other);
      return null;
    }
    const pkg = kaynakPaket(raw.src);
    if (!pkg) {
      this.warnings.add(RAW_WARN.source);
      return null;
    }
    const inst = this.carry.instance(pkg);
    const linked = this.relink(raw.xml, raw.refs, raw.src, inst, ctx);
    if (!linked) {
      this.warnings.add(RAW_WARN.source);
      return null;
    }
    // SmartArt: veri parçasındaki dsp:dataModelExt relId slaydın çizim ilişkisini anar.
    for (const e of raw.extra) {
      const host = inst.mapped(e.patchIn);
      const to = inst.copy(e.target);
      if (!host || !to) continue;
      const rid = ctx.rels.add(e.type, relativeTarget(ctx.part, to));
      const re = new RegExp(`(\\brelId=")${e.id.replace(/[.*+?^${}()|[\]\\]/g, "\\$&")}"`);
      this.carry.patch(host, (t) => t.replace(re, (_m, a: string) => `${a}${rid}"`));
    }
    if (raw.kind === "smartart") {
      this.regenerateCopies(raw, inst);
      // Boyutu değişen SmartArt: önceden çizilmiş dsp çizimi de aynı oranla ölçeklenir (çerçeveyle uyumlu kalsın).
      const k = rawScale(raw.xml, s);
      if (k) for (const e of raw.extra) {
        const to = inst.mapped(e.target);
        if (to) this.carry.patch(to, scaleXfrms(k.sx, k.sy));
      }
    }
    const id = ctx.nextId();
    ctx.idMap.set(s.id, id);
    const key = `${raw.src}|${raw.xml}`;
    const again = this.shapeSeen.has(key);
    this.shapeSeen.add(key);
    return placeXml(again ? freshCreationIds(linked.xml) : linked.xml, s, id);
  }

  /** Aynı SmartArt bir yazmada ikinci kez yazılıyorsa (çoğaltma) kopyanın model kimlikleri yenilenir. */
  private regenerateCopies(raw: RawXml, inst: CarryInstance): void {
    const data = raw.refs.find((r) => relTypeName(r.type) === "diagramData");
    if (!data) return;
    const key = `${raw.src}|${data.target}`.toLowerCase(); // yerel-ayar: ascii
    if (!this.dataSeen.has(key)) {
      this.dataSeen.add(key);
      return;
    }
    const fn = regenerateModelIds(new Map());
    for (const part of [data.target, ...raw.extra.map((e) => e.target)]) {
      const to = inst.mapped(part);
      if (to) this.carry.patch(to, fn);
    }
  }

  /**
   * İçe aktarılan tema değişmemişse özgün tema parçası (yeni yolda; resimleri kopyalanır): XML ve
   * .rels'i. Değiştiyse, uygulamanın kendi teması ise ya da kaynak yoksa null (tema yeniden üretilir).
   */
  async theme(t: Theme, part: string): Promise<{ xml: string; rels: string | null } | null> {
    const s = t.src;
    if (!s || t.id || themeSig(t) !== s.sig) return null;
    const pkg = kaynakPaket(s.src);
    const text = pkg ? await pkg.text(s.part).catch(() => null) : null;
    if (!pkg || !text) return null;
    const rels = new RelsBuilder();
    const byId = new Map((await pkg.rels(s.part)).map((r) => [r.id, r]));
    const refs: RawRef[] = [];
    for (const id of explicitRelIds(text)) {
      const r = byId.get(id);
      if (r) refs.push({ id, type: r.type, target: r.external ? r.target : (r.part ?? r.target), external: r.external });
    }
    const linked = this.relink(text, refs, s.src, this.carry.instance(pkg), { rels, part, nextId: () => 0, idMap: new Map() });
    if (!linked) return null;
    return { xml: linked.xml.replace(/^﻿?<\?xml[^>]*\?>\s*/, ""), rels: rels.size ? rels.xml() : null };
  }

  /** Slaydın tanınmayan öğeleri (ilişkileri yeniden kurulmuş): cSld sonuna ve p:extLst'e eklenecekler. */
  slide(slide: Slide, ctx: RawCtx): { cSld: string; ext: string[] } {
    const r: SlideRaw | undefined = slide.raw;
    if (!r || (!r.cSld?.length && !r.ext?.length)) return { cSld: "", ext: [] };
    const pkg = kaynakPaket(r.src);
    if (!pkg) return { cSld: "", ext: [] };
    const inst = this.carry.instance(pkg);
    const ids = new Map<string, string>();
    const bad = new Set<string>();
    for (const ref of r.refs) {
      if (ids.has(ref.id) || bad.has(ref.id)) continue;
      const to = this.target(ref, r.src, inst);
      if (to === null) bad.add(ref.id);
      else ids.set(ref.id, ctx.rels.add(ref.type, ref.external ? to : relativeTarget(ctx.part, to), ref.external));
    }
    const fix = (x: string) => ([...explicitRelIds(x)].some((id) => bad.has(id)) ? "" : rewriteRelIds(x, ids));
    const key = `${r.src}|${r.part}`.toLowerCase(); // yerel-ayar: ascii
    // Çoğaltılan slaytta oluşturma kimlikleri yinelenmesin.
    const again = this.slideSeen.has(key);
    this.slideSeen.add(key);
    const out = (x: string) => (again ? freshCreationIds(fix(x)) : fix(x));
    return { cSld: (r.cSld ?? []).map(out).join(""), ext: (r.ext ?? []).map(out).filter(Boolean) };
  }
}

/** Çoğaltılan öğede yinelenmemesi gereken oluşturma kimlikleri: p14:creationId (sayı), a16:creationId (GUID). */
function freshCreationIds(xml: string): string {
  return xml
    .replace(/(<(?:[\w.-]+:)?creationId\b[^>]*\sval=")\d+"/g, (_m, a: string) => `${a}${(Math.random() * 0xffffffff) >>> 0}"`)
    .replace(/(<(?:[\w.-]+:)?creationId\b[^>]*\sid=")\{[0-9A-Fa-f-]{36}\}"/g, (_m, a: string) => `${a}${newGuid()}"`);
}

/* ------------------------------------------------------------ sunu bölümü yaması */

/** Kaynaktan taşınan, modelin yazmadığı sunu öğeleri (sahip olunanlar: asıl slayt/slayt listeleri, boyutlar, özel gösteriler, metin stili). */
const PRES_CARRY = ["handoutMasterIdLst", "smartTags", "embeddedFontLst", "photoAlbum", "custDataLst", "kinsoku", "modifyVerifier"] as const;
/** Kök etiketten taşınan öznitelikler (yazıcınınki önceliklidir). */
const PRES_ATTRS = ["serverZoom", "firstSlideNum", "showSpecialPlsOnTitleSld", "rtl", "removePersonalInfoOnSave", "compatMode", "strictFirstAndLastChars", "embedTrueTypeFonts", "saveSubsetFonts", "autoCompressPictures", "bookmarkIdSeed"];

/**
 * Yazıcının ürettiği presentation.xml'e kaynaktaki öğeleri ekler: dinleyici notu asıl sayfası,
 * gömülü yazı tipleri, akıllı etiketler, fotoğraf albümü, kinsoku, değişiklik doğrulayıcı ve extLst'teki
 * tanınmayan uzantılar (p15:sldGuideLst…). İlişkileri `pctx.rels`'e yeniden kurulur.
 */
export async function patchPresentationXml(xml: string, source: string | undefined, raw: RawWriter, pctx: RawCtx): Promise<string> {
  const pkg = kaynakPaket(source);
  const main = pkg?.mainPart;
  if (!pkg || !main) return xml;
  let srcXml: string | null;
  try {
    srcXml = await pkg.text(main);
  } catch {
    return xml;
  }
  if (!srcXml) return xml;
  let src: XmlPart;
  let out: XmlPart;
  try {
    src = new XmlPart(srcXml, PRESENTATION_ORDER);
    out = new XmlPart(xml, PRESENTATION_ORDER);
  } catch {
    return xml;
  }
  const rels = new Map((await pkg.rels(main)).map((r) => [r.id, r]));
  const inst = raw.carry.instance(pkg);
  const relinked = (frag: string): string | null => {
    const ids = new Map<string, string>();
    for (const id of explicitRelIds(frag)) {
      const r = rels.get(id);
      if (!r) return null;
      const target = r.external ? r.target : r.part;
      // Yapısal parçalar yazıcınındır; dinleyici notu asıl sayfasını yazıcı üretmez, kopyalanır.
      if (!target || (STRUCTURAL.has(relTypeName(r.type)) && relTypeName(r.type) !== "handoutMaster")) return null;
      const to = r.external ? target : inst.copy(target);
      if (!to) return null;
      ids.set(id, pctx.rels.add(r.type, r.external ? to : relativeTarget(pctx.part, to), r.external));
    }
    return rewriteRelIds(frag, ids);
  };
  let changed = false;
  for (const name of PRES_CARRY) {
    const frag = src.get(name);
    if (!frag || out.has(name)) continue;
    const fixed = relinked(frag);
    if (fixed === null) continue;
    out.set(name, fixed);
    changed = true;
  }
  // extLst: yazıcının uzantıları (bölümler) geçerli; kaynaktaki diğer adresler eklenir.
  const srcExt = src.get("extLst");
  if (srcExt) {
    const ours = out.get("extLst");
    const have = new Set((ours ? childElements(ours) : []).map((c) => (attrOf(c.xml, "uri") ?? "").toUpperCase())); // yerel-ayar: ascii
    const add: string[] = [];
    for (const c of childElements(srcExt)) {
      const uri = (attrOf(c.xml, "uri") ?? "").toUpperCase(); // yerel-ayar: ascii
      // Bölümler (p14:sectionLst) modelindir: kullanıcı sildiyse eski slayt kimlikleriyle geri gelmesin.
      if (!uri || have.has(uri) || uri === SECTION_EXT_URI.toUpperCase()) continue; // yerel-ayar: ascii
      const fixed = relinked(c.xml);
      if (fixed) add.push(fixed);
    }
    if (add.length) {
      const prefix = /^<([\w.-]+:)?extLst/.exec(ours ?? srcExt)?.[1] ?? "p:";
      const inner = ours ? childElements(ours).map((c) => c.xml) : [];
      out.set("extLst", `<${prefix}extLst>${[...add, ...inner].join("")}</${prefix}extLst>`);
      changed = true;
    }
  }
  // Kök öznitelikleri ve kaynaktaki öğelerin kullandığı ad alanı bildirimleri
  const at = srcXml.search(/<([\w.-]+:)?presentation\b/);
  const srcOpen = at >= 0 ? srcXml.slice(at, openTagEnd(srcXml, at)) : "";
  for (const a of PRES_ATTRS) {
    const v = attrOf(srcOpen, a);
    if (v !== undefined && attrOf(out.open, a) === undefined) {
      out.open = setAttr(out.open, a, v);
      changed = true;
    }
  }
  if (!changed) return xml;
  for (const m of srcOpen.matchAll(/\sxmlns:([\w.-]+)="([^"]*)"/g)) {
    const cur = new RegExp(`\\sxmlns:${m[1].replace(/[.-]/g, "\\$&")}="([^"]*)"`).exec(out.open)?.[1];
    if (cur === undefined) out.ensureNamespace(m[1], m[2]);
  }
  return out.toString();
}

/** Sunudaki ham şekiller var mı (ODP'ye kayıtta uyarı için)? */
export function hasRawShapes(slides: { shapes: Shape[] }[]): boolean {
  const any = (l: Shape[]): boolean => l.some((s) => !!s.raw || (s.kind === "group" && any(s.children)));
  return slides.some((s) => any(s.shapes));
}
