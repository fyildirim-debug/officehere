/**
 * 97-2003 sunu (.ppt, MS-PPT) okuyucusu. Zincir: Current User →
 * UserEditAtom → PersistDirectoryAtom (kalıcı nesne tablosu) →
 * DocumentContainer → slayt/asıl slayt/not listeleri → her kabın PPDrawing
 * şekilleri. Yalnız okunur; şifreli dosyalar anlaşılır bir hatayla reddedilir.
 */
import { isCfb, readCfb, type CfbFile } from "@officehere/core";
import { uid } from "../../model/factory";
import { newPresentation } from "../../model/presentation";
import type { LevelStyle, Presentation, Shape, Slide, SpShape, TextBody, Theme } from "../../model/types";
import { applyPrimaryMaster, chooseLayout, layoutResolver, type ImportMaster } from "../importBuild";
import { blipStore } from "./blips";
import { readDrawing, readProps, type DrawCtx } from "./drawing";
import { PptError, RT, Stream, type Rec } from "./records";
import { hex, masterLevels, textItems, type TextItem } from "./text";

export { PptError };

const MAX_SLIDES = 2000;
const HEADER_TOKEN = 0xe391c05f;
const HEADER_TOKEN_CRYPT = 0xf3d1c4df;
const ENCRYPTED = "Bu sunu parola ile korunuyor (şifreli); açılamıyor.";

interface PersistEntry {
  persistId: number;
  /** Kayıt kimliği (slideId / masterId / notesId). */
  id: number;
  texts: TextItem[];
}

function openCfb(data: Uint8Array): CfbFile {
  if (!isCfb(data)) throw new PptError("Dosya bir 97-2003 sunusu değil (bileşik dosya imzası yok).");
  try {
    return readCfb(data);
  } catch {
    throw new PptError("Dosya bozuk; bileşik dosya yapısı okunamadı.");
  }
}

/** Kalıcı nesne tablosu (persistId → akış konumu) ve belge kimliği. */
function persistDirectory(doc: Stream, cu: Stream | undefined): { dir: Map<number, number>; docRef: number } {
  let off = -1;
  if (cu) {
    const h = cu.header(0);
    if (h?.type === RT.CurrentUserAtom) {
      const token = cu.u32(h.start + 4);
      if (token === HEADER_TOKEN_CRYPT) throw new PptError(ENCRYPTED);
      if (token === HEADER_TOKEN) off = cu.u32(h.start + 8);
    }
  }
  const edits: Rec[] = [];
  const seen = new Set<number>();
  while (off >= 0 && off < doc.length && !seen.has(off) && edits.length < 1000) {
    seen.add(off);
    const r = doc.header(off);
    if (!r || r.type !== RT.UserEditAtom) break;
    edits.push(r);
    const prev = doc.u32(r.start + 8);
    if (!prev || prev === off) break;
    off = prev;
  }
  if (!edits.length) {
    // Geçerli kullanıcı bilgisi yoksa akıştaki son düzenleme kaydı kullanılır.
    const last = doc.list(0, doc.length).filter((r) => r.type === RT.UserEditAtom).pop();
    if (last) edits.push(last);
  }
  if (!edits.length) throw new PptError("Sunu akışında düzenleme kaydı bulunamadı; dosya bozuk olabilir.");
  const dir = new Map<number, number>();
  for (const e of [...edits].reverse()) {
    const pd = doc.header(doc.u32(e.start + 12));
    if (!pd || pd.type !== RT.PersistDirectoryAtom) continue;
    let o = pd.start;
    const end = pd.start + pd.len;
    while (o + 4 <= end) {
      const v = doc.u32(o);
      const id = v & 0xfffff;
      const n = v >>> 20;
      o += 4;
      for (let k = 0; k < n && o + 4 <= end; k++, o += 4) dir.set(id + k, doc.u32(o));
    }
  }
  const newest = edits[0];
  if (newest.len >= 32) {
    const cryptRef = doc.u32(newest.start + 28);
    const at = cryptRef ? dir.get(cryptRef) : undefined;
    if (at !== undefined && doc.header(at)?.type === RT.CryptSession10Container) throw new PptError(ENCRYPTED);
  }
  return { dir, docRef: doc.u32(newest.start + 16) };
}

/** SlideListWithText: kalıcı nesne girdileri ve her birinin anahat metinleri. */
function persistList(doc: Stream, list: Rec | undefined): PersistEntry[] {
  const out: PersistEntry[] = [];
  let cur: { e: PersistEntry; recs: Rec[] } | undefined;
  const flush = () => {
    if (cur) cur.e.texts = textItems(doc, cur.recs);
  };
  for (const r of doc.kids(list)) {
    if (r.type === RT.SlidePersistAtom) {
      flush();
      cur = { e: { persistId: doc.u32(r.start), id: doc.u32(r.start + 12), texts: [] }, recs: [] };
      out.push(cur.e);
      if (out.length > MAX_SLIDES) break;
    } else cur?.recs.push(r);
  }
  flush();
  return out;
}

/** Renk düzeni atomu → 8 renk. */
function schemeOf(doc: Stream, c: Rec | undefined): string[] | undefined {
  const r = doc.kids(c).find((k) => k.type === RT.ColorSchemeAtom && k.inst === 1);
  if (!r || r.len < 32) return undefined;
  return Array.from({ length: 8 }, (_, i) => hex(doc.u8(r.start + i * 4), doc.u8(r.start + i * 4 + 1), doc.u8(r.start + i * 4 + 2)));
}

const DEFAULT_SCHEME = ["#FFFFFF", "#000000", "#808080", "#000000", "#BBE0E3", "#333399", "#009999", "#99CC00"];

/** Metin türü → model metin stili grubu. */
const KIND = (t: number): "title" | "body" | "other" => (t === 0 || t === 6 ? "title" : t === 1 || t === 5 || t === 7 || t === 8 ? "body" : "other");

export async function readPpt(data: ArrayBuffer | Uint8Array): Promise<Presentation> {
  const bytes = data instanceof Uint8Array ? data : new Uint8Array(data);
  const cfb = openCfb(bytes);
  if (!cfb.has("PowerPoint Document")) {
    if (cfb.has("EncryptedPackage")) throw new PptError(ENCRYPTED);
    if (cfb.has("WordDocument")) throw new PptError("Bu dosya bir sunu değil; bir kelime işlemci belgesi.");
    if (cfb.has("Workbook") || cfb.has("Book")) throw new PptError("Bu dosya bir sunu değil; bir hesap tablosu.");
    throw new PptError("Bu sunu desteklenmeyen çok eski bir biçimde ya da dosya bozuk.");
  }
  const doc = new Stream(cfb.read("PowerPoint Document")!);
  const cuBytes = cfb.read("Current User");
  const { dir, docRef } = persistDirectory(doc, cuBytes ? new Stream(cuBytes) : undefined);
  const at = (persistId: number, ...types: number[]) => {
    const o = dir.get(persistId);
    const r = o !== undefined ? doc.header(o) : undefined;
    return r && (!types.length || types.includes(r.type)) ? r : undefined;
  };
  const docRec = at(docRef, RT.Document) ?? doc.list(0, doc.length).find((r) => r.type === RT.Document);
  if (!docRec) throw new PptError("Sunu belgesi kaydı bulunamadı; dosya bozuk olabilir.");
  if (doc.list(0, doc.length).some((r) => r.type === RT.CryptSession10Container)) throw new PptError(ENCRYPTED);

  // Slayt boyutu (576/inç → 96 dpi: /6).
  const da = doc.kid(docRec, RT.DocumentAtom);
  const sw = da ? doc.i32(da.start) / 6 : 960;
  const shh = da ? doc.i32(da.start + 4) / 6 : 720;
  const width = sw >= 96 && sw <= 96 * 56 ? sw : 960;
  const height = shh >= 96 && shh <= 96 * 56 ? shh : 720;

  // Yazı tipleri.
  const env = doc.kid(docRec, RT.Environment);
  const fonts: string[] = [];
  for (const f of doc.kids(doc.kid(env, RT.FontCollection))) {
    if (f.type !== RT.FontEntityAtom) continue;
    let name = doc.utf16(f.start, 32);
    const z = name.indexOf("\0");
    if (z >= 0) name = name.slice(0, z);
    fonts[f.inst] = name.trim();
  }

  // Çizim grubu: resim deposu ve varsayılan şekil özellikleri.
  const dgg = doc.kid(doc.kid(docRec, RT.DrawingGroup), RT.DggContainer);
  const blip = blipStore(doc, dgg, cfb.read("Pictures"));
  const defaultsProps = readProps(doc, [doc.kid(dgg, RT.FOPT)]);
  const defaults = new Map<number, number>();
  for (const pid of [0x180, 0x181, 0x183, 0x1bf, 0x1c0, 0x1cb, 0x1ff]) {
    const v = defaultsProps.get(pid);
    if (v !== undefined) defaults.set(pid, v);
  }

  const lists = doc.kids(docRec).filter((r) => r.type === RT.SlideListWithText);
  const slideEntries = persistList(doc, lists.find((r) => r.inst === 0));
  const masterEntries = persistList(doc, lists.find((r) => r.inst === 1));
  const notesEntries = persistList(doc, lists.find((r) => r.inst === 2));

  const pres = newPresentation("varsayilan", 0, { width, height });
  pres.slides = [];
  pres.props = {};
  const count = { n: 0 };

  // Asıl slaytlar.
  interface MasterData extends ImportMaster {
    scheme: string[];
    styles: Map<number, LevelStyle[]>;
  }
  const masters = new Map<string, MasterData>();
  const docStyles = new Map<number, LevelStyle[]>();
  for (const r of doc.kids(env)) if (r.type === RT.TextMasterStyleAtom) docStyles.set(r.inst, masterLevels(doc, r, { fonts, scheme: DEFAULT_SCHEME }));
  for (const me of masterEntries) {
    const rec = at(me.persistId, RT.MainMaster, RT.Slide);
    if (!rec) continue;
    const scheme = schemeOf(doc, rec) ?? DEFAULT_SCHEME;
    const styles = new Map(docStyles);
    for (const r of doc.kids(rec)) if (r.type === RT.TextMasterStyleAtom) styles.set(r.inst, masterLevels(doc, r, { fonts, scheme }));
    const ctx: DrawCtx = { s: doc, fonts, scheme, blip, defaults, outline: me.texts, levels: () => undefined, count };
    const d = readDrawing(doc, doc.kid(rec, RT.Drawing), ctx);
    const phs = d.shapes.filter((s): s is SpShape => s.kind === "sp" && !!s.ph);
    masters.set(String(me.id), {
      key: String(me.id),
      display: `Asıl ${masters.size + 1}`,
      bg: d.background ?? { type: "solid", color: { rgb: scheme[0] } },
      deco: d.shapes.filter((s) => !(s.kind === "sp" && s.ph)),
      title: phs.find((s) => s.ph!.type === "title" || s.ph!.type === "ctrTitle"),
      body: phs.find((s) => s.ph!.type === "body" || s.ph!.type === "obj"),
      scheme,
      styles,
    });
  }

  // Not sayfaları: slayt kimliği → not metni.
  const notesBySlide = new Map<number, string>();
  for (const ne of notesEntries) {
    const rec = at(ne.persistId, RT.Notes);
    if (!rec) continue;
    const na = doc.kid(rec, RT.NotesAtom);
    const slideRef = na ? doc.u32(na.start) : 0;
    const ctx: DrawCtx = { s: doc, fonts, scheme: DEFAULT_SCHEME, blip, defaults, outline: ne.texts, levels: () => undefined, count: { n: 0 } };
    const d = readDrawing(doc, doc.kid(rec, RT.Drawing), ctx);
    const hasText = (s: Shape | undefined) => s?.kind === "sp" && !!s.text?.paragraphs.some((p) => p.runs.some((r) => r.text.trim()));
    const body = d.all.find((r) => r.ph === 12 && hasText(r.shape))?.shape ?? d.all.find((r) => r.ph !== 11 && r.ph !== 5 && hasText(r.shape))?.shape;
    const text = body?.kind === "sp" && body.text ? body.text.paragraphs.map((p) => p.runs.map((r) => r.text.replace(/\v/g, "\n")).join("")).join("\n").replace(/\s+$/, "") : ne.texts.filter((t) => t.type === 2).map((t) => t.text.replace(/\r/g, "\n")).join("\n");
    if (slideRef && text) notesBySlide.set(slideRef, text);
  }

  // Slaytlar.
  const firstSlide = slideEntries.map((e) => at(e.persistId, RT.Slide)).find(Boolean);
  const firstMasterRef = firstSlide ? doc.u32((doc.kid(firstSlide, RT.SlideAtom)?.start ?? 0) + 12) : 0;
  const primary = masters.get(String(firstMasterRef)) ?? masters.values().next().value;
  const toLevels = (m: MasterData | undefined) => (t: number): TextBody["levels"] => {
    const st = m?.styles.get(t) ?? (KIND(t) === "title" ? m?.styles.get(0) : KIND(t) === "body" ? m?.styles.get(1) : m?.styles.get(4));
    if (!st?.length) return undefined;
    return Array.from({ length: 9 }, (_, i) => structuredClone(st[Math.min(i, st.length - 1)]));
  };
  if (primary) {
    applyPrimaryMaster(pres, primary);
    const ts = pres.master.textStyles;
    const merge = (base: LevelStyle[], src: LevelStyle[] | undefined) => base.map((b, i) => (src?.[Math.min(i, src.length - 1)] ? { para: { ...b.para, ...src[Math.min(i, src.length - 1)].para }, run: { ...b.run, ...src[Math.min(i, src.length - 1)].run } } : b));
    ts.title = merge(ts.title, primary.styles.get(0));
    ts.body = merge(ts.body, primary.styles.get(1));
    const theme: Theme = {
      ...pres.theme,
      name: "Asıl",
      id: undefined,
      variant: undefined,
      colors: { ...pres.theme.colors, lt1: primary.scheme[0], dk1: primary.scheme[1], lt2: primary.scheme[2], dk2: primary.scheme[3], accent1: primary.scheme[4], accent2: primary.scheme[5], accent3: primary.scheme[6], hlink: primary.scheme[6], folHlink: primary.scheme[7] },
    };
    const tf = primary.styles.get(0)?.[0]?.run.font;
    const bf = primary.styles.get(1)?.[0]?.run.font;
    if (tf) theme.fonts = { ...theme.fonts, major: tf };
    if (bf) theme.fonts = { ...theme.fonts, minor: bf };
    pres.theme = theme;
  }
  const layoutFor = layoutResolver(pres, masters, primary);

  for (const se of slideEntries.slice(0, MAX_SLIDES)) {
    const rec = at(se.persistId, RT.Slide);
    if (!rec) continue;
    const sa = doc.kid(rec, RT.SlideAtom);
    const masterRef = sa ? doc.u32(sa.start + 12) : 0;
    const flags = sa ? doc.u16(sa.start + 20) : 7;
    const m = masters.get(String(masterRef)) ?? primary;
    const scheme = (flags & 2 ? undefined : schemeOf(doc, rec)) ?? m?.scheme ?? DEFAULT_SCHEME;
    const ctx: DrawCtx = { s: doc, fonts, scheme, blip, defaults, outline: se.texts, levels: toLevels(m), count };
    const d = readDrawing(doc, doc.kid(rec, RT.Drawing), ctx);
    const layout = layoutFor(m?.key, chooseLayout(d.shapes));
    const slide: Slide = { id: uid("sl"), layoutId: layout.id, shapes: d.shapes, notes: notesBySlide.get(se.id) ?? "", animations: [] };
    if (!(flags & 4) && d.background) slide.background = d.background;
    if (!(flags & 1)) slide.hideMasterShapes = true;
    const info = doc.kid(rec, RT.SlideShowSlideInfoAtom);
    if (info && doc.u16(info.start + 10) & 0x4) slide.hidden = true;
    pres.slides.push(slide);
  }
  if (!pres.slides.length) throw new PptError("Sunuda slayt bulunamadı.");
  return pres;
}
