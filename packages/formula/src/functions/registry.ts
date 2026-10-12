/**
 * Fonksiyon kaydı: kanonik İngilizce ad, Türkçe ad, kategori, bağımsız
 * değişken imzası (Türkçe) ve açıklama.
 */
import type { Node } from "../ast";
import type { FnContext } from "../context";
import type { Value } from "../values";

export type Category = "math" | "stats" | "logical" | "text" | "lookup" | "date" | "info" | "financial" | "engineering" | "database";

export const CATEGORY_LABELS: Record<Category, string> = {
  financial: "Finansal",
  date: "Tarih ve Saat",
  math: "Matematik ve Trig",
  stats: "İstatistik",
  lookup: "Arama ve Başvuru",
  database: "Veritabanı",
  text: "Metin",
  logical: "Mantıksal",
  info: "Bilgi",
  engineering: "Mühendislik",
};

export type Thunk = () => Value;
export type FnImpl = (args: Value[], ctx: FnContext) => Value;
export type LazyImpl = (args: Thunk[], ctx: FnContext, nodes: Node[]) => Value;

export interface FnDef {
  name: string;
  tr: string;
  category: Category;
  /** Türkçe bağımsız değişken imzası: "sayı1; [sayı2]; ...". */
  args: string;
  desc: string;
  min: number;
  max: number;
  impl?: FnImpl;
  lazy?: LazyImpl;
  /** Skaler bağımsız değişkenler: dizi/aralık verilirse fonksiyon öğe öğe uygulanır. */
  lift?: true | number[];
  volatile?: boolean;
  /** Dosyaya yazarken "_xlfn." öneki gerekir. */
  xlfn?: boolean;
}

export const FUNCTIONS = new Map<string, FnDef>();
const TR_INDEX = new Map<string, string>();

export function register(defs: FnDef[]): void {
  for (const d of defs) {
    FUNCTIONS.set(d.name, d);
    TR_INDEX.set(d.tr.toLocaleUpperCase("tr-TR"), d.name);
  }
}

export function getFunction(name: string): FnDef | undefined {
  return FUNCTIONS.get(name);
}

/** Türkçe ya da İngilizce adı kanonik ada çevirir. */
export function resolveFunctionName(raw: string, preferTr = true): string | undefined {
  const tr = raw.toLocaleUpperCase("tr-TR");
  const en = raw.toUpperCase();
  if (preferTr) {
    const t = TR_INDEX.get(tr) ?? TR_INDEX.get(en);
    if (t) return t;
  }
  if (FUNCTIONS.has(en)) return en;
  if (FUNCTIONS.has(tr)) return tr;
  if (!preferTr) return TR_INDEX.get(tr) ?? TR_INDEX.get(en);
  return undefined;
}

export function trNameOf(canonical: string): string {
  return FUNCTIONS.get(canonical)?.tr ?? canonical;
}

/** Kısa tanım yardımcısı. */
export function fn(
  name: string,
  tr: string,
  category: Category,
  args: string,
  desc: string,
  min: number,
  max: number,
  impl: FnImpl,
  opts: Partial<FnDef> = {},
): FnDef {
  return { name, tr, category, args, desc, min, max, impl, ...opts };
}

export function lazyFn(name: string, tr: string, category: Category, args: string, desc: string, min: number, max: number, lazy: LazyImpl, opts: Partial<FnDef> = {}): FnDef {
  return { name, tr, category, args, desc, min, max, lazy, ...opts };
}

/** İmzadaki bağımsız değişken adlarını listeler. */
export function argNames(def: FnDef): string[] {
  return def.args
    .split(";")
    .map((s) => s.trim())
    .filter(Boolean);
}
