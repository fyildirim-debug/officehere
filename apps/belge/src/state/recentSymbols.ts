// Son kullanılan simgeler (Ekle › Simge menüsü ve Simge iletişim kutusu): bu tarayıcıda/kullanıcıda saklanır.
import { COMMON_SYMBOLS } from "../templates/parts";

const KEY = "officehere.belge.sonSimgeler";
const MAX = 20;

export function recentSymbols(): string[] {
  try {
    const v = JSON.parse(localStorage.getItem(KEY) ?? "[]") as unknown;
    return Array.isArray(v) ? v.filter((x): x is string => typeof x === "string" && !!x).slice(0, MAX) : [];
  } catch {
    return [];
  }
}

/** Kullanılan simgeyi listenin başına alır. */
export function rememberSymbol(s: string) {
  const next = [s, ...recentSymbols().filter((x) => x !== s)].slice(0, MAX);
  try {
    localStorage.setItem(KEY, JSON.stringify(next));
  } catch {
    // Depolama kapalıysa liste yalnız varsayılan simgelerden oluşur.
  }
}

/** Simge menüsündeki ızgara: önce son kullanılanlar, kalan yer sık kullanılan simgelerle dolar. */
export function symbolGallery(size = COMMON_SYMBOLS.length): string[] {
  const out = recentSymbols();
  for (const s of COMMON_SYMBOLS) {
    if (out.length >= size) break;
    if (!out.includes(s)) out.push(s);
  }
  return out.slice(0, size);
}
