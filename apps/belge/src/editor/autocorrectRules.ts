// Otomatik Düzelt kuralları (saf işlevler): değiştirme listesi, kesirler, tireler, İKi BÜyük harf,
// cümle başı büyük harf (Türkçe i/İ) ve akıllı tırnaklar. Düzenleyici eklentisi bunları kullanır.

export interface AutoCorrectEntry {
  from: string;
  to: string;
}

export interface AutoCorrectOptions {
  /** İKi BÜyük harfle başlayan sözcükleri düzelt. */
  twoInitialCaps: boolean;
  /** Cümlelerin ilk harfini büyük yap. */
  capFirstSentence: boolean;
  /** Tablo hücrelerinin ilk harfini büyük yap. */
  capFirstCell: boolean;
  /** Yazarken metni değiştir (değiştirme listesi). */
  replaceText: boolean;
  /** "Düz tırnakları" “akıllı tırnaklarla” değiştir. */
  smartQuotes: boolean;
  /** Kesirleri (1/2) kesir karakteriyle (½) değiştir. */
  fractions: boolean;
  /** Kısa çizgileri (--) uzun çizgiyle değiştir. */
  dashes: boolean;
  /** Paragraf başında Sekme / Geri Sil ile sol ve ilk satır girintisini ayarla. */
  tabIndent: boolean;
  entries: AutoCorrectEntry[];
}

/** Varsayılan değiştirme listesi (Türkçe yazım yanlışları ve simgeler). */
export const DEFAULT_ENTRIES: AutoCorrectEntry[] = [
  { from: "(c)", to: "©" },
  { from: "(r)", to: "®" },
  { from: "(tm)", to: "™" },
  { from: "...", to: "…" },
  { from: ":)", to: "☺" },
  { from: ":-)", to: "☺" },
  { from: ":(", to: "☹" },
  { from: ":-(", to: "☹" },
  { from: "-->", to: "→" },
  { from: "<--", to: "←" },
  { from: "==>", to: "⇒" },
  { from: "<==", to: "⇐" },
  { from: "<=>", to: "⇔" },
  { from: "birşey", to: "bir şey" },
  { from: "herşey", to: "her şey" },
  { from: "hiçbirşey", to: "hiçbir şey" },
  { from: "yalnış", to: "yanlış" },
  { from: "yanlız", to: "yalnız" },
  { from: "malesef", to: "maalesef" },
  { from: "şuanda", to: "şu anda" },
  { from: "orjinal", to: "orijinal" },
  { from: "herkez", to: "herkes" },
  { from: "birsürü", to: "bir sürü" },
  { from: "yada", to: "ya da" },
  { from: "laboratuar", to: "laboratuvar" },
  { from: "eşortman", to: "eşofman" },
  { from: "traş", to: "tıraş" },
  { from: "teşekürler", to: "teşekkürler" },
];

export function defaultAutoCorrect(): AutoCorrectOptions {
  return {
    twoInitialCaps: true,
    capFirstSentence: true,
    capFirstCell: true,
    replaceText: true,
    smartQuotes: true,
    fractions: true,
    dashes: true,
    tabIndent: true,
    entries: DEFAULT_ENTRIES.slice(),
  };
}

const STORE = "officehere.belge.autocorrect";
let current: AutoCorrectOptions | null = null;

export function getAutoCorrect(): AutoCorrectOptions {
  if (current) return current;
  let saved: Partial<AutoCorrectOptions> = {};
  try {
    saved = JSON.parse(localStorage.getItem(STORE) ?? "{}") as Partial<AutoCorrectOptions>;
  } catch {
    saved = {};
  }
  current = { ...defaultAutoCorrect(), ...saved };
  return current;
}

export function setAutoCorrect(next: AutoCorrectOptions) {
  current = next;
  try {
    localStorage.setItem(STORE, JSON.stringify(next));
  } catch {
    // tarayıcı depolaması yoksa yalnız bu oturumda geçerli
  }
}

/** Otomatik düzeltmeyi tetikleyen karakterler (sözcük sonu). */
export const DELIMITER = /^[\s.,;:!?)\]}"”’»]$/;

const FRACTIONS: Record<string, string> = { "1/2": "½", "1/4": "¼", "3/4": "¾" };

/** Cümle sonu sayılmayan kısaltmalar (küçük harfle). */
const ABBREVIATIONS = new Set(["vb", "vs", "örn", "bkz", "dr", "prof", "doç", "yrd", "sn", "no", "s", "c", "a.g.e", "age", "vd", "mah", "cad", "sok", "apt", "tel", "st", "av", "müh", "ltd", "şti", "inc", "yy", "bk", "sf", "ör", "çev", "ed", "haz"]);

const lowerTr = (s: string) => s.toLocaleLowerCase("tr-TR");
const upperTr = (s: string) => s.toLocaleUpperCase("tr-TR");
const isLetter = (c: string) => /\p{L}/u.test(c);
const isLower = (c: string) => isLetter(c) && c === lowerTr(c) && c !== upperTr(c);

export interface Correction {
  /** Paragraf içi başlangıç / bitiş (karakter). */
  start: number;
  end: number;
  text: string;
}

/** Sözcüğün büyük/küçük harf biçimini değiştirilen metne uygular. */
function matchCase(src: string, to: string): string {
  const letters = [...src].filter(isLetter);
  if (letters.length > 1 && letters.every((c) => c === upperTr(c) && c !== lowerTr(c))) return upperTr(to);
  const first = letters[0];
  if (first && first === upperTr(first) && first !== lowerTr(first)) return upperTr(to.charAt(0)) + to.slice(1);
  return to;
}

/** Sözcük cümle başında mı (öncesindeki metne bakarak). */
export function isSentenceStart(prefix: string, cellStart: boolean, o: AutoCorrectOptions): boolean {
  if (!prefix.trim()) return cellStart ? o.capFirstCell : o.capFirstSentence;
  if (!o.capFirstSentence) return false;
  const m = /(\S*?)([.!?])["”’)»]*\s+$/u.exec(prefix);
  if (!m) return false;
  if (m[2] !== ".") return true;
  const word = lowerTr(m[1].replace(/^["“‘(«]+/, ""));
  // Sayıdan sonra nokta Türkçede sıra sayısıdır ("3. sınıf"); kısaltmalar cümleyi bitirmez.
  if (/^\d+$/.test(word) || ABBREVIATIONS.has(word) || /^\p{L}$/u.test(word) || word.includes("@")) return false;
  return true;
}

/**
 * Paragrafın imlece kadarki metnindeki son sözcük için düzeltme. `before` imleçten önceki metindir
 * (yazılan ayraç hariç). Değişiklik yoksa null.
 */
export function correctBefore(before: string, o: AutoCorrectOptions, cellStart = false): Correction | null {
  const m = /(\S+)$/u.exec(before);
  if (!m) return null;
  const tokStart = m.index;
  const token = m[1];
  const prefix = before.slice(0, tokStart);
  // Tire: "sözcük -- sözcük" → en tire; "sözcük--sözcük" → uzun tire.
  if (o.dashes) {
    const d = /(\S) -- (\S+)$/u.exec(before);
    if (d) {
      const at = d.index + 1;
      return { start: at, end: at + 4, text: " – " };
    }
    const e = /^(\p{L}+)--(\p{L}+)$/u.exec(token);
    if (e) return { start: tokStart, end: before.length, text: `${e[1]}—${e[2]}` };
  }
  if (/[@/\\]|:\/\/|^www\./i.test(token) && !FRACTIONS[token]) return null;
  let word = token;
  // Değiştirme listesi: tam eşleşme, sonra sondaki noktalama işaretleri ayrılarak.
  if (o.replaceText) {
    const tail = /[.,;:!?"”’)]+$/u.exec(word)?.[0] ?? "";
    const core = tail && tail.length < word.length ? word.slice(0, -tail.length) : word;
    const hit = o.entries.find((x) => x.from === word) ?? o.entries.find((x) => x.from === core) ?? o.entries.find((x) => /\p{L}/u.test(x.from) && lowerTr(x.from) === lowerTr(core));
    if (hit) word = hit.from === word ? hit.to : matchCase(core, hit.to) + (hit.from === core || lowerTr(hit.from) === lowerTr(core) ? tail : "");
    // Harf/rakamla başlamayan simge girdileri sözcüğe bitişik de düzeltilir ("Bekle..." → "Bekle…", "Marka(c)" → "Marka©").
    else {
      const sym = o.entries.find((x) => x.from.length >= 2 && !/^[\p{L}\p{N}]/u.test(x.from) && word.endsWith(x.from) && word.length > x.from.length);
      const punct = /[.,;:!?]+$/u.exec(word)?.[0] ?? "";
      const body = word.slice(0, word.length - punct.length);
      const symCore = !sym && punct ? o.entries.find((x) => x.from.length >= 2 && !/^[\p{L}\p{N}]/u.test(x.from) && body.endsWith(x.from) && body.length > x.from.length) : undefined;
      if (sym) word = word.slice(0, -sym.from.length) + sym.to;
      else if (symCore) word = body.slice(0, -symCore.from.length) + symCore.to + punct;
    }
  }
  if (o.fractions && FRACTIONS[word]) word = FRACTIONS[word];
  // İKi BÜyük harf
  if (o.twoInitialCaps) {
    const lead = /^["“‘(«]*/u.exec(word)?.[0] ?? "";
    const core = word.slice(lead.length);
    const chars = [...core];
    if (chars.length >= 3 && /^\p{Lu}\p{Lu}\p{Ll}+[.,;:!?]*$/u.test(core)) word = lead + chars[0] + lowerTr(chars[1]) + chars.slice(2).join("");
  }
  // Cümle başı büyük harf
  if (isSentenceStart(prefix, cellStart, o)) {
    const i = [...word].findIndex(isLetter);
    const chars = [...word];
    if (i >= 0 && isLower(chars[i]) && chars.slice(0, i).every((c) => /["“‘(«¿¡]/u.test(c))) {
      chars[i] = upperTr(chars[i]);
      word = chars.join("");
    }
  }
  if (word === token) return null;
  return { start: tokStart, end: before.length, text: word };
}

/** Düz tırnak yerine açılış ya da kapanış akıllı tırnağı. */
export function smartQuote(prev: string | undefined, quote: '"' | "'"): string {
  const opening = !prev || /[\s([{“‘«\-–— ￼\n]/u.test(prev);
  if (quote === '"') return opening ? "“" : "”";
  return opening ? "‘" : "’";
}
