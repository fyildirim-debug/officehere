/** Hücreleri Biçimlendir > Sayı sekmesinin kategori ve kod üretimi (saf, test edilebilir). */
import { isDateFormat, isTimeOnlyFormat } from "@officehere/formula";

export type NumCategory = "general" | "number" | "currency" | "accounting" | "date" | "time" | "percent" | "fraction" | "scientific" | "text" | "special" | "custom";

export const CATEGORIES: { id: NumCategory; label: string; desc: string }[] = [
  { id: "general", label: "Genel", desc: "Genel biçimli hücrelerin belirli bir sayı biçimi yoktur." },
  { id: "number", label: "Sayı", desc: "Sayı, genel sayıların görüntülenmesinde kullanılır. Parasal değerler için Para Birimi ve Finansal biçimler kullanılır." },
  { id: "currency", label: "Para Birimi", desc: "Para birimi biçimleri genel parasal değerler için kullanılır." },
  { id: "accounting", label: "Finansal", desc: "Finansal biçimler para birimi simgelerini ve ondalık virgüllerini sütunda hizalar." },
  { id: "date", label: "Tarih", desc: "Tarih biçimleri, tarih ve saat seri numaralarını tarih değerleri olarak görüntüler." },
  { id: "time", label: "Saat", desc: "Saat biçimleri, tarih ve saat seri numaralarını saat değerleri olarak görüntüler." },
  { id: "percent", label: "Yüzde", desc: "Yüzde biçimleri hücre değerini 100 ile çarpar ve sonucu yüzde simgesiyle görüntüler." },
  { id: "fraction", label: "Kesir", desc: "Değeri kesir olarak görüntüler." },
  { id: "scientific", label: "Bilimsel", desc: "Sayıyı bilimsel gösterimle görüntüler." },
  { id: "text", label: "Metin", desc: "Metin biçimli hücreler, sayı girilse bile metin olarak ele alınır. Hücre, girildiği gibi görüntülenir." },
  { id: "special", label: "Özel", desc: "Özel biçimler liste ve veritabanı değerlerinin izlenmesinde kullanılır." },
  { id: "custom", label: "İsteğe Uyarlanmış", desc: "Var olan kodlardan birini başlangıç noktası olarak kullanarak sayı biçimi kodunu yazın." },
];

export const CURRENCY_SYMBOLS: { id: string; label: string }[] = [
  { id: "", label: "Yok" },
  { id: "₺", label: "₺ Türkçe (Türkiye)" },
  { id: "$", label: "$ İngilizce (ABD)" },
  { id: "€", label: "€ Euro" },
  { id: "£", label: "£ İngilizce (Birleşik Krallık)" },
  { id: "TL", label: "TL" },
];

export const DATE_TYPES = ["dd.mm.yyyy", "d.m.yyyy", "d mmmm yyyy", "d mmmm yyyy dddd", "dd.mm.yy", "d.m.yy", "d mmm yy", "mmmm yy", "mmm.yy", "d mmm", "yyyy-mm-dd", "d.m.yyyy h:mm", "dd.mm.yyyy hh:mm"];
export const TIME_TYPES = ["hh:mm:ss", "h:mm", "hh:mm", "h:mm AM/PM", "h:mm:ss AM/PM", "[h]:mm:ss", "mm:ss", "mm:ss.0", "dd.mm.yyyy hh:mm"];
export const FRACTION_TYPES: { code: string; label: string }[] = [
  { code: "# ?/?", label: "En çok bir basamak (1/4)" },
  { code: "# ??/??", label: "En çok iki basamak (21/25)" },
  { code: "# ???/???", label: "En çok üç basamak (312/943)" },
  { code: "# ?/2", label: "Yarımlar olarak (1/2)" },
  { code: "# ?/4", label: "Çeyrekler olarak (2/4)" },
  { code: "# ?/8", label: "Sekizde birler olarak (4/8)" },
  { code: "# ??/16", label: "On altıda birler olarak (8/16)" },
  { code: "# ?/10", label: "Onda birler olarak (3/10)" },
  { code: "# ??/100", label: "Yüzde birler olarak (30/100)" },
];
export const SPECIAL_TYPES: { code: string; label: string }[] = [
  { code: "00000", label: "Posta Kodu" },
  { code: "(000) 000 00 00", label: "Telefon Numarası" },
  { code: "0 (000) 000 00 00", label: "Telefon Numarası (0 ile)" },
  { code: "00000000000", label: "T.C. Kimlik Numarası" },
  { code: "TR00 0000 0000 0000 0000 0000 00", label: "IBAN" },
];
export const CUSTOM_CODES = [
  "General",
  "0",
  "0.00",
  "#,##0",
  "#,##0.00",
  "#,##0;-#,##0",
  "#,##0;[Red]-#,##0",
  "#,##0.00;[Red]-#,##0.00",
  '#,##0 "₺"',
  '#,##0.00 "₺"',
  '#,##0.00 "₺";[Red]-#,##0.00 "₺"',
  "0%",
  "0.00%",
  "0.00E+00",
  "##0.0E+0",
  "# ?/?",
  "# ??/??",
  "dd.mm.yyyy",
  "d mmmm yyyy",
  "d mmm yy",
  "mmmm yy",
  "hh:mm",
  "hh:mm:ss",
  "[h]:mm:ss",
  "dd.mm.yyyy hh:mm",
  "mm:ss.0",
  "@",
  '_-* #,##0 "₺"_-;-* #,##0 "₺"_-;_-* "-" "₺"_-;_-@_-',
  '_-* #,##0.00 "₺"_-;-* #,##0.00 "₺"_-;_-* "-"?? "₺"_-;_-@_-',
];

export interface NumOpts {
  decimals: number;
  thousands: boolean;
  negative: number;
  symbol: string;
}

export const NEGATIVE_SAMPLES = ["-1.234,10", "1.234,10 (kırmızı)", "-1.234,10 (kırmızı)", "(1.234,10)", "(1.234,10) (kırmızı)"];

function base(decimals: number, thousands: boolean): string {
  return (thousands ? "#,##0" : "0") + (decimals > 0 ? "." + "0".repeat(decimals) : "");
}

function withSymbol(b: string, symbol: string): string {
  if (!symbol) return b;
  // ₺ ve TL sağda, diğerleri solda
  if (symbol === "₺" || symbol === "TL") return `${b} "${symbol}"`;
  return `"${symbol}"${b}`;
}

function negatives(x: string, n: number): string {
  switch (n) {
    case 1:
      return `${x};[Red]${x}`;
    case 2:
      return `${x};[Red]-${x}`;
    case 3:
      return `${x}_);(${x})`;
    case 4:
      return `${x}_);[Red](${x})`;
    default:
      return x;
  }
}

/** Kategori ve seçeneklerden biçim kodu. */
export function buildCode(cat: NumCategory, o: NumOpts, typeCode?: string): string {
  switch (cat) {
    case "general":
      return "General";
    case "number":
      return negatives(base(o.decimals, o.thousands), o.negative);
    case "currency":
      return negatives(withSymbol(base(o.decimals, true), o.symbol), o.negative);
    case "accounting": {
      const b = base(o.decimals, true);
      const q = o.decimals > 0 ? "?".repeat(o.decimals) : "";
      const s = o.symbol ? ` "${o.symbol}"` : "";
      return `_-* ${b}${s}_-;-* ${b}${s}_-;_-* "-"${q}${s}_-;_-@_-`;
    }
    case "percent":
      return base(o.decimals, false) + "%";
    case "scientific":
      return base(o.decimals, false) + "E+00";
    case "text":
      return "@";
    default:
      return typeCode ?? "General";
  }
}

/** Var olan koddan kategori ve seçenekleri çıkarır. */
export function detectCategory(code: string | undefined): { cat: NumCategory; opts: NumOpts } {
  const opts: NumOpts = { decimals: 2, thousands: false, negative: 0, symbol: "₺" };
  const c = code ?? "General";
  if (c === "General") return { cat: "general", opts };
  if (c === "@") return { cat: "text", opts };
  if (DATE_TYPES.includes(c) && !TIME_TYPES.includes(c)) return { cat: "date", opts };
  if (TIME_TYPES.includes(c)) return { cat: "time", opts };
  if (FRACTION_TYPES.some((f) => f.code === c)) return { cat: "fraction", opts };
  if (SPECIAL_TYPES.some((f) => f.code === c)) return { cat: "special", opts };
  const decOf = (s: string) => {
    const m = /\.(0+)/.exec(s);
    return m ? m[1].length : 0;
  };
  for (const cat of ["number", "currency", "accounting", "percent", "scientific"] as const) {
    for (let d = 0; d <= 10; d++) {
      for (const th of cat === "number" ? [false, true] : [true]) {
        for (const sym of cat === "currency" || cat === "accounting" ? CURRENCY_SYMBOLS.map((s) => s.id) : ["₺"]) {
          for (let n = 0; n < (cat === "number" || cat === "currency" ? 5 : 1); n++) {
            const o = { decimals: d, thousands: th, negative: n, symbol: sym };
            if (buildCode(cat, o) === c) return { cat, opts: o };
          }
        }
      }
    }
  }
  if (isTimeOnlyFormat(c)) return { cat: "time", opts };
  if (isDateFormat(c)) return { cat: "custom", opts: { ...opts, decimals: decOf(c) } };
  return { cat: "custom", opts };
}
