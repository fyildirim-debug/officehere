// Resmî yazı şablonlarının form alanları, varsayılanları ve başlık (antet) türleri.
import type { YaziTipi, ZarfBoyu } from "./olcu";

/** m.10: başlık türleri. İl/ilçe teşkilatında 2. satır mülki idare (m.10/3); bağlı/ilgili idarede 4 satır (m.10/2). */
export type BaslikTuru = "merkez" | "bagli" | "il" | "ilMem" | "ilceMem" | "okul";

export const BASLIK_TURLERI: { id: BaslikTuru; etiket: string; idare: string; araIdare?: string; birim: string }[] = [
  { id: "merkez", etiket: "Bakanlık ya da merkez birimi", idare: "ÖRNEK BAKANLIĞI", birim: "Destek Hizmetleri Dairesi Başkanlığı" },
  { id: "bagli", etiket: "Bağlı ya da ilgili kurum (4 satır)", idare: "ÖRNEK BAKANLIĞI", araIdare: "Örnek Kurumu Başkanlığı", birim: "Personel Dairesi Başkanlığı" },
  { id: "il", etiket: "Valilik (il müdürlüğü)", idare: "ÖRNEK VALİLİĞİ", birim: "İl Örnek Müdürlüğü" },
  { id: "ilMem", etiket: "İl Millî Eğitim Müdürlüğü", idare: "ÖRNEK VALİLİĞİ", birim: "İl Millî Eğitim Müdürlüğü" },
  { id: "ilceMem", etiket: "İlçe Millî Eğitim Müdürlüğü", idare: "ÖRNEK KAYMAKAMLIĞI", birim: "İlçe Millî Eğitim Müdürlüğü" },
  { id: "okul", etiket: "Okul müdürlüğü", idare: "ÖRNEK KAYMAKAMLIĞI", birim: "Örnek Ortaokulu Müdürlüğü" },
];

/** Muhatap düzeyi: kapanış ibaresini belirler (m.16/12). */
export type Yon = "ust" | "alt" | "karma" | "kisi";

export const YONLER: { id: Yon; etiket: string }[] = [
  { id: "ust", etiket: "Üst ya da aynı düzey makam (arz ederim)" },
  { id: "alt", etiket: "Alt makam (rica ederim)" },
  { id: "karma", etiket: "Üst ve alt makamlar birlikte (arz ve rica ederim)" },
  { id: "kisi", etiket: "Gerçek kişi (Saygılarımla)" },
];

/** m.16/12 a, b, e. */
export const KAPANIS: Record<Yon, string> = {
  ust: "Bilgilerinizi ve gereğini arz ederim.",
  alt: "Bilgilerinizi ve gereğini rica ederim.",
  karma: "Bilgilerinizi ve gereğini arz ve rica ederim.",
  kisi: "Saygılarımla.",
};

/** m.11/1: E (elektronik), Z (zorunlu hâl), O (olağanüstü durum). */
export type Ortam = "E" | "Z" | "O";
export type Ivedilik = "" | "ACELE" | "GÜNLÜDÜR" | "KİŞİYE ÖZEL";

export interface ResmiAlanlar {
  baslikTuru: BaslikTuru;
  idare: string;
  araIdare: string;
  birim: string;
  logo: boolean;
  yaziTipi: YaziTipi;
  ortam: Ortam;
  detsis: string;
  sdp: string;
  kayitNo: string;
  tarih: string;
  ivedilik: Ivedilik;
  konu: string;
  muhatap: string;
  muhatapBirim: string;
  muhatapAdres: string;
  /** Satır satır; birden çoksa a), b)… ile sıralanır (m.15/5). */
  ilgiler: string;
  /** Satır satır paragraflar. */
  metin: string;
  yon: Yon;
  kapanis: string;
  imzaAd: string;
  imzaUnvan: string;
  /** Yetki devri ya da vekâlet satırı ("Vali a.", "Genel Müdür V."; m.17/9-10). */
  imzaYetki: string;
  /** Olur: ara makam görüşü (m.20/4) ve oluru veren makam. */
  araMakam: boolean;
  araMakamAd: string;
  araMakamUnvan: string;
  olurAd: string;
  olurUnvan: string;
  /** Satır satır ekler (m.18). */
  ekler: string;
  /** Satır satır dağıtım (m.19). */
  geregi: string;
  bilgi: string;
  /** İletişim bilgileri (m.24). */
  adres: string;
  telefon: string;
  faks: string;
  eposta: string;
  kep: string;
  internet: string;
  bilgiKisi: string;
  bilgiUnvan: string;
  bilgiTelefon: string;
  /** Dilekçe (kişi → kurum). */
  dilekceAd: string;
  tcKimlik: string;
  dilekceAdres: string;
  dilekceTelefon: string;
  dilekceEposta: string;
  /** Tutanak / rapor. */
  belgeTuru: "tutanak" | "rapor";
  belgeBasligi: string;
  /** Satır satır "Ad SOYAD, Unvan". */
  imzacilar: string;
  /** Zarf. */
  zarfBoyu: ZarfBoyu;
  gonderenAdres: string;
}

/** GG.AA.YYYY (m.12/1). */
export function tarihMetni(d = new Date()): string {
  const iki = (n: number) => String(n).padStart(2, "0");
  return `${iki(d.getDate())}.${iki(d.getMonth() + 1)}.${d.getFullYear()}`;
}

export function varsayilanAlanlar(): ResmiAlanlar {
  const b = BASLIK_TURLERI[0];
  return {
    baslikTuru: b.id,
    idare: b.idare,
    araIdare: "",
    birim: b.birim,
    logo: false,
    yaziTipi: "tnr",
    ortam: "E",
    detsis: "12345678",
    sdp: "903.07.02",
    kayitNo: "",
    tarih: tarihMetni(),
    ivedilik: "",
    konu: "Hizmet İçi Eğitim Programı",
    muhatap: "ÖRNEK VALİLİĞİNE",
    muhatapBirim: "İl Millî Eğitim Müdürlüğü",
    muhatapAdres: "",
    ilgiler: "Örnek Valiliğinin 01.09.2026 tarihli ve E-87654321-903.07.02-1234 sayılı yazısı.",
    metin: [
      "İlgi yazı ile birimimiz personeline yönelik hizmet içi eğitim programının hazırlanması istenmiştir.",
      "Program kapsamında 2027 yılında toplam 1.452 personele eğitim verilmesi planlanmış olup programın tahminî maliyeti 10.545,72 TL'dir. Program taslağı ekte sunulmuştur.",
    ].join("\n"),
    yon: "ust",
    kapanis: KAPANIS.ust,
    imzaAd: "Ad SOYAD",
    imzaUnvan: "Unvan",
    imzaYetki: "",
    araMakam: true,
    araMakamAd: "Ad SOYAD",
    araMakamUnvan: "Unvan",
    olurAd: "Ad SOYAD",
    olurUnvan: "Unvan",
    ekler: "Program Taslağı (3 sayfa)",
    geregi: "",
    bilgi: "",
    adres: "Örnek Mahallesi Örnek Caddesi No: 1 00000 Örnek",
    telefon: "0 (000) 000 00 00",
    faks: "0 (000) 000 00 01",
    eposta: "ornek@ornek.gov.tr",
    kep: "ornek@hs01.kep.tr",
    internet: "www.ornek.gov.tr",
    bilgiKisi: "Ad SOYAD",
    bilgiUnvan: "Unvan",
    bilgiTelefon: "0 (000) 000 00 02",
    dilekceAd: "Ad SOYAD",
    tcKimlik: "",
    dilekceAdres: "Örnek Mahallesi Örnek Sokak No: 1 Daire: 2 Örnek / ÖRNEK",
    dilekceTelefon: "0 (500) 000 00 00",
    dilekceEposta: "",
    belgeTuru: "tutanak",
    belgeBasligi: "TUTANAK",
    imzacilar: ["Ad SOYAD, Unvan", "Ad SOYAD, Unvan", "Ad SOYAD, Unvan"].join("\n"),
    zarfBoyu: "DL",
    gonderenAdres: "",
  };
}

/** Satır satır alanı listeye çevirir (boş satırlar atlanır). */
export const satirlar = (s: string) =>
  s
    .split(/\r?\n/)
    .map((x) => x.trim())
    .filter(Boolean);

/** Türkçe büyük harf (m.10/2 idare adı, m.14/2 muhatap). */
export const buyukHarf = (s: string) => s.toLocaleUpperCase("tr-TR");

/** m.14/2: muhatap idare adı yönelme hâl ekiyle biter (… MÜDÜRLÜĞÜNE, … BAKANLIĞINA). Gerçek kişide ek yok. */
export function yonelmeEkiVar(muhatap: string): boolean {
  const m = muhatap.trim();
  if (!m || /^Sayın\s/u.test(m)) return true;
  const son = buyukHarf(m.split(/\s+/).pop() ?? "").replace(/[^\p{L}]/gu, "");
  return /(A|E)$/u.test(son);
}

/* ------------------------------------------------------------------ */
/* Bekleyen form değerleri                                              */
/* ------------------------------------------------------------------ */

// Dosya › Yeni'deki form doldurulunca değerler burada bekler; şablonun build() çağrısı (useFileOps.newDocument)
// onları bir kez alır. Böylece yeni belge akışı (kaydedilmemiş değişiklik sorusu, ad, taslak) aynı kalır.
let bekleyen: { sablon: string; alanlar: Partial<ResmiAlanlar> } | null = null;

export function alanlariBeklet(sablon: string, alanlar: Partial<ResmiAlanlar>): void {
  bekleyen = { sablon, alanlar };
}

export function bekleyenAlanlar(sablon: string): Partial<ResmiAlanlar> {
  const b = bekleyen;
  bekleyen = null;
  return b && b.sablon === sablon ? b.alanlar : {};
}
