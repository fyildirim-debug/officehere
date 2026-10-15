// Resmî Yazışma (2020) şablon kümesi: üst yazı, dağıtımlı yazı, olur, tekit, ek ve dağıtım listesi sayfaları,
// dilekçe, tutanak/rapor ve zarf. Her şablon form alanlarından belge üretir; ölçüler olcu.ts, bölümler parcalar.ts.
import type { JSONContent } from "@tiptap/core";
import type { BelgeDocument } from "../../model/document";
import { KAPANIS, satirlar, varsayilanAlanlar, type ResmiAlanlar } from "./alanlar";
import { BOSLUK, KENAR_TW, KIRMIZI, ZARF, cmPt } from "./olcu";
import { altBilgi, baslik, belge, bos, dagitimBolumu, ekBolumu, imza, metinVeKapanis, muhatap, olurBolumu, p, resmiAyarlar, sayiMetni, sonKisim, t, ustKisim } from "./parcalar";

/** Formda gösterilen alan kümeleri. */
export type AlanGrubu = "baslik" | "sayi" | "muhatap" | "ilgi" | "metin" | "imza" | "olur" | "ek" | "dagitim" | "iletisim" | "dilekce" | "tutanak" | "zarf";

export interface ResmiSablon {
  id: string;
  etiket: string;
  aciklama: string;
  gruplar: AlanGrubu[];
  /** Şablona özgü varsayılanlar (ortak varsayılanların üzerine). */
  varsayilan: Partial<ResmiAlanlar>;
  uret: (a: ResmiAlanlar) => BelgeDocument;
}

const ustYaziBelgesi = (a: ResmiAlanlar, ad: string, secenek: Parameters<typeof sonKisim>[1] = {}, ek: JSONContent[] = []) => {
  const s = resmiAyarlar(a, ad);
  s.footer = altBilgi(a);
  return belge(s, [...ustKisim(a), ...metinVeKapanis(a), ...sonKisim(a, secenek), ...ek]);
};

const sayfaSonu: JSONContent = { type: "pageBreak" };
const ortali = (s: string) => p(s, { textAlign: "center" });

export const RESMI_SABLONLAR: ResmiSablon[] = [
  {
    id: "resmi-ust-yazi",
    etiket: "Üst yazı",
    aciklama: "Tek muhataplı resmî yazı: başlık, sayı ve tarih, konu, muhatap, ilgi, metin, imza ve ek.",
    gruplar: ["baslik", "sayi", "muhatap", "ilgi", "metin", "imza", "ek", "iletisim"],
    varsayilan: {},
    uret: (a) => ustYaziBelgesi(a, "Üst yazı"),
  },
  {
    id: "resmi-dagitimli",
    etiket: "Dağıtımlı yazı",
    aciklama: "Birden çok muhataba giden yazı: muhatap “DAĞITIM YERLERİNE”, sonda Gereği ve Bilgi listeleri.",
    gruplar: ["baslik", "sayi", "muhatap", "ilgi", "metin", "imza", "ek", "dagitim", "iletisim"],
    varsayilan: {
      muhatap: "DAĞITIM YERLERİNE",
      muhatapBirim: "",
      ilgiler: "",
      yon: "karma",
      kapanis: KAPANIS.karma,
      metin:
        "Birimlerimizde yürütülecek hizmet içi eğitim programının uygulama takvimi ekte gönderilmiştir. Programa katılacak personelin listesinin 30.10.2026 tarihine kadar Müdürlüğümüze bildirilmesi gerekmektedir.",
      ekler: "Uygulama Takvimi (2 sayfa)",
      geregi: ["Örnek Kaymakamlığına", "Örnek Kaymakamlığına"].join("\n"),
      bilgi: "Örnek Bakanlığına",
    },
    uret: (a) => ustYaziBelgesi(a, "Dağıtımlı yazı"),
  },
  {
    id: "resmi-olur",
    etiket: "Olur yazısı",
    aciklama: "Makam oluruna sunulan yazı: teklif eden imza, isteğe bağlı ara makam görüşü ve ortada OLUR.",
    gruplar: ["baslik", "sayi", "muhatap", "ilgi", "metin", "imza", "olur", "ek", "iletisim"],
    varsayilan: {
      muhatap: "VALİLİK MAKAMINA",
      muhatapBirim: "",
      ilgiler: "",
      metin: "Birimimiz personeline yönelik hizmet içi eğitim programının ekteki takvime göre 02.11.2026 ile 06.11.2026 tarihleri arasında uygulanması planlanmaktadır.",
      kapanis: "Makamlarınızca da uygun görülmesi hâlinde olurlarınıza arz ederim.",
      ekler: "Eğitim Takvimi (1 sayfa)",
    },
    uret: (a) => {
      const s = resmiAyarlar(a, "Olur yazısı");
      s.footer = altBilgi(a);
      // m.18/1: olur belgelerinde "Ek:" oluru alınan makamın imza bölümünden sonra.
      const ek = ekBolumu(satirlar(a.ekler));
      return belge(s, [
        ...ustKisim(a),
        ...metinVeKapanis(a),
        ...bos(BOSLUK.metinImza),
        ...imza(a.imzaAd, a.imzaUnvan, a.imzaYetki),
        ...olurBolumu(a),
        ...(ek.length ? [...bos(BOSLUK.imzaEk), ...ek] : []),
      ]);
    },
  },
  {
    id: "resmi-tekit",
    etiket: "Tekit yazısı",
    aciklama: "Cevap verilmeyen yazının yinelenmesi: ilgide önceki yazı, sağ üstte kırmızı GÜNLÜDÜR.",
    gruplar: ["baslik", "sayi", "muhatap", "ilgi", "metin", "imza", "ek", "iletisim"],
    varsayilan: {
      ivedilik: "GÜNLÜDÜR",
      muhatap: "ÖRNEK KAYMAKAMLIĞINA",
      muhatapBirim: "İlçe Millî Eğitim Müdürlüğü",
      ilgiler: "01.10.2026 tarihli ve E-12345678-903.07.02-2345 sayılı yazımız.",
      metin:
        "İlgi yazımız ile hizmet içi eğitim programına katılacak personelin listesi istenmiş, ancak bugüne kadar cevap verilmemiştir. Söz konusu listenin en geç 25.10.2026 tarihine kadar Müdürlüğümüze gönderilmesi gerekmektedir.",
      yon: "alt",
      kapanis: "Bilgilerinizi ve gereğini önemle rica ederim.",
      ekler: "",
    },
    uret: (a) => ustYaziBelgesi(a, "Tekit yazısı"),
  },
  {
    id: "resmi-ek-dagitim",
    etiket: "Ek ve dağıtım listesi",
    aciklama: "Uzun ek ve dağıtım listeleri ayrı sayfalarda: üst yazıda “Ek: Ek Listesi” ve “DAĞITIM LİSTESİ”.",
    gruplar: ["baslik", "sayi", "muhatap", "ilgi", "metin", "imza", "ek", "dagitim", "iletisim"],
    varsayilan: {
      muhatap: "DAĞITIM YERLERİNE",
      muhatapBirim: "",
      ilgiler: "",
      yon: "karma",
      kapanis: KAPANIS.karma,
      metin: "Hizmet içi eğitim programının uygulanmasına ilişkin belgeler ekte gönderilmiştir.",
      ekler: ["Eğitim Programı (4 sayfa)", "Uygulama Takvimi (2 sayfa)", "Katılımcı Listesi (3 sayfa)", "Değerlendirme Formu (1 sayfa)"].join("\n"),
      geregi: ["Örnek Valiliğine", "Örnek Kaymakamlığına", "Örnek Kaymakamlığına"].join("\n"),
      bilgi: ["Örnek Bakanlığına", "Örnek Genel Müdürlüğüne"].join("\n"),
    },
    uret: (a) => {
      const ekler = satirlar(a.ekler);
      const sayfalar: JSONContent[] = [
        sayfaSonu,
        ortali("EK LİSTESİ"),
        ...bos(1),
        ...ekler.map((e, i) => p(`${i + 1}- ${e}`)),
        sayfaSonu,
        ortali("DAĞITIM LİSTESİ"),
        ...bos(1),
        ...dagitimBolumu(satirlar(a.geregi), satirlar(a.bilgi)).slice(1),
      ];
      return ustYaziBelgesi(a, "Ek ve dağıtım listesi", { ekListesi: true, dagitimListesi: true }, sayfalar);
    },
  },
  {
    id: "resmi-dilekce",
    etiket: "Dilekçe",
    aciklama: "Kişiden kuruma dilekçe: tarih, muhatap kurum, metin, ad soyad ve imza, adres ve iletişim (3071 sayılı Kanun).",
    gruplar: ["dilekce", "muhatap", "metin", "ek"],
    varsayilan: {
      muhatap: "ÖRNEK ORTAOKULU MÜDÜRLÜĞÜNE",
      muhatapBirim: "",
      metin: "Okulunuz 7/A sınıfı öğrencisi oğlum/kızım …………………'ın 20.10.2026 ile 24.10.2026 tarihleri arasında sağlık nedeniyle okula devam edemeyeceğini bildiriyorum.",
      kapanis: "Gereğini bilgilerinize arz ederim.",
      ekler: "Sağlık Raporu (1 sayfa)",
    },
    uret: (a) => {
      const s = resmiAyarlar({ ...a, logo: false }, "Dilekçe");
      const ek = ekBolumu(satirlar(a.ekler));
      const iletisim = [
        `Adres: ${a.dilekceAdres.trim() || "…"}`,
        ...(a.dilekceTelefon.trim() ? [`Telefon: ${a.dilekceTelefon.trim()}`] : []),
        ...(a.dilekceEposta.trim() ? [`e-Posta: ${a.dilekceEposta.trim()}`] : []),
        ...(a.tcKimlik.trim() ? [`T.C. Kimlik No: ${a.tcKimlik.trim()}`] : []),
      ];
      return belge(s, [
        p(a.tarih.trim(), { textAlign: "right" }),
        ...bos(2),
        ...muhatap(a),
        ...bos(2),
        ...metinVeKapanis(a),
        ...bos(BOSLUK.metinImza),
        ...imza(a.dilekceAd, "İmza"),
        ...bos(2),
        ...iletisim.map((x) => p(x)),
        ...(ek.length ? [...bos(1), ...ek] : []),
      ]);
    },
  },
  {
    id: "resmi-tutanak",
    etiket: "Tutanak ve rapor",
    aciklama: "Tutanak ya da rapor: başlık, metin, metnin bitiminde tarih (m.12/4) ve Ad SOYAD / Unvan / İmza tablosu.",
    gruplar: ["baslik", "tutanak", "metin"],
    varsayilan: {
      baslikTuru: "okul",
      idare: "ÖRNEK KAYMAKAMLIĞI",
      birim: "Örnek Ortaokulu Müdürlüğü",
      metin: [
        "Okulumuz zümre öğretmenler kurulu 15.10.2026 tarihinde saat 14.00'te okul müdürü başkanlığında toplanmıştır.",
        "Toplantıda öğretim programlarının uygulanması, ortak sınav takvimi ve ölçme değerlendirme çalışmaları görüşülmüş; ortak sınavların 2. ve 3. haftalarda yapılmasına oy birliğiyle karar verilmiştir.",
        "İşbu tutanak üç nüsha olarak düzenlenmiş ve imza altına alınmıştır.",
      ].join("\n"),
      kapanis: "",
    },
    uret: (a) => {
      const s = resmiAyarlar({ ...a, logo: false }, a.belgeTuru === "rapor" ? "Rapor" : "Tutanak");
      const kisiler = satirlar(a.imzacilar).map((x) => {
        const [ad, ...unvan] = x.split(",");
        return [ad.trim(), unvan.join(",").trim()];
      });
      const hucre = (s2: string, attrs: Record<string, unknown> = {}): JSONContent => ({ type: "tableCell", attrs: {}, content: [p(s2, { spaceBefore: 4, spaceAfter: 4, ...attrs })] });
      const tablo: JSONContent = {
        type: "table",
        attrs: { tableStyle: "grid", headerRow: true, bandedRows: false },
        content: [
          { type: "tableRow", content: [hucre("Ad SOYAD", { textAlign: "center" }), hucre("Unvan", { textAlign: "center" }), hucre("İmza", { textAlign: "center" })] },
          ...kisiler.map(([ad, unvan]) => ({ type: "tableRow", content: [hucre(ad), hucre(unvan), hucre("")] })),
        ],
      };
      const baslikSatirlari = baslik({ ...a, logo: false, ivedilik: "" });
      return belge(s, [
        ...baslikSatirlari,
        ...bos(2),
        p((a.belgeBasligi.trim() || (a.belgeTuru === "rapor" ? "RAPOR" : "TUTANAK")).toLocaleUpperCase("tr-TR"), { textAlign: "center" }),
        ...bos(1),
        ...metinVeKapanis(a),
        ...bos(1),
        // m.12/4: tutanak ve raporda tarih metnin bitiminde yer alabilir.
        p(a.tarih.trim(), { textAlign: "right" }),
        ...bos(1),
        tablo,
        p(),
      ]);
    },
  },
  {
    id: "resmi-zarf",
    etiket: "Zarf",
    aciklama: "DL ya da C4 zarf: sol üstte başlık, sayı ve tarih; ortada muhatap ve adres; sağ üstte kırmızı ivedilik (m.31/2).",
    gruplar: ["baslik", "sayi", "muhatap", "zarf"],
    varsayilan: { muhatap: "ÖRNEK BAKANLIĞINA", muhatapBirim: "Personel Genel Müdürlüğü", muhatapAdres: "Örnek Mahallesi Örnek Bulvarı No: 10\n00000 Örnek / ÖRNEK", ivedilik: "ACELE" },
    uret: (a) => {
      const z = ZARF[a.zarfBoyu];
      const s = resmiAyarlar({ ...a, logo: false }, "Zarf");
      s.page = { ...s.page, width: z.w, height: z.h, orientation: "landscape", margin: { ...s.page.margin, top: KENAR_TW, bottom: KENAR_TW, header: KENAR_TW / 2, footer: KENAR_TW / 2 } };
      const alanPt = Math.round(((z.w - 2 * KENAR_TW) / 20) * 100) / 100;
      const sag = [{ pos: alanPt, align: "right" as const, leader: "none" as const }];
      // m.31/2: ivedilik/kişiye özel sağ üst köşede, en üstte, kırmızı büyük harflerle.
      const bas = baslik({ ...a, logo: false, ivedilik: "" });
      const ilk = a.ivedilik ? p([t("T.C.\t"), t(a.ivedilik, [{ type: "textStyle", attrs: { color: KIRMIZI } }])], { tabs: sag }) : bas[0];
      const sol = (x: JSONContent) => ({ ...x, attrs: { ...x.attrs, textAlign: null } });
      const gonderen = satirlar(a.gonderenAdres).map((x) => p(x));
      const muhatapBlok = muhatap(a);
      const ara = a.zarfBoyu === "C4" ? 10 : 2;
      return belge(s, [
        sol(ilk),
        ...bas.slice(1).map(sol),
        ...gonderen,
        p(`Sayı: ${sayiMetni(a)}`),
        p(`Tarih: ${a.tarih.trim()}`),
        ...bos(ara),
        ...muhatapBlok.map((x) => ({ ...x, attrs: { ...x.attrs, indentLeft: cmPt(a.zarfBoyu === "C4" ? 6 : 3), indentRight: cmPt(a.zarfBoyu === "C4" ? 6 : 3) } })),
      ]);
    },
  },
];

export function resmiSablon(id: string): ResmiSablon | undefined {
  return RESMI_SABLONLAR.find((s) => s.id === id);
}

/** Şablonun tüm alanları: ortak varsayılanlar ← şablon varsayılanları ← form değerleri. */
export function resmiAlanlar(id: string, form: Partial<ResmiAlanlar> = {}): ResmiAlanlar {
  return { ...varsayilanAlanlar(), ...(resmiSablon(id)?.varsayilan ?? {}), ...form };
}

export function resmiBelge(id: string, form: Partial<ResmiAlanlar> = {}): BelgeDocument {
  const s = resmiSablon(id) ?? RESMI_SABLONLAR[0];
  return s.uret(resmiAlanlar(s.id, form));
}
