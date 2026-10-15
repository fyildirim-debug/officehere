// Resmî yazının bölümleri (başlık, sayı/tarih, konu, muhatap, ilgi, metin, imza, ek, dağıtım, olur, alt bilgi) için
// belge içeriği üreticileri. Ölçüler ve madde dayanakları olcu.ts'de.
import type { JSONContent } from "@tiptap/core";
import { defaultSettings, type BelgeDocument, type DocSettings } from "../../model/document";
import type { StyleDef } from "../../model/styles";
import type { TabStop } from "../../editor/tabStops";
import { buyukHarf, satirlar, type ResmiAlanlar } from "./alanlar";
import {
  A4,
  ALT_BILGI_CM,
  ALT_KENAR_CM,
  BOSLUK,
  GIRINTI_PT,
  ILETISIM_PUNTO,
  IKI_NOKTA_PT,
  ILGI_METIN_PT,
  IMZA_BLOK_CM,
  KENAR_TW,
  KIRMIZI,
  LOGO_BASLIK_ARASI_CM,
  LOGO_BOY_CM,
  LOGO_UST_TW,
  YAN_METIN_PT,
  YAZI_ALANI_PT,
  YAZI_TIPLERI,
  cmPt,
  type YaziTipi,
} from "./olcu";
import { cmToTw } from "../../model/document";

type Attrs = Record<string, unknown>;

export const t = (text: string, marks?: JSONContent["marks"]): JSONContent => ({ type: "text", text, ...(marks ? { marks } : {}) });
const ts = (attrs: Attrs) => ({ type: "textStyle", attrs });
export const p = (content?: JSONContent[] | string, attrs: Attrs = {}): JSONContent => ({
  type: "paragraph",
  attrs,
  ...(content && content.length ? { content: typeof content === "string" ? [t(content)] : content } : {}),
});
/** n satır boşluk (boş paragraf). */
export const bos = (n: number): JSONContent[] => Array.from({ length: n }, () => p());
const kirmizi = (s: string) => t(s, [ts({ color: KIRMIZI })]);
const sekme = (pos: number, align: TabStop["align"] = "left"): TabStop => ({ pos, align, leader: "none" });

/** m.15/3: Sayı, Konu ve İlgi'nin ":" hizası; tarih için yazı alanının en sağında sağ durak (m.12/1). */
export const YAN_SEKMELER: TabStop[] = [sekme(IKI_NOKTA_PT), sekme(YAN_METIN_PT)];

/** Metin paragrafı: 1,25 cm girinti, iki yana yaslı (m.16/4). */
export const metinP = (s: string): JSONContent => p(s, { indentFirst: GIRINTI_PT, textAlign: "justify" });

/** m.17/1: imza bloğu yazı alanının en sağında ortalanır (sol girinti = alan − blok genişliği). */
export const imzaBlokGirintisi = (alanPt = YAZI_ALANI_PT) => Math.round((alanPt - cmPt(IMZA_BLOK_CM)) * 100) / 100;
const sagBlok = (s: string, alanPt?: number) => p(s, { textAlign: "center", indentLeft: imzaBlokGirintisi(alanPt) });
/** m.20/4: "Uygun görüşle arz ederim." yazı alanının solunda; altındaki ad ve unvan o blokta ortalanır. */
const solBlok = (s: string, alanPt?: number) => p(s, { textAlign: "center", indentRight: imzaBlokGirintisi(alanPt) });

/* ------------------------------------------------------------------ */
/* Belge ayarları                                                       */
/* ------------------------------------------------------------------ */

/** Normal stil: m.7 yazı tipi/punto, satır aralığı tek, paragraflar arası boşluk yok (m.16/4). */
export function normalStil(yazi: YaziTipi): StyleDef {
  const y = YAZI_TIPLERI[yazi];
  return { id: "Normal", name: "Normal", type: "paragraph", props: { font: y.ad, size: y.punto, spaceBefore: 0, spaceAfter: 0, line: 1 } };
}

/** İletişim bilgileri ve sayfa numarası paragraf stili: 8 punto (m.7/1, m.24), aralıksız, tek satır. */
export const ALT_BILGI_STILI = "ResmiAltBilgi";
export function altBilgiStili(): StyleDef {
  return { id: ALT_BILGI_STILI, name: "Resmî Yazı Alt Bilgisi", type: "paragraph", basedOn: "Normal", props: { size: ILETISIM_PUNTO, spaceBefore: 0, spaceAfter: 0, line: 1 }, gallery: false };
}

export function resmiAyarlar(a: Pick<ResmiAlanlar, "yaziTipi" | "logo">, baslik: string): DocSettings {
  const s = defaultSettings();
  s.properties.title = baslik;
  s.page = {
    ...s.page,
    width: A4.w,
    height: A4.h,
    orientation: "portrait",
    // m.8/1: üst, sol, sağ 1,5 cm; logo varsa üst 0,5 cm.
    margin: { top: a.logo ? LOGO_UST_TW : KENAR_TW, left: KENAR_TW, right: KENAR_TW, bottom: cmToTw(ALT_KENAR_CM), header: cmToTw(0.5), footer: cmToTw(ALT_BILGI_CM), gutter: 0 },
  };
  s.customStyles = { Normal: normalStil(a.yaziTipi), [ALT_BILGI_STILI]: altBilgiStili() };
  return s;
}

export const belge = (settings: DocSettings, content: JSONContent[]): BelgeDocument => ({ settings, content: { type: "doc", content } });

/* ------------------------------------------------------------------ */
/* Başlık, sayı, konu, muhatap, ilgi                                    */
/* ------------------------------------------------------------------ */

/** Sol üst logo yer tutucusu (m.9; gerçek logo yok). Dolgusuz, düz kenarlı metin kutusu: saydamlık üretmez. */
export function logoYeri(): JSONContent {
  const px = Math.round((LOGO_BOY_CM / 2.54) * 96);
  return {
    type: "textBox",
    attrs: { x: 0, y: 0, w: px, h: px, border: "#7F7F7F", borderWidth: 0.75, fill: null },
    content: [p([t("LOGO", [ts({ fontSize: "8pt", color: "#7F7F7F" })])], { textAlign: "center", spaceBefore: 26 })],
  };
}

/** m.10: "T.C.", idare adı BÜYÜK HARFLE, (bağlı idarede ara satır), birim adı; ortalı. */
export function baslik(a: ResmiAlanlar): JSONContent[] {
  const satir = ["T.C.", buyukHarf(a.idare.trim() || "…"), ...(a.baslikTuru === "bagli" && a.araIdare.trim() ? [a.araIdare.trim()] : []), a.birim.trim() || "…"];
  const out: JSONContent[] = [];
  if (a.logo) out.push(logoYeri());
  if (a.ivedilik && a.ivedilik !== "KİŞİYE ÖZEL") out.push(p([kirmizi(a.ivedilik)], { textAlign: "right" }));
  satir.forEach((s, i) => out.push(p(s, { textAlign: "center", ...(i === 0 && a.logo && !a.ivedilik ? { spaceBefore: cmPt(LOGO_BASLIK_ARASI_CM) } : {}) })));
  return out;
}

/** m.11/1: E-<DETSİS>-<SDP>-<kayıt no>. */
export const sayiMetni = (a: Pick<ResmiAlanlar, "ortam" | "detsis" | "sdp" | "kayitNo">) => `${a.ortam}-${a.detsis.trim() || "…"}-${a.sdp.trim() || "…"}-${a.kayitNo.trim() || "……"}`;

/** m.11/2 ve m.12/1: "Sayı :" solda; tarih aynı satırda, yazı alanının en sağında. */
export function sayiSatiri(a: ResmiAlanlar, alanPt = YAZI_ALANI_PT): JSONContent {
  return p(`Sayı\t:\t${sayiMetni(a)}\t${a.tarih.trim()}`, { tabs: [...YAN_SEKMELER, sekme(alanPt, "right")] });
}

/** m.13/1: Sayı'nın bir alt satırı; ikinci satır "Konu:" altı boş bırakılarak (asılı girinti). */
export function konuSatiri(konu: string): JSONContent {
  return p(`Konu\t:\t${konu.trim()}`, { tabs: YAN_SEKMELER, indentLeft: YAN_METIN_PT, indentFirst: -YAN_METIN_PT });
}

/** m.14: ortalı, idare BÜYÜK HARFLE + yönelme eki; birim parantez içinde alt satırda; adres altında ortalı. */
export function muhatap(a: Pick<ResmiAlanlar, "muhatap" | "muhatapBirim" | "muhatapAdres">): JSONContent[] {
  const m = a.muhatap.trim();
  const ad = /^Sayın\s/u.test(m) ? m : buyukHarf(m);
  const out = [p(ad || "…", { textAlign: "center" })];
  if (a.muhatapBirim.trim()) out.push(p(`(${a.muhatapBirim.trim().replace(/^\(|\)$/g, "")})`, { textAlign: "center" }));
  for (const s of satirlar(a.muhatapAdres)) out.push(p(s, { textAlign: "center" }));
  return out;
}

/** m.15: "İlgi :" solda; birden çoksa a), b)… (m.15/5); devam satırı harfin altı boş bırakılarak (m.15/4). */
export function ilgiSatirlari(ilgiler: string[]): JSONContent[] {
  if (ilgiler.length === 1) return [p(`İlgi\t:\t${ilgiler[0]}`, { tabs: YAN_SEKMELER, indentLeft: YAN_METIN_PT, indentFirst: -YAN_METIN_PT })];
  const tabs = [...YAN_SEKMELER, sekme(ILGI_METIN_PT)];
  return ilgiler.map((s, i) => p(`${i === 0 ? "İlgi\t:" : "\t"}\t${HARFLER[i] ?? "?"})\t${s}`, { tabs, indentLeft: ILGI_METIN_PT, indentFirst: -ILGI_METIN_PT }));
}

/** m.15/5, m.16/10: Türk alfabesindeki küçük harfler. */
export const HARFLER = "abcçdefgğhıijklmnoöprsştuüvyz".split("");

/** Üst yazının ilk bölümü: başlık → (2) → sayı/tarih → konu → (2) → muhatap → (2) → ilgi → (1) | (2). */
export function ustKisim(a: ResmiAlanlar): JSONContent[] {
  const ilgiler = satirlar(a.ilgiler);
  return [
    ...baslik(a),
    ...bos(BOSLUK.baslikSayi),
    sayiSatiri(a),
    konuSatiri(a.konu),
    ...bos(BOSLUK.konuMuhatap),
    ...muhatap(a),
    ...(ilgiler.length ? [...bos(BOSLUK.muhatapIlgi), ...ilgiSatirlari(ilgiler), ...bos(BOSLUK.ilgiMetin)] : bos(BOSLUK.muhatapMetin)),
  ];
}

/** m.16: metin paragrafları ve kapanış (m.16/12). */
export function metinVeKapanis(a: Pick<ResmiAlanlar, "metin" | "kapanis">): JSONContent[] {
  return [...satirlar(a.metin).map(metinP), ...(a.kapanis.trim() ? [metinP(a.kapanis.trim())] : [])];
}

/* ------------------------------------------------------------------ */
/* İmza, ek, dağıtım, olur                                              */
/* ------------------------------------------------------------------ */

/** m.17: Ad SOYAD, (yetki devri / vekâlet satırı), unvan; yazı alanının en sağında ortalı. */
export function imza(ad: string, unvan: string, yetki = "", alanPt?: number): JSONContent[] {
  return [sagBlok(ad.trim() || "Ad SOYAD", alanPt), ...(yetki.trim() ? [sagBlok(yetki.trim(), alanPt)] : []), sagBlok(unvan.trim() || "Unvan", alanPt)];
}

/** m.18: tek ekte "Ek:" aynı satırda; çok ekte numaralı liste. */
export function ekBolumu(ekler: string[]): JSONContent[] {
  if (!ekler.length) return [];
  if (ekler.length === 1) return [p(`Ek: ${ekler[0]}`)];
  return [p("Ek:"), ...ekler.map((e, i) => p(`${i + 1}- ${e}`))];
}

/** m.19/2: "Gereği:" solda, "Bilgi:" aynı satırda yazı alanının ortasına doğru. */
export function dagitimBolumu(geregi: string[], bilgi: string[], liste = false): JSONContent[] {
  if (liste) return [p("Dağıtım:"), p("DAĞITIM LİSTESİ")];
  if (!geregi.length && !bilgi.length) return [];
  const orta = [sekme(Math.round((YAZI_ALANI_PT / 2) * 100) / 100)];
  if (!bilgi.length) return [p("Dağıtım:"), ...geregi.map((g) => p(g))];
  const out = [p("Dağıtım:"), p("Gereği:\tBilgi:", { tabs: orta })];
  for (let i = 0; i < Math.max(geregi.length, bilgi.length); i++) out.push(p(`${geregi[i] ?? ""}\t${bilgi[i] ?? ""}`, { tabs: orta }));
  return out;
}

/** İmza → (1) → Ek → (1) → Dağıtım. */
export function sonKisim(a: ResmiAlanlar, secenek: { ekListesi?: boolean; dagitimListesi?: boolean } = {}): JSONContent[] {
  const ekler = secenek.ekListesi ? ["Ek Listesi"] : satirlar(a.ekler);
  const ek = ekBolumu(ekler);
  const dag = dagitimBolumu(satirlar(a.geregi), satirlar(a.bilgi), secenek.dagitimListesi);
  return [...bos(BOSLUK.metinImza), ...imza(a.imzaAd, a.imzaUnvan, a.imzaYetki), ...(ek.length ? [...bos(BOSLUK.imzaEk), ...ek] : []), ...(dag.length ? [...bos(BOSLUK.ekDagitim), ...dag] : [])];
}

/** m.20: ara makam (solda "Uygun görüşle arz ederim."), ortada "OLUR", altında ad SOYAD ve unvan. */
export function olurBolumu(a: ResmiAlanlar): JSONContent[] {
  const out: JSONContent[] = [];
  if (a.araMakam)
    out.push(
      ...bos(2),
      solBlok("Uygun görüşle arz ederim."),
      ...(a.ortam === "E" ? [] : [solBlok("…/…/20…")]),
      ...bos(2),
      solBlok(a.araMakamAd.trim() || "Ad SOYAD"),
      solBlok(a.araMakamUnvan.trim() || "Unvan"),
    );
  out.push(...bos(2), p("OLUR", { textAlign: "center" }));
  // m.20/3: el ile imzada OLUR'dan sonra tarih ve imza için boş satır.
  if (a.ortam !== "E") out.push(p("…/…/20…", { textAlign: "center" }), ...bos(2));
  out.push(p(a.olurAd.trim() || "Ad SOYAD", { textAlign: "center" }), p(a.olurUnvan.trim() || "Unvan", { textAlign: "center" }));
  return out;
}

/* ------------------------------------------------------------------ */
/* Alt bilgi: doğrulama yeri, iletişim, sayfa numarası                  */
/* ------------------------------------------------------------------ */

const ALT_SATIR = { styleId: ALT_BILGI_STILI };
const hucre = (paras: JSONContent[], w: number): JSONContent => ({ type: "tableCell", attrs: { colwidth: [w] }, content: paras });

/**
 * Her sayfanın altı (m.16/5): üstte iki satırlık belge doğrulama yeri (m.23; e-imzada EBYS doldurur, şablon boş
 * bırakır), çizgi, solda idarenin iletişim bilgileri, sağda bilgi alınacak kişi (m.24, 8 punto: ALT_BILGI_STILI), altında ortada
 * sayfa numarası "1/3" (m.27).
 */
export function altBilgi(a: ResmiAlanlar): JSONContent {
  const sol = [`Adres: ${a.adres}`, `Telefon No: ${a.telefon}  Faks No: ${a.faks}`, `e-Posta: ${a.eposta}  İnternet Adresi: ${a.internet}`, `Kep Adresi: ${a.kep}`];
  const sag = [`Bilgi için: ${a.bilgiKisi}`, a.bilgiUnvan, `Telefon No: ${a.bilgiTelefon}`];
  const alanPx = Math.round((YAZI_ALANI_PT * 4) / 3);
  const solPx = Math.round(alanPx * 0.64);
  const satir = (s: string) => p(s, ALT_SATIR);
  const alan = (f: "PAGE" | "NUMPAGES"): JSONContent => ({ type: "pageField", attrs: { field: f } });
  return {
    type: "doc",
    content: [
      p(undefined, ALT_SATIR),
      p(undefined, { ...ALT_SATIR, borders: { bottom: { style: "single", size: 0.5, color: "#000000", space: 1 } } }),
      {
        type: "table",
        attrs: { tableStyle: "none", headerRow: false, bandedRows: false },
        content: [{ type: "tableRow", content: [hucre(sol.map(satir), solPx), hucre(sag.map(satir), alanPx - solPx)] }],
      },
      p([alan("PAGE"), t("/"), alan("NUMPAGES")], { textAlign: "center", ...ALT_SATIR }),
    ],
  };
}
