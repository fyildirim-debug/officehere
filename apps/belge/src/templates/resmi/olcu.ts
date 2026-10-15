// Resmî Yazışma şablon ölçüleri. Dayanak: "Resmî Yazışmalarda Uygulanacak Usul ve Esaslar Hakkında Yönetmelik"
// (CB Kararı 2646, Resmî Gazete 10.06.2020/31151; birincil metin RG taramasından okundu). "m." madde, "K"
// Cumhurbaşkanlığı Resmî Yazışma Kılavuzu (2025) demektir. Uzunluklar model birimindedir: sayfa twip, paragraf pt.
import { cmToTw } from "../../model/document";

/** cm → pt (paragraf girintisi ve sekme durakları pt tutulur). */
export const cmPt = (cm: number) => Math.round((cm / 2.54) * 72 * 100) / 100;

/** m.6/1: A4 (210 × 297 mm), twip. */
export const A4 = { w: 11906, h: 16838 } as const;

/** m.8/1: yazı alanı üst, sol ve sağ kenardan 1,5 cm; logo kullanılırsa üst boşluk 0,5 cm. */
export const KENAR_CM = 1.5;
export const LOGO_UST_CM = 0.5;
export const KENAR_TW = cmToTw(KENAR_CM);
export const LOGO_UST_TW = cmToTw(LOGO_UST_CM);
/**
 * Alt kenar Yönetmelik'te belirtilmez (m.8). Alt bilgi (m.23 doğrulama, m.24 iletişim, m.27 sayfa no) sayfanın
 * en altında ve kenardan 0,5 cm'de (K §5); gövde, alt bilginin tamamı sığacak kadar yukarıda biter.
 */
export const ALT_BILGI_CM = 0.5;
export const ALT_KENAR_CM = 3.2;

/** m.9; K §6: sol üst logo üst kenardan 0,5 cm, başlık üst kenardan 1,5 cm (aradaki 1 cm başlığın önünde). */
export const LOGO_BOY_CM = 2;
export const LOGO_BASLIK_ARASI_CM = KENAR_CM - LOGO_UST_CM;

/** m.7/1: Times New Roman 12 ya da Arial 11, normal stil; yazı alanında 9, iletişim bilgilerinde 8 puntoya dek. */
export const YAZI_TIPLERI = {
  tnr: { ad: "Times New Roman", punto: 12, etiket: "Times New Roman 12" },
  arial: { ad: "Arial", punto: 11, etiket: "Arial 11" },
} as const;
export type YaziTipi = keyof typeof YAZI_TIPLERI;
export const EN_KUCUK_PUNTO = 9;
export const ILETISIM_PUNTO = 8;

/** m.16/4: paragraf 1,25 cm içeriden başlar, iki yana yaslı; paragraflar arasında boşluk yok. */
export const GIRINTI_CM = 1.25;
export const GIRINTI_PT = cmPt(GIRINTI_CM);

/** A4'te yazı alanı genişliği (pt): 21 cm − 2 × 1,5 cm = 18 cm. */
export const YAZI_ALANI_PT = cmPt(21 - 2 * KENAR_CM);

/**
 * m.15/3: "Sayı", "Konu" ve "İlgi" yan başlıklarının ":" işareti aynı hizada; m.13/1 ve m.15/4: ikinci satır yan
 * başlığın (ve ilgi harfinin) altı boş bırakılarak başlar. Üç sol sekme durağıyla: ":" → metin → ilgi metni.
 */
export const IKI_NOKTA_PT = cmPt(1.25);
export const YAN_METIN_PT = cmPt(1.6);
export const ILGI_METIN_PT = cmPt(2.2);

/** m.17/1: imza bloğu yazı alanının en sağında ortalanır. Blok genişliği 7 cm; sol girinti = alan − blok. */
export const IMZA_BLOK_CM = 7;

/** Satır boşlukları (boş paragraf sayısı). */
export const BOSLUK = {
  /** m.11/2: "Sayı:" başlığın son satırından iki satır boşluk sonra. */
  baslikSayi: 2,
  /** m.14/1: muhatap konunun son satırından iki satır boşluk sonra. */
  konuMuhatap: 2,
  /** m.15/2: "İlgi:" muhatabın son satırından iki satır boşluk sonra. */
  muhatapIlgi: 2,
  /** m.16/2: ilgi ile metin arasında bir satır. */
  ilgiMetin: 1,
  /** m.16/2: ilgi yoksa muhatap ile metin arasında iki satır. */
  muhatapMetin: 2,
  /** m.17/1: imza metnin bitiminden iki ilâ dört satır boşluk sonra (el ile imza yeri için üç). */
  metinImza: 3,
  /** m.18/1, m.19/1: ek ve dağıtım "uygun satır boşluğu" sonra (bir satır). */
  imzaEk: 1,
  ekDagitim: 1,
} as const;

/** DL (110 × 220 mm) ve C4 (229 × 324 mm) zarf, yatay (twip). m.31/2 zarf düzeni. */
export const ZARF = {
  DL: { w: cmToTw(22), h: cmToTw(11), etiket: "DL (11 × 22 cm)" },
  C4: { w: cmToTw(32.4), h: cmToTw(22.9), etiket: "C4 (22,9 × 32,4 cm)" },
} as const;
export type ZarfBoyu = keyof typeof ZARF;

/** m.26/3 ve m.31/2: ACELE, GÜNLÜDÜR, KİŞİYE ÖZEL kırmızı. Düz renk (saydamlık yok: PDF/A-1b). */
export const KIRMIZI = "#C00000";
