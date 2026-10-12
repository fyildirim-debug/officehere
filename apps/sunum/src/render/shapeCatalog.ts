/** Ekle › Şekiller galerisinin kataloğu (ECMA-376 prstGeom adları ve Türkçe etiketler). */
export interface ShapeCatalogItem {
  geom: string;
  label: string;
  /** Çizgi okları (Ok Çizgisi). */
  tail?: "triangle";
  head?: "triangle";
}

export const SHAPE_CATALOG: { group: string; items: ShapeCatalogItem[] }[] = [
  {
    group: "Çizgiler",
    items: [
      { geom: "line", label: "Çizgi" },
      { geom: "straightConnector1", label: "Ok Çizgisi", tail: "triangle" },
      { geom: "straightConnector1", label: "Çift Ok Çizgisi", tail: "triangle", head: "triangle" },
      { geom: "bentConnector3", label: "Bağlayıcı: Dirsek" },
      { geom: "curvedConnector3", label: "Bağlayıcı: Eğri" },
      { geom: "arc", label: "Yay" },
    ],
  },
  {
    group: "Dikdörtgenler",
    items: [
      { geom: "rect", label: "Dikdörtgen" },
      { geom: "roundRect", label: "Dikdörtgen: Köşeleri Yuvarlatılmış" },
      { geom: "snip1Rect", label: "Dikdörtgen: Tek Köşesi Kesik" },
      { geom: "round2SameRect", label: "Dikdörtgen: Aynı Kenar Köşeleri Yuvarlatılmış" },
    ],
  },
  {
    group: "Temel Şekiller",
    items: [
      { geom: "ellipse", label: "Oval" },
      { geom: "triangle", label: "İkizkenar Üçgen" },
      { geom: "rtTriangle", label: "Dik Üçgen" },
      { geom: "parallelogram", label: "Paralelkenar" },
      { geom: "trapezoid", label: "Yamuk" },
      { geom: "diamond", label: "Elmas" },
      { geom: "pentagon", label: "Normal Beşgen" },
      { geom: "hexagon", label: "Altıgen" },
      { geom: "heptagon", label: "Yedigen" },
      { geom: "octagon", label: "Sekizgen" },
      { geom: "decagon", label: "Ongen" },
      { geom: "dodecagon", label: "Onikigen" },
      { geom: "pie", label: "Pasta" },
      { geom: "teardrop", label: "Gözyaşı" },
      { geom: "frame", label: "Çerçeve" },
      { geom: "corner", label: "L Şekli" },
      { geom: "plus", label: "Artı" },
      { geom: "can", label: "Silindir" },
      { geom: "cube", label: "Küp" },
      { geom: "bevel", label: "Eğim" },
      { geom: "donut", label: "Halka" },
      { geom: "noSmoking", label: "Yasak İşareti" },
      { geom: "foldedCorner", label: "Katlanmış Köşe" },
      { geom: "smileyFace", label: "Gülen Yüz" },
      { geom: "heart", label: "Kalp" },
      { geom: "lightningBolt", label: "Şimşek" },
      { geom: "sun", label: "Güneş" },
      { geom: "moon", label: "Ay" },
      { geom: "cloud", label: "Bulut" },
      { geom: "wave", label: "Dalga" },
    ],
  },
  {
    group: "Blok Oklar",
    items: [
      { geom: "rightArrow", label: "Ok: Sağ" },
      { geom: "leftArrow", label: "Ok: Sol" },
      { geom: "upArrow", label: "Ok: Yukarı" },
      { geom: "downArrow", label: "Ok: Aşağı" },
      { geom: "leftRightArrow", label: "Ok: Sol-Sağ" },
      { geom: "upDownArrow", label: "Ok: Yukarı-Aşağı" },
      { geom: "quadArrow", label: "Ok: Dört Yönlü" },
      { geom: "homePlate", label: "Ok: Beşgen" },
      { geom: "chevron", label: "Ok: Köşeli Çift Ayraç" },
    ],
  },
  {
    group: "Denklem Şekilleri",
    items: [
      { geom: "mathPlus", label: "Artı İşareti" },
      { geom: "mathMinus", label: "Eksi İşareti" },
      { geom: "mathMultiply", label: "Çarpma İşareti" },
      { geom: "mathEqual", label: "Eşittir" },
    ],
  },
  {
    group: "Akış Çizelgesi",
    items: [
      { geom: "flowChartProcess", label: "Akış Çizelgesi: İşlem" },
      { geom: "flowChartDecision", label: "Akış Çizelgesi: Karar" },
      { geom: "flowChartInputOutput", label: "Akış Çizelgesi: Veri" },
      { geom: "flowChartPredefinedProcess", label: "Akış Çizelgesi: Önceden Tanımlı İşlem" },
      { geom: "flowChartDocument", label: "Akış Çizelgesi: Belge" },
      { geom: "flowChartTerminator", label: "Akış Çizelgesi: Sonlandırıcı" },
      { geom: "flowChartConnector", label: "Akış Çizelgesi: Bağlayıcı" },
    ],
  },
  {
    group: "Yıldızlar ve Başlık Sancakları",
    items: [
      { geom: "irregularSeal1", label: "Patlama: 8 Köşeli" },
      { geom: "star4", label: "Yıldız: 4 Köşeli" },
      { geom: "star5", label: "Yıldız: 5 Köşeli" },
      { geom: "star6", label: "Yıldız: 6 Köşeli" },
      { geom: "star8", label: "Yıldız: 8 Köşeli" },
      { geom: "star10", label: "Yıldız: 10 Köşeli" },
      { geom: "star12", label: "Yıldız: 12 Köşeli" },
    ],
  },
  {
    group: "Belirtme Çizgileri",
    items: [
      { geom: "wedgeRectCallout", label: "Konuşma Balonu: Dikdörtgen" },
      { geom: "wedgeRoundRectCallout", label: "Konuşma Balonu: Köşeleri Yuvarlatılmış Dikdörtgen" },
      { geom: "wedgeEllipseCallout", label: "Konuşma Balonu: Oval" },
      { geom: "cloudCallout", label: "Düşünce Balonu: Bulut" },
    ],
  },
];

// Eylem düğmeleri: gösteride tıklanınca eylem (köprü) çalışır; çizilince Eylem Ayarları açılır.
SHAPE_CATALOG.push({
  group: "Eylem Düğmeleri",
  items: [
    { geom: "actionButtonBackPrevious", label: "Eylem Düğmesi: Geri veya Önceki" },
    { geom: "actionButtonForwardNext", label: "Eylem Düğmesi: İleri veya Sonraki" },
    { geom: "actionButtonBeginning", label: "Eylem Düğmesi: Başlangıç" },
    { geom: "actionButtonEnd", label: "Eylem Düğmesi: Son" },
    { geom: "actionButtonHome", label: "Eylem Düğmesi: Giriş" },
    { geom: "actionButtonInformation", label: "Eylem Düğmesi: Bilgi" },
    { geom: "actionButtonReturn", label: "Eylem Düğmesi: Geri Dön" },
    { geom: "actionButtonDocument", label: "Eylem Düğmesi: Belge" },
    { geom: "actionButtonHelp", label: "Eylem Düğmesi: Yardım" },
    { geom: "actionButtonBlank", label: "Eylem Düğmesi: Boş" },
  ],
});

export const SHAPE_LABELS: Record<string, string> = Object.fromEntries(SHAPE_CATALOG.flatMap((g) => g.items.map((i) => [i.geom, i.label])));
