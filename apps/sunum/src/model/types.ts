/**
 * Sunum veri modeli. Geometri 96 dpi piksel (px) cinsindendir: 16:9 slayt
 * 1280×720 px = 33,867×19,05 cm. Yazı tipi boyutları ve çizgi kalınlıkları pt.
 * Yapı ECMA-376 PresentationML/DrawingML kavramlarını izler (yer tutucu,
 * düzen, asıl slayt, tema renk yuvaları, metin düzeyleri).
 */

export type SchemeColorName =
  | "dk1"
  | "lt1"
  | "dk2"
  | "lt2"
  | "accent1"
  | "accent2"
  | "accent3"
  | "accent4"
  | "accent5"
  | "accent6"
  | "hlink"
  | "folHlink";

export const SCHEME_ORDER: SchemeColorName[] = ["lt1", "dk1", "lt2", "dk2", "accent1", "accent2", "accent3", "accent4", "accent5", "accent6"];

export interface Color {
  /** "#RRGGBB" */
  rgb?: string;
  scheme?: SchemeColorName;
  /** 0..1 (DrawingML lumMod/100000) */
  lumMod?: number;
  lumOff?: number;
  shade?: number;
  tint?: number;
  /** 0..1, 1 = opak */
  alpha?: number;
}

export interface GradientStop {
  pos: number; // 0..1
  color: Color;
}

export type Fill =
  | { type: "none" }
  | { type: "solid"; color: Color }
  | { type: "gradient"; stops: GradientStop[]; angle: number; kind: "linear" | "radial" }
  | { type: "image"; src: string; mode: "stretch" | "tile"; alpha?: number };

export type DashType = "solid" | "sysDot" | "sysDash" | "dash" | "dashDot" | "lgDash" | "lgDashDot" | "lgDashDotDot";
export type ArrowType = "none" | "triangle" | "stealth" | "diamond" | "oval" | "arrow";

export interface Line {
  fill: { type: "none" } | { type: "solid"; color: Color };
  /** pt */
  width: number;
  dash: DashType;
  head?: ArrowType;
  tail?: ArrowType;
  /** Çizgi ucu (ECMA-376 ST_LineCap): yuvarlak, kare, düz. */
  cap?: "rnd" | "sq" | "flat";
}

export interface Shadow {
  color: Color;
  /** px */
  blur: number;
  dist: number;
  /** derece */
  dir: number;
}

export interface Effects {
  shadow?: Shadow;
  glow?: { color: Color; radius: number };
  reflection?: boolean;
  softEdge?: number;
}

export interface RunProps {
  /** Yazı tipi adı ya da "+mj" (başlık) / "+mn" (gövde) tema yazı tipi. */
  font?: string;
  /** pt */
  size?: number;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  strike?: boolean;
  color?: Color;
  highlight?: Color;
  shadow?: boolean;
  caps?: "none" | "all" | "small";
  /** Karakter aralığı (pt). */
  spacing?: number;
  /** Üst simge (+) / alt simge (-) yüzdesi. */
  baseline?: number;
  /** Metin anahattı (WordArt). */
  outline?: Color;
  /** Metin parlaması (WordArt). */
  glow?: Color;
  /** Alan (slayt numarası, tarih). */
  field?: "slidenum" | "datetime";
  /** Köprü: dış adres (http, mailto…) ya da "#slide:<kimlik>", "#next", "#prev", "#first", "#last". */
  link?: string;
}

export type AutoNumScheme =
  | "arabicPeriod"
  | "arabicParenR"
  | "romanUcPeriod"
  | "romanLcPeriod"
  | "alphaUcPeriod"
  | "alphaLcPeriod"
  | "alphaLcParenR"
  | "alphaUcParenR";

export type BulletDef =
  | { type: "none" }
  | { type: "char"; char: string; font?: string; color?: Color; size?: number }
  | { type: "auto"; scheme: AutoNumScheme; startAt?: number; color?: Color; size?: number };

export type Align = "l" | "ctr" | "r" | "just";

export interface ParaProps {
  align?: Align;
  bullet?: BulletDef;
  /** Sol kenar boşluğu (px). */
  marL?: number;
  /** İlk satır girintisi (px, asılı için negatif). */
  indent?: number;
  /** Satır aralığı yüzdesi (100 = tek). */
  lineSpacing?: number;
  /** Tam satır aralığı (pt); varsa lineSpacing yerine. */
  lineSpacingPt?: number;
  spaceBefore?: number;
  spaceAfter?: number;
}

export interface Run {
  text: string;
  props: RunProps;
}

export interface Paragraph {
  runs: Run[];
  level: number;
  props: ParaProps;
  /** Paragraf sonu (boş paragrafta yazılacak metnin) biçimi. */
  endProps?: RunProps;
}

export interface LevelStyle {
  para: ParaProps;
  run: RunProps;
}

export type Anchor = "t" | "ctr" | "b";

export interface TextBody {
  paragraphs: Paragraph[];
  anchor?: Anchor;
  wrap?: boolean;
  autofit?: "none" | "shrink" | "shape";
  /** Taşma küçültmesi (0..1). */
  fontScale?: number;
  lnSpcReduction?: number;
  insets?: { l: number; t: number; r: number; b: number };
  vert?: "horz" | "vert" | "vert270" | "stacked";
  /** Şekle özgü düzey stilleri (a:lstStyle). */
  levels?: Partial<LevelStyle>[];
  /** Yer tutucu istem metni (düzenlerde). */
  prompt?: string;
}

export type PlaceholderType = "title" | "ctrTitle" | "subTitle" | "body" | "obj" | "pic" | "tbl" | "dt" | "ftr" | "sldNum";

export interface Placeholder {
  type: PlaceholderType;
  idx?: number;
  /** "half" vb. boyut ipucu; yalnız gidiş-dönüş için. */
  sz?: string;
  orient?: string;
}

export interface ShapeBase {
  id: string;
  name: string;
  x: number;
  y: number;
  w: number;
  h: number;
  /** derece, saat yönünde */
  rot: number;
  flipH?: boolean;
  flipV?: boolean;
  hidden?: boolean;
  descr?: string;
  /**
   * Şekle tıklanınca izlenecek köprü (gösteride). Dış adres ya da
   * "#slide:<kimlik>", "#next", "#prev", "#first", "#last".
   */
  link?: string;
  /** Fare üzerine gelince izlenecek köprü (Eylem Ayarları › Fare Üzerine Gelme). */
  hoverLink?: string;
}

export interface PathCmd {
  c: "M" | "L" | "C" | "Q" | "Z";
  pts: number[];
}

export interface CustomPath {
  w: number;
  h: number;
  cmds: PathCmd[];
  fill?: boolean;
  stroke?: boolean;
}

export interface SpShape extends ShapeBase {
  kind: "sp";
  /** ECMA-376 prstGeom adı ("rect", "roundRect", "ellipse" …) ya da "custom". */
  geom: string;
  adj?: Record<string, number>;
  paths?: CustomPath[];
  /** Çizim sekmesiyle oluşturulmuş serbest mürekkep çizimi. */
  ink?: boolean;
  fill: Fill;
  line: Line;
  effects?: Effects;
  text?: TextBody;
  ph?: Placeholder;
  txBox?: boolean;
  /** Bağlayıcı (cxnSp) olarak yazılsın. */
  connector?: boolean;
}

export interface PicAdjust {
  brightness?: number; // -1..1
  contrast?: number; // -1..1
  grayscale?: boolean;
  alpha?: number;
}

export interface MediaInfo {
  kind: "video" | "audio";
  /** data: URL */
  src: string;
  mime: string;
  /** Gösteride slayt açılınca kendiliğinden başlasın. */
  autoPlay?: boolean;
  loop?: boolean;
}

export interface PicShape extends ShapeBase {
  kind: "pic";
  /** Video/ses: resim, ortamın kapak (poster) görüntüsüdür. */
  media?: MediaInfo;
  src: string;
  crop?: { l: number; t: number; r: number; b: number };
  geom?: string;
  line: Line;
  effects?: Effects;
  adjust?: PicAdjust;
  ph?: Placeholder;
  /** Yakınlaştırma (Slayt / Özet): gösteride tıklanınca hedef slayda yakınlaşarak gidilir. */
  zoom?: ZoomInfo;
}

export interface ZoomInfo {
  slideId: string;
  /** Özet yakınlaştırma: hedef bölüm oynatılır. */
  summary?: boolean;
  /** Hedef (bölüm) bitince yakınlaştırma slaydına dönülür. */
  ret?: boolean;
}

export interface CellBorder {
  color: Color;
  width: number;
}

export interface TableCell {
  text: TextBody;
  fill?: Fill;
  gridSpan?: number;
  rowSpan?: number;
  hMerge?: boolean;
  vMerge?: boolean;
  borders?: { l?: CellBorder | null; r?: CellBorder | null; t?: CellBorder | null; b?: CellBorder | null };
  anchor?: Anchor;
}

export interface TableRow {
  /** Gerçek (içerikle büyümüş olabilen) satır yüksekliği; dosyaya yazılan. */
  h: number;
  /** İçerik büyütmeden önceki en küçük yükseklik (yalnız oturum içinde; yoksa h). */
  minH?: number;
  cells: TableCell[];
}

export interface TableStyleOptions {
  styleId: string;
  firstRow: boolean;
  lastRow: boolean;
  firstCol: boolean;
  lastCol: boolean;
  bandRow: boolean;
  bandCol: boolean;
}

export interface TableShape extends ShapeBase {
  kind: "table";
  cols: number[];
  rows: TableRow[];
  style: TableStyleOptions;
}

export interface GroupShape extends ShapeBase {
  kind: "group";
  /** Alt şekiller slayt koordinatlarında (grup döndürülmemiş varsayılarak). */
  children: Shape[];
}

export type ChartType = "col" | "bar" | "line" | "pie" | "doughnut" | "area" | "scatter" | "radar" | "combo";

/** Sütun/çubuk/çizgi/alan gruplandırması (c:grouping). */
export type ChartGrouping = "clustered" | "stacked" | "percentStacked";
export type DataLabelPos = "ctr" | "inEnd" | "inBase" | "outEnd" | "bestFit" | "t" | "b" | "l" | "r" | "callout";

export interface ChartSeries {
  name: string;
  values: (number | null)[];
  color?: Color;
}

export interface ChartData {
  type: ChartType;
  /** Grafik başlığı; boşsa tek seride seri adı, değilse "Grafik Başlığı". */
  title?: string;
  showTitle: boolean;
  categories: string[];
  series: ChartSeries[];
  legend: "b" | "t" | "l" | "r" | "none";
  dataLabels?: boolean;
  /** Veri etiketi konumu (c:dLblPos; "callout" = veri belirtme çizgisi). Yoksa türün varsayılanı. */
  labelPos?: DataLabelPos;
  /** Dağılım alt türü (yoksa yalnız işaretçiler). */
  scatterStyle?: "lineMarker" | "line" | "smoothMarker" | "smooth";
  /** Radar alt türü (yoksa yalnız çizgi). */
  radarStyle?: "marker" | "filled";
  /** Birleşik: çizgi serisi ikincil eksende. */
  secondaryLine?: boolean;
  /** Yoksa kümelenmiş (çizgi/alan için standart). */
  grouping?: ChartGrouping;
  /** Grafik stili (1–8; c:style). */
  style?: number;
  /** Renkleri Değiştir: renk paleti dizini (0 = Renkli Palet 1). */
  palette?: number;
  /** Eksen başlıkları. */
  axisTitles?: { cat?: string; val?: string };
  /** Birincil ana yatay kılavuz çizgileri (varsayılan açık). */
  gridMajor?: boolean;
  gridMinor?: boolean;
  /** Grafik alanı dolgusu ve anahattı (Biçim sekmesi). */
  fill?: Color;
  line?: Color;
}

export interface ChartShape extends ShapeBase {
  kind: "chart";
  chart: ChartData;
}

export type Shape = SpShape | PicShape | TableShape | GroupShape | ChartShape;

export type TransitionType =
  | "none"
  | "fade"
  | "push"
  | "wipe"
  | "split"
  | "reveal"
  | "cut"
  | "randomBar"
  | "shape"
  | "uncover"
  | "cover"
  | "zoom"
  | "dissolve"
  | "checker"
  | "blinds"
  | "wheel"
  | "wedge"
  | "newsflash"
  /** Dönüştür: aynı adlı/aynı metinli şekiller konum, boyut ve renk değiştirerek geçer. */
  | "morph";

export interface Transition {
  type: TransitionType;
  /** Efekt seçeneği (yön ya da varyant): "l","r","u","d","horz","vert","in","out","circle","diamond","plus"… */
  option?: string;
  /** saniye */
  duration: number;
  advanceOnClick: boolean;
  /** saniye; tanımlıysa otomatik ilerler */
  advanceAfter?: number;
  /** Geçiş sesi (show/sounds kimliği; "stop" önceki sesi durdurur). */
  sound?: string;
}

export type AnimClass = "entr" | "exit" | "emph";

export interface Animation {
  id: string;
  shapeId: string;
  cls: AnimClass;
  effect: string;
  option?: string;
  trigger: "click" | "with" | "after";
  /** saniye */
  duration: number;
  delay: number;
}

export interface Comment {
  id: string;
  author: string;
  text: string;
  x: number;
  y: number;
  date: number;
}

export interface Slide {
  id: string;
  layoutId: string;
  shapes: Shape[];
  background?: Fill;
  /** Arka plan grafiklerini gizle. */
  hideMasterShapes?: boolean;
  hidden?: boolean;
  notes: string;
  transition?: Transition;
  animations: Animation[];
  comments?: Comment[];
  /** Bu slayttan başlayan bölüm (PowerPoint 2010+ bölümleri). */
  section?: { id: string; name: string };
}

export type LayoutType =
  | "title"
  | "obj"
  | "secHead"
  | "twoObj"
  | "twoTxTwoObj"
  | "titleOnly"
  | "blank"
  | "objTx"
  | "picTx"
  | "vertTx"
  | "cust";

export interface Layout {
  id: string;
  type: LayoutType;
  name: string;
  /** Yer tutucular (ph) ve süsleme şekilleri. */
  shapes: Shape[];
  background?: Fill;
  showMasterShapes: boolean;
  /** Seçili slaytlara uygulanan ayrı tema (ikincil asıl slayt); yoksa sununun teması. */
  theme?: Theme;
}

export interface TextStyles {
  title: LevelStyle[];
  body: LevelStyle[];
  other: LevelStyle[];
}

export interface Master {
  background: Fill;
  /** Yer tutucular (ph) ve süsleme şekilleri. */
  shapes: Shape[];
  textStyles: TextStyles;
}

export interface Theme {
  name: string;
  /** Uygulamanın kendi tema kimliği (yoksa içe aktarılmış tema). */
  id?: string;
  variant?: number;
  colors: Record<SchemeColorName, string>;
  fonts: { major: string; minor: string };
}

export interface HeaderFooter {
  dateTime: boolean;
  dateAuto: boolean;
  dateFixed: string;
  slideNumber: boolean;
  footer: boolean;
  footerText: string;
  hideOnTitle: boolean;
}

export interface DocProps {
  title?: string;
  subject?: string;
  creator?: string;
  created?: string;
  modified?: string;
}

/** Özel slayt gösterisi (p:custShow): adlandırılmış slayt listesi. */
export interface CustomShow {
  id: string;
  name: string;
  slideIds: string[];
}

export interface Presentation {
  width: number;
  height: number;
  theme: Theme;
  master: Master;
  layouts: Layout[];
  slides: Slide[];
  headerFooter: HeaderFooter;
  props: DocProps;
  /** Asıl slayt/düzenler elle düzenlendi (slayt boyutu değişince yeniden üretilmez). */
  masterEdited?: boolean;
  /** Özel slayt gösterileri. */
  customShows?: CustomShow[];
}
