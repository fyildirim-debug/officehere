/** Çalışma kitabı modelinin ortak türleri. */
import type { CellStyle } from "./styles";

export interface Rect {
  r1: number;
  c1: number;
  r2: number;
  c2: number;
}

export interface CellPos {
  row: number;
  col: number;
}

export interface Note {
  text: string;
  author?: string;
}

export interface ValidationRule {
  type: "list" | "whole" | "decimal" | "date" | "textLength" | "any" | "custom";
  operator?: "between" | "notBetween" | "equal" | "notEqual" | "greaterThan" | "lessThan" | "greaterThanOrEqual" | "lessThanOrEqual";
  /** Liste için "a;b;c" ya da "=$A$1:$A$5"; diğerleri için değer/formül. */
  formula1?: string;
  formula2?: string;
  allowBlank?: boolean;
  showDropdown?: boolean;
  errorTitle?: string;
  error?: string;
  errorStyle?: "stop" | "warning" | "information";
  promptTitle?: string;
  prompt?: string;
}

export interface Validation {
  id: string;
  ranges: Rect[];
  rule: ValidationRule;
}

export type CondRuleType =
  | "cellIs"
  | "containsText"
  | "notContainsText"
  | "beginsWith"
  | "endsWith"
  | "duplicateValues"
  | "uniqueValues"
  | "top10"
  | "aboveAverage"
  | "dataBar"
  | "colorScale"
  | "expression"
  | "containsBlanks"
  | "notContainsBlanks"
  | "containsErrors"
  | "timePeriod"
  | "iconSet";

export interface CondRule {
  id: string;
  type: CondRuleType;
  ranges: Rect[];
  operator?: "greaterThan" | "lessThan" | "between" | "equal" | "notEqual" | "greaterThanOrEqual" | "lessThanOrEqual" | "notBetween";
  /** Kanonik formül ya da değer metinleri. */
  values?: string[];
  text?: string;
  /** top10: sıra, yüzde, alt. */
  rank?: number;
  percent?: boolean;
  bottom?: boolean;
  /** aboveAverage: altında mı. */
  below?: boolean;
  style?: Partial<CellStyle>;
  /** Veri çubuğu rengi. */
  color?: string;
  /** Renk ölçeği: 2 ya da 3 renk (min, [orta], maks). */
  colors?: string[];
  timePeriod?: string;
  /** Simge kümesi adı (3Arrows, 3TrafficLights1, 5Rating…). */
  iconSet?: string;
  /** Simge eşikleri (ilki her zaman en küçük): yüzde, sayı ya da yüzdebirlik. */
  iconThresholds?: { type: "percent" | "num" | "percentile"; value: number }[];
  reverseIcons?: boolean;
  /** Yalnız simgeyi göster (değer gizli). */
  iconOnly?: boolean;
  stopIfTrue?: boolean;
  priority?: number;
}

export type FilterCriteria =
  | { kind: "values"; values: string[]; blanks: boolean }
  | { kind: "custom"; op1: string; v1: string; join?: "and" | "or"; op2?: string; v2?: string }
  | { kind: "top10"; count: number; bottom: boolean; percent: boolean }
  | { kind: "color"; color: string };

export interface AutoFilter {
  range: Rect;
  columns: Record<number, FilterCriteria>;
  /** Sıralama durumu (başlık okunda gösterilir). */
  sort?: { col: number; desc: boolean };
}

export interface TableDef {
  id: string;
  name: string;
  ref: Rect;
  style: string;
  headerRow: boolean;
  totalRow: boolean;
  bandedRows: boolean;
  bandedCols: boolean;
  firstColumn: boolean;
  lastColumn: boolean;
  filterButton: boolean;
  columns: string[];
  /** Toplam satırındaki fonksiyonlar (sütun dizinine göre). */
  totals?: Record<number, string>;
  filter?: AutoFilter;
}

export type ChartType = "column" | "columnStacked" | "bar" | "barStacked" | "line" | "lineMarkers" | "pie" | "doughnut" | "area" | "areaStacked" | "scatter";

export interface Anchor {
  row: number;
  col: number;
  /** Hücre içi uzaklık (px, %100 yakınlaştırmada). */
  dx: number;
  dy: number;
}

export interface ChartDef {
  id: string;
  type: ChartType;
  /** Veri aralığı ("Sayfa1!$A$1:$C$5"). */
  range: string;
  /** Seriler satırlarda mı (Satır/Sütun Değiştir). */
  seriesInRows: boolean;
  title: string;
  showTitle: boolean;
  showLegend: boolean;
  legendPos: "right" | "bottom" | "top" | "left";
  showDataLabels: boolean;
  showGridlines: boolean;
  palette: number;
  style: number;
  anchor: Anchor;
  width: number;
  height: number;
  xTitle?: string;
  yTitle?: string;
  /**
   * Açık seri tanımları (dosyadan okunan grafiklerde): varsa veri aralığından
   * seri tahmini yapılmaz. Aralıklar nitelenmiş metindir ("Bütçe!$E$3:$E$9").
   */
  series?: ChartSeriesDef[];
  /** Seçim Bölmesinde gizlenmiş. */
  hidden?: boolean;
  /** Seçim Bölmesinde verilen ad. */
  name?: string;
  /** Nesneler arası yığın sırası (yoksa oluşturulma sırası). */
  z?: number;
  /** Grafik alanı dolgusu (Grafik Alanını Biçimlendir). */
  fill?: string;
  /** Grafik alanı kenarlığı rengi; "none" kenarlıksız. */
  border?: string;
  /** Gizli satır ve sütunlardaki verileri göster (dosyada plotVisOnly=0). */
  showHidden?: boolean;
  /** Seri biçimleri (seri sırasıyla): dolgu/çizgi rengi, kenarlık rengi, çizgi kalınlığı (pt). */
  seriesFmt?: ({ color?: string; border?: string; lineWidth?: number } | null)[];
  /** Değer ekseni: en küçük/en büyük sınır, ana birim, sayı biçimi, ikincil kılavuz çizgileri. */
  valAxis?: { min?: number; max?: number; major?: number; numFmt?: string; minorGrid?: boolean };
  /** Veri etiketi konumu (sütun/çubuk). */
  labelPos?: "outEnd" | "inEnd" | "ctr" | "inBase";
}

export interface ChartSeriesDef {
  /** Seri adı hücresi (nitelenmiş başvuru) ya da sabit ad. */
  nameRef?: string;
  name?: string;
  values: string;
  categories?: string;
  /** Dağılım grafiğinde x değerleri. */
  xValues?: string;
}

export interface ImageDef {
  id: string;
  name: string;
  hidden?: boolean;
  /** Nesneler arası yığın sırası (yoksa oluşturulma sırası). */
  z?: number;
  src: string;
  anchor: Anchor;
  width: number;
  height: number;
}

export interface ShapeDef {
  id: string;
  kind: "rect" | "roundRect" | "ellipse" | "triangle" | "arrow" | "line" | "textbox";
  text: string;
  hidden?: boolean;
  /** Seçim Bölmesinde verilen ad. */
  name?: string;
  /** Nesneler arası yığın sırası (yoksa oluşturulma sırası). */
  z?: number;
  fill: string;
  stroke: string;
  anchor: Anchor;
  width: number;
  height: number;
}

export interface PageSetup {
  orientation: "portrait" | "landscape";
  paper: "A4" | "A3" | "A5" | "Letter" | "Legal";
  /** İnç. */
  margins: { top: number; bottom: number; left: number; right: number; header: number; footer: number };
  printArea?: Rect;
  printGridlines: boolean;
  printHeadings: boolean;
  header?: string;
  footer?: string;
  fitToWidth?: boolean;
  /** Ölçek: sayfayı tek sayfaya / tüm sütunları / tüm satırları tek sayfaya sığdır (fitToWidth'in yerine geçer). */
  fitMode?: "sheet" | "cols" | "rows";
  /** El ile eklenen sayfa sonları: yeni sayfanın ilk satırı / sütunu (0 tabanlı, artan sırada). */
  rowBreaks?: number[];
  colBreaks?: number[];
  /** Yazdırma başlıkları: her sayfada yinelenecek üst satırlar / sol sütunlar. */
  titleRows?: { r1: number; r2: number };
  titleCols?: { c1: number; c2: number };
  /** Siyah beyaz, taslak kalitesi, sayfa sırası (Sayfa Yapısı > Çalışma Sayfası). */
  blackAndWhite?: boolean;
  draft?: boolean;
  pageOrder?: "downThenOver" | "overThenDown";
}

export type PivotFn = "sum" | "count" | "average" | "max" | "min" | "product" | "countNums";

export interface PivotValueField {
  /** Kaynak sütun dizini (kaynak aralığın içinde, 0 tabanlı). */
  field: number;
  fn: PivotFn;
  /** Özel ad ("Toplam Satış" gibi); yoksa "Toplam: <alan>". */
  name?: string;
  numFmt?: string;
}

export interface PivotFilter {
  field: number;
  /** Seçili öğeler (metin anahtarları); null = tümü. */
  selected: string[] | null;
}

export interface PivotDef {
  id: string;
  name: string;
  /** Kaynak: sayfa kimliği ve aralık (başlık satırı dahil). */
  sourceSheet: number;
  source: Rect;
  /** Raporun sol üst hücresi (filtreler bunun üstünde gösterilir). */
  anchor: CellPos;
  rows: number[];
  cols: number[];
  values: PivotValueField[];
  filters: PivotFilter[];
  /** Satır/sütun alanlarındaki öğe filtreleri (alan → gizli öğeler). */
  hidden?: Record<number, string[]>;
  /** Birden çok değer alanı varken "Σ Değerler" satırlarda mı. */
  valuesOnRows?: boolean;
  subtotals: "top" | "bottom" | "none";
  grandRows: boolean;
  grandCols: boolean;
  layout: "compact" | "outline" | "tabular";
  /** Öğe etiketlerini yinele (tablo/anahat formu). */
  repeatLabels?: boolean;
  blankRows?: boolean;
  style: string;
  rowHeaders: boolean;
  colHeaders: boolean;
  bandedRows: boolean;
  bandedCols: boolean;
  /** Son çizilen çıktı aralığı (filtre satırları dahil). */
  output?: Rect;
  /** Sayfadaki filtre açılır düğmeleri (rapor filtresi, satır/sütun etiketleri). */
  buttons?: { row: number; col: number; field: number }[];
  /** Sıralama: alan → azalan mı. */
  sort?: Record<number, boolean>;
}

export interface SparklineGroup {
  id: string;
  type: "line" | "column" | "winLoss";
  /** Her mini grafik: hedef hücre ve veri aralığı ("Sayfa1!A1:F1" ya da "A1:F1"). */
  items: { row: number; col: number; range: string }[];
  /** Mini grafik rengi. */
  color: string;
  /** Çizgi kalınlığı (punto). */
  weight?: number;
  /** Göster: işaretçiler (yalnız çizgi), yüksek/düşük/ilk/son/negatif noktalar. */
  markers?: boolean;
  highPoint?: boolean;
  lowPoint?: boolean;
  firstPoint?: boolean;
  lastPoint?: boolean;
  negative?: boolean;
  /** İşaretçi renkleri. */
  markerColor?: string;
  negativeColor?: string;
  highColor?: string;
  lowColor?: string;
  firstColor?: string;
  lastColor?: string;
  /** Yatay ekseni göster (sıfır çizgisi). */
  showAxis?: boolean;
  axisColor?: string;
  /** Dikey eksen en küçük/en büyük: her mini grafik için otomatik, grupta aynı ya da özel değer. */
  minAxis?: "individual" | "group" | "custom";
  maxAxis?: "individual" | "group" | "custom";
  minValue?: number;
  maxValue?: number;
  /** Boş hücreler: boşluk, sıfır ya da veri noktalarını çizgiyle bağla. */
  emptyAs?: "gap" | "zero" | "span";
  /** Gizli satır ve sütunlardaki verileri göster. */
  showHidden?: boolean;
  /** Sağdan sola çizim (Eksen > Sağdan Sola Çizim Verileri). */
  rightToLeft?: boolean;
}

export type ViewMode = "normal" | "pageLayout" | "pageBreak";

export interface Selection {
  ranges: Rect[];
  active: CellPos;
  /** Shift ile genişletmenin bağlantı noktası (son aralık için). */
  anchor: CellPos;
}

/** Senaryo Yöneticisi: değişen hücreler ve senaryodaki değerleri (ECMA-376 §18.3.1.75 scenario). */
export interface Scenario {
  name: string;
  cells: { row: number; col: number; value: string | number | boolean | null }[];
  comment?: string;
  /** Koruma: Değişiklikleri engelle / Gizle (sayfa korunurken etkilidir). */
  locked?: boolean;
  hidden?: boolean;
  user?: string;
}

/** Durum Çözümlemesi > Veri Tablosu: sonuç aralığı ve giriş hücreleri (TABLE işlevi). */
export interface DataTableDef {
  id: string;
  /** Sonuç hücreleri (sol üst köşedeki formül satırı/sütunu hariç). */
  ref: Rect;
  /** Satır giriş hücresi (değerler üst satırdan) ve sütun giriş hücresi (değerler sol sütundan). */
  rowInput?: CellPos;
  colInput?: CellPos;
}

/** Çalışma kitabı yapı koruması (Gözden Geçir > Çalışma Kitabını Koru). */
export interface WorkbookProtection {
  structure: boolean;
  windows?: boolean;
  /** Parola karması (ECMA-376 Bölüm 1 §18.2.29). */
  hash?: { algorithm: string; value: string; salt: string; spin: number };
  /** Eski 16 bit parola karması (workbookPassword, onaltılık). */
  legacy?: string;
}

/** Dilimleyici: tablo ya da özet tablo alanına bağlı düğmeli filtre. */
export interface SlicerDef {
  id: string;
  name: string;
  caption: string;
  source: { kind: "table"; tableId: string; column: string } | { kind: "pivot"; pivotId: string; field: number };
  /** Seçili öğeler (görünen metin); boşsa tümü. */
  selected: string[];
  anchor: { row: number; col: number; dx: number; dy: number };
  width: number;
  height: number;
  columns: number;
  style: string;
}
