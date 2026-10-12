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
  | "timePeriod";

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
}

export interface ImageDef {
  id: string;
  name: string;
  src: string;
  anchor: Anchor;
  width: number;
  height: number;
}

export interface ShapeDef {
  id: string;
  kind: "rect" | "roundRect" | "ellipse" | "triangle" | "arrow" | "line" | "textbox";
  text: string;
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
}

export type ViewMode = "normal" | "pageLayout" | "pageBreak";

export interface Selection {
  ranges: Rect[];
  active: CellPos;
  /** Shift ile genişletmenin bağlantı noktası (son aralık için). */
  anchor: CellPos;
}
