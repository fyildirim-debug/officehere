import type { Node } from "./ast";
import type { LocaleInfo } from "./coerce";
import type { RefValue, Scalar, Value } from "./values";

/** Değerlendiricinin çalışma kitabından istediği hizmetler. */
export interface EvalHost {
  getCellValue(sheet: number, row: number, col: number): Scalar;
  /** Sayfadaki kullanılan alanın boyutu (tüm sütun/satır başvurularını kırpmak için). */
  getUsedExtent(sheet: number): { rows: number; cols: number };
  sheetIdByName(name: string): number | undefined;
  sheetName(id: number): string | undefined;
  /** Sayfanın sıradaki konumu (0 tabanlı). */
  sheetIndex(id: number): number;
  sheetCount(): number;
  getDefinedName(name: string, sheet: number): Node | undefined;
  isRowHidden(sheet: number, row: number): boolean;
  /** Hücredeki formül metni ("=" ile) ya da undefined. */
  getFormulaText(sheet: number, row: number, col: number): string | undefined;
  /** Hücre ALTTOPLAM/TOPLAMA formülü mü (iç içe alt toplamları atlamak için)? */
  isSubtotalCell(sheet: number, row: number, col: number): boolean;
  /** Geçerli zamanın seri numarası. */
  now(): number;
  random(): number;
}

export interface FnContext {
  host: EvalHost;
  /** Formülün bulunduğu hücre. */
  sheet: number;
  row: number;
  col: number;
  locale: LocaleInfo;
  /** Bir düğümü bu bağlamda değerlendirir (LET gibi tembel fonksiyonlar için). */
  evalNode(node: Node, scope?: Map<string, Value>): Value;
  scope?: Map<string, Value>;
  /** DOLAYLI/KAYDIR gibi dinamik başvuru kullanıldığında çağrılır. */
  onDynamicRef?: (ref: RefValue) => void;
}
