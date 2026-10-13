/**
 * Açık çalışma kitabının dosya biçimi (kaydetme davranışı için). Her yüklenen
 * kitap yeni bir model nesnesi olduğundan durum modele bağlı tutulur; yeni
 * kitap, şablon ya da kurtarma yüklemesinde kendiliğinden .xlsx'e döner.
 */
import type { BookFormat } from "../io/formats";
import type { WorkbookModel } from "../model/workbook";

export interface FileFormatState {
  format: BookFormat;
  /** CSV'de dosyadaki ayırıcı (aynı ayırıcıyla kaydedilir). */
  delimiter?: string;
  /** "Bu biçimde bazı özellikler kaybolabilir" sorusu bu kitap için yanıtlandı. */
  confirmed?: boolean;
}

const states = new WeakMap<WorkbookModel, FileFormatState>();

export function fileFormatOf(wb: WorkbookModel): FileFormatState {
  return states.get(wb) ?? { format: "xlsx" };
}

export function setFileFormat(wb: WorkbookModel, s: FileFormatState): void {
  states.set(wb, s);
}
