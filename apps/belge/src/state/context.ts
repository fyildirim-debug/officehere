// Uygulama bağlamı: editör, belge ayarları ve kabuk eylemleri şerit/iletişim kutularına buradan ulaşır.
import { createContext, useContext } from "react";
import type { Editor } from "@tiptap/core";
import type { DocSettings } from "../model/document";
import type { StyleDef } from "../model/styles";
import type { PaginationController } from "../editor/pagination";
import type { FormatState } from "../editor/formatState";

export type DialogKind =
  | "font"
  | "paragraph"
  | "findReplace"
  | "goto"
  | "link"
  | "insertTable"
  | "zoom"
  | "wordCount"
  | "symbol"
  | "dateTime"
  | "pageSetup"
  | "shortcuts"
  | "watermark"
  | "pageBorders"
  | "tableProps"
  | "splitCells"
  | "sort"
  | "bookmark"
  | "columns"
  | "imageSize"
  | "altText"
  | "about"
  | "autocorrect"
  | "style"
  | "tabs"
  | "caption"
  | "tableOfFigures"
  | "crossRef"
  | "source"
  | "sources"
  | "recipients"
  | "merge"
  | "tableSort"
  | "textToTable"
  | "tableToText"
  | "cellMargins"
  | "lineNumbers"
  | "crop"
  | "equation"
  | "chart"
  | "markEntry"
  | "insertIndex"
  | "compare"
  | "protect"
  | "numberingValue"
  | "dropCap";

export interface DialogState {
  kind: DialogKind;
  /** İletişim kutusuna özel başlangıç verisi (ör. Bul ve Değiştir'de açılacak sekme). */
  arg?: string;
}

export type ViewMode = "print" | "web" | "read" | "draft" | "outline";

export interface ViewOptions {
  mode: ViewMode;
  zoom: number;
  ruler: boolean;
  gridlines: boolean;
  navPane: boolean;
  /** Yazdırma Düzeni'nde sayfalar arası boşluk gizli. */
  focusMode: boolean;
  showMarks: boolean;
  /** Açıklamalar kenar bölmesi. */
  showComments: boolean;
  /** Görünüm > Tek Sayfa: sığsa da sayfalar yan yana dizilmez. */
  onePage?: boolean;
  /** Düzenleyici bölmesi (Yazım ve Dil Bilgisi, F7). */
  proofPane?: boolean;
  /** Stiller bölmesi. */
  stylesPane?: boolean;
  /** Gözden Geçir > İşaretleme görünümü. */
  markup?: "all" | "simple" | "none" | "original";
  /** Gözden Geçir > İşaretlemeyi Göster > Balonlar. */
  balloons?: "revisions" | "inline" | "formatting";
}

export interface BelgeCtx {
  /** Şeridin hedefi (belge gövdesi, üst/alt bilgi ya da not düzenleyicisi). */
  editor: Editor;
  /** Belge gövdesi düzenleyicisi (bölümler, sayfa yapısı). */
  mainEditor: Editor;
  /** Şeridin hedefi belge gövdesi mi (üst/alt bilgi ya da not düzenlenmiyor). */
  inBody: boolean;
  fmt: FormatState;
  settings: DocSettings;
  updateSettings: (fn: (s: DocSettings) => DocSettings) => void;
  styles: Record<string, StyleDef>;
  controller: PaginationController;
  view: ViewOptions;
  setView: (patch: Partial<ViewOptions>) => void;
  openDialog: (kind: DialogKind, arg?: string) => void;
  /** Gezinti bölmesini arama sekmesiyle açar. */
  openFind: () => void;
  editHeaderFooter: (which: "header" | "footer" | null) => void;
  /** Verilen sayfanın üst/alt bilgisini düzenlemeye geçer (Önceki/Sonraki). */
  editHeaderFooterAt: (which: "header" | "footer", page: number) => void;
  headerFooterEditing: "header" | "footer" | null;
  /** Üst/alt bilgisi düzenlenen sayfa (0 tabanlı). */
  headerFooterPage: number;
  newComment: () => void;
  toggleTrack: () => void;
  tracking: boolean;
  formatPainter: { active: boolean; locked: boolean };
  startFormatPainter: (lock: boolean) => void;
  insertImageFromFile: () => void;
  updateToc: () => void;
  /** Dipnot / son not ekler ve metnine odaklanır. */
  insertNote: (kind: "footnote" | "endnote") => void;
  notify: (text: string, kind?: "info" | "warning" | "error" | "success") => void;
  /** Dosya > Hakkında'yı açar. */
  openBackstage: (page?: "about" | "info" | "new" | "open" | "export" | "print") => void;
  /** Şerit sekmesine geçer (ör. resim eklenince Resim Biçimi). */
  showTab: (id: string) => void;
  /** Sesli Oku açık mı; aç/kapat (Ctrl+Alt+Boşluk). */
  readAloud: boolean;
  toggleReadAloud: () => void;
}

export const BelgeContext = createContext<BelgeCtx | null>(null);

export function useBelge(): BelgeCtx {
  const c = useContext(BelgeContext);
  if (!c) throw new Error("BelgeContext yok");
  return c;
}
