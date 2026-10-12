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
  | "sort"
  | "bookmark"
  | "columns"
  | "imageSize"
  | "altText"
  | "about";

export interface DialogState {
  kind: DialogKind;
  /** İletişim kutusuna özel başlangıç verisi (ör. Bul ve Değiştir'de açılacak sekme). */
  arg?: string;
}

export type ViewMode = "print" | "web" | "read";

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
}

export interface BelgeCtx {
  editor: Editor;
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
  headerFooterEditing: "header" | "footer" | null;
  newComment: () => void;
  toggleTrack: () => void;
  tracking: boolean;
  formatPainter: { active: boolean; locked: boolean };
  startFormatPainter: (lock: boolean) => void;
  insertImageFromFile: () => void;
  updateToc: () => void;
  /** Dipnot / son not ekler ve metnine odaklanır. */
  insertNote: (kind: "footnote" | "endnote") => void;
  notify: (text: string) => void;
  /** Dosya > Hakkında'yı açar. */
  openBackstage: (page?: "about" | "info" | "new" | "open" | "export" | "print") => void;
}

export const BelgeContext = createContext<BelgeCtx | null>(null);

export function useBelge(): BelgeCtx {
  const c = useContext(BelgeContext);
  if (!c) throw new Error("BelgeContext yok");
  return c;
}
