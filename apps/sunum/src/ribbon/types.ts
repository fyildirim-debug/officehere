import type { Actions } from "../state/actions";
import type { AppState, PaneKind, ViewMode } from "../state/store";
import type { TextInfo } from "../state/textActions";
import type { ShapeCatalogItem } from "../render/geometry";

export type DialogKind =
  | "font"
  | "paragraph"
  | "insertTable"
  | "headerFooter"
  | "slideSize"
  | "find"
  | "replace"
  | "replaceFonts"
  | "symbol"
  | "dateTime"
  | "zoom"
  | "sizePos"
  | "showSetup"
  | "shortcuts"
  | "about"
  | "lineSpacing"
  | "icons";

export interface AppCommands {
  paste: (mode?: "keep" | "text" | "dest") => void;
  cut: () => void;
  copy: () => void;
  insertPictures: () => void;
  startShow: (from: "start" | "current") => void;
  presenterView: () => void;
  print: () => void;
  /** Dosya > Dışarı Aktar > PDF (masaüstünde doğrudan, tarayıcıda yazdırma penceresiyle). */
  exportPdf: () => void;
  previewTransition: () => void;
  previewAnimations: () => void;
  setView: (v: ViewMode) => void;
  togglePane: (p: PaneKind) => void;
  pickShape: (item: ShapeCatalogItem) => void;
  insertWordArt: (index: number) => void;
  insertDateTime: () => void;
  insertSlideNumber: () => void;
  spelling: () => void;
  zoomTo: (z: number, fit?: boolean) => void;
  /** Ekle > Grafik; grafik türünü değiştir; verileri düzenle. */
  insertChart: () => void;
  changeChartType: () => void;
  editChartData: () => void;
  /** Ekle > Medya > Video / Ses. */
  insertMedia: (kind: "video" | "audio") => void;
  playMedia: () => void;
  /** Ekle > Bağlantı (Ctrl+K). */
  insertLink: () => void;
  openLink: () => void;
  removeLink: () => void;
  /** Bölüm Ekle (ardından yeniden adlandırma kutusu açılır). */
  addSection: () => void;
  renameSection: (id: string) => void;
  /** Asıl görünümde seçili düzeni yeniden adlandırır. */
  renameLayout: () => void;
}

export interface RibbonCtx {
  state: AppState;
  actions: Actions;
  text: TextInfo | null;
  openDialog: (d: DialogKind) => void;
  cmd: AppCommands;
}
