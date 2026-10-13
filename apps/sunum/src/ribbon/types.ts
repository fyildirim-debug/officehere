import type { Actions } from "../state/actions";
import type { AppState, PaneKind, ViewMode } from "../state/store";
import type { TextInfo } from "../state/textActions";
import type { ShapeCatalogItem } from "../render/geometry";
import type { ShowSettings } from "../show/settings";

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
  | "icons"
  | "action"
  | "slideZoom"
  | "summaryZoom"
  | "compress";

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
  /** Seçili nesnenin alternatif metni (Alternatif Metni Görüntüle). */
  altText: () => void;
  /** Ekle > Bağlantı (Ctrl+K). */
  insertLink: () => void;
  openLink: () => void;
  removeLink: () => void;
  /** Bölüm Ekle (ardından yeniden adlandırma kutusu açılır). */
  addSection: () => void;
  renameSection: (id: string) => void;
  /** Asıl görünümde seçili düzeni yeniden adlandırır. */
  renameLayout: () => void;
  /** Slayt Gösterisi > Zamanlama Provası. */
  rehearseTimings: () => void;
  /** Özel slayt gösterisini başlatır. */
  startCustomShow: (id: string) => void;
  /** "Özel Gösteriler" iletişim kutusu. */
  customShows: () => void;
  /** Bul iletişim kutusu (isteğe bağlı aranacak metinle). */
  find: (q?: string) => void;
  /** Slayt boyutunu değiştirir (ölçeklendirme sorusuyla). */
  setSlideSize: (w: number, h: number) => void;
  /** Slaytları Yeniden Kullan bölmesi. */
  reuseSlides: () => void;
  /** Anahattan Slaytlar (.txt/.md/.docx). */
  slidesFromOutline: () => void;
}

export interface RibbonCtx {
  state: AppState;
  actions: Actions;
  text: TextInfo | null;
  openDialog: (d: DialogKind) => void;
  cmd: AppCommands;
  /** Gösteriyi Ayarla değerleri (Slayt Gösterisi sekmesi). */
  show?: { settings: ShowSettings; set: (s: ShowSettings) => void };
}
