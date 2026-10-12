/**
 * Masaüstü (Electron) kabuğunun preload betiğiyle sayfaya açtığı köprü.
 * Tarayıcıda çalışırken `window.officehereNative` tanımsızdır; çekirdek
 * dosya işlemleri bu durumda Web API'lerine düşer.
 */

export type AppKind = "belge" | "tablo" | "sunum";

export interface FileFilter {
  name: string;
  extensions: string[];
}

export interface NativeOpenResult {
  path: string;
  name: string;
  data: ArrayBuffer;
}

export interface NativeSaveResult {
  path: string;
  name: string;
}

export interface RecentFile {
  /** Masaüstünde tam yol, tarayıcıda yalnız dosya adı. */
  path: string;
  name: string;
  app: AppKind;
  openedAt: number;
}

/** Güncelleştirme durumu (masaüstü). */
export type UpdatePhase = "idle" | "checking" | "available" | "downloading" | "ready" | "latest" | "error";

export interface UpdateState {
  phase: UpdatePhase;
  /** Çalışan sürüm. */
  current: string;
  /** "auto": arka planda indirip yeniden başlatınca kurar; "manual": indirme sayfasını açar; "off": denetlenmez. */
  mode: "auto" | "manual" | "off";
  /** Bulunan yeni sürüm. */
  version?: string;
  /** İndirme yüzdesi (0–100). */
  percent?: number;
  /** Sürümün indirme sayfası. */
  url?: string;
  error?: string;
  /** Son başarılı denetim zamanı (ms). */
  checkedAt?: number;
  /** Açılışta ve düzenli aralıklarla kendiliğinden denetlensin mi. */
  autoCheck: boolean;
}

export interface UpdatesApi {
  getState(): Promise<UpdateState>;
  check(): Promise<UpdateState>;
  /** İndirilmiş güncelleştirmeyi kurmak için yeniden başlatır; elle kipte indirme sayfasını açar. */
  install(): Promise<void>;
  setAutoCheck(on: boolean): Promise<UpdateState>;
  onChange(callback: (state: UpdateState) => void): () => void;
}

export interface OfficehereNative {
  platform: string;
  openFile(filters: FileFilter[]): Promise<NativeOpenResult | null>;
  readFile(path: string): Promise<NativeOpenResult | null>;
  saveFileAs(suggestedName: string, filters: FileFilter[], data: ArrayBuffer): Promise<NativeSaveResult | null>;
  writeFile(path: string, data: ArrayBuffer): Promise<boolean>;
  getRecent(): Promise<RecentFile[]>;
  addRecent(file: RecentFile): Promise<void>;
  newWindow(app: AppKind, path?: string): Promise<void>;
  setDocumentState(state: { title: string; dirty: boolean }): void;
  /** Ana süreç bir dosya açılmasını istediğinde (dosya ilişkilendirme, son kullanılanlar). */
  onOpenPath(callback: (path: string) => void): () => void;
  /** Pencere kapatılmak istendiğinde; false dönerse kapanma iptal edilir. */
  onBeforeClose(callback: () => boolean | Promise<boolean>): () => void;
  /** Başlık çubuğu üstündeki yerel pencere düğmelerinin renkleri (Windows/Linux). */
  setTitleBarOverlay?(color: string, symbolColor: string): void;
  /** Geçerli sayfayı yazdırma CSS'iyle PDF olarak kaydeder (konum sorar). */
  exportPdf?(suggestedName: string, options?: { landscape?: boolean; pageSize?: "A4" | "A3" | "A5" | "Letter" | "Legal" | { width: number; height: number } }): Promise<NativeSaveResult | null>;
  /** GitHub Sürümleri üzerinden güncelleştirme denetimi ve kurulumu. */
  updates?: UpdatesApi;
}

declare global {
  interface Window {
    officehereNative?: OfficehereNative;
  }
}

export function getNative(): OfficehereNative | undefined {
  return typeof window === "undefined" ? undefined : window.officehereNative;
}

export function isDesktop(): boolean {
  return getNative() !== undefined;
}
