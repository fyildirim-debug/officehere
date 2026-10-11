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
