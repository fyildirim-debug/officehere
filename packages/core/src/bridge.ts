/**
 * Masaüstü (Electron) kabuğunun preload betiğiyle sayfaya açtığı köprü.
 * Tarayıcıda çalışırken `window.officehereNative` tanımsızdır; çekirdek
 * dosya işlemleri bu durumda Web API'lerine düşer.
 */

import type { WriteOutcome } from "./saveErrors";

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

/** Ana süreç dosyayı yazamadı: işletim sistemi hata kodu ve Türkçe nedeni. */
export interface NativeWriteFailure {
  error: { code: string; reason?: string; message?: string };
}

/** Güvenli yazma seçenekleri (otomatik kaydetme). */
export interface NativeWriteOptions {
  /** "<ad>.bak" yan yedeği yoksa önce o alınır; alınamazsa dosya yazılmaz (EBACKUP). */
  sidecar?: boolean;
  /** Üzerine yazılan içerik saklanır; yazılan dosya doğrulanamazsa revertFile ile geri gelir. */
  keepPrevious?: boolean;
}

/** Dosyanın kullanıcı veri dizininde saklanan önceki bir sürümü. */
export interface FileVersion {
  id: string;
  savedAt: number;
  size: number;
}

/** Hata günlüğüne yazılacak kayıt (belge içeriği yazılmaz). */
export interface NativeLogEntry {
  level: "bilgi" | "uyari" | "hata";
  message: string;
  stack?: string;
  /** Olay türü (ör. "window.error", "unhandledrejection"). */
  kind?: string;
}

/** "Hata bildir" penceresinde gösterilen ortam bilgisi. */
export interface AppInfo {
  version: string;
  platform: string;
  arch: string;
  osVersion: string;
  electron?: string;
  chrome?: string;
}

export interface RecentFile {
  /** Masaüstünde tam yol, tarayıcıda yalnız dosya adı. */
  path: string;
  name: string;
  app: AppKind;
  openedAt: number;
  /** Sabitlenen dosya listenin başında durur ve sınır aşılınca atılmaz. */
  pinned?: boolean;
}

/** Belgenin pencere başlığı, kaydedilmemiş değişiklik durumu ve (varsa) diskteki tam yolu. */
export interface DocumentState {
  title: string;
  dirty: boolean;
  path?: string;
}

/** Masaüstü ayarları (kullanıcı verisi dizininde ayarlar.json). */
export interface DesktopSettings {
  /** Uygulama dosyasız açılınca son oturumdaki belgeler yeniden açılsın. */
  oturumuGeriYukle?: boolean;
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
  /** Konum sorup yazar; iptalde null, yazılamazsa { error } (eski kabuklar hata fırlatır). */
  saveFileAs(suggestedName: string, filters: FileFilter[], data: ArrayBuffer): Promise<NativeSaveResult | NativeWriteFailure | null>;
  writeFile(path: string, data: ArrayBuffer): Promise<boolean>;
  /** Güvenli yazma (geçici dosya + fsync + taşıma, yol başına kuyruk); başarısızlıkta hata kodu ve Türkçe neden. */
  writeFileSafe?(path: string, data: ArrayBuffer, options?: NativeWriteOptions): Promise<WriteOutcome>;
  /** Otomatik kaydetmenin son yazmasından önceki içeriği geri yazar (keepPrevious ile yazılmışsa). */
  revertFile?(path: string): Promise<WriteOutcome>;
  /** Dosyanın önceki sürümleri (en yeni önce, en çok 3). */
  listVersions?(path: string): Promise<FileVersion[]>;
  /** Önceki bir sürümü dosyaya geri yazar (geçerli içerik de sürüm olarak saklanır). */
  restoreVersion?(path: string, id: string): Promise<WriteOutcome>;
  /** Yerel hata günlüğüne yazar. */
  log?(entry: NativeLogEntry): void;
  /** Günlüğün son satırları. */
  getLogTail?(lines?: number): Promise<string[]>;
  getAppInfo?(): Promise<AppInfo>;
  /** Bu pencerede açılan dosya başka bir pencerede de açıksa (ana süreç bildirir). */
  onPathConflict?(callback: (path: string) => void): () => void;
  getRecent(): Promise<RecentFile[]>;
  addRecent(file: RecentFile): Promise<void>;
  /** Son kullanılanlarda dosyayı sabitler ya da sabitlemeyi kaldırır. */
  pinRecent?(path: string, pinned: boolean): Promise<void>;
  /** Dosyayı son kullanılanlar listesinden çıkarır (dosyanın kendisine dokunmaz). */
  removeRecent?(path: string): Promise<void>;
  /** Dosyayı sistemin dosya yöneticisinde, klasörü açıp seçili gösterir. */
  showInFolder?(path: string): Promise<void>;
  /** Verilen yollardaki dosyalar hâlâ var mı. */
  recentStatus?(paths: string[]): Promise<Record<string, boolean>>;
  /** Masaüstü ayarları. */
  getSettings?(): Promise<DesktopSettings>;
  setSetting?<K extends keyof DesktopSettings>(key: K, value: DesktopSettings[K]): Promise<DesktopSettings>;
  newWindow(app: AppKind, path?: string): Promise<void>;
  /** Başlık, kirli durum ve belge yolu; ana süreç oturumu bu yolla hatırlar. */
  setDocumentState(state: DocumentState): void;
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
