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
  /** İndirilenler ve e-posta eklerinden (internetten) gelen dosyalar salt okunur açılsın (varsayılan kapalı). */
  indirilenleriSaltOkunurAc?: boolean;
}

/** Görüntüleyici kipi: internetten gelen dosyanın salt okunur açılması (kaynak tespiti ana süreçte). */
export interface ViewerApi {
  /** Dosya salt okunur açılmalı mı: seçenek açık, kaynağı internet ve bu oturumda düzenlemeye açılmamış. */
  check(path: string): Promise<boolean>;
  /** "Düzenlemeyi etkinleştir": dosya bu oturum boyunca düzenlenebilir açılır. */
  allowEditing(path: string): Promise<void>;
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
  /** Güncelleme kanalı: "beta" ön sürümleri de alır (eski kabuklarda yok: kararlı). */
  channel?: UpdateChannel;
}

export type UpdateChannel = "kararli" | "beta";

export interface UpdatesApi {
  getState(): Promise<UpdateState>;
  check(): Promise<UpdateState>;
  /** İndirilmiş güncelleştirmeyi kurmak için yeniden başlatır; elle kipte indirme sayfasını açar. */
  install(): Promise<void>;
  setAutoCheck(on: boolean): Promise<UpdateState>;
  /** Kanalı değiştirir; betadan kararlıya dönüşte sürüm geri alınmaz, sonraki kararlı beklenir. */
  setChannel?(channel: UpdateChannel): Promise<UpdateState>;
  onChange(callback: (state: UpdateState) => void): () => void;
}

/**
 * Yazım denetimi sözlükleri (masaüstü). "sistem": macOS'un kendi denetleyicisi; "yok": indirilmedi
 * (onay yok); "indiriliyor"; "hazir"; "hata": indirilemedi.
 */
export interface SpellcheckState {
  status: "sistem" | "yok" | "indiriliyor" | "hazir" | "hata";
  /** İlk açılış sorusu gösterilsin mi (hiç sorulmadı ve sözlük yerelde yok). */
  ask: boolean;
  /** İndirilecek yaklaşık boyut (MB). */
  sizeMb: number;
}

/** PDF türü: standart PDF ya da arşiv (PDF/A) uyumluluk düzeyi. */
export type PdfStandard = "pdf" | "pdfa-2b" | "pdfa-1b";

/** Dışa aktarmadan sonra: hiçbir şey, dosyayı varsayılan uygulamada aç ya da klasörünü göster. */
export type PdfAfterExport = "none" | "open" | "folder";

/** Masaüstünde PDF dışa aktarma seçenekleri (eski kabuklar yalnız landscape/pageSize'ı okur). */
export interface PdfExportOptions {
  landscape?: boolean;
  pageSize?: "A4" | "A3" | "A5" | "Letter" | "Legal" | { width: number; height: number };
  /** Çıktı sayfaları ("1-3, 5"); boşsa tümü. */
  pageRanges?: string;
  standard?: PdfStandard;
  /** Başlıklardan yer imleri (etiketli PDF gerektirir; açıkken etiket de üretilir). */
  outline?: boolean;
  /** Erişilebilirlik için etiketli PDF. */
  tagged?: boolean;
  title?: string;
  author?: string;
  subject?: string;
  keywords?: string;
  /** Oluşturan uygulama (ör. "OfficeHere Belge"). */
  creator?: string;
  after?: PdfAfterExport;
}

/** PDF kaydedildi; standart, saydamlık nedeniyle 1b'den 2b'ye geçildiyse "pdfa-2b"dir. */
export interface PdfExportResult extends NativeSaveResult {
  standard?: PdfStandard;
}

export interface SpellcheckApi {
  getState(): Promise<SpellcheckState>;
  /** true: sözlükleri indir (onay); false: indirme / kaldır. */
  setDownload(on: boolean): Promise<SpellcheckState>;
  onChange(callback: (state: SpellcheckState) => void): () => void;
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
  /** Geçerli sayfayı yazdırma CSS'iyle PDF (ya da PDF/A) olarak kaydeder (konum sorar); iptalde null. */
  exportPdf?(suggestedName: string, options?: PdfExportOptions): Promise<PdfExportResult | null>;
  /** GitHub Sürümleri üzerinden güncelleştirme denetimi ve kurulumu. */
  updates?: UpdatesApi;
  /** Yazım denetimi sözlüklerinin indirme onayı ve durumu. */
  spellcheck?: SpellcheckApi;
  /** İnternetten gelen dosyaları salt okunur açma (görüntüleyici kipi). */
  viewer?: ViewerApi;
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
