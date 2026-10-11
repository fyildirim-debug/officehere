/** Ortak arayüz metinleri (Türkçe). Uygulamalar kendi metinlerini yanında tutar. */
export const tr = {
  file: "Dosya",
  home: "Giriş",
  insert: "Ekle",
  draw: "Çizim",
  design: "Tasarım",
  layout: "Düzen",
  view: "Görünüm",
  help: "Yardım",
  review: "Gözden Geçir",
  save: "Kaydet",
  saveAs: "Farklı Kaydet",
  open: "Aç",
  new: "Yeni",
  print: "Yazdır",
  export: "Dışarı Aktar",
  close: "Kapat",
  info: "Bilgi",
  options: "Seçenekler",
  undo: "Geri Al",
  redo: "Yinele",
  cut: "Kes",
  copy: "Kopyala",
  paste: "Yapıştır",
  search: "Ara",
  ok: "Tamam",
  cancel: "İptal",
  apply: "Uygula",
  yes: "Evet",
  no: "Hayır",
  dontSave: "Kaydetme",
  unsavedChanges: "Değişiklikleri kaydetmek istiyor musunuz?",
  recent: "Son Kullanılanlar",
  zoom: "Yakınlaştır",
} as const;

/** Türkçe sayı biçimi (1.234,56). */
export const trNumber = new Intl.NumberFormat("tr-TR");

export function formatDateTime(ts: number): string {
  return new Intl.DateTimeFormat("tr-TR", { dateStyle: "medium", timeStyle: "short" }).format(ts);
}
