/**
 * Komut aramasına (Alt+Q) şeritte düğmesi olmayan komutlar: Dosya menüsü
 * (Kaydet, Yazdır, Dışarı Aktar…) ve galeri öğeleri (geçiş ve animasyon
 * efektleri). Bunlar doğrudan çalıştırılır.
 */
import { ANIMATIONS, CLASS_LABELS, TRANSITIONS, transitionDef } from "../show/catalog";
import type { Actions } from "../state/actions";
import type { CommandEntry } from "./commandIndex";

export interface FileCommandHandlers {
  save: () => void;
  saveAs: () => void;
  print: () => void;
  exportPdf: () => void;
  newDoc: () => void;
  open: () => void;
  /** Dosya menüsünü belirli sayfada açar. */
  backstage: (page: "info" | "export" | "options" | "print" | "new" | "open") => void;
  /** İletişim kutusu açar (bağlamsal sekmelerdeki komutlar için). */
  dialog: (kind: "compress") => void;
}

/** Arama eş anlamlıları (ör. geçişin başka adlarla da bulunması). */
const SYNONYMS: Record<string, string> = { morph: "dönüşüm geçişi şekil dönüştürme", zoom: "yakınlaştırma geçişi" };

const entry = (id: string, title: string, tabLabel: string, group: string, run: () => void, extra: Partial<CommandEntry> = {}): CommandEntry => ({ id, title, tabId: "", tabLabel, group, nth: 0, run, ...extra });

export function extraCommands(a: Actions, h: FileCommandHandlers, previewTransition: () => void): CommandEntry[] {
  const file: CommandEntry[] = [
    entry("file:save", "Kaydet", "Dosya", "", h.save, { shortcut: "Ctrl+S" }),
    entry("file:saveas", "Farklı Kaydet", "Dosya", "", h.saveAs, { shortcut: "F12" }),
    entry("file:print", "Yazdır", "Dosya", "", h.print, { shortcut: "Ctrl+P" }),
    entry("file:pdf", "PDF olarak dışarı aktar", "Dosya", "Dışarı Aktar", h.exportPdf, { desc: "PDF belgesi oluştur" }),
    entry("file:export", "Dışarı Aktar", "Dosya", "", () => h.backstage("export"), { desc: "Resim, ana hat, PDF" }),
    entry("file:new", "Yeni Sunu", "Dosya", "", h.newDoc, { shortcut: "Ctrl+N", desc: "Yeni boş sunu" }),
    entry("file:open", "Aç", "Dosya", "", h.open, { shortcut: "Ctrl+O" }),
    entry("file:info", "Bilgi", "Dosya", "", () => h.backstage("info"), { desc: "Sunu özellikleri" }),
    entry("file:options", "Seçenekler", "Dosya", "", () => h.backstage("options"), { desc: "Office teması, yazım denetimi" }),
    entry("pic:compress", "Resimleri Sıkıştır", "Resim Biçimi", "Ayarla", () => h.dialog("compress"), { desc: "Resimlerin çözünürlüğünü düşür, kırpılan alanları sil" }),
  ];
  const trans = TRANSITIONS.filter((t) => t.type !== "none").map((t) =>
    entry(`tr:${t.type}`, t.label, "Geçişler", "Bu Slayda Geçiş", () => {
      const d = transitionDef(t.type);
      a.setTransition({ type: d.type, option: d.defaultOption, duration: d.type === "cut" ? 0.1 : 1 });
      window.setTimeout(previewTransition, 50);
    }, { desc: `geçiş efekti ${SYNONYMS[t.type] ?? ""}`.trim() }),
  );
  const anims = ANIMATIONS.map((d) =>
    entry(`an:${d.id}`, d.label, "Animasyonlar", `Animasyon › ${CLASS_LABELS[d.cls]}`, () => {
      if (!a.s.sel.length) {
        a.message("Animasyon uygulamak için önce bir nesne seçin.", "info");
        return;
      }
      a.applyAnimation(d.id, d.cls, false, d.defaultOption);
    }, { desc: "animasyon efekti" }),
  );
  return [...file, ...trans, ...anims];
}
