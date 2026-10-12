/**
 * Geçiş ve animasyon kataloğu. Her öğe: Türkçe ad, seçenekler, gösteride
 * oynatma için anahtar kareler ve PPTX yazma/okuma eşlemesi (ECMA-376
 * p:transition öğeleri, p:animEffect filtreleri / p:anim özellikleri).
 */
import type { AnimClass, TransitionType } from "../model/types";

export interface Option {
  id: string;
  label: string;
}

export interface TransitionDef {
  type: TransitionType;
  label: string;
  /** Galeri kategorisi. */
  group?: "Hafif" | "Heyecan Verici" | "Dinamik İçerik";
  options?: Option[];
  defaultOption?: string;
}

/** Sil/Aç/Kapla yön sırası: Sağdan, Üstten, Soldan, Alttan. */
const DIR4: Option[] = [
  { id: "l", label: "Sağdan" },
  { id: "d", label: "Üstten" },
  { id: "r", label: "Soldan" },
  { id: "u", label: "Alttan" },
];

/** İt yön sırası: Alttan, Soldan, Sağdan, Üstten. */
const PUSH_DIRS: Option[] = [
  { id: "u", label: "Alttan" },
  { id: "r", label: "Soldan" },
  { id: "l", label: "Sağdan" },
  { id: "d", label: "Üstten" },
];

export const TRANSITIONS: TransitionDef[] = [
  { type: "none", label: "Yok" },
  { type: "fade", label: "Soluklaştır", options: [{ id: "smooth", label: "Düzgün" }, { id: "black", label: "Siyahtan" }], defaultOption: "smooth" },
  { type: "push", label: "İt", options: PUSH_DIRS, defaultOption: "u" },
  { type: "wipe", label: "Sil", options: DIR4, defaultOption: "l" },
  {
    type: "split",
    label: "Böl",
    options: [
      { id: "vertOut", label: "Dikey Dışa" },
      { id: "vertIn", label: "Dikey İçe" },
      { id: "horzOut", label: "Yatay Dışa" },
      { id: "horzIn", label: "Yatay İçe" },
    ],
    defaultOption: "vertOut",
  },
  { type: "reveal", label: "Göster" },
  { type: "cut", label: "Kes" },
  { type: "randomBar", label: "Gelişigüzel Çubuklar", options: [{ id: "vert", label: "Dikey" }, { id: "horz", label: "Yatay" }], defaultOption: "vert" },
  {
    type: "shape",
    label: "Şekil",
    options: [
      { id: "circle", label: "Daire" },
      { id: "diamond", label: "Elmas" },
      { id: "plus", label: "Artı" },
    ],
    defaultOption: "circle",
  },
  { type: "uncover", label: "Aç", options: DIR4, defaultOption: "l" },
  { type: "cover", label: "Kapla", options: DIR4, defaultOption: "l" },
  { type: "zoom", label: "Yakınlaştır", options: [{ id: "in", label: "İçeri" }, { id: "out", label: "Dışarı" }], defaultOption: "in" },
  { type: "dissolve", label: "Çözülme" },
  { type: "checker", label: "Dama Tahtası", options: [{ id: "horz", label: "Soldan" }, { id: "vert", label: "Üstten" }], defaultOption: "horz" },
  { type: "blinds", label: "Jaluziler", options: [{ id: "vert", label: "Dikey" }, { id: "horz", label: "Yatay" }], defaultOption: "vert" },
  { type: "wheel", label: "Saat" },
  { type: "wedge", label: "Kama" },
  { type: "newsflash", label: "Haber Flaşı" },
];

const EXCITING = new Set<TransitionType>(["zoom", "dissolve", "checker", "blinds", "wheel", "wedge", "newsflash"]);
for (const d of TRANSITIONS) d.group = EXCITING.has(d.type) ? "Heyecan Verici" : "Hafif";

export function transitionDef(t: TransitionType): TransitionDef {
  return TRANSITIONS.find((x) => x.type === t) ?? TRANSITIONS[0];
}

export interface AnimDef {
  id: string;
  cls: AnimClass;
  label: string;
  options?: Option[];
  defaultOption?: string;
  /** PPTX presetID */
  presetId: number;
  /** p:animEffect filtresi (varsa) */
  filter?: string;
}

const FLY_DIRS: Option[] = [
  { id: "b", label: "Alttan" },
  { id: "bl", label: "Sol Alttan" },
  { id: "l", label: "Soldan" },
  { id: "tl", label: "Sol Üstten" },
  { id: "t", label: "Üstten" },
  { id: "tr", label: "Sağ Üstten" },
  { id: "r", label: "Sağdan" },
  { id: "br", label: "Sağ Alttan" },
];

const WIPE_DIRS: Option[] = [
  { id: "b", label: "Alttan" },
  { id: "l", label: "Soldan" },
  { id: "r", label: "Sağdan" },
  { id: "t", label: "Üstten" },
];

const SPLIT_OPTS: Option[] = [
  { id: "horzIn", label: "Yatay İçe" },
  { id: "horzOut", label: "Yatay Dışa" },
  { id: "vertIn", label: "Dikey İçe" },
  { id: "vertOut", label: "Dikey Dışa" },
];

export const ANIMATIONS: AnimDef[] = [
  { id: "appear", cls: "entr", label: "Belir", presetId: 1 },
  { id: "fade", cls: "entr", label: "Soluklaştır", presetId: 10, filter: "fade" },
  { id: "flyIn", cls: "entr", label: "Uçarak Gir", options: FLY_DIRS, defaultOption: "b", presetId: 2 },
  { id: "floatIn", cls: "entr", label: "Yüzerek Gir", options: [{ id: "up", label: "Yukarı Yüzerek" }, { id: "down", label: "Aşağı Yüzerek" }], defaultOption: "up", presetId: 42 },
  { id: "split", cls: "entr", label: "Böl", options: SPLIT_OPTS, defaultOption: "vertOut", presetId: 16, filter: "barn" },
  { id: "wipe", cls: "entr", label: "Sil", options: WIPE_DIRS, defaultOption: "b", presetId: 22, filter: "wipe" },
  { id: "shape", cls: "entr", label: "Şekil", options: [{ id: "circle", label: "Daire" }, { id: "diamond", label: "Elmas" }, { id: "plus", label: "Artı" }], defaultOption: "circle", presetId: 6, filter: "circle" },
  { id: "wheel", cls: "entr", label: "Tekerlek", presetId: 21, filter: "wheel(1)" },
  { id: "randomBars", cls: "entr", label: "Gelişigüzel Çubuklar", options: [{ id: "horz", label: "Yatay" }, { id: "vert", label: "Dikey" }], defaultOption: "horz", presetId: 14, filter: "randombar" },
  { id: "growTurn", cls: "entr", label: "Büyüyerek Dön", presetId: 31 },
  { id: "zoom", cls: "entr", label: "Yakınlaştır", presetId: 53 },
  { id: "swivel", cls: "entr", label: "Döndür", presetId: 45 },
  { id: "bounce", cls: "entr", label: "Zıpla", presetId: 26 },
  { id: "pulse", cls: "emph", label: "Darbe", presetId: 26 },
  { id: "colorPulse", cls: "emph", label: "Renk Darbesi", presetId: 27 },
  { id: "teeter", cls: "emph", label: "Tahterevalli", presetId: 32 },
  { id: "spin", cls: "emph", label: "Döndür", presetId: 8 },
  { id: "growShrink", cls: "emph", label: "Büyüt/Küçült", presetId: 6 },
  { id: "transparency", cls: "emph", label: "Saydamlık", presetId: 9 },
  { id: "disappear", cls: "exit", label: "Kaybol", presetId: 1 },
  { id: "fadeOut", cls: "exit", label: "Soluklaştır", presetId: 10, filter: "fade" },
  { id: "flyOut", cls: "exit", label: "Uçarak Çık", options: FLY_DIRS, defaultOption: "b", presetId: 2 },
  { id: "floatOut", cls: "exit", label: "Yüzerek Çık", options: [{ id: "down", label: "Aşağı Yüzerek" }, { id: "up", label: "Yukarı Yüzerek" }], defaultOption: "down", presetId: 42 },
  { id: "splitOut", cls: "exit", label: "Böl", options: SPLIT_OPTS, defaultOption: "vertIn", presetId: 16, filter: "barn" },
  { id: "wipeOut", cls: "exit", label: "Sil", options: WIPE_DIRS, defaultOption: "b", presetId: 22, filter: "wipe" },
  { id: "zoomOut", cls: "exit", label: "Uzaklaştır", presetId: 53 },
  { id: "shrinkTurn", cls: "exit", label: "Küçülerek Dön", presetId: 31 },
];

export function animDef(id: string): AnimDef | undefined {
  return ANIMATIONS.find((a) => a.id === id);
}

export const TRIGGER_LABELS: Record<string, string> = {
  click: "Tıklandığında",
  with: "Öncekiyle",
  after: "Öncekinden Sonra",
};

export const CLASS_LABELS: Record<AnimClass, string> = {
  entr: "Giriş",
  emph: "Vurgu",
  exit: "Çıkış",
};

export const CLASS_COLORS: Record<AnimClass, string> = {
  entr: "#3a9d23",
  emph: "#d8a200",
  exit: "#c8102e",
};
