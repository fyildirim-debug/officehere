/** Uygulamanın başlangıç durumu, oturum deposu ve önizleme alanı yardımcıları. */
import { downloadBlob } from "@officehere/core";
import { blankPresentation } from "../model/presentation";
import { writePresentationFormat } from "../io/formats";
import { Actions } from "../state/actions";
import { documentPres } from "../state/masterActions";
import { Store, type AppState } from "../state/store";
import type { ShowMode } from "../show/SlideShow";
import type { ShowRunOptions } from "../show/settings";
import type { PresenterLink } from "../show/showSync";
import type { RecoveryHandlers } from "./ErrorBoundary";
import { save } from "./fileOps";

export function initialState(): AppState {
  const pres = blankPresentation();
  return {
    pres,
    slideIdx: 0,
    slideSel: [pres.slides[0].id],
    sel: [],
    editing: null,
    fileName: "Sunu1",
    fileRef: null,
    dirty: false,
    view: "normal",
    zoom: 60,
    fit: true,
    showNotes: true,
    notesHeight: 64,
    thumbWidth: 220,
    showRuler: false,
    showGrid: false,
    showGuides: false,
    pane: null,
    insert: null,
    focus: "canvas",
    cellSel: null,
    painter: null,
    spellcheck: true,
    cropId: null,
    ink: null,
    textTick: 0,
    ribbonTab: "home",
    collapsedSections: [],
    masterBase: null,
    message: null,
    colorMode: "color",
    proofLang: "tr-TR",
    pointEdit: null,
  };
}

/**
 * Pencere başına tek depo: kök bileşen bir hata sonrası yeniden kurulsa da
 * sunu kaybolmaz (aynı depo yeniden kullanılır).
 */
let session: { store: Store; actions: Actions } | null = null;

export function appSession(): { store: Store; actions: Actions } {
  if (!session) {
    const store = new Store(initialState());
    session = { store, actions: new Actions(store) };
  }
  return session;
}

/** Sunuyu (kayıt yolundan bağımsız) .pptx olarak indirir; olmazsa JSON yedeği. */
export async function downloadRecovery(a: Actions): Promise<void> {
  const name = a.s.fileName || "Sunu";
  try {
    const bytes = await writePresentationFormat(documentPres(a), "pptx");
    downloadBlob(new Blob([bytes as BlobPart], { type: "application/vnd.openxmlformats-officedocument.presentationml.presentation" }), `${name} (Kurtarılan).pptx`);
  } catch {
    downloadBlob(new Blob([JSON.stringify(documentPres(a))], { type: "application/json" }), `${name} (Kurtarılan).json`);
  }
}

export function recoveryHandlers(): RecoveryHandlers {
  const { actions, store } = appSession();
  return {
    save: () => void save(actions),
    download: () => void downloadRecovery(actions),
    undo: () => store.undo(),
  };
}

export interface ShowState {
  mode: ShowMode;
  start: number;
  preview?: "transition" | "animations";
  frame?: { x: number; y: number; w: number; h: number };
  opts?: ShowRunOptions;
  /** Sunucu Görünümü: izleyici penceresi bağlantısı. */
  link?: PresenterLink;
  /** Önizleme kimliği: yeni önizleme öncekinin yerine yeniden başlar. */
  seq?: number;
}

let previewSeq = 0;
/** Önizleme durumları için artan kimlik. */
export const nextPreviewSeq = () => ++previewSeq;

/** Önizleme için slayt alanının .app-center içindeki konumu. */
export function stageFrame(): ShowState["frame"] {
  const stage = document.querySelector(".app-center .stage");
  const center = document.querySelector(".app-center");
  if (!stage || !center) return undefined;
  const a = stage.getBoundingClientRect();
  const c = center.getBoundingClientRect();
  const x = Math.max(a.left, c.left);
  const y = Math.max(a.top, c.top);
  const r = Math.min(a.right, c.right);
  const b = Math.min(a.bottom, c.bottom);
  if (r - x < 40 || b - y < 40 || a.width > c.width || a.height > c.height) return undefined;
  return { x: x - c.left, y: y - c.top, w: r - x, h: b - y };
}
