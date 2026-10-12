/**
 * Görünüm > Yeni Pencere: aynı çalışma kitabı ikinci bir pencerede açılır ve
 * pencereler BroadcastChannel üzerinden eşzamanlı tutulur. Kitap içeriği
 * (hücreler, biçimler, sayfa nesneleri, adlar) paylaşılır; etkin sayfa, seçim,
 * kaydırma ve yakınlaştırma her pencerenin kendisinindir (Excel'deki gibi).
 * Pencere başlığı "Kitap1:1", "Kitap1:2" biçimini alır.
 */
import type { Controller } from "../controller";
import { StyleTable } from "../model/styles";
import type { WorkbookSnapshot } from "../model/workbook";
import { newSessionId } from "./draftStore";

const CHANNEL = "officehere-tablo-windows";
export const WINDOW_PARAM = "pencere";

interface StateMsg {
  type: "state";
  doc: string;
  from: string;
  index: number;
  snapshot: WorkbookSnapshot;
  styles: unknown[];
  fileName: string;
  dirty: boolean;
  protection: unknown;
}
type Msg = StateMsg | { type: "hello"; doc: string; from: string } | { type: "bye"; doc: string; from: string; index: number } | { type: "here"; doc: string; from: string; index: number; assign?: { to: string; index: number } };

const me = newSessionId();

export interface WindowLink {
  doc: string;
  index: number;
  /** Bağlı diğer pencereler (kimlik → numara). */
  peers: Map<string, number>;
}

function stateOf(ctl: Controller, link: WindowLink): StateMsg {
  return {
    type: "state",
    doc: link.doc,
    from: me,
    index: link.index,
    snapshot: ctl.wb.snapshot(),
    styles: ctl.wb.styles.all().map((s) => ({ ...s })),
    fileName: ctl.fileName,
    dirty: ctl.dirty,
    protection: ctl.wb.protection,
  };
}

/** Uzak pencereden gelen kitap durumunu bu pencerenin görünümünü koruyarak uygular. */
export function applyState(ctl: Controller, m: StateMsg): void {
  const wb = ctl.wb;
  const activeId = wb.active?.id;
  const styles = new StyleTable();
  for (const s of (m.styles as Parameters<StyleTable["intern"]>[0][]).slice(1)) styles.intern(s);
  wb.styles = styles;
  if (ctl.edit) ctl.edit = null;
  wb.restore(m.snapshot);
  wb.protection = (m.protection as typeof wb.protection) ?? null;
  // Etkin sayfa ve seçim bu pencerenin kendisinin
  const idx = wb.sheets.findIndex((s) => s.id === activeId);
  wb.activeIndex = idx >= 0 ? idx : Math.min(wb.activeIndex, wb.sheets.length - 1);
  wb.sheets.forEach((s) => ctl.ops.syncHiddenRows(s));
  ctl.fileName = m.fileName;
  ctl.dirty = m.dirty;
  ctl.pristine = false;
  ctl.dataVersion++;
  ctl.emit();
}

/** Pencere eşitlemesini kurar; bağlantı bilgisi ctl.windowLink'te tutulur. */
export function installWindowSync(ctl: Controller): () => void {
  if (typeof BroadcastChannel === "undefined") return () => undefined;
  let ch: BroadcastChannel;
  try {
    ch = new BroadcastChannel(CHANNEL);
  } catch {
    return () => undefined;
  }
  let applying = false;
  let lastVersion = ctl.dataVersion;
  let lastDirty = ctl.dirty;
  let lastName = ctl.fileName;
  let timer: ReturnType<typeof setTimeout> | null = null;
  const post = (m: Msg) => {
    try {
      ch.postMessage(m);
    } catch {
      /* gönderilemeyen ileti */
    }
  };
  const broadcast = () => {
    timer = null;
    const link = ctl.windowLink;
    if (!link || !link.peers.size) return;
    post(stateOf(ctl, link));
  };
  ch.onmessage = (e: MessageEvent<Msg>) => {
    const m = e.data;
    const link = ctl.windowLink;
    if (!m || m.from === me || !link || m.doc !== link.doc) return;
    if (m.type === "hello") {
      // Yeni pencere: numarasını ve kitabın durumunu bildir
      const index = Math.max(link.index, ...link.peers.values()) + 1;
      link.peers.set(m.from, index);
      post({ type: "here", doc: link.doc, from: me, index: link.index, assign: { to: m.from, index } });
      post(stateOf(ctl, link));
      ctl.emit();
    } else if (m.type === "here") {
      link.peers.set(m.from, m.index);
      if (m.assign?.to === me) link.index = m.assign.index;
      ctl.emit();
    } else if (m.type === "bye") {
      link.peers.delete(m.from);
      ctl.emit();
    } else if (m.type === "state") {
      if (!link.peers.has(m.from)) link.peers.set(m.from, m.index);
      applying = true;
      try {
        applyState(ctl, m);
      } finally {
        applying = false;
        lastVersion = ctl.dataVersion;
        lastDirty = ctl.dirty;
        lastName = ctl.fileName;
      }
    }
  };
  // Bu pencerede değişiklik olunca (kısa bir gecikmeyle) diğerlerine gönderilir
  const unsub = ctl.subscribe(() => {
    if (applying || !ctl.windowLink?.peers.size) return;
    if (ctl.dataVersion === lastVersion && ctl.dirty === lastDirty && ctl.fileName === lastName) return;
    lastVersion = ctl.dataVersion;
    lastDirty = ctl.dirty;
    lastName = ctl.fileName;
    if (!timer) timer = setTimeout(broadcast, 120);
  });
  const bye = () => {
    const link = ctl.windowLink;
    if (link) post({ type: "bye", doc: link.doc, from: me, index: link.index });
  };
  window.addEventListener("pagehide", bye);
  // Yeni pencere olarak açıldıysa bağlan
  const params = new URLSearchParams(window.location.search);
  const doc = params.get(WINDOW_PARAM);
  if (doc) {
    params.delete(WINDOW_PARAM);
    const q = params.toString();
    window.history.replaceState(null, "", `${window.location.pathname}${q ? `?${q}` : ""}${window.location.hash}`);
    ctl.windowLink = { doc, index: 2, peers: new Map() };
    post({ type: "hello", doc, from: me });
  }
  // Not: bileşen yeniden kurulurken (geliştirme kipinde çift kurulum) "bye" gönderilmez; yalnız sayfa kapanınca
  return () => {
    unsub();
    window.removeEventListener("pagehide", bye);
    if (timer) clearTimeout(timer);
    ch.close();
  };
}

/** Yeni Pencere: aynı kitabı ikinci pencerede açar. */
export function openNewWindow(ctl: Controller): void {
  if (!ctl.windowLink) ctl.windowLink = { doc: newSessionId(), index: 1, peers: new Map() };
  const url = new URL(window.location.href);
  url.searchParams.set(WINDOW_PARAM, ctl.windowLink.doc);
  if (!window.open(url.toString(), "_blank")) ctl.alert("Yeni pencere açılamadı. Tarayıcının açılır pencere engelleyicisini kontrol edin.");
  ctl.emit();
}

/** Başlıkta gösterilecek ad: bağlı pencereler varsa "Kitap1:2". */
export function windowTitle(ctl: Controller, base: string): string {
  const l = ctl.windowLink;
  return l && l.peers.size ? `${base}:${l.index}` : base;
}
