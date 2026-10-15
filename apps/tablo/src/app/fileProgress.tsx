/**
 * Süren uzun dosya işleminin (açma, kaydetme) ilerlemesi: durum çubuğunda "Açılıyor… %40 [İptal]"
 * olarak gösterilir (packages/ui Progress). Aynı anda tek işlem gösterilir; sonra başlayan öne geçer.
 */
import { useSyncExternalStore } from "react";
import { Progress } from "@officehere/ui";

interface FileTask {
  id: number;
  label: string;
  value: number | null;
  detail?: string;
  cancel?: () => void;
}

let current: FileTask | null = null;
let seq = 0;
const listeners = new Set<() => void>();

function set(t: FileTask | null): void {
  current = t;
  for (const l of listeners) l();
}

function subscribe(fn: () => void): () => void {
  listeners.add(fn);
  return () => listeners.delete(fn);
}

export interface FileTaskHandle {
  /** İlerleme (0..1) ve adım adı; yüzde ve adım değişmedikçe yeniden çizilmez. */
  update(p: { ratio: number; step?: string }): void;
  end(): void;
}

/** Kısa işlemde gösterge yanıp sönmesin: bu süreden uzun sürerse görünür. */
const SHOW_AFTER_MS = 250;

/** Durum çubuğunda ilerleme gösterir; `cancel` verilirse İptal düğmesi çıkar. */
export function startFileTask(label: string, cancel?: () => void, showAfterMs = SHOW_AFTER_MS): FileTaskHandle {
  const id = ++seq;
  let task: FileTask = { id, label, value: null, cancel };
  let ended = false;
  let timer: ReturnType<typeof setTimeout> | undefined = setTimeout(() => {
    timer = undefined;
    if (!ended) set(task);
  }, showAfterMs);
  return {
    update({ ratio, step }) {
      const value = Math.floor(Math.min(1, Math.max(0, ratio)) * 100) / 100;
      if (ended || (task.value === value && task.detail === step)) return;
      task = { ...task, value, detail: step };
      if (current?.id === id) set(task);
    },
    end() {
      ended = true;
      if (timer !== undefined) clearTimeout(timer);
      if (current?.id === id) set(null);
    },
  };
}

/** Durum çubuğu öğesi (işlem yoksa boş). */
export function FileTaskProgress() {
  const t = useSyncExternalStore(subscribe, () => current);
  if (!t) return null;
  return <Progress label={t.label} value={t.value} detail={t.detail} onCancel={t.cancel} />;
}
