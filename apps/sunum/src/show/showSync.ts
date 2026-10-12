/**
 * Sunucu Görünümü için iki pencere eşitlemesi: sunucu penceresi (notlar,
 * süre, sonraki slayt) gösteriyi yönetir; izleyici penceresi aynı uygulamayı
 * "?izleyici=<kimlik>" ile açar ve BroadcastChannel üzerinden gelen durumu
 * (slayt, animasyon adımı, siyah/beyaz ekran, mürekkep, lazer) izler.
 */
import type { Presentation } from "../model/types";
import type { ShowRunOptions } from "./settings";
import type { ShowStroke } from "./ShowInk";

export interface SyncState {
  index: number;
  step: number;
  ended: boolean;
  screen: "none" | "black" | "white";
  ink: ShowStroke[];
  laser: [number, number] | null;
  laserColor?: string;
  seq: number;
}

export type RemoteCmd = "next" | "prev" | "end";

export type SyncMsg =
  | { type: "hello" }
  | { type: "init"; pres: Presentation; opts: ShowRunOptions; state: SyncState | null }
  | { type: "state"; state: SyncState }
  | { type: "cmd"; cmd: RemoteCmd }
  | { type: "end" }
  | { type: "bye" };

export const channelName = (id: string) => `officehere-sunum-show-${id}`;

export function newChannelId(): string {
  return Math.random().toString(36).slice(2, 10);
}

/** İzleyici penceresini açar (açılır pencere engellenirse null). */
export function openAudienceWindow(id: string): Window | null {
  try {
    const url = new URL(window.location.href);
    url.hash = "";
    url.search = "";
    url.searchParams.set("izleyici", id);
    return window.open(url.toString(), `officehere-sunum-izleyici-${id}`, "popup=yes,width=1100,height=660");
  } catch {
    return null;
  }
}

/** Sunucu tarafı bağlantısı. */
export class PresenterLink {
  private ch: BroadcastChannel | null;
  private last: SyncState | null = null;
  private listeners = new Set<(c: RemoteCmd) => void>();
  private seq = 0;
  /** İzleyici penceresi bağlandı mı? */
  connected = false;
  onConnect: (() => void) | null = null;

  constructor(
    readonly id: string,
    private readonly init: () => { pres: Presentation; opts: ShowRunOptions },
  ) {
    this.ch = typeof BroadcastChannel !== "undefined" ? new BroadcastChannel(channelName(id)) : null;
    if (this.ch)
      this.ch.onmessage = (e: MessageEvent<SyncMsg>) => {
        const m = e.data;
        if (m?.type === "hello") {
          this.connected = true;
          const { pres, opts } = this.init();
          this.post({ type: "init", pres, opts, state: this.last });
          this.onConnect?.();
        } else if (m?.type === "cmd") this.listeners.forEach((fn) => fn(m.cmd));
        else if (m?.type === "bye") this.connected = false;
      };
  }

  private post(m: SyncMsg) {
    try {
      this.ch?.postMessage(m);
    } catch {
      /* kanal kapalı */
    }
  }

  publish(st: Omit<SyncState, "seq">) {
    this.last = { ...st, seq: ++this.seq };
    this.post({ type: "state", state: this.last });
  }

  onCommand(fn: (c: RemoteCmd) => void): () => void {
    this.listeners.add(fn);
    return () => this.listeners.delete(fn);
  }

  end() {
    this.post({ type: "end" });
    this.ch?.close();
    this.ch = null;
    this.listeners.clear();
  }
}
