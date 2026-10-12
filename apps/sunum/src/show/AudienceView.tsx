/**
 * İzleyici penceresi (Sunucu Görünümü'nün ikinci penceresi): sunucu
 * penceresinden gelen sunuyu ve gösteri durumunu izler; tıklama ve tuşlar
 * sunucuya gezinme komutu olarak iletilir.
 */
import { useEffect, useRef, useState } from "react";
import type { Presentation } from "../model/types";
import type { ShowRunOptions } from "./settings";
import { channelName, type RemoteCmd, type SyncMsg, type SyncState } from "./showSync";
import { SlideShow } from "./SlideShow";

export function AudienceView({ id }: { id: string }) {
  const [init, setInit] = useState<{ pres: Presentation; opts: ShowRunOptions } | null>(null);
  const [state, setState] = useState<SyncState | null>(null);
  const [closed, setClosed] = useState(false);
  const chRef = useRef<BroadcastChannel | null>(null);
  const gotInit = useRef(false);

  useEffect(() => {
    document.title = "Slayt Gösterisi - OfficeHere Sunum";
    if (typeof BroadcastChannel === "undefined") return;
    const ch = new BroadcastChannel(channelName(id));
    chRef.current = ch;
    ch.onmessage = (e: MessageEvent<SyncMsg>) => {
      const m = e.data;
      if (m?.type === "init") {
        gotInit.current = true;
        setInit({ pres: m.pres, opts: m.opts });
        if (m.state) setState(m.state);
      } else if (m?.type === "state") setState(m.state);
      else if (m?.type === "end") {
        setClosed(true);
        window.close();
      }
    };
    const hello = () => ch.postMessage({ type: "hello" } satisfies SyncMsg);
    hello();
    // Sunucu penceresi henüz dinlemiyorsa yeniden sor.
    const t = window.setInterval(() => {
      if (!gotInit.current) hello();
    }, 700);
    const bye = () => ch.postMessage({ type: "bye" } satisfies SyncMsg);
    window.addEventListener("pagehide", bye);
    return () => {
      window.clearInterval(t);
      window.removeEventListener("pagehide", bye);
      ch.close();
    };
  }, [id]);

  const send = (cmd: RemoteCmd) => chRef.current?.postMessage({ type: "cmd", cmd } satisfies SyncMsg);

  if (closed) return <div className="audience-wait">Slayt gösterisi bitti. Bu pencereyi kapatabilirsiniz.</div>;
  if (!init) return <div className="audience-wait">Sunucu penceresi bekleniyor…</div>;
  return <SlideShow key={init.pres.slides.length} pres={init.pres} startIndex={state?.index ?? 0} mode="full" opts={{ ...init.opts, manual: true, rehearse: false }} follow={state} onRemote={send} onExit={() => send("end")} />;
}
