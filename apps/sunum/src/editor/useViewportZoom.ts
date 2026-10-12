/** Slayt alanının görünür boyutu, "Pencereye Sığdır" yakınlaştırması ve Ctrl+tekerlek. */
import { useEffect, useLayoutEffect, useMemo, useState, type RefObject } from "react";
import type { AppState } from "../state/store";
import { MARGIN } from "./canvasUtils";

interface ZoomHost {
  state: AppState;
  onZoomChange: (zoom: number, fit: boolean) => void;
}

export function useViewportZoom(scrollRef: RefObject<HTMLDivElement | null>, slideW: number, slideH: number, host: () => ZoomHost): { w: number; h: number } {
  const [viewport, setViewport] = useState({ w: 800, h: 500 });
  const { state } = host();

  // Pencere boyutu ve "Pencereye Sığdır".
  useLayoutEffect(() => {
    const el = scrollRef.current;
    if (!el) return;
    const ro = new ResizeObserver(() => setViewport({ w: el.clientWidth, h: el.clientHeight }));
    ro.observe(el);
    setViewport({ w: el.clientWidth, h: el.clientHeight });
    return () => ro.disconnect();
  }, []);

  const fitZoom = useMemo(() => {
    const zz = Math.min((viewport.w - MARGIN * 2) / slideW, (viewport.h - MARGIN * 2) / slideH);
    return Math.max(10, Math.min(400, Math.floor(zz * 100)));
  }, [viewport, slideW, slideH]);

  useEffect(() => {
    if (state.fit && Math.abs(fitZoom - state.zoom) > 0.5) host().onZoomChange(fitZoom, true);
  }, [state.fit, fitZoom, state.zoom]);

  // Ctrl + tekerlek yakınlaştırma.
  useEffect(() => {
    const el = scrollRef.current;
    if (!el) return;
    const onWheel = (e: WheelEvent) => {
      if (!e.ctrlKey && !e.metaKey) return;
      e.preventDefault();
      const cur = host().state.zoom;
      const next = Math.max(10, Math.min(400, Math.round(cur * (e.deltaY < 0 ? 1.1 : 1 / 1.1))));
      host().onZoomChange(next, false);
    };
    el.addEventListener("wheel", onWheel, { passive: false });
    return () => el.removeEventListener("wheel", onWheel);
  }, []);

  return viewport;
}
