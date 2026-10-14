/**
 * Kod bölme: seyrek kullanılan ya da ilk çizimde gerekmeyen ağır bölümler
 * (dosya okuyucu/yazıcıları ve JSZip, slayt gösterisi, iletişim kutuları,
 * yazdırma düzenleri, video dışa aktarma) ayrı parçalarda tutulur. Uygulama
 * açıldıktan sonra tarayıcı boşta kalınca önceden yüklenirler; böylece ilk
 * açılış hızlanır, sonradan kaydetme/gösteri ağ bağlantısına bağlı kalmaz.
 *
 * Bileşen sarmalayıcıları: parça yüklüyse bileşen doğrudan çizilir (gecikmesiz,
 * ör. gösteri tam ekranı kullanıcı hareketine bağlı kalır); değilse Suspense ile
 * yüklenir. Seçim bağlama anında sabitlenir (yükleme bitince yeniden kurulmaz).
 * Yükleme başarısız olursa sonraki bağlamada yeniden denenir.
 */
import { createElement, lazy, Suspense, useRef, useState, type ComponentType, type ComponentProps } from "react";
import { preloadIo } from "../io/formats";
import type { SlideShow as SlideShowT } from "../show/SlideShow";
import type { AudienceView as AudienceViewT } from "../show/AudienceView";
import type { AppDialogs as AppDialogsT } from "./AppDialogs";
import type { PrintRoot as PrintRootT } from "./PrintPages";
import type { PrintPage as PrintPageT } from "./PrintPage";

function lazyComponent<P extends object>(load: () => Promise<ComponentType<P>>) {
  let loaded: ComponentType<P> | undefined;
  let pending: Promise<ComponentType<P>> | undefined;
  const get = () =>
    (pending ??= load().then(
      (c) => (loaded = c),
      (e: unknown) => {
        pending = undefined;
        throw e;
      },
    ));
  function Wrapper(props: P) {
    const [direct] = useState(() => loaded);
    const [Lazy] = useState(() => lazy(() => get().then((c) => ({ default: c }))));
    if (direct) return createElement(direct, props);
    return <Suspense fallback={null}>{createElement(Lazy as unknown as ComponentType<P>, props)}</Suspense>;
  }
  return { Component: Wrapper, preload: () => void get().catch(() => {}), isLoaded: () => !!loaded };
}

const slideShow = lazyComponent<ComponentProps<typeof SlideShowT>>(() => import("../show/SlideShow").then((m) => m.SlideShow));
const audience = lazyComponent<ComponentProps<typeof AudienceViewT>>(() => import("../show/AudienceView").then((m) => m.AudienceView));
const dialogs = lazyComponent<ComponentProps<typeof AppDialogsT>>(() => import("./AppDialogs").then((m) => m.AppDialogs));
const printRoot = lazyComponent<ComponentProps<typeof PrintRootT>>(() => import("./PrintPages").then((m) => m.PrintRoot));
const printPage = lazyComponent<ComponentProps<typeof PrintPageT>>(() => import("./PrintPage").then((m) => m.PrintPage));

/** Slayt gösterisi (okuma görünümü, önizleme, sunucu görünümü). */
export const SlideShow = slideShow.Component;
/** Sunucu Görünümü'nün izleyici penceresi. */
export const AudienceView = audience.Component;
/**
 * Uygulamanın iletişim kutuları (video dışa aktarma dahil). İlk çizimde parça
 * istenmez: bir kutu açılınca ya da boşta önceden yüklendikten sonraki ilk
 * çizimde bağlanır ve sonra hep bağlı kalır (kutuların kendi durumu korunur).
 */
export function AppDialogs(props: ComponentProps<typeof AppDialogsT>) {
  const on = useRef(false);
  if (props.dialog || props.extra || dialogs.isLoaded()) on.current = true;
  return on.current ? createElement(dialogs.Component, props) : null;
}
/** Yazdırma/PDF sayfa düzenleri. */
export const PrintRoot = printRoot.Component;
/** Dosya › Yazdır sayfası. */
export const PrintPage = printPage.Component;

type IdleWindow = Window & { requestIdleCallback?: (cb: () => void, opts?: { timeout: number }) => number };

/** Ayrı parçaları boşta önceden yükler (hata yutulur; ilk gerçek kullanımda yeniden denenir). */
export function preloadChunks(): void {
  if (typeof window === "undefined") return;
  const run = () => {
    preloadIo();
    dialogs.preload();
    slideShow.preload();
    printPage.preload();
    printRoot.preload();
  };
  const w = window as IdleWindow;
  // İlk çizim ve sayfa yüklemesi bittikten sonra, tarayıcı boşta kalınca.
  const start = () => window.setTimeout(() => (w.requestIdleCallback ? w.requestIdleCallback(run, { timeout: 4000 }) : run()), 800);
  if (document.readyState === "complete") start();
  else window.addEventListener("load", start, { once: true });
}
