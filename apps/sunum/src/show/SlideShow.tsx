/**
 * Slayt gösterisi motoru: tam ekran gösteri, Okuma Görünümü (pencerede),
 * Sunucu Görünümü (izleyici ayrı pencerede) ve Zamanlama Provası. Geçişler
 * ve animasyon adımları Web Animations API ile oynatılır; gizli slaytlar ve
 * ayarlanmış aralık/özel gösteri dışındaki slaytlar atlanır.
 */
import { useCallback, useEffect, useLayoutEffect, useMemo, useRef, useState } from "react";
import { Menu, MenuItem, MenuSeparator, useContextMenu } from "@officehere/ui";
import type { Presentation, Shape, Slide } from "../model/types";
import { findShape } from "../model/ops";
import { linkSlideIndex } from "../model/links";
import { SlideView } from "../render/SlideView";
import { animationKeyframes, buildSteps, initiallyHidden, transitionPlan, type StepEffect } from "./effects";
import { InkOverlay, type ShowStroke } from "./ShowInk";
import { ShowSlideGrid, ShowToolbar, useShowTools } from "./ShowControls";
import { showOrder, stepInOrder, type ShowRunOptions } from "./settings";
import { handleShowKey } from "./showKeys";
import { PresenterLayout } from "./PresenterLayout";
import { RehearsalBar, useRehearsal, type RehearsalResult } from "./Rehearsal";
import type { PresenterLink, RemoteCmd, SyncState } from "./showSync";
import { WindowBar } from "./WindowBar";
import { playTransitionSound } from "./sounds";

export type ShowMode = "full" | "window" | "presenter";

export interface SlideShowProps {
  pres: Presentation;
  startIndex: number;
  mode: ShowMode;
  /** Gösteri bittiğinde; gösteride çizilen mürekkep (slayt kimliğine göre) ve prova süreleriyle. */
  onExit: (lastIndex: number, ink?: Record<string, ShowStroke[]>, rehearsal?: RehearsalResult) => void;
  /** Gösteriyi Ayarla seçenekleri (döngü, animasyonsuz, aralık, renkler…). */
  opts?: ShowRunOptions;
  /** Pencere modunda alt çubuk için görünüm değiştirme. */
  onViewChange?: (v: "normal" | "sorter") => void;
  /** Şeritten önizleme: geçişi ya da animasyonları oynatıp kapanır. */
  preview?: "transition" | "animations";
  /** Önizlemenin çizileceği alan (kapsayıcıya göre px); verilmezse tüm kapsayıcı. */
  frame?: { x: number; y: number; w: number; h: number };
  /** Sunucu tarafı: durumu izleyici penceresine yayınlar. */
  link?: PresenterLink | null;
  /** İzleyici tarafı: sunucudan gelen durum izlenir. */
  follow?: SyncState | null;
  /** İzleyici tarafı: gezinme komutları sunucuya gönderilir. */
  onRemote?: (cmd: RemoteCmd) => void;
}

interface Layer {
  key: string;
  index: number;
  /** Sondan (tüm adımlar oynatılmış) başla. */
  fromEnd: boolean;
}

let layerSeq = 0;

export function SlideShow({ pres, startIndex, mode, onExit, opts = {}, onViewChange, preview, frame, link, follow, onRemote }: SlideShowProps) {
  const rootRef = useRef<HTMLDivElement>(null);
  const areaRef = useRef<HTMLDivElement>(null);
  const layerEls = useRef(new Map<string, HTMLDivElement>());
  const running = useRef<globalThis.Animation[]>([]);
  const noAnimations = !!opts.noAnimations;
  const order = useMemo(() => opts.order ?? showOrder(pres, { kind: "all" }), [opts.order, pres]);
  const [layers, setLayers] = useState<Layer[]>(() => {
    const idx = Math.max(0, Math.min(pres.slides.length - 1, startIndex));
    // Geçiş önizlemesi önceki slayttan başlar.
    if (preview === "transition") return [{ key: `l${++layerSeq}`, index: Math.max(0, idx - 1), fromEnd: true }];
    return [{ key: `l${++layerSeq}`, index: idx, fromEnd: false }];
  });
  const previewStage = useRef(0);
  const [step, setStep] = useState(0);
  const [ended, setEnded] = useState(false);
  const [screen, setScreen] = useState<"none" | "black" | "white">("none");
  const [area, setArea] = useState({ w: window.innerWidth, h: window.innerHeight });
  const [numBuf, setNumBuf] = useState("");
  const [cursorHidden, setCursorHidden] = useState(false);
  const [laser, setLaser] = useState<[number, number] | null>(null);
  const [audience, setAudience] = useState<"none" | "waiting" | "connected">(link ? (link.connected ? "connected" : "waiting") : "none");
  const ctx = useContextMenu();
  const wasFullscreen = useRef(false);
  const autoTimer = useRef<number | undefined>(undefined);

  const current = layers[layers.length - 1];
  const index = current.index;
  const slide: Slide | undefined = pres.slides[index];
  const steps = useMemo(() => (slide && !noAnimations ? buildSteps(slide) : { auto: [], clicks: [] }), [slide, noAnimations]);
  const tools = useShowTools(slide?.id, opts.penColor);
  const rehearsal = useRehearsal(!!opts.rehearse, slide?.id);
  const [grid, setGrid] = useState(false);
  const exitRef = useRef((i: number) => onExit(i, tools.allInk(), opts.rehearse ? rehearsal.result() : undefined));
  exitRef.current = (i: number) => onExit(i, tools.allInk(), opts.rehearse ? rehearsal.result() : undefined);
  const exit = (i: number) => (onRemote ? onRemote("end") : exitRef.current(i));
  const stateRef = useRef({ index, step, ended, steps, layers });
  stateRef.current = { index, step, ended, steps, layers };
  const followRef = useRef(follow);
  followRef.current = follow;

  // Alan boyutu.
  useLayoutEffect(() => {
    const el = areaRef.current;
    if (!el) return;
    const ro = new ResizeObserver(() => setArea({ w: el.clientWidth, h: el.clientHeight }));
    ro.observe(el);
    setArea({ w: el.clientWidth, h: el.clientHeight });
    return () => ro.disconnect();
  }, [mode]);

  const scale = Math.min(area.w / pres.width, area.h / pres.height) || 0.1;

  // Tam ekran.
  useEffect(() => {
    if (mode === "window") return;
    const el = rootRef.current;
    const onFs = () => {
      if (document.fullscreenElement) wasFullscreen.current = true;
      else if (wasFullscreen.current) exitRef.current(stateRef.current.index);
    };
    document.addEventListener("fullscreenchange", onFs);
    if (document.fullscreenElement) wasFullscreen.current = true;
    else if (navigator.userActivation?.isActive !== false) el?.requestFullscreen?.().catch(() => undefined);
    return () => {
      document.removeEventListener("fullscreenchange", onFs);
      if (document.fullscreenElement) document.exitFullscreen().catch(() => undefined);
    };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  // Önizleme odağı almaz (şerit kısayolları çalışmaya devam etsin); ızgara kapanınca odak geri döner.
  useEffect(() => {
    if (!preview && !grid) rootRef.current?.focus();
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [preview, grid]);

  // Odak gösterinin dışına düşerse (ör. kapanan menü/ızgara) tuşlar yine gösteriye gider.
  const keyRef = useRef<(e: React.KeyboardEvent) => void>(() => undefined);
  useEffect(() => {
    if (preview) return;
    const onKey = (e: KeyboardEvent) => {
      const root = rootRef.current;
      const t = e.target as Element | null;
      if (e.defaultPrevented || !root || (t && (root.contains(t) || t.closest?.("[data-oh-popover], input, textarea, select")))) return;
      keyRef.current(e as unknown as React.KeyboardEvent);
    };
    window.addEventListener("keydown", onKey);
    return () => window.removeEventListener("keydown", onKey);
  }, [preview]);

  // İmleç gizleme.
  useEffect(() => {
    if (mode !== "full") return;
    let t = window.setTimeout(() => setCursorHidden(true), 3000);
    const onMove = () => {
      setCursorHidden(false);
      window.clearTimeout(t);
      t = window.setTimeout(() => setCursorHidden(true), 3000);
    };
    window.addEventListener("mousemove", onMove);
    return () => {
      window.removeEventListener("mousemove", onMove);
      window.clearTimeout(t);
    };
  }, [mode]);

  const finishRunning = () => {
    let any = false;
    for (const a of running.current) {
      if (a.playState === "running" || a.pending) {
        any = true;
        a.finish();
      }
    }
    running.current = running.current.filter((a) => a.playState !== "finished" || a.effect?.getComputedTiming().fill !== "none");
    return any;
  };

  const runEffects = useCallback(
    (layerKey: string, effects: StepEffect[], instant: boolean): Promise<void> => {
      const el = layerEls.current.get(layerKey);
      const sl = pres.slides[stateRef.current.layers.find((l) => l.key === layerKey)?.index ?? -1];
      if (!el || !sl || !effects.length) return Promise.resolve();
      const promises: Promise<unknown>[] = [];
      for (const e of effects) {
        const target = el.querySelector(`[data-shape-id="${e.anim.shapeId}"]`) as HTMLElement | null;
        if (!target) continue;
        const sh = findShape(sl.shapes, e.anim.shapeId);
        const g = sh ?? { x: 0, y: 0, w: 100, h: 100 };
        const { kf, easing } = animationKeyframes(e.anim, g, pres.width, pres.height);
        const dur = Math.max(1, e.anim.duration * 1000);
        if (e.anim.cls === "entr") target.style.visibility = "visible";
        const anim = target.animate(kf, { duration: dur, delay: e.start * 1000, fill: "forwards", easing });
        running.current.push(anim);
        if (instant) anim.finish();
        promises.push(anim.finished.catch(() => undefined));
      }
      return Promise.all(promises).then(() => undefined);
    },
    [pres],
  );

  /** Katmanı belirli adım sayısına getirir (anında). */
  const applyState = useCallback(
    (layerKey: string, sl: Slide, upto: number) => {
      const el = layerEls.current.get(layerKey);
      if (!el) return;
      el.querySelectorAll<HTMLElement>("[data-shape-id]").forEach((n) => {
        n.getAnimations().forEach((a) => a.cancel());
        n.style.visibility = "";
      });
      if (noAnimations) return;
      const st = buildSteps(sl);
      void runEffects(layerKey, st.auto, true);
      for (let i = 0; i < upto && i < st.clicks.length; i++) void runEffects(layerKey, st.clicks[i], true);
    },
    [runEffects, noAnimations],
  );

  // Yeni katman eklenince: başlangıç durumu, geçiş ve otomatik adım.
  const lastSetup = useRef<string>("");
  useLayoutEffect(() => {
    const top = layers[layers.length - 1];
    if (lastSetup.current === top.key) return;
    lastSetup.current = top.key;
    const sl = pres.slides[top.index];
    if (!sl) return;
    window.clearTimeout(autoTimer.current);
    const st = noAnimations ? { auto: [], clicks: [] as StepEffect[][] } : buildSteps(sl);
    let startStep = top.fromEnd ? st.clicks.length : 0;
    // İzleyici: sunucunun adımına getir.
    const f = followRef.current;
    if (f && f.index === top.index && f.step !== startStep) startStep = Math.min(f.step, st.clicks.length);
    if (startStep > 0) applyState(top.key, sl, startStep);
    setStep(startStep);
    const layerEl = layerEls.current.get(top.key);
    // Anlatımsız: kendiliğinden başlayan sesler çalınmaz; sunucu penceresi (izleyici bağlıyken) sessizdir.
    if (layerEl && (opts.noNarration || link)) {
      layerEl.querySelectorAll<HTMLMediaElement>("audio, video").forEach((m) => {
        if (link) m.muted = true;
        if (opts.noNarration && m.tagName === "AUDIO") {
          m.autoplay = false;
          m.pause();
        }
      });
    }
    // Geçiş sesi (izleyici penceresine bağlı sunucu penceresi sessizdir).
    if (!top.fromEnd && !link && sl.transition?.sound && (layers.length > 1 || preview !== "animations")) playTransitionSound(sl.transition.sound);
    const plan = layers.length > 1 ? transitionPlan(sl.transition) : null;
    const inEl = layerEl;
    const outLayer = layers.length > 1 ? layers[layers.length - 2] : undefined;
    const outEl = outLayer ? layerEls.current.get(outLayer.key) : undefined;
    const dur = Math.max(0.05, sl.transition?.duration ?? 1) * 1000;
    const afterTransition = () => {
      setLayers((ls) => (ls.length > 1 ? [ls[ls.length - 1]] : ls));
      if (preview === "transition") {
        previewStage.current += 1;
        if (previewStage.current === 1) window.setTimeout(() => goTo(Math.max(0, Math.min(pres.slides.length - 1, startIndex))), 150);
        else window.setTimeout(() => onExit(top.index), 400);
        return;
      }
      if (preview === "animations") {
        void (async () => {
          await runEffects(top.key, st.auto, false);
          for (const step of st.clicks) await runEffects(top.key, step, false);
          window.setTimeout(() => onExit(top.index), 500);
        })();
        return;
      }
      if (startStep === 0) {
        void runEffects(top.key, st.auto, false).then(() => {
          const adv = sl.transition?.advanceAfter;
          const auto = adv !== undefined && !opts.manual && !opts.rehearse && !followRef.current;
          if (auto) autoTimer.current = window.setTimeout(() => advance(true), adv * 1000);
        });
      }
    };
    if (plan && inEl && outEl) {
      if (plan.outgoingOnTop) {
        outEl.style.zIndex = "2";
        inEl.style.zIndex = "1";
      }
      const anims: globalThis.Animation[] = [];
      if (plan.incoming) anims.push(inEl.animate(plan.incoming, { duration: dur, easing: plan.easing ?? "linear", fill: "both" }));
      if (plan.outgoing) anims.push(outEl.animate(plan.outgoing, { duration: dur, easing: plan.easing ?? "linear", fill: "both" }));
      Promise.all(anims.map((a) => a.finished.catch(() => undefined))).then(() => {
        anims.forEach((a) => a.cancel());
        inEl.style.zIndex = "";
        afterTransition();
      });
    } else afterTransition();
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [layers]);

  const goTo = useCallback((i: number, fromEnd = false, withTransition = true) => {
    window.clearTimeout(autoTimer.current);
    finishRunning();
    setEnded(false);
    setLayers((ls) => {
      const top = ls[ls.length - 1];
      const nl: Layer = { key: `l${++layerSeq}`, index: i, fromEnd };
      return withTransition ? [top, nl] : [nl];
    });
  }, []);

  const advanceLocal = useCallback(
    (auto = false) => {
      const st = stateRef.current;
      if (st.ended) {
        if (!auto) exitRef.current(st.index);
        return;
      }
      if (!auto && finishRunning()) return;
      const top = st.layers[st.layers.length - 1];
      if (st.step < st.steps.clicks.length) {
        const k = st.step;
        setStep(k + 1);
        void runEffects(top.key, st.steps.clicks[k], false);
        return;
      }
      const ni = stepInOrder(order, st.index, 1);
      if (ni === null) {
        if (opts.loop && order.length) {
          goTo(order[0]);
          return;
        }
        setEnded(true);
        return;
      }
      goTo(ni);
    },
    [order, goTo, runEffects, opts.loop],
  );

  const backLocal = useCallback(() => {
    const st = stateRef.current;
    window.clearTimeout(autoTimer.current);
    if (st.ended) {
      setEnded(false);
      return;
    }
    finishRunning();
    if (st.step > 0) {
      const top = st.layers[st.layers.length - 1];
      const sl = pres.slides[top.index];
      applyState(top.key, sl, st.step - 1);
      setStep(st.step - 1);
      return;
    }
    const pi = stepInOrder(order, st.index, -1);
    if (pi === null) return;
    goTo(pi, true, false);
  }, [pres, order, goTo, applyState]);

  // İzleyici penceresinde gezinme sunucuya iletilir.
  const advance = useCallback(
    (auto = false) => {
      if (!onRemote) advanceLocal(auto);
      else if (!auto) onRemote("next");
    },
    [onRemote, advanceLocal],
  );
  const back = useCallback(() => {
    if (onRemote) onRemote("prev");
    else backLocal();
  }, [onRemote, backLocal]);

  // Sunucu: izleyiciden gelen komutlar.
  useEffect(() => {
    if (!link) return;
    link.onConnect = () => setAudience("connected");
    const off = link.onCommand((c) => {
      if (c === "next") advanceLocal();
      else if (c === "prev") backLocal();
      else exitRef.current(stateRef.current.index);
    });
    return () => {
      off();
      link.onConnect = null;
    };
  }, [link, advanceLocal, backLocal]);

  // Sunucu: durum değişince yayınla.
  useEffect(() => {
    link?.publish({ index, step, ended, screen, ink: tools.visible ? tools.strokes : [], laser: tools.tool === "laser" ? laser : null, laserColor: opts.laserColor });
  }, [link, index, step, ended, screen, tools.strokes, tools.visible, tools.tool, laser, opts.laserColor]);

  // İzleyici: sunucunun durumuna uy.
  useEffect(() => {
    if (!follow) return;
    const st = stateRef.current;
    setScreen(follow.screen);
    if (follow.index !== st.index) {
      const forward = stepInOrder(order, st.index, 1) === follow.index;
      if (follow.step === 0) goTo(follow.index, false, forward);
      else goTo(follow.index, true, false);
      if (follow.ended) setEnded(true);
      return;
    }
    if (follow.ended !== st.ended) setEnded(follow.ended);
    if (follow.step === st.step + 1) advanceLocal();
    else if (follow.step !== st.step) {
      const top = st.layers[st.layers.length - 1];
      const sl = pres.slides[top.index];
      if (sl) applyState(top.key, sl, follow.step);
      setStep(follow.step);
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [follow?.seq]);

  const gotoIndex = (i: number) => {
    if (onRemote) return;
    if (i >= 0 && i < pres.slides.length) goTo(i, false, false);
  };

  const onKeyDown = (e: React.KeyboardEvent) => {
    e.stopPropagation();
    if (!preview && !follow && tools.handleKey(e)) return;
    handleShowKey(e, {
      advance: () => advance(),
      back,
      first: () => order.length && gotoIndex(order[0]),
      last: () => order.length && gotoIndex(order[order.length - 1]),
      gotoNumber: (n) => gotoIndex(n - 1),
      screen,
      setScreen,
      exit: () => exit(index),
      grid,
      setGrid,
      numBuf,
      setNumBuf,
      kiosk: !!opts.kiosk,
      preview: !!preview || !!follow,
    });
  };

  keyRef.current = onKeyDown;

  /** Gösteride köprüye tıklanınca: dış adres yeni sekmede, sunu içi hedef slayta gidilir. */
  const followShowLink = (lnk: string) => {
    const i = linkSlideIndex(lnk, pres, index);
    if (i === null) {
      window.open(lnk, "_blank", "noopener,noreferrer");
      return;
    }
    if (i >= 0 && i < pres.slides.length) goTo(i);
  };

  const onClick = (e: React.MouseEvent) => {
    if (e.button !== 0) return;
    if ((e.target as Element).closest("button, .show-bar, .presenter-side, .show-toolbar, .show-grid, .rehearse-bar, video, audio, [data-media]")) return;
    if (tools.tool === "pen" || tools.tool === "highlighter" || tools.tool === "eraser") return;
    const linkEl = (e.target as Element).closest("[data-link]") as HTMLElement | null;
    if (linkEl?.dataset.link && screen === "none" && !onRemote) {
      followShowLink(linkEl.dataset.link);
      return;
    }
    if (screen !== "none") {
      setScreen("none");
      return;
    }
    // Bilgi noktasında tıklama ilerletmez.
    if (opts.kiosk && !ended) return;
    advance();
  };

  const openMenu = (e: React.MouseEvent) => {
    if (opts.kiosk) return;
    ctx.open(e, mainMenu());
  };

  const mainMenu = () => (
    <Menu>
      <MenuItem label="Sonraki" onClick={() => advance()} />
      <MenuItem label="Önceki" onClick={() => back()} />
      <MenuSeparator />
      <MenuItem
        label="Slayta Git"
        disabled={!!onRemote}
        submenu={
          <Menu>
            {pres.slides.map((s, i) => (
              <MenuItem key={s.id} label={`${i + 1}. ${slideTitle(s) || "Slayt"}${s.hidden ? " (gizli)" : ""}`} checked={i === index} onClick={() => gotoIndex(i)} />
            ))}
          </Menu>
        }
      />
      {!preview && !follow && <MenuItem label="Tüm Slaytları Gör" shortcut="G" onClick={() => setGrid(true)} />}
      {!preview && !follow && tools.pointerMenu()}
      <MenuItem
        label="Ekran"
        submenu={
          <Menu>
            <MenuItem label="Siyah Ekran" shortcut="B" onClick={() => setScreen("black")} />
            <MenuItem label="Beyaz Ekran" shortcut="W" onClick={() => setScreen("white")} />
          </Menu>
        }
      />
      <MenuSeparator />
      <MenuItem label="Gösteriyi Bitir" shortcut="Esc" onClick={() => exit(index)} />
    </Menu>
  );

  const hiddenSets = useMemo(() => new Map(layers.map((l) => [l.key, noAnimations ? new Set<string>() : initiallyHidden(pres.slides[l.index] ?? ({ animations: [] } as unknown as Slide))])), [layers, pres, noAnimations]);

  const remoteInk = follow ? follow.ink : null;
  const stage = (
    <div className="show-stage" style={{ width: pres.width * scale, height: pres.height * scale, background: layers.length > 1 && transitionPlan(slide?.transition)?.thruBlack ? "#000" : undefined }}>
      <div className="show-scale" style={{ width: pres.width, height: pres.height, transform: `scale(${scale})` }}>
        {layers.map((l) => {
          const sl = pres.slides[l.index];
          if (!sl) return null;
          const hidden = hiddenSets.get(l.key) ?? new Set<string>();
          return (
            <div
              key={l.key}
              className="show-layer"
              ref={(el) => {
                if (el) layerEls.current.set(l.key, el);
                else layerEls.current.delete(l.key);
              }}
            >
              <SlideView pres={pres} slide={sl} slideNumber={l.index + 1} mode="show" shapeClass={(s: Shape) => (hidden.has(s.id) ? "anim-init-hidden" : undefined)} />
            </div>
          );
        })}
        {!preview && !follow && (
          <InkOverlay width={pres.width} height={pres.height} tool={tools.tool} color={tools.color} strokes={tools.strokes} visible={tools.visible} onAdd={tools.add} onErase={tools.erase} onLaser={link ? setLaser : undefined} laserColor={opts.laserColor} />
        )}
        {remoteInk && <InkOverlay width={pres.width} height={pres.height} tool="pointer" color="#000" strokes={remoteInk} visible onAdd={() => undefined} onErase={() => undefined} remoteLaser={follow?.laser ?? null} laserColor={follow?.laserColor} />}
      </div>
      {ended && (
        <div className="show-end" onClick={() => exit(index)}>
          Slayt gösterisinin sonu. Çıkmak için tıklayın.
        </div>
      )}
      {screen !== "none" && <div className={`show-screen is-${screen}`} />}
    </div>
  );

  const gridEl = grid && (
    <ShowSlideGrid
      pres={pres}
      current={index}
      onPick={(i) => {
        setGrid(false);
        gotoIndex(i);
      }}
      onClose={() => setGrid(false)}
    />
  );

  if (mode === "presenter") {
    const ni = stepInOrder(order, index, 1);
    const nextSlide = step < steps.clicks.length ? slide : ni !== null ? pres.slides[ni] : undefined;
    return (
      <PresenterLayout
        pres={pres}
        index={index}
        slide={slide}
        nextSlide={nextSlide}
        nextNumber={(nextSlide === slide ? index : (ni ?? index)) + 1}
        stage={stage}
        areaRef={areaRef}
        rootRef={rootRef}
        onKeyDown={onKeyDown}
        onClick={onClick}
        onMenu={openMenu}
        onOpenMenu={(e, m) => ctx.open(e, m)}
        onExit={() => exit(index)}
        onBack={back}
        onNext={() => advance()}
        onGrid={() => setGrid(true)}
        onBlack={() => setScreen((s) => (s === "black" ? "none" : "black"))}
        tools={tools}
        audience={audience}
        overlay={
          <>
            {gridEl}
            {ctx.element}
          </>
        }
      />
    );
  }

  return (
    <div
      ref={rootRef}
      className={`show-root${mode === "window" ? " is-window" : ""}${mode === "window" && !preview && !onViewChange ? " is-browsed" : ""}${cursorHidden ? " no-cursor" : ""}${preview ? " is-preview" : ""}`}
      style={frame ? { inset: "auto", left: frame.x, top: frame.y, width: frame.w, height: frame.h, background: "transparent" } : undefined}
      tabIndex={preview ? -1 : 0}
      onKeyDown={onKeyDown}
      onClick={onClick}
      onContextMenu={(e) => {
        e.preventDefault();
        openMenu(e);
      }}
    >
      <div className="show-area" ref={areaRef}>
        {stage}
      </div>
      {opts.rehearse && <RehearsalBar r={rehearsal} onNext={() => advance()} />}
      {mode === "full" && !grid && !opts.kiosk && !follow && <ShowToolbar tools={tools} onPrev={back} onNext={() => advance()} onOpenMenu={(e, m) => ctx.open(e, m)} onAllSlides={() => setGrid(true)} moreMenu={mainMenu} />}
      {gridEl}
      {mode === "window" && !preview && <WindowBar index={index} count={pres.slides.length} onBack={back} onNext={() => advance()} onMenu={openMenu} onViewChange={onViewChange} onExit={() => exit(index)} />}
      {ctx.element}
    </div>
  );
}

export function slideTitle(s: Slide): string {
  const t = s.shapes.find((x) => x.kind === "sp" && (x.ph?.type === "title" || x.ph?.type === "ctrTitle"));
  if (!t || t.kind !== "sp" || !t.text) return "";
  return t.text.paragraphs.map((p) => p.runs.map((r) => r.text).join("")).join(" ").trim();
}
