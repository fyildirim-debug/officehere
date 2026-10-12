/**
 * Slayt gösterisi motoru: tam ekran gösteri, Okuma Görünümü (pencerede) ve
 * Sunucu Görünümü. Geçişler ve animasyon adımları Web Animations API ile
 * oynatılır; gizli slaytlar atlanır.
 */
import { useCallback, useEffect, useLayoutEffect, useMemo, useRef, useState } from "react";
import { Menu, MenuItem, MenuSeparator, useContextMenu } from "@officehere/ui";
import type { Presentation, Shape, Slide } from "../model/types";
import { findShape } from "../model/ops";
import { ScaledSlide, SlideView } from "../render/SlideView";
import { animationKeyframes, buildSteps, initiallyHidden, transitionPlan, type StepEffect } from "./effects";

export type ShowMode = "full" | "window" | "presenter";

export interface SlideShowProps {
  pres: Presentation;
  startIndex: number;
  mode: ShowMode;
  onExit: (lastIndex: number) => void;
  loop?: boolean;
  noAnimations?: boolean;
  /** Pencere modunda alt çubuk için görünüm değiştirme. */
  onViewChange?: (v: "normal" | "sorter") => void;
  /** Şeritten önizleme: geçişi ya da animasyonları oynatıp kapanır. */
  preview?: "transition" | "animations";
  /** Önizlemenin çizileceği alan (kapsayıcıya göre px); verilmezse tüm kapsayıcı. */
  frame?: { x: number; y: number; w: number; h: number };
}

interface Layer {
  key: string;
  index: number;
  /** Sondan (tüm adımlar oynatılmış) başla. */
  fromEnd: boolean;
}

let layerSeq = 0;

function nextVisible(pres: Presentation, from: number, dir: 1 | -1): number {
  let i = from + dir;
  while (i >= 0 && i < pres.slides.length && pres.slides[i].hidden) i += dir;
  return i;
}

function formatTime(ms: number): string {
  const s = Math.floor(ms / 1000);
  const h = Math.floor(s / 3600);
  const m = Math.floor((s % 3600) / 60);
  const ss = s % 60;
  return `${h > 0 ? `${h}:` : ""}${String(m).padStart(2, "0")}:${String(ss).padStart(2, "0")}`;
}

export function SlideShow({ pres, startIndex, mode, onExit, loop, noAnimations, onViewChange, preview, frame }: SlideShowProps) {
  const rootRef = useRef<HTMLDivElement>(null);
  const areaRef = useRef<HTMLDivElement>(null);
  const layerEls = useRef(new Map<string, HTMLDivElement>());
  const running = useRef<globalThis.Animation[]>([]);
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
  const [startTime] = useState(() => Date.now());
  const [now, setNow] = useState(() => Date.now());
  const [cursorHidden, setCursorHidden] = useState(false);
  const [notesZoom, setNotesZoom] = useState(1);
  const ctx = useContextMenu();
  const wasFullscreen = useRef(false);
  const autoTimer = useRef<number | undefined>(undefined);

  const current = layers[layers.length - 1];
  const index = current.index;
  const slide: Slide | undefined = pres.slides[index];
  const steps = useMemo(() => (slide && !noAnimations ? buildSteps(slide) : { auto: [], clicks: [] }), [slide, noAnimations]);
  const stateRef = useRef({ index, step, ended, steps, layers });
  stateRef.current = { index, step, ended, steps, layers };

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
      else if (wasFullscreen.current) onExit(stateRef.current.index);
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

  useEffect(() => {
    rootRef.current?.focus();
  }, []);

  // Sunucu görünümü süre sayacı.
  useEffect(() => {
    if (mode !== "presenter") return;
    const t = window.setInterval(() => setNow(Date.now()), 1000);
    return () => window.clearInterval(t);
  }, [mode]);

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
    if (top.fromEnd) {
      applyState(top.key, sl, st.clicks.length);
      setStep(st.clicks.length);
    } else setStep(0);
    const plan = layers.length > 1 ? transitionPlan(sl.transition) : null;
    const inEl = layerEls.current.get(top.key);
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
      if (!top.fromEnd) {
        void runEffects(top.key, st.auto, false).then(() => {
          const adv = sl.transition?.advanceAfter;
          if (adv !== undefined && mode !== "presenter") {
            autoTimer.current = window.setTimeout(() => advance(true), adv * 1000);
          }
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

  const goTo = useCallback(
    (i: number, fromEnd = false, withTransition = true) => {
      window.clearTimeout(autoTimer.current);
      finishRunning();
      setEnded(false);
      setLayers((ls) => {
        const top = ls[ls.length - 1];
        const nl: Layer = { key: `l${++layerSeq}`, index: i, fromEnd };
        return withTransition ? [top, nl] : [nl];
      });
    },
    [],
  );

  const advance = useCallback(
    (auto = false) => {
      const st = stateRef.current;
      if (st.ended) {
        if (!auto) onExit(st.index);
        return;
      }
      if (!auto && finishRunning()) return;
      const top = st.layers[st.layers.length - 1];
      if (st.step < st.steps.clicks.length) {
        const k = st.step;
        setStep(k + 1);
        void runEffects(top.key, st.steps.clicks[k], false).then(() => {
          // Son adımdan sonra otomatik ilerleme süresi yeniden başlamaz.
        });
        return;
      }
      const ni = nextVisible(pres, st.index, 1);
      if (ni >= pres.slides.length) {
        if (loop) {
          const first = nextVisible(pres, -1, 1);
          if (first < pres.slides.length) goTo(first);
          return;
        }
        setEnded(true);
        return;
      }
      goTo(ni);
    },
    [pres, goTo, runEffects, loop, onExit],
  );

  const back = useCallback(() => {
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
    const pi = nextVisible(pres, st.index, -1);
    if (pi < 0) return;
    goTo(pi, true, false);
  }, [pres, goTo, applyState]);

  const onKeyDown = (e: React.KeyboardEvent) => {
    const k = e.key;
    e.stopPropagation();
    if (/^[0-9]$/.test(k)) {
      setNumBuf((b) => (b + k).slice(-4));
      return;
    }
    if (k === "Enter" && numBuf) {
      e.preventDefault();
      const n = parseInt(numBuf, 10) - 1;
      setNumBuf("");
      if (n >= 0 && n < pres.slides.length) goTo(n, false, false);
      return;
    }
    setNumBuf("");
    if (screen !== "none" && k !== "Escape") {
      setScreen("none");
      if (k === "b" || k === "B" || k === "w" || k === "W" || k === "." || k === ",") return;
    }
    switch (k) {
      case "ArrowRight":
      case "ArrowDown":
      case " ":
      case "PageDown":
      case "Enter":
      case "n":
      case "N":
        e.preventDefault();
        advance();
        break;
      case "ArrowLeft":
      case "ArrowUp":
      case "PageUp":
      case "Backspace":
      case "p":
      case "P":
        e.preventDefault();
        back();
        break;
      case "Home":
        goTo(nextVisible(pres, -1, 1), false, false);
        break;
      case "End":
        goTo(nextVisible(pres, pres.slides.length, -1), false, false);
        break;
      case "b":
      case "B":
      case ".":
        setScreen("black");
        break;
      case "w":
      case "W":
      case ",":
        setScreen("white");
        break;
      case "Escape":
      case "-":
        e.preventDefault();
        if (screen !== "none") setScreen("none");
        else onExit(index);
        break;
      default:
        break;
    }
  };

  const onClick = (e: React.MouseEvent) => {
    if (e.button !== 0) return;
    if ((e.target as Element).closest("button, .show-bar, .presenter-side")) return;
    if (screen !== "none") {
      setScreen("none");
      return;
    }
    advance();
  };

  const openMenu = (e: React.MouseEvent) => {
    ctx.open(
      e,
      <Menu>
        <MenuItem label="Sonraki" onClick={() => advance()} />
        <MenuItem label="Önceki" onClick={() => back()} />
        <MenuSeparator />
        <MenuItem
          label="Slayta Git"
          submenu={
            <Menu>
              {pres.slides.map((s, i) => (
                <MenuItem key={s.id} label={`${i + 1}. ${slideTitle(s) || "Slayt"}${s.hidden ? " (gizli)" : ""}`} checked={i === index} onClick={() => goTo(i, false, false)} />
              ))}
            </Menu>
          }
        />
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
        <MenuItem label="Gösteriyi Bitir" shortcut="Esc" onClick={() => onExit(index)} />
      </Menu>,
    );
  };

  const hiddenSets = useMemo(() => new Map(layers.map((l) => [l.key, noAnimations ? new Set<string>() : initiallyHidden(pres.slides[l.index] ?? { animations: [] } as unknown as Slide)])), [layers, pres, noAnimations]);

  const stage = (
    <div
      className="show-stage"
      style={{ width: pres.width * scale, height: pres.height * scale, background: layers.length > 1 && transitionPlan(slide?.transition)?.thruBlack ? "#000" : undefined }}
    >
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
      </div>
      {ended && (
        <div className="show-end" onClick={() => onExit(index)}>
          Slayt gösterisinin sonu. Çıkmak için tıklayın.
        </div>
      )}
      {screen !== "none" && <div className={`show-screen is-${screen}`} />}
    </div>
  );

  if (mode === "presenter") {
    const ni = nextVisible(pres, index, 1);
    const nextSlide = step < steps.clicks.length ? slide : pres.slides[ni];
    return (
      <div ref={rootRef} className="show-root presenter" tabIndex={0} onKeyDown={onKeyDown} onContextMenu={(e) => e.preventDefault()}>
        <div className="presenter-top">
          <span className="presenter-timer">{formatTime(now - startTime)}</span>
          <span>{new Intl.DateTimeFormat("tr-TR", { hour: "2-digit", minute: "2-digit" }).format(now)}</span>
          <button type="button" className="presenter-btn" onClick={() => onExit(index)}>
            Gösteriyi Bitir
          </button>
        </div>
        <div className="presenter-main">
          <div className="presenter-current" ref={areaRef} onClick={onClick} onContextMenu={openMenu}>
            {stage}
          </div>
          <div className="presenter-side">
            <div className="presenter-label">Sonraki slayt</div>
            {nextSlide ? (
              <ScaledSlide pres={pres} slide={nextSlide} slideNumber={(nextSlide === slide ? index : ni) + 1} width={340} mode="show" />
            ) : (
              <div className="presenter-endbox">Slayt gösterisinin sonu</div>
            )}
            <div className="presenter-label">
              Notlar
              <span>
                <button type="button" className="presenter-btn" onClick={() => setNotesZoom((z) => Math.min(2.5, z + 0.2))} title="Metni büyüt">
                  A+
                </button>
                <button type="button" className="presenter-btn" onClick={() => setNotesZoom((z) => Math.max(0.6, z - 0.2))} title="Metni küçült">
                  A−
                </button>
              </span>
            </div>
            <div className="presenter-notes" style={{ fontSize: 16 * notesZoom }}>
              {slide?.notes || <em>Bu slayt için not yok.</em>}
            </div>
          </div>
        </div>
        <div className="presenter-bottom">
          <button type="button" className="presenter-btn" onClick={() => back()} title="Önceki">
            ◀
          </button>
          <span>
            Slayt {index + 1} / {pres.slides.length}
          </span>
          <button type="button" className="presenter-btn" onClick={() => advance()} title="Sonraki">
            ▶
          </button>
          <button type="button" className="presenter-btn" onClick={() => setScreen((s) => (s === "black" ? "none" : "black"))}>
            Siyah Ekran
          </button>
        </div>
        {ctx.element}
      </div>
    );
  }

  return (
    <div
      ref={rootRef}
      className={`show-root${mode === "window" ? " is-window" : ""}${cursorHidden ? " no-cursor" : ""}`}
      style={frame ? { inset: "auto", left: frame.x, top: frame.y, width: frame.w, height: frame.h, background: "transparent" } : undefined}
      tabIndex={0}
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
      {mode === "window" && !preview && (
        <div className="show-bar" onClick={(e) => e.stopPropagation()}>
          <span className="show-bar-title">{`Slayt ${index + 1} / ${pres.slides.length}`}</span>
          <span className="grow" />
          <button type="button" className="show-bar-btn" onClick={() => back()} title="Önceki">
            ◀
          </button>
          <button type="button" className="show-bar-btn" onClick={openMenu} title="Menü">
            ☰
          </button>
          <button type="button" className="show-bar-btn" onClick={() => advance()} title="Sonraki">
            ▶
          </button>
          {onViewChange && (
            <>
              <button type="button" className="show-bar-btn" onClick={() => onViewChange("normal")} title="Normal">
                Normal
              </button>
              <button type="button" className="show-bar-btn" onClick={() => onViewChange("sorter")} title="Slayt Sıralayıcısı">
                Sıralayıcı
              </button>
            </>
          )}
          <button type="button" className="show-bar-btn" onClick={() => onExit(index)} title="Okuma Görünümünü Kapat">
            ✕
          </button>
        </div>
      )}
      {ctx.element}
    </div>
  );
}

export function slideTitle(s: Slide): string {
  const t = s.shapes.find((x) => x.kind === "sp" && (x.ph?.type === "title" || x.ph?.type === "ctrTitle"));
  if (!t || t.kind !== "sp" || !t.text) return "";
  return t.text.paragraphs.map((p) => p.runs.map((r) => r.text).join("")).join(" ").trim();
}
