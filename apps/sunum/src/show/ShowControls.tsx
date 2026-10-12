/**
 * Gösteri denetimleri: işaretçi araçlarının durumu ve kısayolları
 * (Ctrl+L lazer, Ctrl+P kalem, Ctrl+I vurgulayıcı, Ctrl+E silgi, Ctrl+A ok,
 * E mürekkebi sil, Ctrl+M mürekkebi göster/gizle), sol alttaki gösteri
 * araç çubuğu ve "Tüm slaytları gör" ızgarası.
 */
import { useCallback, useRef, useState, type ReactElement } from "react";
import { ArrowLeft20Regular, ArrowRight20Regular, Grid20Regular, MoreHorizontal20Regular, Pen20Regular } from "@fluentui/react-icons";
import { Menu, MenuItem, MenuSeparator } from "@officehere/ui";
import type { Presentation } from "../model/types";
import { ScaledSlide } from "../render/SlideView";
import { SHOW_INK_COLORS, type ShowStroke, type ShowTool } from "./ShowInk";

export interface ShowTools {
  tool: ShowTool;
  color: string;
  visible: boolean;
  strokes: ShowStroke[];
  setTool: (t: ShowTool) => void;
  add: (s: ShowStroke) => void;
  erase: (indices: number[]) => void;
  clearSlide: () => void;
  /** Gösteri tuşlarından önce çağrılır; tuşu işlediyse true. */
  handleKey: (e: React.KeyboardEvent) => boolean;
  /** Araç menüsü (araç çubuğundaki kalem düğmesi). */
  toolMenu: () => ReactElement;
  /** Sağ tık menüsündeki "İşaretçi Seçenekleri". */
  pointerMenu: () => ReactElement;
  /** Tüm slaytlardaki mürekkep (gösteri sonunda "Tut/At" için). */
  allInk: () => Record<string, ShowStroke[]>;
}

export function useShowTools(slideId: string | undefined): ShowTools {
  const [tool, setTool] = useState<ShowTool>("pointer");
  const [penColor, setPenColor] = useState(SHOW_INK_COLORS[0].hex);
  const [hlColor, setHlColor] = useState(SHOW_INK_COLORS[5].hex);
  const color = tool === "highlighter" ? hlColor : penColor;
  const [visible, setVisible] = useState(true);
  const [ink, setInk] = useState<Record<string, ShowStroke[]>>({});
  const inkRef = useRef(ink);
  inkRef.current = ink;

  const add = useCallback(
    (s: ShowStroke) => {
      if (!slideId) return;
      setVisible(true);
      setInk((m) => ({ ...m, [slideId]: [...(m[slideId] ?? []), s] }));
    },
    [slideId],
  );
  const erase = useCallback(
    (indices: number[]) => {
      if (!slideId || !indices.length) return;
      const set = new Set(indices);
      setInk((m) => ({ ...m, [slideId]: (m[slideId] ?? []).filter((_, i) => !set.has(i)) }));
    },
    [slideId],
  );
  const clearSlide = useCallback(() => {
    if (!slideId) return;
    setInk((m) => ({ ...m, [slideId]: [] }));
  }, [slideId]);

  const pick = (t: ShowTool) => setTool((cur) => (cur === t && t !== "pointer" ? "pointer" : t));

  const handleKey = (e: React.KeyboardEvent): boolean => {
    const mod = e.ctrlKey || e.metaKey;
    const k = e.code.startsWith("Key") ? e.code.slice(3) : e.key.toLocaleUpperCase("tr-TR");
    if (mod && !e.altKey) {
      const map: Record<string, ShowTool> = { L: "laser", P: "pen", I: "highlighter", E: "eraser", A: "pointer" };
      if (map[k]) {
        e.preventDefault();
        pick(map[k]);
        return true;
      }
      if (k === "M") {
        e.preventDefault();
        setVisible((v) => !v);
        return true;
      }
      return false;
    }
    if (!e.altKey && (k === "E" || e.key === "e" || e.key === "E")) {
      e.preventDefault();
      clearSlide();
      return true;
    }
    if (e.key === "Escape" && tool !== "pointer") {
      e.preventDefault();
      setTool("pointer");
      return true;
    }
    return false;
  };

  const strokes = (slideId && ink[slideId]) || [];

  const toolMenu = () => (
        <Menu>
          <MenuItem label="Lazer İşaretçisi" shortcut="Ctrl+L" checked={tool === "laser"} onClick={() => pick("laser")} />
          <MenuItem label="Kalem" shortcut="Ctrl+P" checked={tool === "pen"} onClick={() => pick("pen")} />
          <MenuItem label="Vurgulayıcı" shortcut="Ctrl+I" checked={tool === "highlighter"} onClick={() => pick("highlighter")} />
          <MenuItem
            label="Mürekkep Rengi"
            submenu={
              <Menu>
                {SHOW_INK_COLORS.map((c) => (
                  <MenuItem
                    key={c.hex}
                    icon={<span className="show-swatch" style={{ background: c.hex }} />}
                    label={c.name}
                    checked={color === c.hex}
                    onClick={() => {
                      if (tool === "highlighter") setHlColor(c.hex);
                      else {
                        setPenColor(c.hex);
                        if (tool !== "pen") setTool("pen");
                      }
                    }}
                  />
                ))}
              </Menu>
            }
          />
          <MenuSeparator />
          <MenuItem label="Silgi" shortcut="Ctrl+E" checked={tool === "eraser"} onClick={() => pick("eraser")} />
          <MenuItem label="Slayttaki Tüm Mürekkebi Sil" shortcut="E" disabled={!strokes.length} onClick={clearSlide} />
          <MenuSeparator />
          <MenuItem label="Ok" shortcut="Ctrl+A" checked={tool === "pointer"} onClick={() => setTool("pointer")} />
          <MenuItem label="Mürekkep İşaretlemesini Göster" shortcut="Ctrl+M" checked={visible} onClick={() => setVisible((v) => !v)} />
        </Menu>
  );
  const pointerMenu = () => <MenuItem label="İşaretçi Seçenekleri" submenu={toolMenu()} />;

  return {
    tool,
    color,
    visible,
    strokes,
    setTool,
    add,
    erase,
    clearSlide,
    handleKey,
    toolMenu,
    pointerMenu,
    allInk: () => Object.fromEntries(Object.entries(inkRef.current).filter(([, l]) => l.length)),
  };
}

/** Sol alttaki yarı saydam gösteri araç çubuğu. */
export function ShowToolbar({ tools, onPrev, onNext, onOpenMenu, onAllSlides, moreMenu }: { tools: ShowTools; onPrev: () => void; onNext: () => void; onOpenMenu: (e: React.MouseEvent, menu: ReactElement) => void; onAllSlides: () => void; moreMenu: () => ReactElement }) {
  return (
    <div className={`show-toolbar${tools.tool !== "pointer" ? " is-tool" : ""}`} onClick={(e) => e.stopPropagation()} onPointerDown={(e) => e.stopPropagation()}>
      <button type="button" className="show-tb-btn" aria-label="Önceki" title="Önceki" onClick={onPrev}>
        <ArrowLeft20Regular />
      </button>
      <button type="button" className="show-tb-btn" aria-label="Sonraki" title="Sonraki" onClick={onNext}>
        <ArrowRight20Regular />
      </button>
      <button
        type="button"
        className={`show-tb-btn${tools.tool !== "pointer" ? " is-on" : ""}`}
        aria-label="Kalem ve lazer işaretçisi araçları"
        title="Kalem ve lazer işaretçisi araçları"
        onClick={(e) => onOpenMenu(e, tools.toolMenu())}
      >
        <Pen20Regular />
      </button>
      <button type="button" className="show-tb-btn" aria-label="Tüm slaytları gör" title="Tüm slaytları gör" onClick={onAllSlides}>
        <Grid20Regular />
      </button>
      <button type="button" className="show-tb-btn" aria-label="Diğer slayt gösterisi seçenekleri" title="Diğer slayt gösterisi seçenekleri" onClick={(e) => onOpenMenu(e, moreMenu())}>
        <MoreHorizontal20Regular />
      </button>
    </div>
  );
}

/** "Tüm slaytları gör": küçük resim ızgarası. */
export function ShowSlideGrid({ pres, current, onPick, onClose }: { pres: Presentation; current: number; onPick: (i: number) => void; onClose: () => void }) {
  return (
    <div
      className="show-grid oh-scroll"
      role="dialog"
      aria-label="Tüm slaytlar"
      onClick={(e) => {
        e.stopPropagation();
        if (e.target === e.currentTarget) onClose();
      }}
    >
      <div className="show-grid-head">
        <button type="button" className="show-tb-btn" aria-label="Geri" onClick={onClose}>
          <ArrowLeft20Regular />
        </button>
        <span>Tüm slaytlar</span>
      </div>
      <div className="show-grid-items">
        {pres.slides.map((s, i) => (
          <button key={s.id} type="button" className={`show-grid-item${i === current ? " is-current" : ""}${s.hidden ? " is-hidden" : ""}`} onClick={() => onPick(i)}>
            <ScaledSlide pres={pres} slide={s} slideNumber={i + 1} width={220} mode="show" />
            <span>{i + 1}</span>
          </button>
        ))}
      </div>
    </div>
  );
}
