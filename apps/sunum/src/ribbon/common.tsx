/** Şerit sekmelerinde ortak kullanılan bileşenler: tema renk paleti, şekil galerisi, düzen galerisi. */
import { useMemo, useRef, type ReactElement, type ReactNode } from "react";
import { MenuHeader, MenuItem, MenuSeparator, RibbonSplitButton, RibbonMenuButton, useMenuClose, type RibbonButtonSize } from "@officehere/ui";
import { STANDARD_COLORS } from "@officehere/core";
import { colorLabel, cssColor, themeShadeMods } from "../model/color";
import { instantiatePlaceholder } from "../model/presentation";
import type { Color, Layout, Presentation, SchemeColorName, Slide, Theme } from "../model/types";
import { SCHEME_ORDER } from "../model/types";
import { presetGeometry, SHAPE_CATALOG, type ShapeCatalogItem } from "../render/geometry";
import { SlideView } from "../render/SlideView";
import { eyedropperAvailable, startEyedropper } from "../editor/Eyedropper";
import { Eyedropper20Regular } from "@fluentui/react-icons";

/* ------------------------------------------------------------------ */
/* Tema renk paleti                                                     */
/* ------------------------------------------------------------------ */

export interface ThemePaletteProps {
  theme: Theme;
  onSelect: (c: Color | null) => void;
  noneLabel?: string;
  autoLabel?: string;
  extra?: ReactNode;
  value?: Color | null;
}

function Swatch({ color, css, title, onPick, selected }: { color: Color; css: string; title: string; onPick: (c: Color) => void; selected?: boolean }) {
  return (
    <button
      type="button"
      className={`oh-swatch${selected ? " is-selected" : ""}`}
      style={{ background: css }}
      title={title}
      aria-label={title}
      onMouseDown={(e) => e.preventDefault()}
      onClick={() => onPick(color)}
    />
  );
}

export function ThemePalette({ theme, onSelect, noneLabel, autoLabel, extra, value }: ThemePaletteProps) {
  const close = useMenuClose();
  const inputRef = useRef<HTMLInputElement>(null);
  const pick = (c: Color | null) => {
    onSelect(c);
    close();
  };
  const same = (c: Color) => !!value && JSON.stringify(c) === JSON.stringify(value);
  return (
    <div className="oh-palette">
      {autoLabel && (
        <>
          <MenuItem icon={<span className="oh-swatch-none" style={{ background: "#000" }} />} label={autoLabel} onClick={() => onSelect(null)} />
          <MenuSeparator />
        </>
      )}
      <MenuHeader>Tema Renkleri</MenuHeader>
      <div className="oh-palette-theme">
        <div className="oh-palette-grid">
          {SCHEME_ORDER.map((n) => {
            const c: Color = { scheme: n };
            return <Swatch key={n} color={c} css={cssColor(c, theme)} title={colorLabel(c)} onPick={pick} selected={same(c)} />;
          })}
        </div>
        <div className="oh-palette-grid is-shades">
          {[0, 1, 2, 3, 4].map((row) =>
            SCHEME_ORDER.map((n: SchemeColorName) => {
              const mods = themeShadeMods(theme.colors[n])[row];
              const c: Color = { scheme: n, ...mods };
              return <Swatch key={`${row}-${n}`} color={c} css={cssColor(c, theme)} title={colorLabel(c)} onPick={pick} selected={same(c)} />;
            }),
          )}
        </div>
      </div>
      <MenuHeader>Standart Renkler</MenuHeader>
      <div className="oh-palette-grid">
        {STANDARD_COLORS.map((h) => {
          const c: Color = { rgb: h };
          return <Swatch key={h} color={c} css={h} title={h} onPick={pick} selected={same(c)} />;
        })}
      </div>
      <MenuSeparator />
      {noneLabel && <MenuItem icon={<span className="oh-swatch-none" />} label={noneLabel} onClick={() => onSelect(null)} />}
      <MenuItem
        label="Diğer Renkler..."
        keepOpen
        icon={
          <svg width="16" height="16" viewBox="0 0 16 16" aria-hidden="true">
            <circle cx="5" cy="6" r="3" fill="#e81123" />
            <circle cx="11" cy="6" r="3" fill="#0078d4" />
            <circle cx="8" cy="11" r="3" fill="#16c60c" />
          </svg>
        }
        onClick={() => inputRef.current?.click()}
      />
      <input ref={inputRef} type="color" className="oh-hidden-input" onChange={(e) => pick({ rgb: e.target.value.toUpperCase() })} />
      <MenuItem
        icon={<Eyedropper20Regular />}
        label="Damlalık"
        disabled={!eyedropperAvailable()}
        onClick={() => startEyedropper((hex) => onSelect({ rgb: hex }))}
      />
      {extra}
    </div>
  );
}

export interface ColorSplitProps {
  icon: ReactElement;
  label?: string;
  tip: string;
  theme: Theme;
  color: Color | null;
  onApply: (c: Color | null) => void;
  onColorChange: (c: Color | null) => void;
  noneLabel?: string;
  autoLabel?: string;
  size?: RibbonButtonSize;
  disabled?: boolean;
  extra?: ReactNode;
  description?: string;
}

/** Yazı tipi rengi, dolgu, anahat gibi renk şeritli bölünmüş düğme (tema rengi döndürür). */
export function ThemeColorSplit({ icon, label, tip, theme, color, onApply, onColorChange, noneLabel, autoLabel, size = "small", disabled, extra, description }: ColorSplitProps) {
  return (
    <RibbonSplitButton
      icon={icon}
      label={label}
      size={size}
      tip={tip}
      description={description}
      disabled={disabled}
      colorBar={color ? cssColor(color, theme) : null}
      onClick={() => onApply(color)}
      menu={
        <div className="oh-menu oh-menu-palette">
          <ThemePalette
            theme={theme}
            noneLabel={noneLabel}
            autoLabel={autoLabel}
            value={color}
            extra={extra}
            onSelect={(c) => {
              onColorChange(c);
              onApply(c);
            }}
          />
        </div>
      }
    />
  );
}

/* ------------------------------------------------------------------ */
/* Şekil simgeleri ve galerisi                                          */
/* ------------------------------------------------------------------ */

export function ShapeIcon({ geom, size = 18, item }: { geom: string; size?: number; item?: ShapeCatalogItem }) {
  const s = size - 2;
  const isLine = geom === "line" || geom === "straightConnector1" || geom === "bentConnector3" || geom === "curvedConnector3";
  const g = presetGeometry(geom, s, isLine ? s : geom.includes("Arrow") && !geom.startsWith("up") && !geom.startsWith("down") ? s * 0.7 : s);
  return (
    <svg width={size} height={size} viewBox={`-1 -1 ${size} ${size}`} aria-hidden="true">
      {item?.tail && (
        <defs>
          <marker id={`mk-${geom}`} viewBox="0 0 10 10" refX="7" refY="5" markerWidth="5" markerHeight="5" orient="auto-start-reverse">
            <path d="M0 0L10 5L0 10z" fill="currentColor" />
          </marker>
        </defs>
      )}
      {g.paths.map((p, i) => (
        <path
          key={i}
          d={p.d}
          fill="none"
          stroke="currentColor"
          strokeWidth={1}
          fillRule={p.evenodd ? "evenodd" : undefined}
          markerEnd={item?.tail ? `url(#mk-${geom})` : undefined}
          markerStart={item?.head ? `url(#mk-${geom})` : undefined}
          vectorEffect="non-scaling-stroke"
        />
      ))}
    </svg>
  );
}

export function ShapeGalleryMenu({ onPick, recent }: { onPick: (item: ShapeCatalogItem) => void; recent?: ShapeCatalogItem[] }) {
  const close = useMenuClose();
  const groups = recent?.length ? [{ group: "Son Kullanılan Şekiller", items: recent }, ...SHAPE_CATALOG] : SHAPE_CATALOG;
  return (
    <div className="oh-menu" style={{ maxHeight: "70vh", width: 300 }}>
      {groups.map((g) => (
        <div key={g.group}>
          <MenuHeader>{g.group}</MenuHeader>
          <div className="shape-grid">
            {g.items.map((it, i) => (
              <button
                key={`${it.geom}-${i}`}
                type="button"
                className="shape-btn"
                title={it.label}
                onMouseDown={(e) => e.preventDefault()}
                onClick={() => {
                  onPick(it);
                  close();
                }}
              >
                <ShapeIcon geom={it.geom} item={it} />
              </button>
            ))}
          </div>
        </div>
      ))}
    </div>
  );
}

/** Giriş > Çizim grubundaki küçük şekil galerisi. */
export function InlineShapeGallery({ onPick }: { onPick: (item: ShapeCatalogItem) => void }) {
  const items = useMemo(() => {
    const all = SHAPE_CATALOG.flatMap((g) => g.items);
    const pick = ["textbox", "line", "straightConnector1", "rect", "ellipse", "triangle", "rtTriangle", "rightArrow", "downArrow", "roundRect", "diamond", "pentagon", "hexagon", "star5", "heart", "lightningBolt", "wedgeRectCallout", "leftArrow", "upArrow", "chevron", "homePlate", "can", "cloud", "parallelogram"];
    return pick.map((g) => (g === "textbox" ? { geom: "textbox", label: "Metin Kutusu" } : all.find((i) => i.geom === g)!)).filter(Boolean) as ShapeCatalogItem[];
  }, []);
  return (
    <div className="shape-gallery-wrap">
      <div className="shape-gallery-inline">
        {items.map((it, i) => (
          <button key={i} type="button" className="shape-btn" title={it.label} style={{ width: 20, height: 20 }} onMouseDown={(e) => e.preventDefault()} onClick={() => onPick(it)}>
            {it.geom === "textbox" ? <TextBoxIcon /> : <ShapeIcon geom={it.geom} item={it} size={16} />}
          </button>
        ))}
      </div>
      <RibbonMenuButton size="small" tip="Şekiller" className="shape-gallery-more" menu={<ShapeGalleryMenu onPick={onPick} />} />
    </div>
  );
}

export function TextBoxIcon() {
  return (
    <svg width="16" height="16" viewBox="0 0 16 16" aria-hidden="true">
      <rect x="1.5" y="2.5" width="13" height="11" fill="none" stroke="currentColor" />
      <path d="M5 5.5h6M8 5.5v6" stroke="currentColor" strokeWidth="1.2" />
    </svg>
  );
}

/* ------------------------------------------------------------------ */
/* Düzen galerisi                                                        */
/* ------------------------------------------------------------------ */

export function layoutPreviewSlide(layout: Layout): Slide {
  const shapes = layout.shapes.filter((s) => s.kind === "sp" && s.ph && !["dt", "ftr", "sldNum"].includes(s.ph.type)).map((s) => instantiatePlaceholder(s as never));
  return { id: `prev-${layout.id}`, layoutId: layout.id, shapes, notes: "", animations: [] };
}

export function LayoutThumb({ pres, layout, width }: { pres: Presentation; layout: Layout; width: number }) {
  const slide = useMemo(() => layoutPreviewSlide(layout), [layout]);
  const scale = width / pres.width;
  return (
    <div style={{ width, height: pres.height * scale, overflow: "hidden", position: "relative", pointerEvents: "none" }}>
      <SlideView pres={pres} slide={slide} slideNumber={1} mode="edit" style={{ transform: `scale(${scale})`, transformOrigin: "0 0", position: "absolute", left: 0, top: 0 }} className="sv-layout-prev" />
    </div>
  );
}

export function LayoutGallery({ pres, onPick, selectedId, header = "Düzen" }: { pres: Presentation; onPick: (id: string) => void; selectedId?: string; header?: string }) {
  const close = useMenuClose();
  // Seçili slaytlara uygulanmış temaların düzenleri ayrı başlık altında (asıl slayt başına bir grup).
  const groups: { name: string; layouts: Layout[] }[] = [{ name: header, layouts: pres.layouts.filter((l) => !l.theme) }];
  for (const l of pres.layouts) {
    if (!l.theme) continue;
    const g = groups.find((x) => x.name === l.theme!.name && x !== groups[0]);
    if (g) g.layouts.push(l);
    else groups.push({ name: l.theme.name, layouts: [l] });
  }
  return (
    <div className="oh-menu" style={{ maxHeight: "75vh" }}>
      {groups.map((g, gi) => (
        <div key={gi}>
          <MenuHeader>{g.name}</MenuHeader>
          <div className="layout-grid">
            {g.layouts.map((l) => (
              <button
                key={l.id}
                type="button"
                className={`layout-tile${l.id === selectedId ? " is-selected" : ""}`}
                onMouseDown={(e) => e.preventDefault()}
                onClick={() => {
                  onPick(l.id);
                  close();
                }}
              >
                <span className="layout-tile-prev">
                  <LayoutThumb pres={pres} layout={l} width={104} />
                </span>
                <span>{l.name}</span>
              </button>
            ))}
          </div>
        </div>
      ))}
    </div>
  );
}

export function cm(px: number): number {
  return Math.round((px / (96 / 2.54)) * 100) / 100;
}

export function fromCm(v: number): number {
  return (v * 96) / 2.54;
}
