/** Şekil Dolgusu, Şekil Anahattı, Şekil Efektleri, Hızlı Stiller ve Yerleştir menüleri. */
import { useState } from "react";
import {
  ArrowRotateClockwise20Regular,
  ArrowRotateCounterclockwise20Regular,
  ColorFill20Regular,
  ColorLine20Regular,
  FlipHorizontal20Regular,
  FlipVertical20Regular,
  Group20Regular,
  GroupDismiss20Regular,
  Image20Regular,
  PositionBackward20Regular,
  PositionForward20Regular,
  PositionToBack20Regular,
  PositionToFront20Regular,
  Shapes32Regular,
  TextEffects20Regular,
  AlignLeft20Regular,
  AlignCenterHorizontal20Regular,
  AlignRight20Regular,
  AlignTop20Regular,
  AlignCenterVertical20Regular,
  AlignBottom20Regular,
  AlignSpaceEvenlyHorizontal20Regular,
  AlignSpaceEvenlyVertical20Regular,
  LayerDiagonal20Regular,
  ArrowSort20Regular,
} from "@fluentui/react-icons";
import { Menu, MenuHeader, MenuItem, MenuSeparator, RibbonMenuButton, useMenuClose } from "@officehere/ui";
import { cssColor } from "../model/color";
import type { Color, Fill, Line, Shape, Theme } from "../model/types";
import { fillToCss } from "../render/style";
import type { Actions } from "../state/actions";
import { readImageFile } from "../app/clipboard";
import { ARROW_PRESETS, DASHES, fillBaseColor, GLOW_PRESETS, gradientPresets, LINE_WEIGHTS, QUICK_STYLES, SHADOW_PRESETS, type QuickStyle } from "./presets";
import { ThemeColorSplit } from "./common";
import { applyRunPropsAll } from "../model/text";

function deepMap(s: Shape, fn: (s: Shape) => Shape): Shape {
  return s.kind === "group" ? { ...s, children: s.children.map((c) => deepMap(c, fn)) } : fn(s);
}

export function setFill(a: Actions, fill: Fill) {
  a.updateShapes((s) =>
    deepMap(s, (x) => {
      if (x.kind === "sp") return { ...x, fill };
      if (x.kind === "table") {
        const cs = a.s.cellSel;
        return {
          ...x,
          rows: x.rows.map((r, ri) => ({
            ...r,
            cells: r.cells.map((c, ci) => (!cs || (ri >= Math.min(cs.r1, cs.r2) && ri <= Math.max(cs.r1, cs.r2) && ci >= Math.min(cs.c1, cs.c2) && ci <= Math.max(cs.c1, cs.c2)) ? { ...c, fill } : c)),
          })),
        };
      }
      return x;
    }),
  );
}

export function setLine(a: Actions, patch: (l: Line) => Line) {
  a.updateShapes((s) => deepMap(s, (x) => (x.kind === "sp" || x.kind === "pic" ? { ...x, line: patch(x.line) } : x)));
}

export function setEffects(a: Actions, patch: (fx: NonNullable<Extract<Shape, { kind: "sp" }>["effects"]>) => Extract<Shape, { kind: "sp" }>["effects"]) {
  a.updateShapes((s) => deepMap(s, (x) => (x.kind === "sp" || x.kind === "pic" ? { ...x, effects: patch(x.effects ?? {}) } : x)));
}

export function applyQuickStyle(a: Actions, qs: QuickStyle) {
  a.updateShapes((s) =>
    deepMap(s, (x) => {
      if (x.kind !== "sp") return x;
      const text = x.text ? { ...applyRunPropsAll(x.text, (p) => ({ ...p, color: undefined })), levels: [{ ...(x.text.levels?.[0] ?? {}), run: { ...(x.text.levels?.[0]?.run ?? {}), color: qs.text } }] } : x.text;
      return { ...x, fill: qs.fill, line: qs.line, effects: qs.effects, text };
    }),
  );
}

export function currentSp(a: Actions) {
  const s = a.selectedShapes()[0];
  if (!s) return undefined;
  if (s.kind === "group") return s.children.find((c) => c.kind === "sp") as Extract<Shape, { kind: "sp" }> | undefined;
  return s.kind === "sp" ? s : undefined;
}

export function ShapeFillButton({ a, theme, size = "small", label }: { a: Actions; theme: Theme; size?: "small" | "medium" | "large"; label?: string }) {
  const [color, setColor] = useState<Color | null>({ scheme: "accent1" });
  const disabled = a.s.sel.length === 0;
  const base = fillBaseColor(currentSp(a)?.fill);
  return (
    <ThemeColorSplit
      icon={<ColorFill20Regular />}
      label={label}
      size={size}
      tip="Şekil Dolgusu"
      description="Seçili şekli düz renk, gradyan, resim veya dokuyla doldurun."
      theme={theme}
      color={color}
      disabled={disabled}
      noneLabel="Dolgu Yok"
      onColorChange={setColor}
      onApply={(c) => setFill(a, c ? { type: "solid", color: c } : { type: "none" })}
      extra={
        <>
          <MenuItem
            icon={<Image20Regular />}
            label="Resim..."
            onClick={() => {
              const input = document.createElement("input");
              input.type = "file";
              input.accept = "image/*";
              input.onchange = async () => {
                const f = input.files?.[0];
                if (!f) return;
                const img = await readImageFile(f);
                setFill(a, { type: "image", src: img.src, mode: "stretch" });
              };
              input.click();
            }}
          />
          <MenuItem
            label="Gradyan"
            icon={<span className="oh-swatch-none" style={{ background: `linear-gradient(90deg, #fff, ${cssColor(base, theme)})`, border: "1px solid #999" }} />}
            submenu={
              <Menu>
                <MenuHeader>Çeşitlemeler</MenuHeader>
                <div className="preset-grid" style={{ gridTemplateColumns: "repeat(4, 40px)" }}>
                  {gradientPresets(base).map((g) => (
                    <GradientTile key={g.label} label={g.label} fill={g.fill} theme={theme} onPick={() => setFill(a, g.fill)} />
                  ))}
                </div>
                <MenuSeparator />
                <MenuItem label="Gradyan Yok" onClick={() => setFill(a, { type: "solid", color: base })} />
              </Menu>
            }
          />
        </>
      }
    />
  );
}

function GradientTile({ label, fill, theme, onPick }: { label: string; fill: Fill; theme: Theme; onPick: () => void }) {
  const close = useMenuClose();
  return (
    <button
      type="button"
      className="preset-btn"
      title={label}
      onMouseDown={(e) => e.preventDefault()}
      onClick={() => {
        onPick();
        close();
      }}
    >
      <span style={{ display: "block", width: 34, height: 34, background: fillToCss(fill, theme) }} />
    </button>
  );
}

export function ShapeOutlineButton({ a, theme, size = "small", label }: { a: Actions; theme: Theme; size?: "small" | "medium" | "large"; label?: string }) {
  const [color, setColor] = useState<Color | null>({ scheme: "accent1", lumMod: 0.5 });
  const disabled = a.s.sel.length === 0;
  const cur = a.selectedShapes()[0];
  const curLine = cur && (cur.kind === "sp" || cur.kind === "pic") ? cur.line : undefined;
  const ensure = (l: Line): Line => (l.fill.type === "none" ? { ...l, fill: { type: "solid", color: color ?? { scheme: "dk1" } } } : l);
  return (
    <ThemeColorSplit
      icon={<ColorLine20Regular />}
      label={label}
      size={size}
      tip="Şekil Anahattı"
      description="Seçili şeklin anahattının rengini, kalınlığını ve çizgi stilini seçin."
      theme={theme}
      color={color}
      disabled={disabled}
      noneLabel="Anahat Yok"
      onColorChange={setColor}
      onApply={(c) => setLine(a, (l) => ({ ...l, fill: c ? { type: "solid", color: c } : { type: "none" } }))}
      extra={
        <>
          <MenuItem
            label="Kalınlık"
            submenu={
              <Menu>
                {LINE_WEIGHTS.map((w) => (
                  <MenuItem
                    key={w}
                    checked={curLine?.width === w && curLine.fill.type !== "none"}
                    label={
                      <span style={{ display: "flex", alignItems: "center", gap: 10 }}>
                        <span style={{ width: 40 }}>{w.toLocaleString("tr-TR")} nk</span>
                        <span style={{ display: "inline-block", width: 90, borderTop: `${Math.max(1, w * 1.33)}px solid currentColor` }} />
                      </span>
                    }
                    onClick={() => setLine(a, (l) => ({ ...ensure(l), width: w }))}
                  />
                ))}
              </Menu>
            }
          />
          <MenuItem
            label="Çizgiler"
            submenu={
              <Menu>
                {DASHES.map((d) => (
                  <MenuItem
                    key={d.id}
                    checked={curLine?.dash === d.id}
                    label={
                      <span title={d.label} style={{ display: "inline-block", width: 120, borderTop: `2px ${d.id === "solid" ? "solid" : d.id.includes("Dot") || d.id === "sysDot" ? "dotted" : "dashed"} currentColor` }} />
                    }
                    onClick={() => setLine(a, (l) => ({ ...ensure(l), dash: d.id }))}
                  />
                ))}
              </Menu>
            }
          />
          <MenuItem
            label="Oklar"
            submenu={
              <Menu>
                {ARROW_PRESETS.map((p) => (
                  <MenuItem key={p.label} label={p.label} onClick={() => setLine(a, (l) => ({ ...ensure(l), head: p.head, tail: p.tail }))} />
                ))}
              </Menu>
            }
          />
        </>
      }
    />
  );
}

export function ShapeEffectsMenu({ a, theme, size = "small", label }: { a: Actions; theme: Theme; size?: "small" | "medium" | "large"; label?: string }) {
  const disabled = a.s.sel.length === 0;
  return (
    <RibbonMenuButton
      size={size}
      icon={<TextEffects20Regular />}
      label={label}
      tip="Şekil Efektleri"
      description="Seçili şekle gölge, parlama, yansıma gibi görsel efektler uygulayın."
      disabled={disabled}
      menu={
        <Menu>
          <MenuItem
            label="Gölge"
            submenu={
              <Menu>
                {SHADOW_PRESETS.map((g) => (
                  <div key={g.group}>
                    <MenuHeader>{g.group}</MenuHeader>
                    {g.items.map((it) => (
                      <MenuItem key={it.label} label={it.label} onClick={() => setEffects(a, (fx) => ({ ...fx, shadow: it.shadow ?? undefined }))} />
                    ))}
                  </div>
                ))}
              </Menu>
            }
          />
          <MenuItem
            label="Yansıma"
            submenu={
              <Menu>
                <MenuItem label="Yansıma Yok" onClick={() => setEffects(a, (fx) => ({ ...fx, reflection: undefined }))} />
                <MenuItem label="Sıkı Yansıma: Dokunma" onClick={() => setEffects(a, (fx) => ({ ...fx, reflection: true }))} />
              </Menu>
            }
          />
          <MenuItem
            label="Parlama"
            submenu={
              <Menu>
                <MenuItem label="Parlama Yok" onClick={() => setEffects(a, (fx) => ({ ...fx, glow: undefined }))} />
                <MenuHeader>Parlama Çeşitlemeleri</MenuHeader>
                {GLOW_PRESETS.map((g) => (
                  <MenuItem
                    key={g.label}
                    label={g.label}
                    icon={<span className="oh-swatch-none" style={{ background: cssColor({ scheme: "accent1", alpha: 0.6 }, theme), boxShadow: `0 0 4px ${cssColor({ scheme: "accent1" }, theme)}` }} />}
                    onClick={() => setEffects(a, (fx) => ({ ...fx, glow: { color: { scheme: "accent1", alpha: 0.6 }, radius: g.radius } }))}
                  />
                ))}
              </Menu>
            }
          />
        </Menu>
      }
    />
  );
}

export function QuickStylesMenu({ a, theme, size = "large" }: { a: Actions; theme: Theme; size?: "large" | "medium" | "small" }) {
  return (
    <RibbonMenuButton
      size={size}
      icon={<Shapes32Regular />}
      label="Hızlı Stiller"
      tip="Hızlı Stiller"
      description="Şekle hazır bir görsel stil uygulayın."
      disabled={a.s.sel.length === 0}
      menu={<QuickStyleGrid a={a} theme={theme} />}
    />
  );
}

export function QuickStyleGrid({ a, theme }: { a: Actions; theme: Theme }) {
  const close = useMenuClose();
  return (
    <div className="oh-menu">
      <MenuHeader>Tema Stilleri</MenuHeader>
      <div className="preset-grid" style={{ gridTemplateColumns: "repeat(7, 56px)" }}>
        {QUICK_STYLES.map((q) => (
          <button
            key={q.id}
            type="button"
            className="preset-btn"
            title={q.label}
            onMouseDown={(e) => e.preventDefault()}
            onClick={() => {
              applyQuickStyle(a, q);
              close();
            }}
          >
            <QuickStylePreview q={q} theme={theme} />
          </button>
        ))}
      </div>
    </div>
  );
}

export function QuickStylePreview({ q, theme, w = 50, h = 36 }: { q: QuickStyle; theme: Theme; w?: number; h?: number }) {
  const border = q.line.fill.type === "solid" ? `${Math.max(1, q.line.width)}px solid ${cssColor(q.line.fill.color, theme)}` : "1px solid transparent";
  return (
    <span
      style={{
        display: "flex",
        alignItems: "center",
        justifyContent: "center",
        width: w,
        height: h,
        background: fillToCss(q.fill, theme),
        border,
        borderRadius: 3,
        color: cssColor(q.text, theme),
        fontSize: 14,
        boxShadow: q.effects?.shadow ? "2px 2px 4px rgba(0,0,0,0.4)" : undefined,
      }}
    >
      Abc
    </span>
  );
}

/** Hizala alt menüsü öğeleri. */
export function AlignItems({ a }: { a: Actions }) {
  const [toSlide, setToSlide] = useState(a.s.sel.length < 2);
  return (
    <>
      <MenuItem icon={<AlignLeft20Regular />} label="Sola Hizala" onClick={() => a.alignSelection("left", toSlide)} />
      <MenuItem icon={<AlignCenterHorizontal20Regular />} label="Ortala" onClick={() => a.alignSelection("center", toSlide)} />
      <MenuItem icon={<AlignRight20Regular />} label="Sağa Hizala" onClick={() => a.alignSelection("right", toSlide)} />
      <MenuSeparator />
      <MenuItem icon={<AlignTop20Regular />} label="Üste Hizala" onClick={() => a.alignSelection("top", toSlide)} />
      <MenuItem icon={<AlignCenterVertical20Regular />} label="Ortaya Hizala" onClick={() => a.alignSelection("middle", toSlide)} />
      <MenuItem icon={<AlignBottom20Regular />} label="Alta Hizala" onClick={() => a.alignSelection("bottom", toSlide)} />
      <MenuSeparator />
      <MenuItem icon={<AlignSpaceEvenlyHorizontal20Regular />} label="Yatay Olarak Dağıt" onClick={() => a.alignSelection("distH", toSlide)} />
      <MenuItem icon={<AlignSpaceEvenlyVertical20Regular />} label="Dikey Olarak Dağıt" onClick={() => a.alignSelection("distV", toSlide)} />
      <MenuSeparator />
      <MenuItem label="Slayda Hizala" checked={toSlide} keepOpen onClick={() => setToSlide(true)} />
      <MenuItem label="Seçili Nesneleri Hizala" checked={!toSlide} disabled={a.s.sel.length < 2} keepOpen onClick={() => setToSlide(false)} />
    </>
  );
}

export function RotateItems({ a, onMore }: { a: Actions; onMore?: () => void }) {
  return (
    <>
      <MenuItem icon={<ArrowRotateClockwise20Regular />} label="Sağa 90° Döndür" onClick={() => a.rotateSelection(90)} />
      <MenuItem icon={<ArrowRotateCounterclockwise20Regular />} label="Sola 90° Döndür" onClick={() => a.rotateSelection(-90)} />
      <MenuItem icon={<FlipVertical20Regular />} label="Dikey Çevir" onClick={() => a.rotateSelection(null, "v")} />
      <MenuItem icon={<FlipHorizontal20Regular />} label="Yatay Çevir" onClick={() => a.rotateSelection(null, "h")} />
      {onMore && (
        <>
          <MenuSeparator />
          <MenuItem label="Diğer Döndürme Seçenekleri..." onClick={onMore} />
        </>
      )}
    </>
  );
}

/** Giriş > Çizim > Yerleştir menüsü. */
export function ArrangeMenu({ a, onSelectionPane, onMoreRotate }: { a: Actions; onSelectionPane: () => void; onMoreRotate: () => void }) {
  const has = a.s.sel.length > 0;
  const shapes = a.selectedShapes();
  return (
    <Menu className="menu-wide">
      <MenuHeader>Nesneleri Sırala</MenuHeader>
      <MenuItem icon={<PositionToFront20Regular />} label="En Öne Getir" disabled={!has} onClick={() => a.reorderSelection("front")} />
      <MenuItem icon={<PositionToBack20Regular />} label="En Arkaya Gönder" disabled={!has} onClick={() => a.reorderSelection("back")} />
      <MenuItem icon={<PositionForward20Regular />} label="Öne Getir" disabled={!has} onClick={() => a.reorderSelection("forward")} />
      <MenuItem icon={<PositionBackward20Regular />} label="Arkaya Gönder" disabled={!has} onClick={() => a.reorderSelection("backward")} />
      <MenuHeader>Grup Nesneleri</MenuHeader>
      <MenuItem icon={<Group20Regular />} label="Grupla" shortcut="Ctrl+G" disabled={!a.canGroup()} onClick={() => a.groupSelection()} />
      <MenuItem icon={<GroupDismiss20Regular />} label="Grubu Çöz" shortcut="Ctrl+Shift+G" disabled={!shapes.some((s) => s.kind === "group")} onClick={() => a.ungroupSelection()} />
      <MenuHeader>Nesneleri Konumlandır</MenuHeader>
      <MenuItem
        icon={<ArrowSort20Regular />}
        label="Hizala"
        disabled={!has}
        submenu={
          <Menu>
            <AlignItems a={a} />
          </Menu>
        }
      />
      <MenuItem
        icon={<ArrowRotateClockwise20Regular />}
        label="Döndür"
        disabled={!has}
        submenu={
          <Menu>
            <RotateItems a={a} onMore={onMoreRotate} />
          </Menu>
        }
      />
      <MenuSeparator />
      <MenuItem icon={<LayerDiagonal20Regular />} label="Seçim Bölmesi..." onClick={onSelectionPane} />
    </Menu>
  );
}
