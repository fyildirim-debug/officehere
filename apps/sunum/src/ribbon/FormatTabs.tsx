/** Bağlamsal sekmeler: Şekil Biçimi ve Resim Biçimi. */
import { useState } from "react";
import {
  AlignLeft24Regular,
  ArrowRotateClockwise24Regular,
  BrightnessHigh24Regular,
  Color24Regular,
  Crop24Regular,
  Edit24Regular,
  Group24Regular,
  ImageArrowCounterclockwise20Regular,
  ImageEdit20Regular,
  LayerDiagonal20Regular,
  PositionBackward24Regular,
  PositionForward24Regular,
  TextColor20Regular,
  TextEffects20Regular,
  ColorLine20Regular,
  Textbox24Regular,
  ImageBorder20Regular,
  DarkTheme24Regular,
  Image24Regular,
  ImageEdit24Regular,
  PaintBucket24Regular,
  ResizeLarge24Regular,
  Shapes24Regular,
  TextEffects24Regular,
} from "@fluentui/react-icons";
import { Menu, MenuHeader, MenuItem, MenuSeparator, RibbonButton, RibbonGallery, RibbonGroup, RibbonMenuButton, RibbonSpinner, RibbonSplitButton, RibbonStack, useMenuClose } from "@officehere/ui";
import { cssColor } from "../model/color";
import type { Color, PicShape, Shape } from "../model/types";
import { applyRun } from "../state/textActions";
import { readImageFile } from "../app/clipboard";
import { cm, fromCm, InlineShapeGallery, ShapeGalleryMenu, ThemeColorSplit } from "./common";
import { QUICK_STYLES, WORDART_PRESETS } from "./presets";
import { AlignItems, applyQuickStyle, QuickStylePreview, RotateItems, ShapeEffectsMenu, ShapeFillButton, ShapeOutlineButton } from "./shapeMenus";
import type { RibbonCtx } from "./types";
import { toCustomShape } from "../model/editPoints";
import { mergeSelection } from "../editor/mergeShapes";
import { ShapeExclude20Regular, ShapeIntersect20Regular, ShapeSubtract20Regular, ShapeUnion20Regular, ShapeUnion24Regular } from "@fluentui/react-icons";
import { isLineGeom } from "../render/geometry";

function ArrangeGroup({ ctx }: { ctx: RibbonCtx }) {
  const { actions: a, cmd } = ctx;
  const shapes = a.selectedShapes();
  return (
    <RibbonGroup label="Yerleştir" icon={<Shapes24Regular />}>
      <RibbonSplitButton
        size="large"
        icon={<PositionForward24Regular />}
        label="Öne Getir"
        onClick={() => a.reorderSelection("forward")}
        menu={
          <Menu>
            <MenuItem label="Öne Getir" onClick={() => a.reorderSelection("forward")} />
            <MenuItem label="En Öne Getir" onClick={() => a.reorderSelection("front")} />
          </Menu>
        }
      />
      <RibbonSplitButton
        size="large"
        icon={<PositionBackward24Regular />}
        label="Arkaya Gönder"
        onClick={() => a.reorderSelection("backward")}
        menu={
          <Menu>
            <MenuItem label="Arkaya Gönder" onClick={() => a.reorderSelection("backward")} />
            <MenuItem label="En Arkaya Gönder" onClick={() => a.reorderSelection("back")} />
          </Menu>
        }
      />
      <RibbonStack>
        <RibbonButton icon={<LayerDiagonal20Regular />} label="Seçim Bölmesi" checked={ctx.state.pane === "selection"} onClick={() => cmd.togglePane("selection")} />
        <RibbonMenuButton
          icon={<AlignLeft24Regular />}
          label="Hizala"
          menu={
            <Menu>
              <AlignItems a={a} />
            </Menu>
          }
        />
        <RibbonMenuButton
          icon={<Group24Regular />}
          label="Grupla"
          menu={
            <Menu>
              <MenuItem label="Grupla" shortcut="Ctrl+G" disabled={!a.canGroup()} onClick={() => a.groupSelection()} />
              <MenuItem label="Grubu Çöz" shortcut="Ctrl+Shift+G" disabled={!shapes.some((s) => s.kind === "group")} onClick={() => a.ungroupSelection()} />
            </Menu>
          }
        />
      </RibbonStack>
      <RibbonStack>
        <RibbonMenuButton
          icon={<ArrowRotateClockwise24Regular />}
          label="Döndür"
          menu={
            <Menu>
              <RotateItems a={a} onMore={() => ctx.openDialog("sizePos")} />
            </Menu>
          }
        />
      </RibbonStack>
    </RibbonGroup>
  );
}

function SizeGroup({ ctx, lock }: { ctx: RibbonCtx; lock?: boolean }) {
  const { actions: a } = ctx;
  const s: Shape | undefined = a.selectedShapes()[0];
  return (
    <RibbonGroup label="Boyut" icon={<ResizeLarge24Regular />} onLauncher={() => ctx.openDialog("sizePos")} launcherTip="Boyut ve Konum">
      <RibbonStack>
        <RibbonSpinner label="Yükseklik:" value={s ? cm(s.h) : 0} min={0} max={142} step={0.1} decimals={2} unit="cm" width={84} onChange={(v) => a.setSize(undefined, fromCm(v), lock)} />
        <RibbonSpinner label="Genişlik:" value={s ? cm(s.w) : 0} min={0} max={142} step={0.1} decimals={2} unit="cm" width={84} onChange={(v) => a.setSize(fromCm(v), undefined, lock)} />
      </RibbonStack>
    </RibbonGroup>
  );
}

function WordArtGallery({ ctx }: { ctx: RibbonCtx }) {
  const theme = ctx.state.pres.theme;
  return (
    <RibbonGallery
      columns={3}
      itemWidth={44}
      itemHeight={40}
      expandedColumns={4}
      items={WORDART_PRESETS.map((w, i) => ({
        id: String(i),
        label: w.label,
        preview: (
          <span className="wordart-tile" style={{ color: cssColor(w.color, theme), WebkitTextStroke: w.outline ? `1px ${cssColor(w.outline, theme)}` : undefined, textShadow: w.shadow ? "2px 2px 2px rgba(0,0,0,.4)" : undefined, fontSize: 22 }}>
            A
          </span>
        ),
      }))}
      onSelect={(id) => {
        const w = WORDART_PRESETS[Number(id)];
        applyRun(ctx.actions, (p) => ({ ...p, color: w.color, outline: w.outline, shadow: w.shadow || undefined, glow: w.glow ? w.color : undefined, bold: w.bold || p.bold }));
      }}
    />
  );
}

export function ShapeFormatTab({ ctx }: { ctx: RibbonCtx }) {
  const { state, actions: a, cmd } = ctx;
  const theme = state.pres.theme;
  const [textFill, setTextFill] = useState<Color | null>({ scheme: "accent1" });
  const [textLine, setTextLine] = useState<Color | null>({ scheme: "dk1" });
  const sp = a.selectedShapes().find((s) => s.kind === "sp");
  return (
    <>
      <RibbonGroup label="Şekil Ekle" icon={<Shapes24Regular />}>
        <InlineShapeGallery onPick={cmd.pickShape} />
        <RibbonStack>
          <RibbonMenuButton
            icon={<Edit24Regular />}
            label="Şekli Düzenle"
            disabled={!sp}
            menu={
              <Menu>
                <MenuItem
                  label="Şekli Değiştir"
                  submenu={
                    <ShapeGalleryMenu
                      onPick={(it) => {
                        if (it.geom === "textbox") return;
                        a.updateShapes((s) => (s.kind === "sp" ? { ...s, geom: it.geom, adj: undefined, paths: undefined } : s));
                        a.store.set({ pointEdit: null });
                      }}
                    />
                  }
                />
                <MenuItem
                  label="Noktaları Düzenle"
                  disabled={!sp || sp.kind !== "sp" || isLineGeom(sp.geom) || !!sp.ink}
                  checked={!!sp && state.pointEdit === sp.id}
                  onClick={() => {
                    if (!sp || sp.kind !== "sp") return;
                    if (state.pointEdit === sp.id) return a.store.set({ pointEdit: null });
                    a.updateShapes((s) => (s.kind === "sp" ? toCustomShape(s) : s), [sp.id]);
                    a.store.set({ pointEdit: sp.id, editing: null });
                  }}
                />
              </Menu>
            }
          />
          <RibbonButton icon={<Textbox24Regular />} label="Metin Kutusu" onClick={() => a.store.set({ insert: { kind: "textbox" } })} />
          <RibbonMenuButton
            icon={<ShapeUnion24Regular />}
            label="Şekilleri Birleştir"
            description="Seçili şekilleri tek bir yeni şekilde birleştirin (en az iki şekil seçin)."
            disabled={a.selectedShapes().filter((s) => s.kind === "sp").length < 2}
            menu={
              <Menu>
                <MenuItem icon={<ShapeUnion20Regular />} label="Birleşim" onClick={() => mergeSelection(a, "union")} />
                <MenuItem icon={<ShapeExclude20Regular />} label="Birleştir" onClick={() => mergeSelection(a, "combine")} />
                <MenuItem icon={<ShapeIntersect20Regular />} label="Kesişim" onClick={() => mergeSelection(a, "intersect")} />
                <MenuItem icon={<ShapeSubtract20Regular />} label="Çıkar" onClick={() => mergeSelection(a, "subtract")} />
              </Menu>
            }
          />
        </RibbonStack>
      </RibbonGroup>
      <RibbonGroup label="Şekil Stilleri" icon={<PaintBucket24Regular />} onLauncher={() => cmd.togglePane("shape")} launcherTip="Şekli Biçimlendir">
        <RibbonGallery
          columns={6}
          itemWidth={56}
          itemHeight={42}
          expandedColumns={7}
          items={QUICK_STYLES.map((q) => ({ id: q.id, label: q.label, preview: <QuickStylePreview q={q} theme={theme} w={48} h={32} /> }))}
          onSelect={(id) => {
            const q = QUICK_STYLES.find((x) => x.id === id);
            if (q) applyQuickStyle(a, q);
          }}
        />
        <RibbonStack>
          <ShapeFillButton a={a} theme={theme} label="Şekil Dolgusu" size="medium" />
          <ShapeOutlineButton a={a} theme={theme} label="Şekil Anahattı" size="medium" />
          <ShapeEffectsMenu a={a} theme={theme} label="Şekil Efektleri" size="medium" />
        </RibbonStack>
      </RibbonGroup>
      <RibbonGroup label="Sanatsal Metin Stilleri" icon={<TextEffects24Regular />} onLauncher={() => cmd.togglePane("shape")} launcherTip="Metin Efektlerini Biçimlendir">
        <WordArtGallery ctx={ctx} />
        <RibbonStack>
          <ThemeColorSplit icon={<TextColor20Regular />} tip="Metin Dolgusu" label="Metin Dolgusu" size="medium" theme={theme} color={textFill} onColorChange={setTextFill} onApply={(c) => applyRun(a, (p) => ({ ...p, color: c ?? undefined }))} />
          <ThemeColorSplit icon={<ColorLine20Regular />} tip="Metin Anahattı" label="Metin Anahattı" size="medium" theme={theme} color={textLine} noneLabel="Anahat Yok" onColorChange={setTextLine} onApply={(c) => applyRun(a, (p) => ({ ...p, outline: c ?? undefined }))} />
          <RibbonMenuButton
            icon={<TextEffects20Regular />}
            label="Metin Efektleri"
            menu={
              <Menu>
                <MenuItem label="Gölge" onClick={() => applyRun(a, (p, _pp, eff) => ({ ...p, shadow: !eff.shadow }))} />
                <MenuItem label="Parlama" onClick={() => applyRun(a, (p, _pp, eff) => ({ ...p, glow: eff.glow ? undefined : { scheme: "accent1", alpha: 0.6 } }))} />
                <MenuSeparator />
                <MenuItem label="Efektleri Temizle" onClick={() => applyRun(a, (p) => ({ ...p, shadow: undefined, glow: undefined, outline: undefined }))} />
              </Menu>
            }
          />
        </RibbonStack>
      </RibbonGroup>
      <ArrangeGroup ctx={ctx} />
      <SizeGroup ctx={ctx} />
    </>
  );
}

/* ------------------------------------------------------------------ */
/* Resim Biçimi                                                         */
/* ------------------------------------------------------------------ */

const PIC_STYLES: { label: string; apply: (p: PicShape) => PicShape }[] = [
  { label: "Basit Çerçeve, Beyaz", apply: (p) => ({ ...p, geom: undefined, line: { fill: { type: "solid", color: { rgb: "#FFFFFF" } }, width: 6, dash: "solid" }, effects: { shadow: { color: { rgb: "#000000", alpha: 0.35 }, blur: 6, dist: 3, dir: 45 } } }) },
  { label: "Gölgeli Dikdörtgen", apply: (p) => ({ ...p, geom: undefined, line: { fill: { type: "none" }, width: 1, dash: "solid" }, effects: { shadow: { color: { rgb: "#000000", alpha: 0.5 }, blur: 12, dist: 6, dir: 45 } } }) },
  { label: "Yuvarlatılmış Köşeli Dikdörtgen", apply: (p) => ({ ...p, geom: "roundRect", line: { fill: { type: "none" }, width: 1, dash: "solid" }, effects: undefined }) },
  { label: "Oval, Siyah", apply: (p) => ({ ...p, geom: "ellipse", line: { fill: { type: "solid", color: { rgb: "#000000" } }, width: 3, dash: "solid" }, effects: undefined }) },
  { label: "Kalın Çerçeve, Siyah", apply: (p) => ({ ...p, geom: undefined, line: { fill: { type: "solid", color: { rgb: "#000000" } }, width: 8, dash: "solid" }, effects: undefined }) },
  { label: "Yansımalı Yuvarlatılmış Dikdörtgen", apply: (p) => ({ ...p, geom: "roundRect", line: { fill: { type: "none" }, width: 1, dash: "solid" }, effects: { reflection: true } }) },
  { label: "Kesik Köşe", apply: (p) => ({ ...p, geom: "snip1Rect", line: { fill: { type: "none" }, width: 1, dash: "solid" }, effects: undefined }) },
  { label: "Yumuşak Kenar Oval", apply: (p) => ({ ...p, geom: "ellipse", line: { fill: { type: "none" }, width: 1, dash: "solid" }, effects: { glow: { color: { rgb: "#FFFFFF", alpha: 0.8 }, radius: 8 } } }) },
];

function PicStyleIcon({ i }: { i: number }) {
  const r = i === 2 || i === 5 ? 6 : i === 3 || i === 7 ? 20 : 0;
  return (
    <span
      style={{
        display: "block",
        width: 40,
        height: 30,
        borderRadius: r,
        background: "linear-gradient(135deg, #8ec5fc, #3a7bd5)",
        border: i === 0 ? "3px solid #fff" : i === 4 ? "3px solid #000" : i === 3 ? "2px solid #000" : undefined,
        boxShadow: i === 0 || i === 1 ? "2px 2px 4px rgba(0,0,0,.45)" : undefined,
        clipPath: i === 6 ? "polygon(0 0, 80% 0, 100% 25%, 100% 100%, 0 100%)" : undefined,
      }}
    />
  );
}

function AdjustGrid({ ctx }: { ctx: RibbonCtx }) {
  const close = useMenuClose();
  const vals = [-0.4, -0.2, 0, 0.2, 0.4];
  return (
    <div className="oh-menu">
      <MenuHeader>Parlaklık/Karşıtlık</MenuHeader>
      <div className="preset-grid" style={{ gridTemplateColumns: "repeat(5, 50px)" }}>
        {vals.map((c) =>
          vals.map((b) => (
            <button
              key={`${b}${c}`}
              type="button"
              className="preset-btn"
              title={`Parlaklık: %${Math.round(b * 100)} Karşıtlık: %${Math.round(c * 100)}`}
              onMouseDown={(e) => e.preventDefault()}
              onClick={() => {
                ctx.actions.updateShapes((s) => (s.kind === "pic" ? { ...s, adjust: { ...s.adjust, brightness: b || undefined, contrast: c || undefined } } : s));
                close();
              }}
            >
              <span style={{ display: "block", width: 44, height: 30, background: "linear-gradient(135deg, #f6d365, #fda085)", filter: `brightness(${1 + b}) contrast(${1 + c})` }} />
            </button>
          )),
        )}
      </div>
    </div>
  );
}

export function PictureFormatTab({ ctx }: { ctx: RibbonCtx }) {
  const { state, actions: a, cmd } = ctx;
  const theme = state.pres.theme;
  const pic = a.selectedShapes().find((s): s is PicShape => s.kind === "pic");
  const cropping = state.cropId === pic?.id && !!pic;
  return (
    <>
      <RibbonGroup label="Ayarla" icon={<ImageEdit24Regular />}>
        <RibbonMenuButton size="large" icon={<BrightnessHigh24Regular />} label="Düzeltmeler" disabled={!pic} menu={<AdjustGrid ctx={ctx} />} />
        <RibbonMenuButton
          size="large"
          icon={<Color24Regular />}
          label="Renk"
          disabled={!pic}
          menu={
            <Menu>
              <MenuItem label="Renk Yok (Özgün)" checked={!pic?.adjust?.grayscale} onClick={() => a.updateShapes((s) => (s.kind === "pic" ? { ...s, adjust: { ...s.adjust, grayscale: undefined } } : s))} />
              <MenuItem label="Gri Tonlama" checked={!!pic?.adjust?.grayscale} onClick={() => a.updateShapes((s) => (s.kind === "pic" ? { ...s, adjust: { ...s.adjust, grayscale: true } } : s))} />
            </Menu>
          }
        />
        <RibbonMenuButton
          size="large"
          icon={<DarkTheme24Regular />}
          label="Saydamlık"
          disabled={!pic}
          menu={
            <Menu>
              {[0, 15, 30, 50, 65, 80, 95].map((v) => (
                <MenuItem key={v} label={`Saydamlık: %${v}`} checked={Math.round((1 - (pic?.adjust?.alpha ?? 1)) * 100) === v} onClick={() => a.updateShapes((s) => (s.kind === "pic" ? { ...s, adjust: { ...s.adjust, alpha: v ? 1 - v / 100 : undefined } } : s))} />
              ))}
            </Menu>
          }
        />
        <RibbonStack>
          <RibbonButton
            icon={<ImageEdit20Regular />}
            label="Resmi Değiştir"
            disabled={!pic}
            onClick={() => {
              const input = document.createElement("input");
              input.type = "file";
              input.accept = "image/*";
              input.onchange = async () => {
                const f = input.files?.[0];
                if (!f) return;
                const img = await readImageFile(f);
                a.updateShapes((s) => {
                  if (s.kind !== "pic") return s;
                  const sc = Math.min(s.w / img.w, s.h / img.h);
                  return { ...s, src: img.src, crop: undefined, w: img.w * sc, h: img.h * sc };
                });
              };
              input.click();
            }}
          />
          <RibbonButton icon={<ImageArrowCounterclockwise20Regular />} label="Resmi Sıfırla" disabled={!pic} onClick={() => a.updateShapes((s) => (s.kind === "pic" ? { ...s, adjust: undefined, crop: undefined, effects: undefined, geom: undefined, line: { fill: { type: "none" }, width: 1, dash: "solid" } } : s))} />
        </RibbonStack>
      </RibbonGroup>
      <RibbonGroup label="Resim Stilleri" icon={<Image24Regular />} onLauncher={() => cmd.togglePane("shape")} launcherTip="Resmi Biçimlendir">
        <RibbonGallery
          columns={5}
          itemWidth={50}
          itemHeight={42}
          items={PIC_STYLES.map((p, i) => ({ id: String(i), label: p.label, preview: <PicStyleIcon i={i} /> }))}
          onSelect={(id) => a.updateShapes((s) => (s.kind === "pic" ? PIC_STYLES[Number(id)].apply(s) : s))}
        />
        <RibbonStack>
          <ShapeOutlineButton a={a} theme={theme} label="Resim Kenarlığı" size="medium" />
          <ShapeEffectsMenu a={a} theme={theme} label="Resim Efektleri" size="medium" />
          <RibbonButton icon={<ImageBorder20Regular />} label="Kenarlığı Kaldır" disabled={!pic} onClick={() => a.updateShapes((s) => (s.kind === "pic" ? { ...s, line: { ...s.line, fill: { type: "none" } } } : s))} />
        </RibbonStack>
      </RibbonGroup>
      <ArrangeGroup ctx={ctx} />
      <RibbonGroup label="Boyut" icon={<ResizeLarge24Regular />} onLauncher={() => ctx.openDialog("sizePos")} launcherTip="Boyut ve Konum">
        <RibbonButton size="large" icon={<Crop24Regular />} label="Kırp" checked={cropping} disabled={!pic} description="Resmin istenmeyen kısımlarını kırpın: kenar tutamaçlarını içeri sürükleyin." onClick={() => a.store.set({ cropId: cropping ? null : (pic?.id ?? null) })} />
        <RibbonStack>
          <RibbonSpinner label="Yükseklik:" value={pic ? cm(pic.h) : 0} min={0} max={142} step={0.1} decimals={2} unit="cm" width={84} onChange={(v) => a.setSize(undefined, fromCm(v), true)} />
          <RibbonSpinner label="Genişlik:" value={pic ? cm(pic.w) : 0} min={0} max={142} step={0.1} decimals={2} unit="cm" width={84} onChange={(v) => a.setSize(fromCm(v), undefined, true)} />
        </RibbonStack>
      </RibbonGroup>
    </>
  );
}

