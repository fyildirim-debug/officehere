// Konum menüsü: "Metinle Aynı Hizada" ve 3×3 "Metin Kaydırmalı" sayfa önizlemesi galerisi.
// Menü yalnız açıkken kurulur; seçili öğe resmin yerleşiminden ve sayfadaki yerinden okunur.
import { useMemo } from "react";
import { Menu, MenuGallery, MenuHeader, PositionPreview } from "@officehere/ui";
import { useBelge } from "../state/context";
import { POSITION_PRESETS, nearestVertical, positionPresetId, selectedImage, setPosition, setWrap, type PositionPreset, type WrapAttrs } from "../editor/imageWrap";

const TILE = { itemWidth: 54, itemHeight: 62 };

export function PositionMenu() {
  const ctx = useBelge();
  const { editor } = ctx;
  const img = selectedImage(editor);

  const pageBox = () => {
    if (!img) return null;
    const info = ctx.controller.info;
    const page = info.pages[ctx.controller.pageOf(img.pos)] ?? info.pages[0];
    const pagesEl = editor.view.dom.closest(".bl-pages") as HTMLElement | null;
    return page ? { page, top: pagesEl?.getBoundingClientRect().top ?? 0, scale: ctx.view.zoom / 100 } : null;
  };

  // Açılışta bir kez: resim Konum'un bıraktığı yerde mi (sayfanın üstü / ortası / altı)?
  const selectedId = useMemo(() => {
    if (!img) return null;
    const attrs = img.node.attrs as Partial<WrapAttrs>;
    let v: PositionPreset["v"] | null = null;
    const box = pageBox();
    const host = editor.view.nodeDOM(img.pos) as HTMLElement | null;
    const el = host?.querySelector("img") ?? host;
    if (box && el && (attrs.wrap ?? "inline") !== "inline") {
      const r = el.getBoundingClientRect();
      const { page, top, scale } = box;
      v = nearestVertical(r, { top: top + (page.top + page.mt) * scale, bottom: top + (page.top + page.h - page.mb) * scale });
    }
    return positionPresetId(attrs, v);
  }, []); // eslint-disable-line react-hooks/exhaustive-deps

  const toPosition = (id: string) => {
    const preset = POSITION_PRESETS.find((p) => p.id === id);
    const box = pageBox();
    if (!preset || !box) return;
    setPosition(editor, preset, box.page, box.top, box.scale);
  };

  return (
    <Menu label="Konum">
      <MenuHeader>Metinle Aynı Hizada</MenuHeader>
      <MenuGallery
        label="Metinle Aynı Hizada"
        columns={3}
        {...TILE}
        selectedId={selectedId}
        onSelect={() => setWrap(editor, "inline")}
        items={[{ id: "inline", label: "Metinle Aynı Hizada", description: "Resim bir harf gibi satırın içinde durur.", preview: <PositionPreview pos="inline" /> }]}
      />
      <MenuHeader>Metin Kaydırmalı</MenuHeader>
      <MenuGallery
        label="Metin Kaydırmalı"
        columns={3}
        {...TILE}
        selectedId={selectedId}
        onSelect={toPosition}
        items={POSITION_PRESETS.map((p) => ({ id: p.id, label: p.label, tip: p.short, description: p.wrapLabel, preview: <PositionPreview pos={{ v: p.v, h: p.h }} /> }))}
      />
    </Menu>
  );
}
