/** Slayt alanı, küçük resim bölmesi, bölüm başlığı ve Asıl görünüm sağ tık menüleri. */
import {
  ArrowReset20Regular,
  ClipboardPaste20Regular,
  Copy20Regular,
  Cut20Regular,
  Delete20Regular,
  EyeOff20Regular,
  PaintBucket20Regular,
  SlideAdd20Regular,
  SlideLayout20Regular,
  SlideMultiple20Regular,
  Ruler20Regular,
  Rename20Regular,
  ArrowUp20Regular,
  ArrowDown20Regular,
  ArrowCollapseAll20Regular,
  ArrowExpandAll20Regular,
} from "@fluentui/react-icons";
import { Menu, MenuItem, MenuSeparator } from "@officehere/ui";
import { LayoutGallery } from "../../ribbon/common";
import type { Actions } from "../../state/actions";
import type { AppCommands } from "../../ribbon/types";
import { sectionsOf } from "../../model/sections";
import { MASTER_SLIDE_ID } from "../../model/masterView";
import { collapseAllSections, moveSectionCmd, removeAllSectionsCmd, removeSectionCmd } from "../../state/sectionActions";

export function slideAreaMenu(a: Actions, cmd: AppCommands, toggleRuler: () => void, rulerOn: boolean) {
  const s = a.s;
  const slide = s.pres.slides[s.slideIdx];
  return (
    <Menu>
      <MenuItem icon={<ClipboardPaste20Regular />} label="Yapıştır" shortcut="Ctrl+V" onClick={() => cmd.paste()} />
      <MenuSeparator />
      {s.view !== "master" && (
        <>
          <MenuItem icon={<SlideLayout20Regular />} label="Düzen" disabled={!slide} submenu={<LayoutGallery pres={s.pres} selectedId={slide?.layoutId} onPick={(id) => a.setLayout(id)} />} />
          <MenuItem icon={<ArrowReset20Regular />} label="Slaytı Sıfırla" disabled={!slide} onClick={() => a.resetSlides()} />
          <MenuSeparator />
        </>
      )}
      <MenuItem icon={<Ruler20Regular />} label="Cetvel" checked={rulerOn} onClick={toggleRuler} />
      <MenuItem label="Kılavuz Çizgileri" checked={s.showGrid} onClick={() => a.store.set({ showGrid: !s.showGrid })} />
      <MenuItem label="Kılavuzlar" checked={s.showGuides} onClick={() => a.store.set({ showGuides: !s.showGuides })} />
      <MenuSeparator />
      <MenuItem icon={<PaintBucket20Regular />} label="Arka Planı Biçimlendir..." onClick={() => cmd.togglePane("background")} />
    </Menu>
  );
}

export function thumbMenu(a: Actions, cmd: AppCommands, onSlide: boolean) {
  const s = a.s;
  const slide = s.pres.slides[s.slideIdx];
  const hidden = !!slide?.hidden;
  return (
    <Menu>
      {onSlide && (
        <>
          <MenuItem icon={<Cut20Regular />} label="Kes" shortcut="Ctrl+X" onClick={() => a.copySlides(true)} />
          <MenuItem icon={<Copy20Regular />} label="Kopyala" shortcut="Ctrl+C" onClick={() => a.copySlides(false)} />
        </>
      )}
      <MenuItem icon={<ClipboardPaste20Regular />} label="Yapıştır" shortcut="Ctrl+V" onClick={() => a.pasteInternal()} />
      <MenuSeparator />
      <MenuItem icon={<SlideAdd20Regular />} label="Yeni Slayt" shortcut="Ctrl+M" onClick={() => a.newSlide()} />
      {onSlide && (
        <>
          <MenuItem icon={<SlideMultiple20Regular />} label="Slaytı Çoğalt" shortcut="Ctrl+D" onClick={() => a.duplicateSlides()} />
          <MenuItem icon={<Delete20Regular />} label="Slaytı Sil" shortcut="Delete" onClick={() => a.deleteSlides()} />
          <MenuSeparator />
          <MenuItem icon={<SlideLayout20Regular />} label="Düzen" submenu={<LayoutGallery pres={s.pres} selectedId={slide?.layoutId} onPick={(id) => a.setLayout(id)} />} />
          <MenuItem icon={<ArrowReset20Regular />} label="Slaytı Sıfırla" onClick={() => a.resetSlides()} />
          <MenuItem icon={<PaintBucket20Regular />} label="Arka Planı Biçimlendir..." onClick={() => cmd.togglePane("background")} />
          <MenuItem icon={<EyeOff20Regular />} label={hidden ? "Slaytı Göster" : "Slaytı Gizle"} checked={hidden} onClick={() => a.toggleHideSlides()} />
          <MenuSeparator />
          <MenuItem label="Bölüm Ekle" onClick={cmd.addSection} />
        </>
      )}
    </Menu>
  );
}

/** Bölüm başlığı sağ tık menüsü ve Giriş > Bölüm menüsü. */
export function sectionMenu(a: Actions, cmd: AppCommands, id: string | undefined) {
  const secs = sectionsOf(a.s.pres.slides);
  const i = id ? secs.findIndex((s) => s.id === id) : -1;
  return (
    <Menu>
      <MenuItem icon={<SlideAdd20Regular />} label="Bölüm Ekle" disabled={!a.s.pres.slides.length} onClick={cmd.addSection} />
      <MenuItem icon={<Rename20Regular />} label="Bölümü Yeniden Adlandır..." disabled={i < 0} onClick={() => id && cmd.renameSection(id)} />
      <MenuItem icon={<Delete20Regular />} label="Bölümü Kaldır" disabled={i < 0} onClick={() => id && removeSectionCmd(a, id)} />
      <MenuItem label="Bölümü ve Slaytları Kaldır" disabled={i < 0} onClick={() => id && removeSectionCmd(a, id, true)} />
      <MenuItem label="Tüm Bölümleri Kaldır" disabled={!secs.length} onClick={() => removeAllSectionsCmd(a)} />
      <MenuSeparator />
      <MenuItem icon={<ArrowUp20Regular />} label="Bölümü Yukarı Taşı" disabled={i <= 0} onClick={() => id && moveSectionCmd(a, id, -1)} />
      <MenuItem icon={<ArrowDown20Regular />} label="Bölümü Aşağı Taşı" disabled={i < 0 || i >= secs.length - 1} onClick={() => id && moveSectionCmd(a, id, 1)} />
      <MenuSeparator />
      <MenuItem icon={<ArrowCollapseAll20Regular />} label="Tümünü Daralt" disabled={!secs.length} onClick={() => collapseAllSections(a, true)} />
      <MenuItem icon={<ArrowExpandAll20Regular />} label="Tümünü Genişlet" disabled={!secs.length} onClick={() => collapseAllSections(a, false)} />
    </Menu>
  );
}

/** Asıl görünüm bölmesi sağ tık menüsü. */
export function masterMenu(a: Actions, cmd: AppCommands) {
  const cur = a.s.pres.slides[a.s.slideIdx];
  const isMaster = cur?.id === MASTER_SLIDE_ID;
  return (
    <Menu>
      <MenuItem icon={<SlideAdd20Regular />} label="Düzen Ekle" onClick={() => a.newSlide()} />
      <MenuItem icon={<SlideMultiple20Regular />} label="Düzeni Çoğalt" disabled={isMaster} onClick={() => a.duplicateSlides()} />
      <MenuItem icon={<Delete20Regular />} label="Düzeni Sil" disabled={isMaster} onClick={() => a.deleteSlides()} />
      <MenuItem icon={<Rename20Regular />} label="Düzeni Yeniden Adlandır..." disabled={isMaster} onClick={cmd.renameLayout} />
      <MenuSeparator />
      <MenuItem icon={<PaintBucket20Regular />} label="Arka Planı Biçimlendir..." onClick={() => cmd.togglePane("background")} />
      <MenuSeparator />
      <MenuItem label="Asıl Görünümü Kapat" onClick={() => cmd.setView("normal")} />
    </Menu>
  );
}
