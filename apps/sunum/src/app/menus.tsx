/** Sağ tık menüleri: şekil, slayt alanı ve küçük resim bölmesi. */
import {
  ArrowReset20Regular,
  ClipboardPaste20Regular,
  Copy20Regular,
  Cut20Regular,
  Delete20Regular,
  Edit20Regular,
  EyeOff20Regular,
  Group20Regular,
  PaintBucket20Regular,
  SlideAdd20Regular,
  SlideLayout20Regular,
  SlideMultiple20Regular,
  Ruler20Regular,
  Link20Regular,
  LinkDismiss20Regular,
  LinkEdit20Regular,
  Open20Regular,
  Rename20Regular,
  ArrowUp20Regular,
  ArrowDown20Regular,
  ArrowCollapseAll20Regular,
  ArrowExpandAll20Regular,
} from "@fluentui/react-icons";
import { LayerIcon, Menu, MenuItem, MenuSeparator } from "@officehere/ui";
import { LayoutGallery } from "../ribbon/common";
import type { Actions } from "../state/actions";
import type { AppCommands } from "../ribbon/types";
import { currentLink } from "../state/linkActions";
import { sectionsOf } from "../model/sections";
import { MASTER_SLIDE_ID } from "../model/masterView";
import { collapseAllSections, moveSectionCmd, removeAllSectionsCmd, removeSectionCmd } from "../state/sectionActions";
import { toCustomShape } from "../model/editPoints";
import { isLineGeom } from "../render/geometry";


export function shapeMenu(a: Actions, cmd: AppCommands, openDialog: (d: "sizePos") => void) {
  const shapes = a.selectedShapes();
  const sp = shapes.length === 1 && shapes[0].kind === "sp" ? shapes[0] : undefined;
  const isPic = shapes.length === 1 && shapes[0].kind === "pic";
  const link = currentLink(a);
  return (
    <Menu>
      <MenuItem icon={<Cut20Regular />} label="Kes" shortcut="Ctrl+X" onClick={cmd.cut} />
      <MenuItem icon={<Copy20Regular />} label="Kopyala" shortcut="Ctrl+C" onClick={cmd.copy} />
      <MenuItem icon={<ClipboardPaste20Regular />} label="Yapıştır" shortcut="Ctrl+V" onClick={() => cmd.paste()} />
      <MenuSeparator />
      {sp && <MenuItem icon={<Edit20Regular />} label="Metni Düzenle" onClick={() => a.startEditing({ shapeId: sp.id, selectAll: true })} />}
      {sp && !sp.ink && !isLineGeom(sp.geom) && (
        <MenuItem
          label="Noktaları Düzenle"
          checked={a.s.pointEdit === sp.id}
          onClick={() => {
            a.updateShapes((s) => (s.kind === "sp" ? toCustomShape(s) : s), [sp.id]);
            a.store.set({ pointEdit: sp.id, editing: null });
          }}
        />
      )}
      {shapes.length === 1 && shapes[0].kind === "pic" && shapes[0].zoom && (
        <MenuItem
          label="Yakınlaştırmaya Dön"
          checked={!!shapes[0].zoom.ret}
          onClick={() => a.updateShapes((s) => (s.kind === "pic" && s.zoom ? { ...s, zoom: { ...s.zoom, ret: s.zoom.ret ? undefined : true } } : s), [shapes[0].id])}
        />
      )}
      {shapes.length === 1 && shapes[0].kind === "chart" && (
        <>
          <MenuItem label="Verileri Düzenle..." onClick={cmd.editChartData} />
          <MenuItem label="Grafik Türünü Değiştir..." onClick={cmd.changeChartType} />
        </>
      )}
      <MenuItem
        icon={<Group20Regular />}
        label="Grupla"
        submenu={
          <Menu>
            <MenuItem label="Grupla" disabled={!a.canGroup()} onClick={() => a.groupSelection()} />
            <MenuItem label="Grubu Çöz" disabled={!shapes.some((s) => s.kind === "group")} onClick={() => a.ungroupSelection()} />
          </Menu>
        }
      />
      <MenuItem
        icon={<LayerIcon kind="forward" />}
        label="Öne Getir"
        submenu={
          <Menu>
            <MenuItem icon={<LayerIcon kind="front" />} label="En Öne Getir" onClick={() => a.reorderSelection("front")} />
            <MenuItem icon={<LayerIcon kind="forward" />} label="Öne Getir" onClick={() => a.reorderSelection("forward")} />
          </Menu>
        }
      />
      <MenuItem
        icon={<LayerIcon kind="backward" />}
        label="Arkaya Gönder"
        submenu={
          <Menu>
            <MenuItem icon={<LayerIcon kind="back" />} label="En Arkaya Gönder" onClick={() => a.reorderSelection("back")} />
            <MenuItem icon={<LayerIcon kind="backward" />} label="Arkaya Gönder" onClick={() => a.reorderSelection("backward")} />
          </Menu>
        }
      />
      <MenuSeparator />
      {link ? (
        <>
          <MenuItem icon={<LinkEdit20Regular />} label="Bağlantıyı Düzenle..." onClick={cmd.insertLink} />
          <MenuItem icon={<Open20Regular />} label="Bağlantıyı Aç" onClick={cmd.openLink} />
          <MenuItem icon={<LinkDismiss20Regular />} label="Bağlantıyı Kaldır" onClick={cmd.removeLink} />
        </>
      ) : (
        <MenuItem icon={<Link20Regular />} label="Bağlantı..." shortcut="Ctrl+K" onClick={cmd.insertLink} />
      )}
      <MenuSeparator />
      <MenuItem label="Boyut ve Konum..." onClick={() => openDialog("sizePos")} />
      <MenuItem icon={<PaintBucket20Regular />} label={isPic ? "Resmi Biçimlendir..." : shapes.length === 1 && shapes[0].kind === "chart" ? "Grafik Alanını Biçimlendir..." : "Şekli Biçimlendir..."} onClick={() => cmd.togglePane("shape")} />
      <MenuSeparator />
      <MenuItem icon={<Delete20Regular />} label="Sil" shortcut="Delete" onClick={() => a.deleteSelection()} />
    </Menu>
  );
}

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
