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
  PositionBackward20Regular,
  PositionForward20Regular,
  SlideAdd20Regular,
  SlideLayout20Regular,
  SlideMultiple20Regular,
  Ruler20Regular,
  Link20Regular,
  LinkDismiss20Regular,
  LinkEdit20Regular,
  Open20Regular,
} from "@fluentui/react-icons";
import { Menu, MenuItem, MenuSeparator } from "@officehere/ui";
import { LayoutGallery } from "../ribbon/common";
import type { Actions } from "../state/actions";
import type { AppCommands } from "../ribbon/types";
import { currentLink } from "../state/linkActions";


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
      <MenuItem
        icon={<Group20Regular />}
        label="Grupla"
        submenu={
          <Menu>
            <MenuItem label="Grupla" disabled={shapes.length < 2} onClick={() => a.groupSelection()} />
            <MenuItem label="Grubu Çöz" disabled={!shapes.some((s) => s.kind === "group")} onClick={() => a.ungroupSelection()} />
          </Menu>
        }
      />
      <MenuItem
        icon={<PositionForward20Regular />}
        label="Öne Getir"
        submenu={
          <Menu>
            <MenuItem label="En Öne Getir" onClick={() => a.reorderSelection("front")} />
            <MenuItem label="Öne Getir" onClick={() => a.reorderSelection("forward")} />
          </Menu>
        }
      />
      <MenuItem
        icon={<PositionBackward20Regular />}
        label="Arkaya Gönder"
        submenu={
          <Menu>
            <MenuItem label="En Arkaya Gönder" onClick={() => a.reorderSelection("back")} />
            <MenuItem label="Arkaya Gönder" onClick={() => a.reorderSelection("backward")} />
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
      <MenuItem icon={<PaintBucket20Regular />} label={isPic ? "Resmi Biçimlendir..." : "Şekli Biçimlendir..."} onClick={() => cmd.togglePane("shape")} />
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
      <MenuItem icon={<SlideLayout20Regular />} label="Düzen" disabled={!slide} submenu={<LayoutGallery pres={s.pres} selectedId={slide?.layoutId} onPick={(id) => a.setLayout(id)} />} />
      <MenuItem icon={<ArrowReset20Regular />} label="Slaytı Sıfırla" disabled={!slide} onClick={() => a.resetSlides()} />
      <MenuSeparator />
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
        </>
      )}
    </Menu>
  );
}
