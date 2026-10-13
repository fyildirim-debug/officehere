/** Nesne bağlam menülerinde ortak öğeler: pano, katman sırası, gruplama, bağlantı, boyut, biçim, sil. */
import {
  ClipboardPaste20Regular,
  CommentAdd20Regular,
  Copy20Regular,
  Cut20Regular,
  Delete20Regular,
  Group20Regular,
  GroupDismiss20Regular,
  Link20Regular,
  LinkDismiss20Regular,
  LinkEdit20Regular,
  Open20Regular,
  PaintBucket20Regular,
  ResizeLarge20Regular,
} from "@fluentui/react-icons";
import { getUserName, LayerIcon, Menu, MenuItem } from "@officehere/ui";
import type { AppCommands, DialogKind } from "../../ribbon/types";
import type { Actions } from "../../state/actions";
import { currentLink } from "../../state/linkActions";

/** Menü kurucularının ortak bağımlılıkları. */
export interface MenuDeps {
  a: Actions;
  cmd: AppCommands;
  openDialog: (d: DialogKind) => void;
}

export function ClipItems({ cmd }: { cmd: AppCommands }) {
  return (
    <>
      <MenuItem icon={<Cut20Regular />} label="Kes" shortcut="Ctrl+X" onClick={cmd.cut} />
      <MenuItem icon={<Copy20Regular />} label="Kopyala" shortcut="Ctrl+C" onClick={cmd.copy} />
      <MenuItem icon={<ClipboardPaste20Regular />} label="Yapıştır" shortcut="Ctrl+V" onClick={() => cmd.paste()} />
    </>
  );
}

/** Öne Getir › ve Arkaya Gönder › alt menüleri (katman çizimleriyle). */
export function OrderItems({ a }: { a: Actions }) {
  return (
    <>
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
    </>
  );
}

/** Grupla › (tek nesnede Grupla devre dışıdır; seçimde grup varsa Grubu Çöz etkin). */
export function GroupItem({ a }: { a: Actions }) {
  const shapes = a.selectedShapes();
  return (
    <MenuItem
      icon={<Group20Regular />}
      label="Grupla"
      submenu={
        <Menu>
          <MenuItem icon={<Group20Regular />} label="Grupla" shortcut="Ctrl+G" disabled={!a.canGroup()} onClick={() => a.groupSelection()} />
          <MenuItem icon={<GroupDismiss20Regular />} label="Grubu Çöz" shortcut="Ctrl+Shift+G" disabled={!shapes.some((s) => s.kind === "group")} onClick={() => a.ungroupSelection()} />
        </Menu>
      }
    />
  );
}

/** Şeklin köprüsü: varsa Düzenle/Aç/Kaldır, yoksa Bağlantı... */
export function LinkItems({ a, cmd }: { a: Actions; cmd: AppCommands }) {
  const link = currentLink(a);
  return link ? (
    <>
      <MenuItem icon={<LinkEdit20Regular />} label="Bağlantıyı Düzenle..." onClick={cmd.insertLink} />
      <MenuItem icon={<Open20Regular />} label="Bağlantıyı Aç" onClick={cmd.openLink} />
      <MenuItem icon={<LinkDismiss20Regular />} label="Bağlantıyı Kaldır" onClick={cmd.removeLink} />
    </>
  ) : (
    <MenuItem icon={<Link20Regular />} label="Bağlantı..." shortcut="Ctrl+K" onClick={cmd.insertLink} />
  );
}

export function SizePosItem({ openDialog }: { openDialog: (d: DialogKind) => void }) {
  return <MenuItem icon={<ResizeLarge20Regular />} label="Boyut ve Konum..." onClick={() => openDialog("sizePos")} />;
}

/** Biçim bölmesini açar (açıksa kapatmaz). */
export function FormatItem({ a, label }: { a: Actions; label: string }) {
  return <MenuItem icon={<PaintBucket20Regular />} label={label} onClick={() => a.store.set({ pane: "shape" })} />;
}

export function CommentItem({ a }: { a: Actions }) {
  return <MenuItem icon={<CommentAdd20Regular />} label="Yeni Açıklama" onClick={() => a.addComment(getUserName())} />;
}

export function DeleteItem({ a }: { a: Actions }) {
  return <MenuItem icon={<Delete20Regular />} label="Sil" shortcut="Delete" onClick={() => a.deleteSelection()} />;
}
