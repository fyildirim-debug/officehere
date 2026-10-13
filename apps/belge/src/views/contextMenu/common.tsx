// Bağlam menülerinin ortak parçaları: pano komutları, bağlantı/açıklama, Metni Kaydır ve katman sırası alt menüleri.
import type { Editor } from "@tiptap/core";
import {
  ClipboardPaste20Regular,
  ClipboardTextLtr20Regular,
  CommentAdd20Regular,
  Copy20Regular,
  Cut20Regular,
  Delete20Regular,
  Link20Regular,
  LinkDismiss20Regular,
  Open20Regular,
} from "@fluentui/react-icons";
import { LayerIcon, Menu, MenuItem, MenuSeparator, WrapIcon } from "@officehere/ui";
import { useBelge } from "../../state/context";
import { execClipboard, pasteFromClipboard } from "../../ribbon/clipboard";
import { WRAP_LABEL, arrange, setWrap, type WrapMode } from "../../editor/imageWrap";

export interface MenuDeps {
  /** Menünün hedef düzenleyicisi (gövde, üst/alt bilgi ya da not). */
  editor: Editor;
  openLink: (href: string) => void;
  /** Resmi Değiştir: dosya seçicisini açar, seçilen dosya seçili resmin yerine geçer. */
  replaceImage: () => void;
}

/** Alt menülerdeki yerleşim çizimlerinin boyutu (px). */
export const ARRANGE_ICON = 24;

/** Kes, Kopyala ve (istenirse) Yapıştırma Seçenekleri. */
export function ClipboardItems({ editor, paste = true }: { editor: Editor; paste?: boolean }) {
  const { notify } = useBelge();
  const empty = editor.state.selection.empty;
  return (
    <>
      <MenuItem icon={<Cut20Regular />} label="Kes" shortcut="Ctrl+X" disabled={empty} onClick={() => execClipboard(editor, "cut")} />
      <MenuItem icon={<Copy20Regular />} label="Kopyala" shortcut="Ctrl+C" disabled={empty} onClick={() => execClipboard(editor, "copy")} />
      {paste && (
        <MenuItem
          icon={<ClipboardPaste20Regular />}
          label="Yapıştırma Seçenekleri:"
          submenu={
            <Menu>
              <MenuItem icon={<ClipboardPaste20Regular />} label="Kaynak Biçimlendirmesini Koru (K)" onClick={() => pasteFromClipboard(editor, "keep", notify)} />
              <MenuItem label="Biçimlendirmeyi Birleştir (B)" onClick={() => pasteFromClipboard(editor, "merge", notify)} />
              <MenuItem icon={<ClipboardTextLtr20Regular />} label="Yalnızca Metni Koru (M)" onClick={() => pasteFromClipboard(editor, "text", notify)} />
            </Menu>
          }
        />
      )}
    </>
  );
}

/** Bağlantı (ya da köprü komutları) ve Yeni Açıklama. */
export function LinkCommentItems({ editor, openLink }: Pick<MenuDeps, "editor" | "openLink">) {
  const { openDialog, newComment, inBody } = useBelge();
  const link = editor.isActive("link") ? (editor.getAttributes("link").href as string) : null;
  return (
    <>
      {link ? (
        <>
          <MenuItem icon={<Link20Regular />} label="Köprüyü Düzenle..." onClick={() => openDialog("link")} />
          <MenuItem icon={<Open20Regular />} label="Köprüyü Aç" onClick={() => openLink(link)} />
          <MenuItem icon={<LinkDismiss20Regular />} label="Köprüyü Kaldır" onClick={() => editor.chain().focus().extendMarkRange("link").unsetLink().run()} />
        </>
      ) : (
        <MenuItem icon={<Link20Regular />} label="Bağlantı" shortcut="Ctrl+K" onClick={() => openDialog("link")} />
      )}
      {inBody && <MenuItem icon={<CommentAdd20Regular />} label="Yeni Açıklama" shortcut="Ctrl+Alt+M" onClick={newComment} />}
    </>
  );
}

/** Seçili nesneyi siler. */
export function DeleteItem({ editor, label = "Sil" }: { editor: Editor; label?: string }) {
  return <MenuItem icon={<Delete20Regular />} label={label} shortcut="Del" onClick={() => editor.chain().focus().deleteSelection().run()} />;
}

/** Metni Kaydır alt menüsü (seçili resim). */
export function WrapItem({ editor, wrap }: { editor: Editor; wrap: WrapMode }) {
  return (
    <MenuItem
      icon={<WrapIcon mode={wrap} />}
      label="Metni Kaydır"
      submenu={
        <Menu width={220} label="Metni Kaydır">
          {(Object.keys(WRAP_LABEL) as WrapMode[]).map((w) => (
            <MenuItem key={w} icon={<WrapIcon mode={w} size={ARRANGE_ICON} />} label={WRAP_LABEL[w]} checked={wrap === w} onClick={() => setWrap(editor, w)} />
          ))}
        </Menu>
      }
    />
  );
}

/** Öne Getir / Arkaya Gönder alt menüleri (seçili resim). Satır içi resimde yalnız metnin önüne/arkasına alma etkindir. */
export function LayerItems({ editor, wrap }: { editor: Editor; wrap: WrapMode }) {
  const inline = wrap === "inline";
  return (
    <>
      <MenuItem
        icon={<LayerIcon kind="forward" />}
        label="Öne Getir"
        submenu={
          <Menu label="Öne Getir">
            <MenuItem icon={<LayerIcon kind="front" size={ARRANGE_ICON} />} label="En Öne Getir" disabled={inline} onClick={() => arrange(editor, "front")} />
            <MenuItem icon={<LayerIcon kind="forward" size={ARRANGE_ICON} />} label="Öne Getir" disabled={inline} onClick={() => arrange(editor, "forward")} />
            <MenuSeparator />
            <MenuItem icon={<WrapIcon mode="front" size={ARRANGE_ICON} />} label="Metnin Önüne Getir" onClick={() => setWrap(editor, "front")} />
          </Menu>
        }
      />
      <MenuItem
        icon={<LayerIcon kind="backward" />}
        label="Arkaya Gönder"
        submenu={
          <Menu label="Arkaya Gönder">
            <MenuItem icon={<LayerIcon kind="back" size={ARRANGE_ICON} />} label="En Arkaya Gönder" disabled={inline} onClick={() => arrange(editor, "back")} />
            <MenuItem icon={<LayerIcon kind="backward" size={ARRANGE_ICON} />} label="Arkaya Gönder" disabled={inline} onClick={() => arrange(editor, "backward")} />
            <MenuSeparator />
            <MenuItem icon={<WrapIcon mode="behind" size={ARRANGE_ICON} />} label="Metnin Arkasına Gönder" onClick={() => setWrap(editor, "behind")} />
          </Menu>
        }
      />
    </>
  );
}
