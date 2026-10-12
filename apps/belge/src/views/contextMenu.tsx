// Sağ tık menüsü içeriği (metin, tablo, resim, bağlantı bağlamına göre).
import type { Editor } from "@tiptap/core";
import { CellSelection } from "@tiptap/pm/tables";
import {
  ClipboardPaste20Regular,
  ClipboardTextLtr20Regular,
  CommentAdd20Regular,
  Copy20Regular,
  Cut20Regular,
  Link20Regular,
  LinkDismiss20Regular,
  Open20Regular,
  TableInsertRow20Regular,
  TableInsertColumn20Regular,
  TableDeleteRow20Regular,
  TableDeleteColumn20Regular,
  TableCellsMerge20Regular,
  TableCellsSplit20Regular,
  TableSettings20Regular,
  Delete20Regular,
  TextFont20Regular,
  TextParagraph20Regular,
  ImageAltText20Regular,
  TextBulletListLtr20Regular,
  TextNumberListLtr20Regular,
} from "@fluentui/react-icons";
import { Menu, MenuItem, MenuSeparator } from "@officehere/ui";
import type { DialogKind } from "../state/context";
import { execClipboard, pasteFromClipboard } from "../ribbon/clipboard";

interface Deps {
  editor: Editor;
  openDialog: (k: DialogKind) => void;
  newComment: () => void;
  notify: (t: string) => void;
  openLink: (href: string) => void;
}

export function buildContextMenu({ editor, openDialog, newComment, notify, openLink }: Deps) {
  const sel = editor.state.selection;
  const empty = sel.empty;
  const inTable = editor.isActive("table");
  const image = editor.isActive("image");
  const link = editor.isActive("link") ? (editor.getAttributes("link").href as string) : null;
  const chain = () => editor.chain().focus();
  return (
    <Menu width={230}>
      <MenuItem icon={<Cut20Regular />} label="Kes" shortcut="Ctrl+X" disabled={empty} onClick={() => execClipboard(editor, "cut")} />
      <MenuItem icon={<Copy20Regular />} label="Kopyala" shortcut="Ctrl+C" disabled={empty} onClick={() => execClipboard(editor, "copy")} />
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
      <MenuSeparator />
      {inTable && (
        <>
          <MenuItem
            label="Ekle"
            icon={<TableInsertRow20Regular />}
            submenu={
              <Menu>
                <MenuItem icon={<TableInsertColumn20Regular />} label="Sola Sütun Ekle" onClick={() => chain().addColumnBefore().run()} />
                <MenuItem icon={<TableInsertColumn20Regular />} label="Sağa Sütun Ekle" onClick={() => chain().addColumnAfter().run()} />
                <MenuItem icon={<TableInsertRow20Regular />} label="Üste Satır Ekle" onClick={() => chain().addRowBefore().run()} />
                <MenuItem icon={<TableInsertRow20Regular />} label="Alta Satır Ekle" onClick={() => chain().addRowAfter().run()} />
              </Menu>
            }
          />
          <MenuItem
            label="Sil"
            icon={<Delete20Regular />}
            submenu={
              <Menu>
                <MenuItem icon={<TableDeleteColumn20Regular />} label="Sütunu Sil" onClick={() => chain().deleteColumn().run()} />
                <MenuItem icon={<TableDeleteRow20Regular />} label="Satırı Sil" onClick={() => chain().deleteRow().run()} />
                <MenuItem icon={<Delete20Regular />} label="Tabloyu Sil" onClick={() => chain().deleteTable().run()} />
              </Menu>
            }
          />
          <MenuItem icon={<TableCellsMerge20Regular />} label="Hücreleri Birleştir" disabled={!(sel instanceof CellSelection)} onClick={() => chain().mergeCells().run()} />
          <MenuItem icon={<TableCellsSplit20Regular />} label="Hücreleri Böl..." onClick={() => chain().splitCell().run()} />
          <MenuItem icon={<TableSettings20Regular />} label="Tablo Özellikleri..." onClick={() => openDialog("tableProps")} />
          <MenuSeparator />
        </>
      )}
      {image && (
        <>
          <MenuItem icon={<ImageAltText20Regular />} label="Alternatif Metni Görüntüle..." onClick={() => openDialog("altText")} />
          <MenuItem label="Boyut ve Konum..." onClick={() => openDialog("imageSize")} />
          <MenuSeparator />
        </>
      )}
      <MenuItem icon={<TextFont20Regular />} label="Yazı Tipi..." shortcut="Ctrl+D" onClick={() => openDialog("font")} />
      <MenuItem icon={<TextParagraph20Regular />} label="Paragraf..." onClick={() => openDialog("paragraph")} />
      {!inTable && (
        <>
          <MenuItem icon={<TextBulletListLtr20Regular />} label="Madde İşaretleri" checked={editor.isActive("bulletList")} onClick={() => chain().toggleBulletList().run()} />
          <MenuItem icon={<TextNumberListLtr20Regular />} label="Numaralandırma" checked={editor.isActive("orderedList")} onClick={() => chain().toggleOrderedList().run()} />
        </>
      )}
      <MenuSeparator />
      {link ? (
        <>
          <MenuItem icon={<Link20Regular />} label="Köprüyü Düzenle..." onClick={() => openDialog("link")} />
          <MenuItem icon={<Open20Regular />} label="Köprüyü Aç" onClick={() => openLink(link)} />
          <MenuItem icon={<LinkDismiss20Regular />} label="Köprüyü Kaldır" onClick={() => chain().extendMarkRange("link").unsetLink().run()} />
        </>
      ) : (
        <MenuItem icon={<Link20Regular />} label="Bağlantı" shortcut="Ctrl+K" onClick={() => openDialog("link")} />
      )}
      <MenuItem icon={<CommentAdd20Regular />} label="Yeni Açıklama" shortcut="Ctrl+Alt+M" onClick={newComment} />
    </Menu>
  );
}
