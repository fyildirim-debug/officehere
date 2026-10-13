// Metin bağlam menüsü: pano, Yazı Tipi / Paragraf, madde işaretleri, alan güncelleme, köprü ve açıklama.
// Metin kutusunun içindeyken altta kutunun komutları da bulunur.
import { ArrowSync20Regular, Shapes20Regular, TextBulletListLtr20Regular, TextFont20Regular, TextNumberListLtr20Regular, TextParagraph20Regular } from "@fluentui/react-icons";
import { Menu, MenuItem, MenuSeparator } from "@officehere/ui";
import { useBelge } from "../../state/context";
import { toggleList } from "../../editor/listApply";
import { updateFields } from "../../editor/fields";
import { FIELD_REFRESH_EVENT } from "../../editor/nodes";
import { ClipboardItems, LinkCommentItems, type MenuDeps } from "./common";
import type { MenuTarget } from "./target";

/** Alanı Güncelleştir: resim yazısı numaraları, çapraz başvurular, alıntılar ve sayfa/tarih alanları yenilenir. */
function useUpdateField(deps: MenuDeps) {
  const { controller } = useBelge();
  return () => {
    const { editor } = deps;
    const tr = updateFields(editor.state, (p) => controller.pageOf(p));
    if (tr) editor.view.dispatch(tr);
    // Sayfa numarası alanları sayfalamadan, tarih/saat alanları şimdiki zamandan okunur.
    controller.request();
    window.dispatchEvent(new CustomEvent(FIELD_REFRESH_EVENT));
    editor.commands.focus();
  };
}

export function TextMenu({ deps, target }: { deps: MenuDeps; target: MenuTarget }) {
  const { openDialog, showTab, inBody } = useBelge();
  const { editor } = deps;
  const updateField = useUpdateField(deps);
  return (
    <Menu width={230} label="Metin">
      <ClipboardItems editor={editor} />
      <MenuSeparator />
      <MenuItem icon={<TextFont20Regular />} label="Yazı Tipi..." shortcut="Ctrl+D" onClick={() => openDialog("font")} />
      <MenuItem icon={<TextParagraph20Regular />} label="Paragraf..." onClick={() => openDialog("paragraph")} />
      <MenuItem icon={<TextBulletListLtr20Regular />} label="Madde İşaretleri" checked={editor.isActive("bulletList")} onClick={() => toggleList(editor, "bulletList")} />
      <MenuItem icon={<TextNumberListLtr20Regular />} label="Numaralandırma" checked={editor.isActive("orderedList")} onClick={() => toggleList(editor, "orderedList")} />
      {target.field && (
        <>
          <MenuSeparator />
          <MenuItem icon={<ArrowSync20Regular />} label="Alanı Güncelleştir" onClick={updateField} />
        </>
      )}
      <MenuSeparator />
      <LinkCommentItems editor={editor} openLink={deps.openLink} />
      {target.inTextBox && inBody && (
        <>
          <MenuSeparator />
          <MenuItem icon={<Shapes20Regular />} label="Şekli Biçimlendir..." onClick={() => showTab("shape")} />
        </>
      )}
    </Menu>
  );
}
