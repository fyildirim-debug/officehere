// Metin kutusunun kenarına/çerçevesine sağ tık: pano, Metni Düzenle, Şekli Biçimlendir ve sil.
import { Shapes20Regular, TextEditStyle20Regular } from "@fluentui/react-icons";
import { TextSelection } from "@tiptap/pm/state";
import { Menu, MenuItem, MenuSeparator } from "@officehere/ui";
import { useBelge } from "../../state/context";
import { ClipboardItems, DeleteItem, type MenuDeps } from "./common";
import type { MenuTarget } from "./target";

/** Metin kutusunun kenarına/çerçevesine sağ tık. */
export function TextBoxMenu({ deps, target }: { deps: MenuDeps; target: MenuTarget }) {
  const { showTab, inBody } = useBelge();
  const { editor } = deps;
  // Metni Düzenle: imleç kutudaki metnin sonuna geçer.
  const editText = () => {
    const node = editor.state.doc.nodeAt(target.pos);
    if (!node) return;
    const end = target.pos + node.nodeSize - 2;
    editor.view.dispatch(editor.state.tr.setSelection(TextSelection.near(editor.state.doc.resolve(end), -1)));
    editor.view.focus();
  };
  return (
    <Menu width={230} label="Metin Kutusu">
      <ClipboardItems editor={editor} />
      <MenuSeparator />
      <MenuItem icon={<TextEditStyle20Regular />} label="Metni Düzenle" onClick={editText} />
      {inBody && <MenuItem icon={<Shapes20Regular />} label="Şekli Biçimlendir..." onClick={() => showTab("shape")} />}
      <MenuSeparator />
      <DeleteItem editor={editor} />
    </Menu>
  );
}
