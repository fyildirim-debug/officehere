// İçindekiler, Şekiller Tablosu ya da kaynakça bloğu bağlam menüsü: güncelleştir, kaldır, Kes/Kopyala.
import { ArrowSync20Regular, Delete20Regular } from "@fluentui/react-icons";
import { Menu, MenuItem, MenuSeparator } from "@officehere/ui";
import { useBelge } from "../../state/context";
import { setFieldContext } from "../../editor/fields";
import { ClipboardItems, type MenuDeps } from "./common";
import type { MenuTarget } from "./target";

/** İçindekiler, Şekiller Tablosu ya da kaynakça bloğu: güncelleştir ve kaldır. */
export function TocMenu({ deps, target }: { deps: MenuDeps; target: MenuTarget }) {
  const { updateToc } = useBelge();
  const { editor } = deps;
  const bib = target.kind === "bibliography";
  const name = bib ? "Kaynakçayı" : target.node?.attrs.kind === "tof" ? "Şekiller Tablosunu" : "İçindekileri";
  const update = () => {
    if (bib) setFieldContext(editor.view, {});
    else updateToc();
    editor.commands.focus();
  };
  const remove = () => {
    const node = editor.state.doc.nodeAt(target.pos);
    if (node) editor.view.dispatch(editor.state.tr.delete(target.pos, target.pos + node.nodeSize).scrollIntoView());
    editor.commands.focus();
  };
  return (
    <Menu width={240} label={bib ? "Kaynakça" : "İçindekiler"}>
      <MenuItem icon={<ArrowSync20Regular />} label={`${name} Güncelleştir`} onClick={update} />
      <MenuItem icon={<Delete20Regular />} label={`${name} Kaldır`} onClick={remove} />
      <MenuSeparator />
      <ClipboardItems editor={editor} paste={false} />
    </Menu>
  );
}
