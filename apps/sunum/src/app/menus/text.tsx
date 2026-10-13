/** Metin kutusu / yer tutucu sağ tık menüsü (metin düzenlenmiyorken). */
import { Edit20Regular, TextFont20Regular, TextParagraph20Regular } from "@fluentui/react-icons";
import { Menu, MenuItem, MenuSeparator } from "@officehere/ui";
import type { SpShape } from "../../model/types";
import type { Actions } from "../../state/actions";
import { ClipItems, CommentItem, DeleteItem, FormatItem, GroupItem, LinkItems, OrderItems, SizePosItem, type MenuDeps } from "./common";

export function EditTextItem({ a, sp }: { a: Actions; sp: SpShape }) {
  return <MenuItem icon={<Edit20Regular />} label="Metni Düzenle" onClick={() => a.startEditing({ shapeId: sp.id, selectAll: true })} />;
}

export function textBoxMenu({ a, cmd, openDialog }: MenuDeps, sp: SpShape) {
  return (
    <Menu>
      <ClipItems cmd={cmd} />
      <MenuSeparator />
      <EditTextItem a={a} sp={sp} />
      <MenuItem icon={<TextFont20Regular />} label="Yazı Tipi..." onClick={() => openDialog("font")} />
      <MenuItem icon={<TextParagraph20Regular />} label="Paragraf..." onClick={() => openDialog("paragraph")} />
      <MenuSeparator />
      {!sp.ph && <GroupItem a={a} />}
      <OrderItems a={a} />
      <MenuSeparator />
      <LinkItems a={a} cmd={cmd} />
      <MenuSeparator />
      <SizePosItem openDialog={openDialog} />
      <FormatItem a={a} label="Şekli Biçimlendir..." />
      <CommentItem a={a} />
      <MenuSeparator />
      <DeleteItem a={a} />
    </Menu>
  );
}
