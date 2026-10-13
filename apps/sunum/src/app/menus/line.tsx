/** Çizgi, bağlayıcı ve mürekkep çizimi sağ tık menüsü. */
import { Menu, MenuSeparator } from "@officehere/ui";
import { ClipItems, CommentItem, DeleteItem, FormatItem, GroupItem, LinkItems, OrderItems, SizePosItem, type MenuDeps } from "./common";

export function lineMenu({ a, cmd, openDialog }: MenuDeps) {
  return (
    <Menu>
      <ClipItems cmd={cmd} />
      <MenuSeparator />
      <GroupItem a={a} />
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
