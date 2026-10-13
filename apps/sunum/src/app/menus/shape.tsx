/** Şekil (dikdörtgen, oval, serbest biçim…) sağ tık menüsü. */
import { Menu, MenuItem, MenuSeparator } from "@officehere/ui";
import { toCustomShape } from "../../model/editPoints";
import type { SpShape } from "../../model/types";
import { ClipItems, CommentItem, DeleteItem, FormatItem, GroupItem, LinkItems, OrderItems, SizePosItem, type MenuDeps } from "./common";
import { EditTextItem } from "./text";

export function shapeMenu({ a, cmd, openDialog }: MenuDeps, sp: SpShape) {
  const editingPoints = a.s.pointEdit === sp.id;
  return (
    <Menu>
      <ClipItems cmd={cmd} />
      <MenuSeparator />
      <EditTextItem a={a} sp={sp} />
      <MenuItem
        label="Noktaları Düzenle"
        checked={editingPoints}
        onClick={() => {
          if (editingPoints) return a.store.set({ pointEdit: null });
          a.updateShapes((s) => (s.kind === "sp" ? toCustomShape(s) : s), [sp.id]);
          a.store.set({ pointEdit: sp.id, editing: null });
        }}
      />
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
