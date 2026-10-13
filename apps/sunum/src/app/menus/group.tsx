/** Grup ve çoklu seçim sağ tık menüleri. */
import { Group20Regular, GroupDismiss20Regular } from "@fluentui/react-icons";
import { AlignIcon, Menu, MenuItem, MenuSeparator } from "@officehere/ui";
import { AlignItems } from "../../ribbon/shapeMenus";
import { ClipItems, CommentItem, DeleteItem, FormatItem, LinkItems, OrderItems, SizePosItem, type MenuDeps } from "./common";

export function groupMenu({ a, cmd, openDialog }: MenuDeps) {
  return (
    <Menu>
      <ClipItems cmd={cmd} />
      <MenuSeparator />
      <MenuItem icon={<GroupDismiss20Regular />} label="Grubu Çöz" shortcut="Ctrl+Shift+G" onClick={() => a.ungroupSelection()} />
      <MenuSeparator />
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

/** Birden çok nesne seçiliyken. */
export function multiMenu({ a, cmd }: MenuDeps) {
  return (
    <Menu>
      <ClipItems cmd={cmd} />
      <MenuSeparator />
      <MenuItem icon={<Group20Regular />} label="Grupla" shortcut="Ctrl+G" disabled={!a.canGroup()} onClick={() => a.groupSelection()} />
      <MenuItem
        icon={<AlignIcon kind="left" />}
        label="Hizala"
        submenu={
          <Menu>
            <AlignItems a={a} />
          </Menu>
        }
      />
      <MenuSeparator />
      <OrderItems a={a} />
      <MenuSeparator />
      <DeleteItem a={a} />
    </Menu>
  );
}
