/** Tablo çerçevesi seçiliyken (hücre seçimi yokken) sağ tık menüsü. */
import { TableDismiss20Regular, TableInsertRow20Regular } from "@fluentui/react-icons";
import { Menu, MenuItem, MenuSeparator } from "@officehere/ui";
import { insertCol, insertRow } from "../../model/tableOps";
import type { TableShape } from "../../model/types";
import type { Actions } from "../../state/actions";
import { ClipItems, CommentItem, DeleteItem, OrderItems, SizePosItem, type MenuDeps } from "./common";

function update(a: Actions, t: TableShape, fn: (t: TableShape) => TableShape) {
  a.updateShapes((s) => (s.id === t.id && s.kind === "table" ? fn(s) : s), [t.id]);
}

export function tableMenu({ a, cmd, openDialog }: MenuDeps, t: TableShape) {
  const lastRow = t.rows.length - 1;
  const lastCol = t.cols.length - 1;
  return (
    <Menu>
      <ClipItems cmd={cmd} />
      <MenuSeparator />
      {/* Çerçeve seçiliyken satır/sütun tablonun kenarına eklenir. */}
      <MenuItem
        icon={<TableInsertRow20Regular />}
        label="Ekle"
        submenu={
          <Menu>
            <MenuItem label="Sütunları Sola Ekle" onClick={() => update(a, t, (x) => insertCol(x, 0, false))} />
            <MenuItem label="Sütunları Sağa Ekle" onClick={() => update(a, t, (x) => insertCol(x, lastCol, true))} />
            <MenuItem label="Satırları Üste Ekle" onClick={() => update(a, t, (x) => insertRow(x, 0, false))} />
            <MenuItem label="Satırları Alta Ekle" onClick={() => update(a, t, (x) => insertRow(x, lastRow, true))} />
          </Menu>
        }
      />
      <MenuItem label="Tabloyu Seç" onClick={() => a.store.set({ cellSel: { r1: 0, c1: 0, r2: lastRow, c2: lastCol }, editing: null })} />
      <MenuItem icon={<TableDismiss20Regular />} label="Tabloyu Sil" onClick={() => (a.store.set({ editing: null, cellSel: null }), a.deleteSelection())} />
      <MenuSeparator />
      <OrderItems a={a} />
      <MenuSeparator />
      <SizePosItem openDialog={openDialog} />
      <CommentItem a={a} />
      <MenuSeparator />
      <DeleteItem a={a} />
    </Menu>
  );
}
