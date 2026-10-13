/** Grafik sağ tık menüsü. */
import { ChartMultiple20Regular, TableEdit20Regular } from "@fluentui/react-icons";
import { Menu, MenuItem, MenuSeparator } from "@officehere/ui";
import { ClipItems, CommentItem, DeleteItem, FormatItem, OrderItems, SizePosItem, type MenuDeps } from "./common";

export function chartMenu({ a, cmd, openDialog }: MenuDeps) {
  return (
    <Menu>
      <ClipItems cmd={cmd} />
      <MenuSeparator />
      <MenuItem icon={<TableEdit20Regular />} label="Verileri Düzenle..." onClick={cmd.editChartData} />
      <MenuItem icon={<ChartMultiple20Regular />} label="Grafik Türünü Değiştir..." onClick={cmd.changeChartType} />
      <MenuSeparator />
      <OrderItems a={a} />
      <MenuSeparator />
      <SizePosItem openDialog={openDialog} />
      <FormatItem a={a} label="Grafik Alanını Biçimlendir..." />
      <CommentItem a={a} />
      <MenuSeparator />
      <DeleteItem a={a} />
    </Menu>
  );
}
