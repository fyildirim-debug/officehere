/** PivotTable hücresinde sağ tık menüsü. */
import { ArrowSortDownLines20Regular, ArrowSortUpLines20Regular, ArrowSync20Regular, Copy20Regular, Dismiss16Regular, TableEdit20Regular, TextBulletListSquare20Regular } from "@fluentui/react-icons";
import { Menu, MenuItem, MenuSeparator } from "@officehere/ui";
import type { Controller } from "../controller";
import type { PivotDef } from "../model/types";
import * as E from "../actions/edit";
import { activeField } from "./PivotTabs";
import { applyPivot, refreshPivot } from "./pivotOps";
import { removeAt } from "./fields";

export function PivotCellMenu({ ctl, def }: { ctl: Controller; def: PivotDef }) {
  const sheet = ctl.sheet;
  const af = activeField(ctl, def);
  const field = af ? (af.kind === "row" ? def.rows[af.index] : af.kind === "col" ? def.cols[af.index] : af.kind === "filter" ? def.filters[af.index].field : def.values[af.index]?.field) : undefined;
  const sortable = af && (af.kind === "row" || af.kind === "col") && field !== undefined;
  const sort = (desc: boolean) => field !== undefined && applyPivot(ctl, sheet, def.id, (d) => (d.sort = { ...(d.sort ?? {}), [field]: desc }), "Sırala");
  return (
    <Menu width={250}>
      <MenuItem
        icon={<Copy20Regular />}
        label="Kopyala"
        shortcut="Ctrl+C"
        onClick={() => {
          const d = E.copySelection(ctl, false);
          if (d) void E.writeSystemClipboard(d);
        }}
      />
      <MenuItem icon={<TableEdit20Regular />} label="Hücreleri Biçimlendir..." onClick={() => ctl.openDialog("formatCells")} />
      <MenuSeparator />
      <MenuItem icon={<ArrowSync20Regular />} label="Yenile" onClick={() => refreshPivot(ctl, sheet, def.id)} />
      <MenuItem
        label="Sırala"
        disabled={!sortable}
        submenu={
          <Menu width={220}>
            <MenuItem icon={<ArrowSortDownLines20Regular />} label="A'dan Z'ye Sırala" onClick={() => sort(false)} />
            <MenuItem icon={<ArrowSortUpLines20Regular />} label="Z'den A'ya Sırala" onClick={() => sort(true)} />
          </Menu>
        }
      />
      <MenuSeparator />
      {af && (
        <MenuItem
          icon={<Dismiss16Regular />}
          label={`"${af.label}" Alanını Kaldır`}
          onClick={() => applyPivot(ctl, sheet, def.id, (d) => removeAt(d, af.kind === "row" ? "rows" : af.kind === "col" ? "cols" : af.kind === "filter" ? "filters" : "values", af.index), "Alanı Kaldır")}
        />
      )}
      {af?.kind === "value" && <MenuItem label="Değer Alanı Ayarları..." onClick={() => ctl.openDialog("pivotValueField", { pivotId: def.id, index: af.index })} />}
      <MenuItem label="Özet Tablo Seçenekleri..." onClick={() => ctl.openDialog("pivotOptions", { pivotId: def.id })} />
      <MenuItem
        icon={<TextBulletListSquare20Regular />}
        label={ctl.showFieldList ? "Alan Listesini Gizle" : "Alan Listesini Göster"}
        onClick={() => {
          ctl.showFieldList = !ctl.showFieldList;
          ctl.emit();
        }}
      />
    </Menu>
  );
}
