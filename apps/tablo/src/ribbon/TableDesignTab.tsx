/** Bağlamsal Tablo Tasarımı sekmesi: Özellikler, Araçlar, Tablo Stili Seçenekleri, Tablo Stilleri. */
import { useEffect, useState } from "react";
import { TableDismiss20Regular, TableSettings24Regular, ArrowSync20Regular } from "@fluentui/react-icons";
import { RibbonButton, RibbonCheckbox, RibbonGallery, RibbonGroup, RibbonLabel, RibbonStack } from "@officehere/ui";
import type { Controller } from "../controller";
import * as D from "../actions/data";
import { TABLE_STYLES } from "../model/tablestyles";
import type { TableDef } from "../model/types";
import { TablePreview } from "./HomeTab2";

export function TableDesignTab({ ctl, table }: { ctl: Controller; table: TableDef }) {
  const [name, setName] = useState(table.name);
  useEffect(() => setName(table.name), [table.id, table.name]);
  const commitName = () => {
    const n = name.trim();
    if (n === table.name) return;
    const taken = ctl.wb.sheets.some((s) => s.tables.some((t) => t.id !== table.id && t.name.toLocaleUpperCase("tr-TR") === n.toLocaleUpperCase("tr-TR")));
    if (!/^[\p{L}_][\p{L}\p{N}_.]*$/u.test(n) || taken || ctl.wb.engine.getNames().some((x) => x.name.toLocaleUpperCase("tr-TR") === n.toLocaleUpperCase("tr-TR"))) {
      ctl.alert("Girdiğiniz ad geçerli değil ya da zaten kullanılıyor.");
      setName(table.name);
      return;
    }
    D.updateTable(ctl, table, { name: n });
  };
  const opt = (k: "headerRow" | "totalRow" | "bandedRows" | "firstColumn" | "lastColumn" | "bandedCols" | "filterButton", label: string) => (
    <RibbonCheckbox label={label} checked={!!table[k]} disabled={k === "filterButton" && !table.headerRow} onChange={(v) => D.updateTable(ctl, table, { [k]: v })} />
  );
  return (
    <>
      <RibbonGroup label="Özellikler" icon={<TableSettings24Regular />}>
        <RibbonStack>
          <RibbonLabel>Tablo Adı:</RibbonLabel>
          <input
            className="tb-ribbon-input"
            value={name}
            onChange={(e) => setName(e.target.value)}
            onBlur={commitName}
            onKeyDown={(e) => {
              if (e.key === "Enter") {
                commitName();
                ctl.focusGrid();
              } else if (e.key === "Escape") {
                setName(table.name);
                ctl.focusGrid();
              }
            }}
            aria-label="Tablo Adı"
          />
        </RibbonStack>
      </RibbonGroup>
      <RibbonGroup label="Araçlar" icon={<TableDismiss20Regular />}>
        <RibbonStack>
          <RibbonButton icon={<TableDismiss20Regular />} label="Yinelenenleri Kaldır" onClick={() => (ctl.selectRange(table.ref), ctl.openDialog("removeDuplicates"))} />
          <RibbonButton
            icon={<ArrowSync20Regular />}
            label="Aralığa Dönüştür"
            onClick={() => ctl.openDialog("confirm", { message: "Tabloyu normal aralığa dönüştürmek istiyor musunuz?", onOk: () => D.convertTableToRange(ctl, table), okLabel: "Evet" })}
          />
        </RibbonStack>
      </RibbonGroup>
      <RibbonGroup label="Tablo Stili Seçenekleri">
        <RibbonStack>
          {opt("headerRow", "Üst Bilgi Satırı")}
          {opt("totalRow", "Toplam Satırı")}
          {opt("bandedRows", "Şeritli Satırlar")}
        </RibbonStack>
        <RibbonStack>
          {opt("firstColumn", "İlk Sütun")}
          {opt("lastColumn", "Son Sütun")}
          {opt("bandedCols", "Şeritli Sütunlar")}
        </RibbonStack>
        <RibbonStack>{opt("filterButton", "Filtre Düğmesi")}</RibbonStack>
      </RibbonGroup>
      <RibbonGroup label="Tablo Stilleri">
        <RibbonGallery
          items={TABLE_STYLES.map((p) => ({ id: p.id, label: p.label, preview: <TablePreview p={p} size={50} /> }))}
          selectedId={table.style}
          onSelect={(id) => D.updateTable(ctl, table, { style: id })}
          itemWidth={58}
          itemHeight={44}
          columns={7}
          expandedColumns={7}
        />
      </RibbonGroup>
    </>
  );
}
