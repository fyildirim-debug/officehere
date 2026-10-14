/** Bağlamsal Tablo Tasarımı sekmesi: Özellikler, Araçlar, Tablo Stili Seçenekleri, Tablo Stilleri. */
import { useEffect, useRef, useState } from "react";
import { Filter20Regular, TableDismiss20Regular, TableSettings24Regular, ArrowSync20Regular, Table24Regular, TableResizeRow20Regular } from "@fluentui/react-icons";
import { RibbonButton, RibbonCheckbox, RibbonGallery, RibbonGroup, RibbonLabel, RibbonStack } from "@officehere/ui";
import type { Controller } from "../controller";
import * as D from "../actions/data";
import { TABLE_STYLES } from "../model/tablestyles";
import { tableNameError } from "../model/tableName";
import type { TableDef } from "../model/types";
import { TablePreview } from "./HomeTab2";

export function TableDesignTab({ ctl, table }: { ctl: Controller; table: TableDef }) {
  const [name, setName] = useState(table.name);
  useEffect(() => setName(table.name), [table.id, table.name]);
  // Enter/Esc ile kapanırken gelen odak kaybı ikinci kez işlemesin
  const handled = useRef(false);
  const commitName = () => {
    const n = name.trim();
    if (n === table.name) {
      setName(table.name);
      return;
    }
    const err = tableNameError(ctl.wb, n, table.id);
    if (err) {
      setName(table.name);
      ctl.alert(err, "Tablo Adı");
      return;
    }
    setName(n);
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
            onFocus={(e) => {
              handled.current = false;
              e.currentTarget.select();
            }}
            onBlur={() => {
              if (!handled.current) commitName();
              handled.current = false;
            }}
            onKeyDown={(e) => {
              if (e.key !== "Enter" && e.key !== "Escape") return;
              // Enter'ın varsayılan eylemi etkin hücreye "yazarak düzenleme" olarak ulaşmasın
              e.preventDefault();
              e.stopPropagation();
              handled.current = true;
              if (e.key === "Enter") commitName();
              else setName(table.name);
              ctl.focusGrid();
            }}
            aria-label="Tablo Adı"
          />
          <RibbonButton icon={<TableResizeRow20Regular />} label="Tabloyu Yeniden Boyutlandır" description="Satır ya da sütun ekleyip çıkararak tablonun aralığını değiştirir." onClick={() => ctl.openDialog("resizeTable", { tableId: table.id })} />
        </RibbonStack>
      </RibbonGroup>
      <RibbonGroup label="Araçlar" icon={<TableDismiss20Regular />}>
        <RibbonStack>
          <RibbonButton icon={<TableDismiss20Regular />} label="Yinelenenleri Kaldır" onClick={() => (ctl.selectRange(table.ref), ctl.openDialog("removeDuplicates"))} />
          <RibbonButton icon={<Filter20Regular />} label="Dilimleyici Ekle" onClick={() => ctl.openDialog("insertSlicer")} />
          <RibbonButton
            icon={<ArrowSync20Regular />}
            label="Aralığa Dönüştür"
            onClick={() => ctl.openDialog("confirm", { message: "Tabloyu normal aralığa dönüştürmek istiyor musunuz?", onOk: () => D.convertTableToRange(ctl, table), okLabel: "Evet" })}
          />
        </RibbonStack>
      </RibbonGroup>
      <RibbonGroup label="Tablo Stili Seçenekleri" icon={<TableSettings24Regular />}>
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
      <RibbonGroup label="Tablo Stilleri" icon={<Table24Regular />}>
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
