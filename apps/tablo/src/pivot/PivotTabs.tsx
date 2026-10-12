/** Bağlamsal PivotTable sekmeleri: PivotTable Analiz ve Tasarım. */
import { useEffect, useState } from "react";
import {
  ArrowMove24Regular,
  ArrowSync24Regular,
  DataUsageEdit24Regular,
  Eraser24Regular,
  SelectAllOn24Regular,
  Settings20Regular,
  TableSettings24Regular,
  TextBulletListSquare24Regular,
  TextSortAscending20Regular,
  TextSortDescending20Regular,
  Tag20Regular,
  RowTriple24Regular,
  TableBottomRow24Regular,
  TableSimple24Regular,
  TableCellsSplit24Regular,
} from "@fluentui/react-icons";
import { Menu, MenuItem, RibbonButton, RibbonCheckbox, RibbonGallery, RibbonGroup, RibbonLabel, RibbonMenuButton, RibbonSplitButton, RibbonStack } from "@officehere/ui";
import type { Controller } from "../controller";
import type { PivotDef } from "../model/types";
import { valueCaption } from "./compute";
import { applyPivot, clearPivotFields, PIVOT_STYLES, readSource, refreshAll, refreshPivot, type PivotStylePreset } from "./pivotOps";

/** Etkin hücrenin hangi alana ait olduğu (Etkin Alan). */
export function activeField(ctl: Controller, def: PivotDef): { kind: "row" | "col" | "value" | "filter"; index: number; label: string } | null {
  const src = readSource(ctl.wb, def);
  if (typeof src === "string" || !def.output) return null;
  const { row, col } = ctl.sel.active;
  const out = def.output;
  const headers = src.headers;
  // Filtre satırları
  const fIdx = def.filters.findIndex((_, i) => row === out.r1 + i);
  if (fIdx >= 0 && col <= out.c1 + 1) return { kind: "filter", index: fIdx, label: headers[def.filters[fIdx].field] };
  const relCol = col - out.c1;
  const labelCols = def.layout === "compact" ? (def.rows.length ? 1 : 0) : def.rows.length;
  if (relCol < labelCols && def.rows.length) {
    const i = def.layout === "compact" ? Math.min(def.rows.length - 1, Math.max(0, ctl.wb.styleAt(ctl.sheet, row, col).indent ?? 0)) : relCol;
    return { kind: "row", index: i, label: headers[def.rows[i]] };
  }
  if (def.values.length) {
    const vi = def.values.length === 1 ? 0 : Math.max(0, (relCol - labelCols) % def.values.length);
    return { kind: "value", index: vi, label: valueCaption(def.values[vi], headers) };
  }
  if (def.cols.length) return { kind: "col", index: 0, label: headers[def.cols[0]] };
  return null;
}

export function PivotAnalyzeTab({ ctl, def }: { ctl: Controller; def: PivotDef }) {
  const sheet = ctl.sheet;
  const [name, setName] = useState(def.name);
  useEffect(() => setName(def.name), [def.id, def.name]);
  const af = activeField(ctl, def);
  const commitName = () => {
    const n = name.trim();
    if (!n || n === def.name) return setName(def.name);
    if (ctl.wb.sheets.some((s) => s.pivots.some((p) => p.id !== def.id && p.name.toLocaleUpperCase("tr-TR") === n.toLocaleUpperCase("tr-TR")))) {
      ctl.alert("Bu ada sahip bir PivotTable zaten var.");
      return setName(def.name);
    }
    applyPivot(ctl, sheet, def.id, (d) => (d.name = n), "PivotTable Adı");
  };
  const sortField = (desc: boolean) => {
    if (!af || (af.kind !== "row" && af.kind !== "col")) return;
    const field = af.kind === "row" ? def.rows[af.index] : def.cols[af.index];
    applyPivot(ctl, sheet, def.id, (d) => (d.sort = { ...(d.sort ?? {}), [field]: desc }), "Sırala");
  };
  return (
    <>
      <RibbonGroup label="PivotTable" icon={<TableSettings24Regular />}>
        <RibbonStack>
          <RibbonLabel>PivotTable Adı:</RibbonLabel>
          <input
            className="tb-ribbon-input"
            value={name}
            onChange={(e) => setName(e.target.value)}
            onBlur={commitName}
            onKeyDown={(e) => {
              if (e.key === "Enter") (commitName(), ctl.focusGrid());
              else if (e.key === "Escape") (setName(def.name), ctl.focusGrid());
            }}
            aria-label="PivotTable Adı"
          />
          <RibbonButton icon={<Settings20Regular />} label="Seçenekler" onClick={() => ctl.openDialog("pivotOptions", { pivotId: def.id })} />
        </RibbonStack>
      </RibbonGroup>
      <RibbonGroup label="Etkin Alan" icon={<Tag20Regular />}>
        <RibbonStack>
          <RibbonLabel>Etkin Alan:</RibbonLabel>
          <span className="tb-ribbon-value" aria-label="Etkin Alan">{af?.label ?? ""}</span>
          <RibbonButton
            icon={<Settings20Regular />}
            label="Alan Ayarları"
            disabled={!af}
            onClick={() => (af?.kind === "value" ? ctl.openDialog("pivotValueField", { pivotId: def.id, index: af.index }) : ctl.openDialog("pivotOptions", { pivotId: def.id }))}
          />
        </RibbonStack>
      </RibbonGroup>
      <RibbonGroup label="Sırala" icon={<TextSortAscending20Regular />}>
        <RibbonStack>
          <RibbonButton size="small" icon={<TextSortAscending20Regular />} tip="A'dan Z'ye Sırala" disabled={!af || (af.kind !== "row" && af.kind !== "col")} onClick={() => sortField(false)} />
          <RibbonButton size="small" icon={<TextSortDescending20Regular />} tip="Z'den A'ya Sırala" disabled={!af || (af.kind !== "row" && af.kind !== "col")} onClick={() => sortField(true)} />
        </RibbonStack>
      </RibbonGroup>
      <RibbonGroup label="Veri" icon={<ArrowSync24Regular />}>
        <RibbonSplitButton
          size="large"
          icon={<ArrowSync24Regular />}
          label="Yenile"
          shortcut="Alt+F5"
          description="Kaynak verideki değişiklikleri PivotTable'a yansıtır."
          onClick={() => refreshPivot(ctl, sheet, def.id)}
          menu={
            <Menu width={200}>
              <MenuItem label="Yenile" shortcut="Alt+F5" onClick={() => refreshPivot(ctl, sheet, def.id)} />
              <MenuItem label="Tümünü Yenile" shortcut="Ctrl+Alt+F5" onClick={() => refreshAll(ctl)} />
            </Menu>
          }
        />
        <RibbonButton size="large" icon={<DataUsageEdit24Regular />} label="Veri Kaynağını Değiştir" onClick={() => ctl.openDialog("pivotSource", { pivotId: def.id })} />
      </RibbonGroup>
      <RibbonGroup label="Eylemler" icon={<Eraser24Regular />}>
        <RibbonMenuButton
          size="large"
          icon={<Eraser24Regular />}
          label="Temizle"
          menu={
            <Menu width={200}>
              <MenuItem label="Tümünü Temizle" onClick={() => clearPivotFields(ctl, sheet, def.id)} />
              <MenuItem
                label="Filtreleri Temizle"
                disabled={!def.filters.some((f) => f.selected) && !Object.values(def.hidden ?? {}).some((l) => l.length)}
                onClick={() =>
                  applyPivot(ctl, sheet, def.id, (d) => {
                    d.filters = d.filters.map((f) => ({ ...f, selected: null }));
                    d.hidden = {};
                  }, "Filtreleri Temizle")
                }
              />
            </Menu>
          }
        />
        <RibbonMenuButton
          size="large"
          icon={<SelectAllOn24Regular />}
          label="Seç"
          menu={
            <Menu width={200}>
              <MenuItem label="Tüm PivotTable" disabled={!def.output} onClick={() => def.output && ctl.selectRange(def.output)} />
            </Menu>
          }
        />
        <RibbonButton size="large" icon={<ArrowMove24Regular />} label="PivotTable'ı Taşı" onClick={() => ctl.openDialog("pivotMove", { pivotId: def.id })} />
      </RibbonGroup>
      <RibbonGroup label="Göster" icon={<TextBulletListSquare24Regular />}>
        <RibbonButton
          size="large"
          icon={<TextBulletListSquare24Regular />}
          label="Alan Listesi"
          checked={ctl.showFieldList}
          onClick={() => {
            ctl.showFieldList = !ctl.showFieldList;
            ctl.emit();
          }}
        />
        <RibbonButton size="large" icon={<TableSimple24Regular />} label="Alan Üst Bilgileri" checked={def.rowHeaders && def.colHeaders} onClick={() => applyPivot(ctl, sheet, def.id, (d) => (d.rowHeaders = d.colHeaders = !(def.rowHeaders && def.colHeaders)), "Alan Üst Bilgileri")} />
      </RibbonGroup>
    </>
  );
}

export function PivotPreview({ p, size = 50 }: { p: PivotStylePreset; size?: number }) {
  const h = Math.round(size * 0.72);
  const rows = 5;
  const rh = h / rows;
  return (
    <svg width={size} height={h} viewBox={`0 0 ${size} ${h}`} aria-hidden="true">
      <rect x="0" y="0" width={size} height={h} fill="#FFFFFF" />
      <rect x="0" y="0" width={size} height={rh} fill={p.headerFill} />
      <rect x="0" y={h - rh} width={size} height={rh} fill={p.totalFill} />
      {Array.from({ length: rows + 1 }, (_, i) => (
        <line key={i} x1="0" x2={size} y1={i * rh} y2={i * rh} stroke={p.border} strokeWidth="0.6" />
      ))}
      <rect x="0.5" y="0.5" width={size - 1} height={h - 1} fill="none" stroke={p.border} />
    </svg>
  );
}

export function PivotDesignTab({ ctl, def }: { ctl: Controller; def: PivotDef }) {
  const sheet = ctl.sheet;
  const set = (patch: Partial<PivotDef>, label: string) => applyPivot(ctl, sheet, def.id, (d) => Object.assign(d, patch), label);
  const opt = (k: "rowHeaders" | "colHeaders" | "bandedRows" | "bandedCols", label: string) => <RibbonCheckbox label={label} checked={!!def[k]} onChange={(v) => set({ [k]: v }, "PivotTable Stil Seçenekleri")} />;
  return (
    <>
      <RibbonGroup label="Düzen" icon={<RowTriple24Regular />}>
        <RibbonMenuButton
          size="large"
          icon={<TableCellsSplit24Regular />}
          label="Alt Toplamlar"
          menu={
            <Menu width={300}>
              <MenuItem label="Alt Toplamları Gösterme" checked={def.subtotals === "none"} onClick={() => set({ subtotals: "none" }, "Alt Toplamlar")} />
              <MenuItem label="Tüm Alt Toplamları Grubun Altında Göster" checked={def.subtotals === "bottom"} onClick={() => set({ subtotals: "bottom" }, "Alt Toplamlar")} />
              <MenuItem label="Tüm Alt Toplamları Grubun Üstünde Göster" checked={def.subtotals === "top"} disabled={def.layout === "tabular"} onClick={() => set({ subtotals: "top" }, "Alt Toplamlar")} />
            </Menu>
          }
        />
        <RibbonMenuButton
          size="large"
          icon={<TableBottomRow24Regular />}
          label="Genel Toplamlar"
          menu={
            <Menu width={300}>
              <MenuItem label="Satırlar ve Sütunlar için Kapalı" checked={!def.grandRows && !def.grandCols} onClick={() => set({ grandRows: false, grandCols: false }, "Genel Toplamlar")} />
              <MenuItem label="Satırlar ve Sütunlar için Açık" checked={def.grandRows && def.grandCols} onClick={() => set({ grandRows: true, grandCols: true }, "Genel Toplamlar")} />
              <MenuItem label="Yalnızca Satırlar için Açık" checked={!def.grandRows && def.grandCols} onClick={() => set({ grandRows: false, grandCols: true }, "Genel Toplamlar")} />
              <MenuItem label="Yalnızca Sütunlar için Açık" checked={def.grandRows && !def.grandCols} onClick={() => set({ grandRows: true, grandCols: false }, "Genel Toplamlar")} />
            </Menu>
          }
        />
        <RibbonMenuButton
          size="large"
          icon={<TableSimple24Regular />}
          label="Rapor Düzeni"
          menu={
            <Menu width={280}>
              <MenuItem label="Sıkıştırılmış Biçimde Göster" checked={def.layout === "compact"} onClick={() => set({ layout: "compact" }, "Rapor Düzeni")} />
              <MenuItem label="Anahat Biçiminde Göster" checked={def.layout === "outline"} onClick={() => set({ layout: "outline" }, "Rapor Düzeni")} />
              <MenuItem label="Tablo Biçiminde Göster" checked={def.layout === "tabular"} onClick={() => set({ layout: "tabular", subtotals: def.subtotals === "top" ? "bottom" : def.subtotals }, "Rapor Düzeni")} />
              <MenuItem label="Tüm Öğe Etiketlerini Yinele" checked={!!def.repeatLabels} disabled={def.layout === "compact"} onClick={() => set({ repeatLabels: true }, "Rapor Düzeni")} />
              <MenuItem label="Öğe Etiketlerini Yineleme" checked={!def.repeatLabels} disabled={def.layout === "compact"} onClick={() => set({ repeatLabels: false }, "Rapor Düzeni")} />
            </Menu>
          }
        />
        <RibbonMenuButton
          size="large"
          icon={<RowTriple24Regular />}
          label="Boş Satırlar"
          menu={
            <Menu width={280}>
              <MenuItem label="Her Öğeden Sonra Boş Satır Ekle" checked={!!def.blankRows} onClick={() => set({ blankRows: true }, "Boş Satırlar")} />
              <MenuItem label="Her Öğeden Sonra Boş Satırı Kaldır" checked={!def.blankRows} onClick={() => set({ blankRows: false }, "Boş Satırlar")} />
            </Menu>
          }
        />
      </RibbonGroup>
      <RibbonGroup label="PivotTable Stil Seçenekleri" icon={<TableSettings24Regular />}>
        <RibbonStack>
          {opt("rowHeaders", "Satır Üst Bilgileri")}
          {opt("colHeaders", "Sütun Üst Bilgileri")}
        </RibbonStack>
        <RibbonStack>
          {opt("bandedRows", "Şeritli Satırlar")}
          {opt("bandedCols", "Şeritli Sütunlar")}
        </RibbonStack>
      </RibbonGroup>
      <RibbonGroup label="PivotTable Stilleri" icon={<TableSimple24Regular />}>
        <RibbonGallery items={PIVOT_STYLES.map((p) => ({ id: p.id, label: p.label, preview: <PivotPreview p={p} /> }))} selectedId={def.style} onSelect={(id) => set({ style: id }, "PivotTable Stili")} itemWidth={58} itemHeight={44} columns={6} expandedColumns={6} />
      </RibbonGroup>
    </>
  );
}

