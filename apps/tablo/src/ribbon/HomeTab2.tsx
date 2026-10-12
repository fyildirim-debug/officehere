/** Giriş sekmesi: Stiller, Hücreler, Düzenleme grupları. */
import {
  ArrowSortDownLines20Regular,
  ArrowSortUpLines20Regular,
  Eraser20Regular,
  Filter20Regular,
  ArrowDown20Regular,
  Search20Regular,
  TableAdd24Regular,
  TableDismiss24Regular,
  TableEdit24Regular,
  Table24Regular,
  ColorBackground24Regular,
  DataBarHorizontal24Regular,
  TableInsertRow20Regular,
  TableInsertColumn20Regular,
  TableDeleteRow20Regular,
  TableDeleteColumn20Regular,
  ArrowSort20Regular,
} from "@fluentui/react-icons";
import { Menu, MenuHeader, MenuItem, MenuSeparator, RibbonGroup, RibbonMenuButton, RibbonSplitButton, RibbonStack, ColorPalette, useMenuClose } from "@officehere/ui";
import { paletteOf } from "../model/theme";
import type { Controller } from "../controller";
import { requestDeleteSheet } from "../dialogs/MiscDialogs";
import { IconSetMenu } from "./IconSetMenu";
import * as F from "../actions/format";
import * as E from "../actions/edit";
import * as D from "../actions/data";
import { CELL_STYLES, type CellStyle } from "../model/styles";
import { TABLE_STYLES, previewColors, type TableStylePreset } from "../model/tablestyles";
import { COLOR_SCALES, DATA_BAR_COLORS } from "../model/condfmt";
import { autoFitCols, autoFitRows } from "../grid/autofit";
import { currentRegion } from "../model/selection";
import { SigmaGlyph } from "./glyphs";
import { selectFirstObject } from "../actions/objectClip";
import { deleteTableCols, deleteTableRows, insertTableCols, insertTableRows } from "../actions/tableEdit";

export function TablePreview({ p, size = 56 }: { p: TableStylePreset; size?: number }) {
  const c = previewColors(p);
  const rows = 5;
  const h = (size * 0.7) / rows;
  return (
    <svg width={size} height={size * 0.75} viewBox={`0 0 ${size} ${size * 0.75}`} aria-hidden="true">
      {Array.from({ length: rows }, (_, i) => (
        <rect key={i} x={2} y={2 + i * h} width={size - 4} height={h} fill={i === 0 ? c.header : i % 2 ? c.band : c.body} stroke={c.line} strokeWidth={0.6} />
      ))}
    </svg>
  );
}

function TableGallery({ onPick }: { onPick: (id: string) => void }) {
  const close = useMenuClose();
  const groups = ["Açık", "Orta", "Koyu"] as const;
  return (
    <div className="oh-menu tb-gallery-menu" style={{ width: 430, padding: "4px 6px" }}>
      {groups.map((g) => (
        <div key={g}>
          <MenuHeader>{g}</MenuHeader>
          <div style={{ display: "grid", gridTemplateColumns: "repeat(7, 58px)", gap: 2 }}>
            {TABLE_STYLES.filter((t) => t.group === g).map((t) => (
              <button
                key={t.id}
                type="button"
                className="oh-gallery-item"
                style={{ width: 58, height: 46 }}
                title={t.label}
                onMouseDown={(e) => e.preventDefault()}
                onClick={() => {
                  onPick(t.id);
                  close();
                }}
              >
                <TablePreview p={t} size={52} />
              </button>
            ))}
          </div>
        </div>
      ))}
    </div>
  );
}

function CellStylePreview({ s, label }: { s: CellStyle; label: string }) {
  return (
    <span
      className="tb-gallery-cellstyle"
      style={{
        background: s.fill ?? "#fff",
        color: s.color ?? "#000",
        fontWeight: s.bold ? 700 : 400,
        fontStyle: s.italic ? "italic" : undefined,
        fontSize: s.size && s.size > 12 ? 12 : 11,
        borderBottom: s.bBottom ? `${s.bBottom.style === "thick" ? 3 : s.bBottom.style === "double" ? 3 : 1}px ${s.bBottom.style === "double" ? "double" : "solid"} ${s.bBottom.color}` : undefined,
        outline: s.bTop && s.bLeft ? `1px solid ${s.bTop.color}` : undefined,
        outlineOffset: -1,
      }}
    >
      {label}
    </span>
  );
}

function CellStyleGallery({ ctl }: { ctl: Controller }) {
  const close = useMenuClose();
  const groups = [...new Set(CELL_STYLES.map((s) => s.group))];
  return (
    <div className="oh-menu" style={{ width: 690, maxWidth: "calc(100vw - 16px)", padding: "4px 6px" }}>
      {groups.map((g) => (
        <div key={g}>
          <MenuHeader>{g}</MenuHeader>
          <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fill, minmax(108px, 1fr))", gap: 3, margin: "2px 0 4px" }}>
            {CELL_STYLES.filter((s) => s.group === g).map((s) => (
              <button
                key={s.id}
                type="button"
                className="oh-gallery-item"
                style={{ width: "100%", height: 26 }}
                title={s.label}
                onMouseDown={(e) => e.preventDefault()}
                onClick={() => {
                  F.applyCellStyle(ctl, s.style);
                  close();
                }}
              >
                <CellStylePreview s={s.style} label={s.label} />
              </button>
            ))}
          </div>
        </div>
      ))}
    </div>
  );
}

function ColorBars({ ctl }: { ctl: Controller }) {
  const close = useMenuClose();
  return (
    <div className="oh-menu" style={{ padding: 6 }}>
      <MenuHeader>Dolgu Rengi</MenuHeader>
      <div style={{ display: "grid", gridTemplateColumns: "repeat(3, 44px)", gap: 4 }}>
        {DATA_BAR_COLORS.map((c) => (
          <button
            key={c}
            type="button"
            className="oh-gallery-item"
            style={{ width: 44, height: 40 }}
            onClick={() => {
              D.addCondRule(ctl, { type: "dataBar", color: c });
              close();
            }}
          >
            <svg width="34" height="30" aria-hidden="true">
              {[26, 18, 30, 12].map((w, i) => (
                <rect key={i} x={2} y={2 + i * 7} width={w} height={5} fill={c} />
              ))}
            </svg>
          </button>
        ))}
      </div>
    </div>
  );
}

function ColorScales({ ctl }: { ctl: Controller }) {
  const close = useMenuClose();
  return (
    <div className="oh-menu" style={{ padding: 6 }}>
      <div style={{ display: "grid", gridTemplateColumns: "repeat(4, 44px)", gap: 4 }}>
        {COLOR_SCALES.map((s) => (
          <button
            key={s.id}
            type="button"
            className="oh-gallery-item"
            style={{ width: 44, height: 40 }}
            title={s.label}
            onClick={() => {
              D.addCondRule(ctl, { type: "colorScale", colors: s.colors });
              close();
            }}
          >
            <svg width="30" height="30" aria-hidden="true">
              {[0, 1, 2].map((i) => (
                <rect key={i} x={2} y={2 + i * 9} width={26} height={8} fill={s.colors[s.colors.length === 2 ? Math.min(i, 1) : 2 - i] ?? s.colors[0]} />
              ))}
            </svg>
          </button>
        ))}
      </div>
    </div>
  );
}

export function StylesGroup({ ctl }: { ctl: Controller }) {
  const cf = (rule: string, arg?: unknown) => ctl.openDialog("condRule", { rule, ...(arg as object) });
  return (
    <RibbonGroup label="Stiller" icon={<ColorBackground24Regular />}>
      <RibbonMenuButton
        size="large"
        icon={<ColorBackground24Regular />}
        label="Koşullu Biçimlendirme"
        menu={
          <Menu width={250}>
            <MenuItem
              label="Hücre Vurgulama Kuralları"
              submenu={
                <Menu width={200}>
                  <MenuItem label="Büyüktür..." onClick={() => cf("greaterThan")} />
                  <MenuItem label="Küçüktür..." onClick={() => cf("lessThan")} />
                  <MenuItem label="Arasında..." onClick={() => cf("between")} />
                  <MenuItem label="Eşittir..." onClick={() => cf("equal")} />
                  <MenuItem label="İçeren Metin..." onClick={() => cf("containsText")} />
                  <MenuItem label="Gerçekleşen Tarih..." onClick={() => cf("timePeriod")} />
                  <MenuItem label="Yinelenen Değerler..." onClick={() => cf("duplicateValues")} />
                </Menu>
              }
            />
            <MenuItem
              label="Üst/Alt Kuralları"
              submenu={
                <Menu width={200}>
                  <MenuItem label="İlk 10 Öğe..." onClick={() => cf("top10")} />
                  <MenuItem label="İlk %10..." onClick={() => cf("top10Pct")} />
                  <MenuItem label="Son 10 Öğe..." onClick={() => cf("bottom10")} />
                  <MenuItem label="Son %10..." onClick={() => cf("bottom10Pct")} />
                  <MenuItem label="Ortalamanın Üstünde..." onClick={() => cf("aboveAverage")} />
                  <MenuItem label="Ortalamanın Altında..." onClick={() => cf("belowAverage")} />
                </Menu>
              }
            />
            <MenuItem icon={<DataBarHorizontal24Regular />} label="Veri Çubukları" submenu={<ColorBars ctl={ctl} />} />
            <MenuItem label="Renk Ölçekleri" submenu={<ColorScales ctl={ctl} />} />
            <MenuItem label="Simge Kümeleri" submenu={<IconSetMenu ctl={ctl} />} />
            <MenuSeparator />
            <MenuItem label="Yeni Kural..." onClick={() => cf("new")} />
            <MenuItem
              label="Kuralları Temizle"
              submenu={
                <Menu width={240}>
                  <MenuItem label="Seçili Hücrelerden Kuralları Temizle" onClick={() => D.clearCondRules(ctl, "selection")} />
                  <MenuItem label="Tüm Sayfadan Kuralları Temizle" onClick={() => D.clearCondRules(ctl, "sheet")} />
                </Menu>
              }
            />
            <MenuItem label="Kuralları Yönet..." onClick={() => ctl.openDialog("condManager")} />
          </Menu>
        }
      />
      <RibbonMenuButton size="large" icon={<Table24Regular />} label={"Tablo\u00a0Olarak Biçimlendir"} menu={<TableGallery onPick={(id) => ctl.openDialog("createTable", { style: id })} />} />
      <RibbonMenuButton size="large" icon={<CellStylesGlyph />} label="Hücre Stilleri" menu={<CellStyleGallery ctl={ctl} />} />
    </RibbonGroup>
  );
}

function CellStylesGlyph() {
  return (
    <svg width="32" height="32" viewBox="0 0 32 32" aria-hidden="true">
      <rect x="3" y="6" width="26" height="20" rx="1.5" fill="#fff" stroke="currentColor" strokeWidth="1.2" />
      <rect x="4" y="7" width="12" height="9" fill="#C6EFCE" />
      <rect x="16" y="7" width="12" height="9" fill="#FFC7CE" />
      <rect x="4" y="16" width="12" height="9" fill="#FFEB9C" />
      <path d="M3 16h26M16 6v20" stroke="currentColor" strokeWidth="1" />
    </svg>
  );
}

export function CellsGroup({ ctl }: { ctl: Controller }) {
  const sheet = ctl.sheet;
  const range = ctl.range;
  const table = sheet.tableAt(ctl.sel.active.row, ctl.sel.active.col);
  const lines = (axis: "row" | "col") => {
    const out: number[] = [];
    const [a, b] = axis === "row" ? [range.r1, Math.min(range.r2, range.r1 + 50000)] : [range.c1, range.c2];
    for (let i = a; i <= b; i++) out.push(i);
    return out;
  };
  return (
    <RibbonGroup label="Hücreler" icon={<TableAdd24Regular />}>
      <RibbonSplitButton
        size="large"
        icon={<TableAdd24Regular />}
        label="Ekle"
        onClick={() => E.insertCellsSmart(ctl)}
        menu={
          <Menu width={210}>
            <MenuItem label="Hücre Ekle..." onClick={() => ctl.openDialog("insertCells")} />
            <MenuItem icon={<TableInsertRow20Regular />} label="Sayfa Satırları Ekle" onClick={() => E.insertRows(ctl)} />
            <MenuItem icon={<TableInsertColumn20Regular />} label="Sayfa Sütunları Ekle" onClick={() => E.insertCols(ctl)} />
            {table && (
              <>
                <MenuSeparator />
                <MenuItem label="Tablo Satırları Yukarı Ekle" onClick={() => insertTableRows(ctl, table)} />
                <MenuItem label="Tablo Sütunları Sola Ekle" onClick={() => insertTableCols(ctl, table)} />
              </>
            )}
            <MenuSeparator />
            <MenuItem label="Sayfa Ekle" onClick={() => ctl.ops.addSheet(ctl.wb.activeIndex)} />
          </Menu>
        }
      />
      <RibbonSplitButton
        size="large"
        icon={<TableDismiss24Regular />}
        label="Sil"
        onClick={() => E.deleteCellsSmart(ctl)}
        menu={
          <Menu width={210}>
            <MenuItem label="Hücreleri Sil..." onClick={() => ctl.openDialog("deleteCells")} />
            <MenuItem icon={<TableDeleteRow20Regular />} label="Sayfa Satırlarını Sil" onClick={() => E.deleteRows(ctl)} />
            <MenuItem icon={<TableDeleteColumn20Regular />} label="Sayfa Sütunlarını Sil" onClick={() => E.deleteCols(ctl)} />
            {table && (
              <>
                <MenuSeparator />
                <MenuItem label="Tablo Satırlarını Sil" onClick={() => deleteTableRows(ctl, table)} />
                <MenuItem label="Tablo Sütunlarını Sil" onClick={() => deleteTableCols(ctl, table)} />
              </>
            )}
            <MenuSeparator />
            <MenuItem label="Sayfayı Sil" onClick={() => requestDeleteSheet(ctl)} />
          </Menu>
        }
      />
      <RibbonMenuButton
        size="large"
        icon={<TableEdit24Regular />}
        label="Biçim"
        menu={
          <Menu width={260}>
            <MenuHeader>Hücre Boyutu</MenuHeader>
            <MenuItem label="Satır Yüksekliği..." onClick={() => ctl.openDialog("rowHeight")} />
            <MenuItem label="Satır Yüksekliğini Otomatik Sığdır" onClick={() => autoFitRows(ctl, lines("row"))} />
            <MenuItem label="Sütun Genişliği..." onClick={() => ctl.openDialog("colWidth")} />
            <MenuItem label="Sütun Genişliğini Otomatik Sığdır" onClick={() => autoFitCols(ctl, lines("col"))} />
            <MenuItem label="Varsayılan Genişlik..." onClick={() => ctl.openDialog("defaultWidth")} />
            <MenuHeader>Görünürlük</MenuHeader>
            <MenuItem
              label="Gizle ve Göster"
              submenu={
                <Menu width={200}>
                  <MenuItem label="Satırları Gizle" shortcut="Ctrl+9" onClick={() => ctl.ops.setHidden(sheet, "row", lines("row"), true)} />
                  <MenuItem label="Sütunları Gizle" shortcut="Ctrl+0" onClick={() => ctl.ops.setHidden(sheet, "col", lines("col"), true)} />
                  <MenuItem label="Sayfayı Gizle" onClick={() => ctl.ops.setSheetHidden(ctl.wb.activeIndex, true)} />
                  <MenuSeparator />
                  <MenuItem label="Satırları Göster" shortcut="Ctrl+Shift+9" onClick={() => ctl.ops.setHidden(sheet, "row", sheet.rows.hiddenItems().filter((r) => r >= range.r1 - 1 && r <= range.r2 + 1), false)} />
                  <MenuItem label="Sütunları Göster" shortcut="Ctrl+Shift+0" onClick={() => ctl.ops.setHidden(sheet, "col", sheet.cols.hiddenItems().filter((c) => c >= range.c1 - 1 && c <= range.c2 + 1), false)} />
                  <MenuItem label="Sayfayı Göster..." disabled={!ctl.wb.sheets.some((s) => s.hidden)} onClick={() => ctl.openDialog("unhideSheet")} />
                </Menu>
              }
            />
            <MenuHeader>Sayfaları Düzenle</MenuHeader>
            <MenuItem label="Sayfayı Yeniden Adlandır" onClick={() => window.dispatchEvent(new CustomEvent("tb-rename-sheet", { detail: ctl.wb.activeIndex }))} />
            <MenuItem label="Sayfayı Taşı veya Kopyala..." onClick={() => ctl.openDialog("moveSheet")} />
            <MenuItem
              label="Sekme Rengi"
              submenu={
                <div className="oh-menu oh-menu-palette">
                  <ColorPalette value={sheet.tabColor} noneLabel="Renk Yok" theme={paletteOf(ctl.wb.theme)} onSelect={(c) => ctl.ops.setTabColor(sheet, c ?? undefined)} />
                </div>
              }
            />
            <MenuHeader>Koruma</MenuHeader>
            <MenuItem label={sheet.protected ? "Sayfanın Korumasını Kaldır..." : "Sayfayı Koru..."} onClick={() => ctl.openDialog("protect")} />
            <MenuItem label="Hücreyi Kilitle" checked={F.activeStyle(ctl).locked !== false} onClick={() => F.applyStyle(ctl, { locked: F.activeStyle(ctl).locked === false ? undefined : false })} />
            <MenuItem label="Hücreleri Biçimlendir..." shortcut="Ctrl+1" onClick={() => ctl.openDialog("formatCells")} />
          </Menu>
        }
      />
    </RibbonGroup>
  );
}

export function SortFilterMenu({ ctl }: { ctl: Controller }) {
  const hasFilter = D.filtersOf(ctl.sheet).length > 0;
  return (
    <Menu width={220}>
      <MenuItem icon={<ArrowSortDownLines20Regular />} label="A'dan Z'ye Sırala" onClick={() => D.quickSort(ctl, false)} />
      <MenuItem icon={<ArrowSortUpLines20Regular />} label="Z'den A'ya Sırala" onClick={() => D.quickSort(ctl, true)} />
      <MenuItem icon={<ArrowSort20Regular />} label="Özel Sıralama..." onClick={() => ctl.openDialog("sort")} />
      <MenuSeparator />
      <MenuItem icon={<Filter20Regular />} label="Filtre" shortcut="Ctrl+Shift+L" checked={!!ctl.sheet.autoFilter} onClick={() => D.toggleAutoFilter(ctl)} />
      <MenuItem label="Temizle" disabled={!hasFilter} onClick={() => D.clearAllFilters(ctl)} />
      <MenuItem label="Yeniden Uygula" disabled={!hasFilter} onClick={() => D.reapplyFilters(ctl)} />
    </Menu>
  );
}

export function EditingGroup({ ctl }: { ctl: Controller }) {
  const selectCells = (pred: (row: number, col: number) => boolean) => {
    const sheet = ctl.sheet;
    const ranges: { r1: number; c1: number; r2: number; c2: number }[] = [];
    ctl.wb.engine.forEachCell(sheet.id, (row, col) => {
      if (pred(row, col)) ranges.push({ r1: row, c1: col, r2: row, c2: col });
    });
    if (!ranges.length) {
      ctl.alert("Hiçbir hücre bulunamadı.");
      return;
    }
    ctl.setSelection({ ranges, active: { row: ranges[0].r1, col: ranges[0].c1 }, anchor: { row: ranges[0].r1, col: ranges[0].c1 } });
  };
  return (
    <RibbonGroup label="Düzenleme" icon={<Search20Regular />}>
      <RibbonStack>
        <RibbonSplitButton
          icon={<SigmaGlyph />}
          label="Otomatik Toplam"
          shortcut="Alt+="
          onClick={() => E.autoSum(ctl)}
          menu={
            <Menu>
              <MenuItem label="Topla" onClick={() => E.autoSum(ctl, "SUM")} />
              <MenuItem label="Ortalama" onClick={() => E.autoSum(ctl, "AVERAGE")} />
              <MenuItem label="Sayıları Say" onClick={() => E.autoSum(ctl, "COUNT")} />
              <MenuItem label="En Büyük" onClick={() => E.autoSum(ctl, "MAX")} />
              <MenuItem label="En Küçük" onClick={() => E.autoSum(ctl, "MIN")} />
              <MenuSeparator />
              <MenuItem label="Diğer İşlevler..." onClick={() => ctl.openDialog("insertFunction")} />
            </Menu>
          }
        />
        <RibbonMenuButton
          icon={<ArrowDown20Regular />}
          label="Doldur"
          menu={
            <Menu>
              <MenuItem label="Aşağı" shortcut="Ctrl+D" onClick={() => E.fillDirection(ctl, "down")} />
              <MenuItem label="Sağa" shortcut="Ctrl+R" onClick={() => E.fillDirection(ctl, "right")} />
              <MenuItem label="Yukarı" onClick={() => E.fillDirection(ctl, "up")} />
              <MenuItem label="Sola" onClick={() => E.fillDirection(ctl, "left")} />
              <MenuSeparator />
              <MenuItem label="Seriler..." onClick={() => ctl.openDialog("series")} />
            </Menu>
          }
        />
        <RibbonMenuButton
          icon={<Eraser20Regular />}
          label="Temizle"
          menu={
            <Menu width={230}>
              <MenuItem label="Tümünü Temizle" onClick={() => ctl.ops.clear(ctl.sheet, ctl.sel.ranges, "all")} />
              <MenuItem label="Biçimleri Temizle" onClick={() => ctl.ops.clear(ctl.sheet, ctl.sel.ranges, "formats")} />
              <MenuItem label="İçeriği Temizle" onClick={() => ctl.ops.clear(ctl.sheet, ctl.sel.ranges, "contents")} />
              <MenuItem label="Notları Temizle" onClick={() => ctl.ops.clear(ctl.sheet, ctl.sel.ranges, "notes")} />
              <MenuItem label="Köprüleri Temizle" onClick={() => ctl.ops.clear(ctl.sheet, ctl.sel.ranges, "links")} />
            </Menu>
          }
        />
      </RibbonStack>
      <RibbonMenuButton size="large" icon={<SortFilterGlyph />} label={"Sırala ve Filtre\u00a0Uygula"} menu={<SortFilterMenu ctl={ctl} />} />
      <RibbonMenuButton
        size="large"
        icon={<FindGlyph />}
        label="Bul ve Seç"
        menu={
          <Menu width={220}>
            <MenuItem icon={<Search20Regular />} label="Bul..." shortcut="Ctrl+F" onClick={() => ctl.openDialog("find")} />
            <MenuItem label="Değiştir..." shortcut="Ctrl+H" onClick={() => ctl.openDialog("replace")} />
            <MenuItem label="Git..." shortcut="Ctrl+G" onClick={() => ctl.openDialog("goto")} />
            <MenuItem label="Özel Git..." onClick={() => ctl.openDialog("goSpecial")} />
            <MenuSeparator />
            <MenuItem label="Formüller" onClick={() => selectCells((r, c) => !!ctl.wb.getFormula(ctl.sheet, r, c))} />
            <MenuItem label="Notlar" onClick={() => {
              const ranges = [...ctl.sheet.notes.keys()].map((k) => ({ r1: Math.floor(k / 16384), c1: k % 16384, r2: Math.floor(k / 16384), c2: k % 16384 }));
              if (!ranges.length) ctl.alert("Hiçbir hücre bulunamadı.");
              else ctl.setSelection({ ranges, active: { row: ranges[0].r1, col: ranges[0].c1 }, anchor: { row: ranges[0].r1, col: ranges[0].c1 } });
            }} />
            <MenuItem label="Sabitler" onClick={() => selectCells((r, c) => !ctl.wb.getFormula(ctl.sheet, r, c))} />
            <MenuItem label="Koşullu Biçimlendirme" onClick={() => {
              const ranges = ctl.sheet.condFormats.flatMap((c) => c.ranges);
              if (!ranges.length) ctl.alert("Hiçbir hücre bulunamadı.");
              else ctl.setSelection({ ranges, active: { row: ranges[0].r1, col: ranges[0].c1 }, anchor: { row: ranges[0].r1, col: ranges[0].c1 } });
            }} />
            <MenuItem label="Veri Doğrulama" onClick={() => {
              const ranges = ctl.sheet.validations.flatMap((v) => v.ranges);
              if (!ranges.length) ctl.alert("Hiçbir hücre bulunamadı.");
              else ctl.setSelection({ ranges, active: { row: ranges[0].r1, col: ranges[0].c1 }, anchor: { row: ranges[0].r1, col: ranges[0].c1 } });
            }} />
            <MenuItem label="Geçerli Bölge" onClick={() => ctl.selectRange(currentRegion(ctl.wb, ctl.sheet, ctl.sel.active))} />
            <MenuSeparator />
            <MenuItem label="Nesneleri Seç" disabled={!ctl.sheet.charts.length && !ctl.sheet.images.length && !ctl.sheet.shapes.length} onClick={() => selectFirstObject(ctl)} />
            <MenuItem label="Seçim Bölmesi..." onClick={() => ((ctl.formatPane = "selection"), ctl.emit())} />
          </Menu>
        }
      />
    </RibbonGroup>
  );
}

function SortFilterGlyph() {
  return (
    <svg width="32" height="32" viewBox="0 0 32 32" aria-hidden="true">
      <text x="3" y="13" fontSize="10" fontWeight="700" fill="#107C41" fontFamily="Segoe UI, sans-serif">A</text>
      <text x="3" y="27" fontSize="10" fontWeight="700" fill="currentColor" fontFamily="Segoe UI, sans-serif">Z</text>
      <path d="M13 6v20M10 23l3 3 3-3" fill="none" stroke="currentColor" strokeWidth="1.3" />
      <path d="M18 9h11l-4.2 5.5V22l-2.6 2v-9.5z" fill="none" stroke="currentColor" strokeWidth="1.2" />
    </svg>
  );
}

function FindGlyph() {
  return (
    <svg width="32" height="32" viewBox="0 0 32 32" aria-hidden="true">
      <circle cx="13" cy="13" r="8" fill="none" stroke="currentColor" strokeWidth="1.6" />
      <path d="M19 19l8 8" stroke="currentColor" strokeWidth="2.2" strokeLinecap="round" />
    </svg>
  );
}
