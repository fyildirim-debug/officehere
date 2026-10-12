/** Bağlamsal sekmeler: Tablo Tasarımı ve (Tablo) Düzen. */
import { useState } from "react";
import {
  BorderAll20Regular,
  ColorBackground20Regular,
  TableCellsMerge24Regular,
  TableCellsSplit24Regular,
  TableDeleteRow24Regular,
  TableMoveAbove24Regular,
  TableMoveBelow24Regular,
  TableMoveLeft24Regular,
  TableMoveRight24Regular,
  TableSelectRange24Regular,
  AutoFitHeight20Regular,
  AutoFitWidth20Regular,
  TextAlignLeft20Regular,
  TextAlignCenter20Regular,
  TextAlignRight20Regular,
  TextboxAlignTop20Regular,
  TextboxAlignMiddle20Regular,
  TextboxAlignBottom20Regular,
} from "@fluentui/react-icons";
import { Menu, MenuItem, RibbonButton, RibbonCheckbox, RibbonGallery, RibbonGroup, RibbonMenuButton, RibbonRow, RibbonSpinner, RibbonStack } from "@officehere/ui";
import { cssColor } from "../model/color";
import { deleteCols, deleteRows, distributeCols, distributeRows, insertCol, insertRow, mapCells, mergeCells, normRange, splitCell, type CellRange } from "../model/tableOps";
import type { CellBorder, Color, TableShape, TableStyleOptions } from "../model/types";
import { TABLE_STYLES, cellFormat, tableStyleById } from "../render/tableStyles";
import { fillToCss } from "../render/style";
import { setParaProps } from "../state/textActions";
import { cm, fromCm, ThemeColorSplit } from "./common";
import type { RibbonCtx } from "./types";

function currentTable(ctx: RibbonCtx): TableShape | undefined {
  const s = ctx.actions.selectedShapes();
  return s.length === 1 && s[0].kind === "table" ? s[0] : undefined;
}

function currentRange(ctx: RibbonCtx, t: TableShape): CellRange {
  const st = ctx.state;
  if (st.cellSel) return normRange(st.cellSel);
  const c = st.editing?.shapeId === t.id ? st.editing.cell : undefined;
  if (c) return { r1: c[0], c1: c[1], r2: c[0], c2: c[1] };
  return { r1: 0, c1: 0, r2: 0, c2: 0 };
}

function updateTable(ctx: RibbonCtx, fn: (t: TableShape) => TableShape | null) {
  const t = currentTable(ctx);
  if (!t) return;
  const res = fn(t);
  if (!res) {
    ctx.actions.deleteSelection();
    return;
  }
  ctx.actions.updateShapes((s) => (s.id === t.id ? res : s), [t.id]);
}

function StylePreview({ id, ctx }: { id: string; ctx: RibbonCtx }) {
  const theme = ctx.state.pres.theme;
  const fake: TableShape = {
    kind: "table",
    id: "p",
    name: "",
    x: 0,
    y: 0,
    w: 50,
    h: 36,
    rot: 0,
    cols: [10, 10, 10, 10, 10],
    rows: Array.from({ length: 5 }, () => ({ h: 7, cells: Array.from({ length: 5 }, () => ({ text: { paragraphs: [] } })) })),
    style: { ...(currentTable(ctx)?.style ?? { firstRow: true, bandRow: true, firstCol: false, lastCol: false, lastRow: false, bandCol: false }), styleId: id },
  };
  return (
    <span style={{ display: "grid", gridTemplateColumns: "repeat(5, 9px)", gridAutoRows: "6px", gap: 0 }}>
      {fake.rows.flatMap((r, ri) =>
        r.cells.map((_, ci) => {
          const f = cellFormat(fake, ri, ci);
          const b = f.borders;
          const bc = (x: CellBorder | null) => (x ? cssColor(x.color, theme) : "transparent");
          return <span key={`${ri}-${ci}`} style={{ background: fillToCss(f.fill, theme), borderBottom: `1px solid ${bc(b.b)}`, borderRight: `1px solid ${bc(b.r)}` }} />;
        }),
      )}
    </span>
  );
}

type BorderMode = "bottom" | "top" | "left" | "right" | "none" | "all" | "outside" | "inside";

function applyBorders(t: TableShape, range: CellRange, mode: BorderMode, b: CellBorder): TableShape {
  const r = normRange(range);
  return mapCells(t, r, (c, i, j) => {
    const borders = { ...(c.borders ?? {}) };
    const top = i === r.r1;
    const bottom = i === r.r2;
    const left = j === r.c1;
    const right = j === r.c2;
    switch (mode) {
      case "none":
        return { ...c, borders: { l: null, r: null, t: null, b: null } };
      case "all":
        return { ...c, borders: { l: b, r: b, t: b, b } };
      case "outside":
        if (top) borders.t = b;
        if (bottom) borders.b = b;
        if (left) borders.l = b;
        if (right) borders.r = b;
        return { ...c, borders };
      case "inside":
        if (!top) borders.t = b;
        if (!bottom) borders.b = b;
        if (!left) borders.l = b;
        if (!right) borders.r = b;
        return { ...c, borders };
      case "bottom":
        if (bottom) borders.b = b;
        return { ...c, borders };
      case "top":
        if (top) borders.t = b;
        return { ...c, borders };
      case "left":
        if (left) borders.l = b;
        return { ...c, borders };
      case "right":
        if (right) borders.r = b;
        return { ...c, borders };
    }
  });
}

export function TableDesignTab({ ctx }: { ctx: RibbonCtx }) {
  const t = currentTable(ctx);
  const theme = ctx.state.pres.theme;
  const [shade, setShade] = useState<Color | null>({ scheme: "accent1" });
  const [pen, setPen] = useState<Color | null>({ rgb: "#000000" });
  const [penW, setPenW] = useState(1);
  const opt = (k: keyof Omit<TableStyleOptions, "styleId">, label: string) => (
    <RibbonCheckbox label={label} checked={!!t?.style[k]} disabled={!t} onChange={(v) => updateTable(ctx, (x) => ({ ...x, style: { ...x.style, [k]: v } }))} />
  );
  const range = t ? (ctx.state.cellSel || ctx.state.editing?.cell ? currentRange(ctx, t) : { r1: 0, c1: 0, r2: t.rows.length - 1, c2: t.cols.length - 1 }) : null;
  return (
    <>
      <RibbonGroup label="Tablo Stili Seçenekleri">
        <RibbonStack>
          {opt("firstRow", "Üst Bilgi Satırı")}
          {opt("lastRow", "Toplam Satırı")}
          {opt("bandRow", "Şeritli Satırlar")}
        </RibbonStack>
        <RibbonStack>
          {opt("firstCol", "İlk Sütun")}
          {opt("lastCol", "Son Sütun")}
          {opt("bandCol", "Şeritli Sütunlar")}
        </RibbonStack>
      </RibbonGroup>
      <RibbonGroup label="Tablo Stilleri">
        <RibbonGallery
          columns={7}
          itemWidth={54}
          itemHeight={40}
          expandedColumns={7}
          selectedId={t?.style.styleId}
          items={TABLE_STYLES.map((s) => ({ id: s.id, label: s.name, preview: <StylePreview id={s.id} ctx={ctx} /> }))}
          onSelect={(id) => updateTable(ctx, (x) => ({ ...x, style: { ...x.style, styleId: tableStyleById(id).id } }))}
        />
        <RibbonStack>
          <ThemeColorSplit
            icon={<ColorBackground20Regular />}
            label="Gölgelendirme"
            tip="Gölgelendirme"
            size="medium"
            theme={theme}
            color={shade}
            noneLabel="Dolgu Yok"
            disabled={!t}
            onColorChange={setShade}
            onApply={(c) => range && updateTable(ctx, (x) => mapCells(x, range, (cell) => ({ ...cell, fill: c ? { type: "solid", color: c } : { type: "none" } })))}
          />
          <RibbonMenuButton
            icon={<BorderAll20Regular />}
            label="Kenarlıklar"
            disabled={!t}
            menu={
              <Menu>
                {(
                  [
                    ["bottom", "Alt Kenarlık"],
                    ["top", "Üst Kenarlık"],
                    ["left", "Sol Kenarlık"],
                    ["right", "Sağ Kenarlık"],
                    ["none", "Kenarlık Yok"],
                    ["all", "Tüm Kenarlıklar"],
                    ["outside", "Dış Kenarlıklar"],
                    ["inside", "İç Kenarlıklar"],
                  ] as [BorderMode, string][]
                ).map(([m, l]) => (
                  <MenuItem key={m} label={l} onClick={() => range && updateTable(ctx, (x) => applyBorders(x, range, m, { color: pen ?? { rgb: "#000000" }, width: penW }))} />
                ))}
              </Menu>
            }
          />
        </RibbonStack>
      </RibbonGroup>
      <RibbonGroup label="Kenarlık Çiz">
        <RibbonStack>
          <RibbonSpinner label="Kalem Kalınlığı:" value={penW} min={0.25} max={6} step={0.25} decimals={2} unit="nk" width={70} onChange={setPenW} />
          <ThemeColorSplit icon={<BorderAll20Regular />} label="Kalem Rengi" tip="Kalem Rengi" size="medium" theme={theme} color={pen} onColorChange={setPen} onApply={(c) => setPen(c)} />
        </RibbonStack>
      </RibbonGroup>
    </>
  );
}

export function TableLayoutTab({ ctx }: { ctx: RibbonCtx }) {
  const { actions: a } = ctx;
  const t = currentTable(ctx);
  const r = t ? currentRange(ctx, t) : null;
  const dis = !t || !r;
  const rowH = t && r ? t.rows[r.r1]?.h ?? 0 : 0;
  const colW = t && r ? t.cols[r.c1] ?? 0 : 0;
  return (
    <>
      <RibbonGroup label="Tablo">
        <RibbonMenuButton
          size="large"
          icon={<TableSelectRange24Regular />}
          label="Seç"
          disabled={dis}
          menu={
            <Menu>
              <MenuItem label="Tabloyu Seç" onClick={() => t && a.store.set({ cellSel: { r1: 0, c1: 0, r2: t.rows.length - 1, c2: t.cols.length - 1 }, editing: null })} />
              <MenuItem label="Sütunu Seç" onClick={() => t && r && a.store.set({ cellSel: { r1: 0, c1: r.c1, r2: t.rows.length - 1, c2: r.c2 }, editing: null })} />
              <MenuItem label="Satırı Seç" onClick={() => t && r && a.store.set({ cellSel: { r1: r.r1, c1: 0, r2: r.r2, c2: t.cols.length - 1 }, editing: null })} />
            </Menu>
          }
        />
      </RibbonGroup>
      <RibbonGroup label="Satırlar ve Sütunlar">
        <RibbonMenuButton
          size="large"
          icon={<TableDeleteRow24Regular />}
          label="Sil"
          disabled={dis}
          menu={
            <Menu>
              <MenuItem label="Sütunları Sil" onClick={() => r && (a.store.set({ editing: null, cellSel: null }), updateTable(ctx, (x) => deleteCols(x, r.c1, r.c2)))} />
              <MenuItem label="Satırları Sil" onClick={() => r && (a.store.set({ editing: null, cellSel: null }), updateTable(ctx, (x) => deleteRows(x, r.r1, r.r2)))} />
              <MenuItem label="Tabloyu Sil" onClick={() => a.deleteSelection()} />
            </Menu>
          }
        />
        <RibbonButton size="large" icon={<TableMoveAbove24Regular />} label="Üste Ekle" disabled={dis} onClick={() => r && updateTable(ctx, (x) => insertRow(x, r.r1, false))} />
        <RibbonStack>
          <RibbonButton icon={<TableMoveBelow24Regular />} label="Alta Ekle" disabled={dis} onClick={() => r && updateTable(ctx, (x) => insertRow(x, r.r2, true))} />
          <RibbonButton icon={<TableMoveLeft24Regular />} label="Sola Ekle" disabled={dis} onClick={() => r && updateTable(ctx, (x) => insertCol(x, r.c1, false))} />
          <RibbonButton icon={<TableMoveRight24Regular />} label="Sağa Ekle" disabled={dis} onClick={() => r && updateTable(ctx, (x) => insertCol(x, r.c2, true))} />
        </RibbonStack>
      </RibbonGroup>
      <RibbonGroup label="Birleştir">
        <RibbonButton size="large" icon={<TableCellsMerge24Regular />} label="Hücreleri Birleştir" disabled={dis || !ctx.state.cellSel} onClick={() => r && (updateTable(ctx, (x) => mergeCells(x, r)), a.store.set({ cellSel: null }))} />
        <RibbonButton size="large" icon={<TableCellsSplit24Regular />} label="Hücreleri Böl" disabled={dis || !(t && r && (t.rows[r.r1].cells[r.c1].gridSpan || t.rows[r.r1].cells[r.c1].rowSpan))} onClick={() => r && updateTable(ctx, (x) => splitCell(x, r.r1, r.c1))} />
      </RibbonGroup>
      <RibbonGroup label="Hücre Boyutu">
        <RibbonStack>
          <RibbonSpinner label="Yükseklik:" value={cm(rowH)} min={0.1} max={50} step={0.1} decimals={2} unit="cm" width={84} onChange={(v) => r && updateTable(ctx, (x) => { const rows = x.rows.map((row, i) => (i >= r.r1 && i <= r.r2 ? { ...row, h: fromCm(v) } : row)); return { ...x, rows, h: rows.reduce((s, q) => s + q.h, 0) }; })} />
          <RibbonSpinner label="Genişlik:" value={cm(colW)} min={0.1} max={50} step={0.1} decimals={2} unit="cm" width={84} onChange={(v) => r && updateTable(ctx, (x) => { const cols = x.cols.map((c, j) => (j >= r.c1 && j <= r.c2 ? fromCm(v) : c)); return { ...x, cols, w: cols.reduce((s, q) => s + q, 0) }; })} />
        </RibbonStack>
        <RibbonStack>
          <RibbonButton icon={<AutoFitHeight20Regular />} label="Satırları Dağıt" disabled={dis} onClick={() => updateTable(ctx, distributeRows)} />
          <RibbonButton icon={<AutoFitWidth20Regular />} label="Sütunları Dağıt" disabled={dis} onClick={() => updateTable(ctx, distributeCols)} />
        </RibbonStack>
      </RibbonGroup>
      <RibbonGroup label="Hizalama">
        <RibbonStack>
          <RibbonRow>
            <RibbonButton size="small" icon={<TextAlignLeft20Regular />} tip="Sola Hizala" disabled={dis} onClick={() => setParaProps(a, { align: "l" })} />
            <RibbonButton size="small" icon={<TextAlignCenter20Regular />} tip="Ortala" disabled={dis} onClick={() => setParaProps(a, { align: "ctr" })} />
            <RibbonButton size="small" icon={<TextAlignRight20Regular />} tip="Sağa Hizala" disabled={dis} onClick={() => setParaProps(a, { align: "r" })} />
          </RibbonRow>
          <RibbonRow>
            <RibbonButton size="small" icon={<TextboxAlignTop20Regular />} tip="Üste Hizala" disabled={dis} onClick={() => r && updateTable(ctx, (x) => mapCells(x, ctx.state.cellSel ? r : null, (c) => ({ ...c, anchor: "t" })))} />
            <RibbonButton size="small" icon={<TextboxAlignMiddle20Regular />} tip="Dikey Ortala" disabled={dis} onClick={() => r && updateTable(ctx, (x) => mapCells(x, ctx.state.cellSel ? r : null, (c) => ({ ...c, anchor: "ctr" })))} />
            <RibbonButton size="small" icon={<TextboxAlignBottom20Regular />} tip="Alta Hizala" disabled={dis} onClick={() => r && updateTable(ctx, (x) => mapCells(x, ctx.state.cellSel ? r : null, (c) => ({ ...c, anchor: "b" })))} />
          </RibbonRow>
        </RibbonStack>
      </RibbonGroup>
      <RibbonGroup label="Tablo Boyutu">
        <RibbonStack>
          <RibbonSpinner label="Yükseklik:" value={t ? cm(t.h) : 0} min={0.5} max={142} step={0.1} decimals={2} unit="cm" width={84} onChange={(v) => updateTable(ctx, (x) => { const k = fromCm(v) / (x.h || 1); const rows = x.rows.map((row) => ({ ...row, h: row.h * k })); return { ...x, rows, h: fromCm(v) }; })} />
          <RibbonSpinner label="Genişlik:" value={t ? cm(t.w) : 0} min={0.5} max={142} step={0.1} decimals={2} unit="cm" width={84} onChange={(v) => updateTable(ctx, (x) => { const k = fromCm(v) / (x.w || 1); return { ...x, cols: x.cols.map((c) => c * k), w: fromCm(v) }; })} />
        </RibbonStack>
      </RibbonGroup>
    </>
  );
}
