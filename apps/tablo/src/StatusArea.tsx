/** Durum çubuğu: kip, ileti, seçim istatistikleri, görünüm düğmeleri ve yakınlaştırma. */
import { useSyncExternalStore } from "react";
import { DocumentOnePage20Regular, Grid20Regular, ReadingModeMobile20Regular } from "@fluentui/react-icons";
import { StatusBar, StatusItem, ViewSwitcher, ZoomSlider } from "@officehere/ui";
import { addr, formatValue, isError } from "@officehere/formula";
import type { Controller } from "./controller";

const MODE: Record<string, string> = { ready: "Hazır", enter: "Gir", edit: "Düzenle", point: "İşaret" };

function selectionStats(ctl: Controller): { count: number; nums: number; sum: number } | null {
  const { wb, sheet } = ctl;
  const ext = wb.dataExtent(sheet);
  let count = 0;
  let nums = 0;
  let sum = 0;
  let cells = 0;
  const seen = new Set<number>();
  for (const r of ctl.sel.ranges) {
    const r2 = Math.min(r.r2, ext.rows - 1);
    const c2 = Math.min(r.c2, ext.cols - 1);
    const area = (r2 - r.r1 + 1) * (c2 - r.c1 + 1);
    if (area <= 0) continue;
    if (area > 400000) {
      wb.engine.forEachValue(sheet.id, (row, col, v) => {
        if (row < r.r1 || row > r2 || col < r.c1 || col > c2 || v === null) return;
        const k = row * 16384 + col;
        if (seen.has(k)) return;
        seen.add(k);
        if (sheet.rows.isHidden(row)) return;
        count++;
        if (typeof v === "number") {
          nums++;
          sum += v;
        }
      });
      continue;
    }
    for (let i = r.r1; i <= r2; i++) {
      if (sheet.rows.isHidden(i)) continue;
      for (let j = r.c1; j <= c2; j++) {
        cells++;
        const k = i * 16384 + j;
        if (seen.has(k)) continue;
        seen.add(k);
        const v = wb.getValue(sheet, i, j);
        if (v === null || isError(v)) continue;
        count++;
        if (typeof v === "number") {
          nums++;
          sum += v;
        }
      }
    }
  }
  void cells;
  return count >= 2 ? { count, nums, sum } : null;
}

export function StatusArea({ ctl }: { ctl: Controller }) {
  useSyncExternalStore(ctl.subscribe, ctl.getVersion);
  const sheet = ctl.sheet;
  const stats = ctl.edit ? null : selectionStats(ctl);
  const fmt = ctl.wb.styleAt(sheet, ctl.sel.active.row, ctl.sel.active.col).numFmt;
  const show = (v: number) => formatValue(Number(v.toPrecision(15)), fmt && fmt !== "@" ? fmt : "General").text;
  const circ = ctl.wb.engine.getCircularRefs().find((c) => c.sheet === sheet.id);
  const calcPending = !ctl.wb.engine.autoCalc && ctl.wb.engine.pending;
  return (
    <StatusBar
      left={
        <>
          <StatusItem>{MODE[ctl.mode]}</StatusItem>
          {ctl.status && <StatusItem>{ctl.status}</StatusItem>}
          {circ && <StatusItem tip="Döngüsel başvuru içeren hücre">Döngüsel Başvurular: {addr(circ.row, circ.col)}</StatusItem>}
          {calcPending && (
            <StatusItem onClick={() => (ctl.wb.engine.recalc(true), ctl.emit())} tip="Şimdi Hesapla (F9)">
              Hesapla
            </StatusItem>
          )}
          {ctl.wb.sheets.length > 0 && sheet.autoFilter && Object.keys(sheet.autoFilter.columns).length > 0 && <StatusItem>Filtre Modu</StatusItem>}
        </>
      }
      right={
        <>
          {stats && (
            <span className="tb-status-stats">
              {stats.nums > 0 && <StatusItem>Ortalama: {show(stats.sum / stats.nums)}</StatusItem>}
              <StatusItem>Sayım: {stats.count.toLocaleString("tr-TR")}</StatusItem>
              {stats.nums > 0 && <StatusItem>Toplam: {show(stats.sum)}</StatusItem>}
            </span>
          )}
          <ViewSwitcher
            active={sheet.viewMode}
            onChange={(id) => {
              sheet.viewMode = id as typeof sheet.viewMode;
              ctl.emit();
            }}
            views={[
              { id: "normal", icon: <Grid20Regular />, tip: "Normal" },
              { id: "pageLayout", icon: <DocumentOnePage20Regular />, tip: "Sayfa Düzeni" },
              { id: "pageBreak", icon: <ReadingModeMobile20Regular />, tip: "Sayfa Sonu Önizleme" },
            ]}
          />
          <ZoomSlider
            value={sheet.zoom}
            min={10}
            max={400}
            onChange={(v) => {
              sheet.zoom = Math.max(10, Math.min(400, v));
              ctl.emit();
            }}
            onPercentClick={() => ctl.openDialog("zoom")}
          />
        </>
      }
    />
  );
}
