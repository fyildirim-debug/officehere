/** Grafik Ekle / Grafik Türünü Değiştir ve Verileri Düzenle iletişim kutuları. */
import { useState } from "react";
import { DataArea24Regular, DataBarHorizontal24Regular, DataBarVertical24Regular, DataLine24Regular, DataPie24Regular, Add20Regular, Subtract20Regular } from "@fluentui/react-icons";
import { Dialog, DialogButton } from "@officehere/ui";
import { applySubtype, CHART_TYPES, chartFromGrid, chartSubtypes, chartToGrid, defaultChartData, subtypeIdOf } from "../model/chart";
import type { ChartData, ChartType, Theme } from "../model/types";
import { ChartSvg } from "../render/ChartSvg";

const ICONS: Record<ChartType, React.ReactNode> = {
  col: <DataBarVertical24Regular />,
  bar: <DataBarHorizontal24Regular />,
  line: <DataLine24Regular />,
  pie: <DataPie24Regular />,
  doughnut: (
    <svg viewBox="0 0 24 24" width="24" height="24" aria-hidden="true">
      <circle cx="12" cy="12" r="8" fill="none" stroke="currentColor" strokeWidth="4" />
    </svg>
  ),
  area: <DataArea24Regular />,
  scatter: (
    <svg viewBox="0 0 24 24" width="24" height="24" aria-hidden="true">
      <path d="M3 3v18h18" fill="none" stroke="currentColor" strokeWidth="1.4" />
      {[[7, 15], [10, 9], [13, 13], [16, 6], [19, 10]].map(([x, y]) => (
        <circle key={`${x}-${y}`} cx={x} cy={y} r="1.6" fill="currentColor" />
      ))}
    </svg>
  ),
  combo: (
    <svg viewBox="0 0 24 24" width="24" height="24" aria-hidden="true">
      <rect x="4" y="12" width="3.5" height="8" fill="none" stroke="currentColor" strokeWidth="1.2" />
      <rect x="10.2" y="8" width="3.5" height="12" fill="none" stroke="currentColor" strokeWidth="1.2" />
      <rect x="16.4" y="14" width="3.5" height="6" fill="none" stroke="currentColor" strokeWidth="1.2" />
      <path d="M3 9l6-4 6 3 6-5" fill="none" stroke="currentColor" strokeWidth="1.5" />
    </svg>
  ),
  radar: (
    <svg viewBox="0 0 24 24" width="24" height="24" aria-hidden="true">
      <path d="M12 2.5l9 6.5-3.5 11h-11L3 9z" fill="none" stroke="currentColor" strokeWidth="1.2" />
      <path d="M12 7l5 4-2 6H9l-2-6z" fill="none" stroke="currentColor" strokeWidth="1.4" />
    </svg>
  ),
};

/** Grafik türü seçimi. `current` verilirse "Grafik Türünü Değiştir" kipidir. */
export function InsertChartDialog({ open, current, theme, onClose, onPick }: { open: boolean; current?: ChartData; theme: Theme; onClose: () => void; onPick: (t: ChartType, patch: Partial<ChartData>) => void }) {
  const [type, setTypeRaw] = useState<ChartType>(current?.type ?? "col");
  const [sub, setSub] = useState<string | undefined>(current ? subtypeIdOf(current) : undefined);
  const [last, setLast] = useState(false);
  if (open !== last) {
    setLast(open);
    if (open) {
      setTypeRaw(current?.type ?? "col");
      setSub(current ? subtypeIdOf(current) : undefined);
    }
  }
  if (!open) return null;
  const setType = (t: ChartType) => {
    setTypeRaw(t);
    setSub(current && current.type === t ? subtypeIdOf(current) : t === "doughnut" ? "doughnut" : undefined);
  };
  const subs = chartSubtypes(type);
  const chosen = subs.find((x) => x.id === sub) ?? subs[0];
  const patch = chosen?.patch ?? {};
  const preview: ChartData = current ? applySubtype(current, patch.type ?? type, patch) : defaultChartData(type, patch);
  const subLabel = chosen?.label;
  // Seçim sonrası kapatma/sonraki adım çağırana aittir (ekleme sonrası veri tablosu açılır).
  const ok = () => onPick(type, patch);
  return (
    <Dialog
      open
      title={current ? "Grafik Türünü Değiştir" : "Grafik Ekle"}
      onClose={onClose}
      onSubmit={ok}
      width={640}
      footer={
        <>
          <DialogButton primary onClick={ok}>
            Tamam
          </DialogButton>
          <DialogButton onClick={onClose}>İptal</DialogButton>
        </>
      }
    >
      <div className="chart-dlg">
        <div className="chart-dlg-list" role="listbox" aria-label="Grafik türü">
          {CHART_TYPES.map((c) => (
            <div key={c.type} role="option" aria-selected={type === c.type} className={`list-row${type === c.type ? " is-selected" : ""}`} onClick={() => setType(c.type)} onDoubleClick={() => onPick(c.type, chartSubtypes(c.type)[0]?.patch ?? {})}>
              {ICONS[c.type]}
              <span>{c.label}</span>
            </div>
          ))}
        </div>
        <div className="chart-dlg-preview">
          {subs.length > 0 && (
            <div className="chart-dlg-subs" role="radiogroup" aria-label="Alt tür">
              {subs.map((x) => (
                <button key={x.id} type="button" role="radio" aria-checked={chosen?.id === x.id} className={`chart-dlg-sub${chosen?.id === x.id ? " is-selected" : ""}`} title={x.label} aria-label={x.label} onClick={() => setSub(x.id)}>
                  <span className="chart-thumb" style={{ width: 64, height: 44 }}>
                    <ChartSvg chart={{ ...defaultChartData(type, x.patch), showTitle: false, legend: "none" }} w={64} h={44} theme={theme} />
                  </span>
                </button>
              ))}
            </div>
          )}
          <div className="chart-dlg-cap">{subLabel ?? CHART_TYPES.find((c) => c.type === type)?.label}</div>
          <div className="chart-dlg-box">
            <ChartSvg chart={preview} w={400} h={250} theme={theme} />
          </div>
        </div>
      </div>
    </Dialog>
  );
}

/** Basit veri tablosu: ilk satır seri adları, ilk sütun kategoriler. */
export function ChartDataDialog({ chart, onClose, onApply }: { chart: ChartData | null; onClose: () => void; onApply: (c: ChartData) => void }) {
  const [grid, setGrid] = useState<string[][]>(() => (chart ? chartToGrid(chart) : [[""]]));
  const [last, setLast] = useState<ChartData | null>(null);
  if (chart !== last) {
    setLast(chart);
    if (chart) setGrid(chartToGrid(chart));
  }
  if (!chart) return null;
  const pie = chart.type === "pie" || chart.type === "doughnut";
  const set = (r: number, c: number, v: string) => setGrid((g) => g.map((row, i) => (i === r ? row.map((x, j) => (j === c ? v : x)) : row)));
  const addRow = () => setGrid((g) => [...g, [`Kategori ${g.length}`, ...g[0].slice(1).map(() => "")]]);
  const addCol = () => setGrid((g) => g.map((row, i) => [...row, i === 0 ? `Seri ${row.length}` : ""]));
  const delRow = () => setGrid((g) => (g.length > 2 ? g.slice(0, -1) : g));
  const delCol = () => setGrid((g) => (g[0].length > 2 ? g.map((row) => row.slice(0, -1)) : g));
  const ok = () => {
    onApply(chartFromGrid(chart, grid));
    onClose();
  };
  return (
    <Dialog
      open
      title="Verileri Düzenle"
      onClose={onClose}
      width={Math.min(760, 180 + grid[0].length * 110)}
      footer={
        <>
          <DialogButton primary onClick={ok}>
            Tamam
          </DialogButton>
          <DialogButton onClick={onClose}>İptal</DialogButton>
        </>
      }
    >
      <div className="chart-grid-tools">
        <button type="button" className="oh-btn" onClick={addRow}>
          <Add20Regular /> Satır Ekle
        </button>
        <button type="button" className="oh-btn" onClick={delRow} disabled={grid.length <= 2}>
          <Subtract20Regular /> Satır Sil
        </button>
        {!pie && (
          <>
            <button type="button" className="oh-btn" onClick={addCol}>
              <Add20Regular /> Seri Ekle
            </button>
            <button type="button" className="oh-btn" onClick={delCol} disabled={grid[0].length <= 2}>
              <Subtract20Regular /> Seri Sil
            </button>
          </>
        )}
      </div>
      <div className="chart-grid oh-scroll">
        <table>
          <thead>
            <tr>
              <th />
              {grid[0].map((_, j) => (
                <th key={j}>{String.fromCharCode(65 + j)}</th>
              ))}
            </tr>
          </thead>
          <tbody>
            {grid.map((row, i) => (
              <tr key={i}>
                <th>{i + 1}</th>
                {row.map((v, j) => (
                  <td key={j} className={i === 0 && j > 0 ? "is-series" : j === 0 && i > 0 ? "is-cat" : undefined}>
                    <input
                      type="text"
                      aria-label={`${String.fromCharCode(65 + j)}${i + 1}`}
                      value={v}
                      disabled={i === 0 && j === 0}
                      inputMode={i > 0 && j > 0 ? "decimal" : undefined}
                      onChange={(e) => set(i, j, e.target.value)}
                      onKeyDown={(e) => {
                        if (e.key === "Enter") {
                          e.preventDefault();
                          const next = (e.currentTarget.closest("tbody")?.querySelectorAll("tr")[i + 1]?.querySelectorAll("input")[j] as HTMLInputElement | undefined) ?? undefined;
                          next?.focus();
                          next?.select();
                        }
                      }}
                    />
                  </td>
                ))}
              </tr>
            ))}
          </tbody>
        </table>
      </div>
      <p className="chart-grid-note">İlk satır seri adlarını, ilk sütun kategorileri içerir. Ondalık ayırıcı olarak virgül kullanın.</p>
    </Dialog>
  );
}
