/** Grafik iletişim kutuları: Grafik Ekle / Grafik Türünü Değiştir, Veri Kaynağını Seç, Grafiği Taşı. */
import { useMemo, useState, type ReactNode } from "react";
import { Dialog, Field } from "@officehere/ui";
import { colToName as col, parseQualifiedRange, quoteSheetName } from "@officehere/formula";
import type { Controller } from "../controller";
import { ChartView } from "../charts/ChartView";
import { defaultSeriesInRows, extractChartData } from "../charts/chartData";
import { chartSourceRange, insertChart, moveChartToSheet, updateChart } from "../actions/objects";
import type { ChartDef, ChartType } from "../model/types";
import { OkCancel, Radio, SimpleDialog } from "./common";

type P = { ctl: Controller; onClose: () => void; arg?: unknown };

export const CHART_FAMILIES: { id: string; label: string; types: { type: ChartType; label: string }[] }[] = [
  { id: "column", label: "Sütun", types: [{ type: "column", label: "Kümelenmiş Sütun" }, { type: "columnStacked", label: "Yığılmış Sütun" }] },
  { id: "line", label: "Çizgi", types: [{ type: "line", label: "Çizgi" }, { type: "lineMarkers", label: "İşaretçili Çizgi" }] },
  { id: "pie", label: "Pasta", types: [{ type: "pie", label: "Pasta" }, { type: "doughnut", label: "Halka" }] },
  { id: "bar", label: "Çubuk", types: [{ type: "bar", label: "Kümelenmiş Çubuk" }, { type: "barStacked", label: "Yığılmış Çubuk" }] },
  { id: "area", label: "Alan", types: [{ type: "area", label: "Alan" }, { type: "areaStacked", label: "Yığılmış Alan" }] },
  { id: "scatter", label: "X Y (Dağılım)", types: [{ type: "scatter", label: "Dağılım" }] },
];

export function familyOf(t: ChartType): string {
  return CHART_FAMILIES.find((f) => f.types.some((x) => x.type === t))?.id ?? "column";
}

/** Önizleme için geçici grafik tanımı. */
function previewDef(ctl: Controller, base: ChartDef | null, type: ChartType): ChartDef | null {
  if (base) return { ...base, type, showGridlines: type === "pie" || type === "doughnut" ? false : base.showGridlines };
  const r = chartSourceRange(ctl);
  if (!r) return null;
  return {
    id: "preview",
    type,
    range: `${quoteSheetName(ctl.sheet.name)}!${col(r.c1)}${r.r1 + 1}:${col(r.c2)}${r.r2 + 1}`,
    seriesInRows: defaultSeriesInRows(r),
    title: "Grafik Başlığı",
    showTitle: true,
    showLegend: true,
    legendPos: type === "pie" || type === "doughnut" ? "right" : "bottom",
    showDataLabels: false,
    showGridlines: type !== "pie" && type !== "doughnut",
    palette: 0,
    style: 1,
    anchor: { row: 0, col: 0, dx: 0, dy: 0 },
    width: 400,
    height: 240,
  };
}

/** Grafik Ekle (arg.insert) ya da seçili grafiğin türünü değiştir. */
export function ChartTypeDialog({ ctl, onClose, arg }: P) {
  const a = (arg ?? {}) as { insert?: boolean; type?: ChartType };
  const chart = !a.insert && ctl.selObj?.kind === "chart" ? (ctl.sheet.charts.find((c) => c.id === ctl.selObj!.id) ?? null) : null;
  const [type, setType] = useState<ChartType>(a.type ?? chart?.type ?? "column");
  const [fam, setFam] = useState(familyOf(type));
  const def = useMemo(() => previewDef(ctl, chart, type), [type]); // eslint-disable-line react-hooks/exhaustive-deps
  const data = useMemo(() => (def ? extractChartData(ctl.wb, def) : null), [def]); // eslint-disable-line react-hooks/exhaustive-deps
  const family = CHART_FAMILIES.find((f) => f.id === fam)!;
  const ok = () => {
    if (chart) updateChart(ctl, chart.id, { type, showGridlines: type === "pie" || type === "doughnut" ? false : chart.showGridlines || (chart.type === "pie" || chart.type === "doughnut") }, "Grafik Türünü Değiştir");
    else insertChart(ctl, type);
    onClose();
  };
  return (
    <Dialog open title={chart ? "Grafik Türünü Değiştir" : "Grafik Ekle"} onClose={onClose} width={720} onSubmit={ok} footer={<OkCancel onOk={ok} onCancel={onClose} />}>
      <div className="tb-charttype">
        <div className="tb-list tb-charttype-fams">
          {CHART_FAMILIES.map((f) => (
            <div
              key={f.id}
              className={`tb-list-item${f.id === fam ? " is-active" : ""}`}
              onClick={() => {
                setFam(f.id);
                setType(f.types[0].type);
              }}
            >
              {f.label}
            </div>
          ))}
        </div>
        <div className="tb-charttype-main">
          <div className="tb-charttype-subs">
            {family.types.map((t) => (
              <button key={t.type} type="button" className={`tb-charttype-sub${t.type === type ? " is-on" : ""}`} onClick={() => setType(t.type)} title={t.label}>
                <ChartTypeIcon type={t.type} />
              </button>
            ))}
          </div>
          <div className="tb-charttype-name">{family.types.find((t) => t.type === type)?.label}</div>
          <div className="tb-charttype-preview">
            {def && data ? <ChartView def={def} data={data} width={440} height={260} /> : <div className="tb-print-empty">Önizleme için veri içeren bir aralık seçin.</div>}
          </div>
        </div>
      </div>
    </Dialog>
  );
}

/** Grafik türü için küçük simge (kendi çizimimiz). */
export function ChartTypeIcon({ type, size = 40 }: { type: ChartType; size?: number }) {
  const c1 = "#156082";
  const c2 = "#E97132";
  const body = (() => {
    switch (type) {
      case "column":
        return (
          <>
            <rect x="6" y="18" width="5" height="16" fill={c1} />
            <rect x="12" y="10" width="5" height="24" fill={c2} />
            <rect x="21" y="14" width="5" height="20" fill={c1} />
            <rect x="27" y="6" width="5" height="28" fill={c2} />
          </>
        );
      case "columnStacked":
        return (
          <>
            <rect x="7" y="20" width="8" height="14" fill={c1} />
            <rect x="7" y="12" width="8" height="8" fill={c2} />
            <rect x="22" y="14" width="8" height="20" fill={c1} />
            <rect x="22" y="5" width="8" height="9" fill={c2} />
          </>
        );
      case "bar":
        return (
          <>
            <rect x="6" y="6" width="16" height="5" fill={c1} />
            <rect x="6" y="12" width="24" height="5" fill={c2} />
            <rect x="6" y="21" width="20" height="5" fill={c1} />
            <rect x="6" y="27" width="28" height="5" fill={c2} />
          </>
        );
      case "barStacked":
        return (
          <>
            <rect x="6" y="8" width="14" height="8" fill={c1} />
            <rect x="20" y="8" width="8" height="8" fill={c2} />
            <rect x="6" y="22" width="20" height="8" fill={c1} />
            <rect x="26" y="22" width="9" height="8" fill={c2} />
          </>
        );
      case "line":
      case "lineMarkers":
        return (
          <>
            <polyline points="5,28 14,18 23,22 35,8" fill="none" stroke={c1} strokeWidth="2" />
            <polyline points="5,32 14,26 23,30 35,20" fill="none" stroke={c2} strokeWidth="2" />
            {type === "lineMarkers" && [5, 14, 23, 35].map((x, i) => <circle key={x} cx={x} cy={[28, 18, 22, 8][i]} r="2.2" fill={c1} />)}
          </>
        );
      case "pie":
        return (
          <>
            <circle cx="20" cy="20" r="14" fill={c1} />
            <path d="M20 20 L20 6 A14 14 0 0 1 33 25 Z" fill={c2} />
          </>
        );
      case "doughnut":
        return (
          <>
            <circle cx="20" cy="20" r="14" fill={c1} />
            <path d="M20 20 L20 6 A14 14 0 0 1 33 25 Z" fill={c2} />
            <circle cx="20" cy="20" r="7" fill="#fff" />
          </>
        );
      case "area":
        return (
          <>
            <path d="M5 34 L5 20 L14 12 L23 18 L35 8 L35 34 Z" fill={c1} />
            <path d="M5 34 L5 28 L14 24 L23 28 L35 22 L35 34 Z" fill={c2} />
          </>
        );
      case "areaStacked":
        return (
          <>
            <path d="M5 34 L5 14 L14 10 L23 14 L35 6 L35 34 Z" fill={c2} />
            <path d="M5 34 L5 24 L14 20 L23 24 L35 18 L35 34 Z" fill={c1} />
          </>
        );
      case "scatter":
        return [
          [8, 28],
          [12, 22],
          [17, 24],
          [21, 15],
          [26, 17],
          [31, 9],
        ].map(([x, y]) => <circle key={x} cx={x} cy={y} r="2.4" fill={c1} />);
    }
  })();
  return (
    <svg width={size} height={size} viewBox="0 0 40 40" aria-hidden="true">
      <path d="M4 4V35H36" fill="none" stroke="#8a8a8a" strokeWidth="1" />
      {body}
    </svg>
  );
}

export function ChartDataDialog({ ctl, onClose }: P) {
  const chart = ctl.selObj?.kind === "chart" ? ctl.sheet.charts.find((c) => c.id === ctl.selObj!.id) : undefined;
  const [range, setRange] = useState(chart ? "=" + chart.range : "");
  const [rows, setRows] = useState(chart?.seriesInRows ?? false);
  const draft = useMemo(() => {
    if (!chart) return null;
    const q = parseQualifiedRange(range.replace(/^=/, ""));
    if (!q) return null;
    return { ...chart, range: range.replace(/^=/, ""), seriesInRows: rows };
  }, [range, rows]); // eslint-disable-line react-hooks/exhaustive-deps
  const data = useMemo(() => (draft ? extractChartData(ctl.wb, draft) : null), [draft]); // eslint-disable-line react-hooks/exhaustive-deps
  if (!chart) return null;
  const ok = () => {
    if (!draft) return ctl.alert("Girdiğiniz başvuru geçerli değil.");
    const q = parseQualifiedRange(draft.range);
    const range2 = q && !q.sheet ? `${quoteSheetName(ctl.sheet.name)}!${draft.range}` : draft.range;
    updateChart(ctl, chart.id, { range: range2, seriesInRows: rows }, "Veri Seç");
    onClose();
  };
  return (
    <SimpleDialog title="Veri Kaynağını Seç" onClose={onClose} onOk={ok} width={600}>
      <Field label="Grafik veri aralığı:">
        <input type="text" value={range} onChange={(e) => setRange(e.target.value)} data-autofocus="true" />
      </Field>
      <div className="tb-row-gap">
        <button type="button" className="tb-btn" onClick={() => setRows(!rows)}>
          Satır/Sütun Değiştir
        </button>
      </div>
      <div className="tb-chartdata">
        <FieldSetLike title="Gösterge Girişleri (Seriler)">
          {data?.series.map((s, i) => (
            <div key={i} className="tb-list-item">
              {s.name}
            </div>
          ))}
        </FieldSetLike>
        <FieldSetLike title="Yatay (Kategori) Eksen Etiketleri">
          {data?.categories.map((c, i) => (
            <div key={i} className="tb-list-item">
              {c}
            </div>
          ))}
        </FieldSetLike>
      </div>
    </SimpleDialog>
  );
}

function FieldSetLike({ title, children }: { title: string; children: ReactNode }) {
  return (
    <div className="tb-chartdata-col">
      <div className="tb-chartdata-title">{title}</div>
      <div className="tb-list" style={{ height: 150 }}>
        {children}
      </div>
    </div>
  );
}

export function MoveChartDialog({ ctl, onClose }: P) {
  const chart = ctl.selObj?.kind === "chart" ? ctl.sheet.charts.find((c) => c.id === ctl.selObj!.id) : undefined;
  const others = ctl.wb.sheets.map((s, i) => ({ s, i })).filter((x) => x.s !== ctl.sheet && !x.s.hidden);
  const [mode, setMode] = useState<"new" | "object">(others.length ? "object" : "new");
  const [target, setTarget] = useState(others[0]?.i ?? -1);
  const [name, setName] = useState(ctl.wb.nextSheetName("Grafik"));
  if (!chart) return null;
  const ok = () => {
    if (mode === "new") {
      const err = ctl.wb.isValidSheetName(name.trim());
      if (err) return ctl.alert(err);
      ctl.ops.snapshotOp("Grafiği Taşı", () => {
        ctl.wb.addSheet(name.trim());
      });
      moveChartToSheet(ctl, chart.id, ctl.wb.sheets.length - 1);
    } else if (target >= 0) moveChartToSheet(ctl, chart.id, target);
    onClose();
  };
  return (
    <SimpleDialog title="Grafiği Taşı" onClose={onClose} onOk={ok} width={440}>
      <p className="tb-dlg-text">Grafiğin yerleştirileceği yeri seçin:</p>
      <div className="tb-row-gap">
        <Radio name="mv" value="new" current={mode} onChange={setMode}>
          Yeni sayfa:
        </Radio>
        <input type="text" value={name} disabled={mode !== "new"} onChange={(e) => setName(e.target.value)} />
      </div>
      <div className="tb-row-gap">
        <Radio name="mv" value="object" current={mode} onChange={setMode}>
          Nesne olarak:
        </Radio>
        <select value={target} disabled={mode !== "object" || !others.length} onChange={(e) => setTarget(Number(e.target.value))}>
          {others.map((x) => (
            <option key={x.i} value={x.i}>
              {x.s.name}
            </option>
          ))}
        </select>
      </div>
    </SimpleDialog>
  );
}
