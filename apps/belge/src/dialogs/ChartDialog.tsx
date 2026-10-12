// Grafik Ekle / Verileri Düzenle: grafik türü, başlık, boyut ve veri tablosu (kategoriler × seriler);
// canlı önizleme. Seçili grafik varsa düzenlenir.
import { useMemo, useState } from "react";
import { NodeSelection } from "@tiptap/pm/state";
import { THEMES } from "@officehere/core";
import { Dialog, DialogButton, Field } from "@officehere/ui";
import { useBelge } from "../state/context";
import { CHART_TYPE_LABEL, chartSvg, defaultChart, type ChartData, type ChartType } from "../editor/chart";

const TYPES: ChartType[] = ["column", "bar", "line", "pie"];
const num = (s: string) => {
  const v = parseFloat(s.replace(/\./g, "").replace(",", "."));
  return Number.isFinite(v) ? v : 0;
};
const show = (v: number) => (Number.isFinite(v) ? String(v).replace(".", ",") : "");

export function ChartDialog({ onClose }: { onClose: () => void }) {
  const { editor, settings } = useBelge();
  const sel = editor.state.selection;
  const editing = sel instanceof NodeSelection && sel.node.type.name === "chart" ? (sel.node.attrs.data as ChartData | null) : null;
  const accents = useMemo(() => (THEMES.find((t) => t.name === settings.theme) ?? THEMES[0]).colors.slice(4), [settings.theme]);
  const [data, setData] = useState<ChartData>(() => (editing ? JSON.parse(JSON.stringify(editing)) : defaultChart("column", accents)));
  const set = (patch: Partial<ChartData>) => setData((d) => ({ ...d, ...patch }));
  const setCell = (si: number, ci: number, v: string) => setData((d) => ({ ...d, series: d.series.map((s, k) => (k === si ? { ...s, values: d.categories.map((_, j) => (j === ci ? num(v) : (s.values[j] ?? 0))) } : s)) }));
  const addCategory = () => setData((d) => ({ ...d, categories: [...d.categories, `Kategori ${d.categories.length + 1}`], series: d.series.map((s) => ({ ...s, values: [...s.values, 0] })) }));
  const removeCategory = () => setData((d) => (d.categories.length > 1 ? { ...d, categories: d.categories.slice(0, -1), series: d.series.map((s) => ({ ...s, values: s.values.slice(0, -1) })) } : d));
  const addSeries = () => setData((d) => ({ ...d, series: [...d.series, { name: `Seri ${d.series.length + 1}`, values: d.categories.map(() => 0), color: accents[d.series.length % accents.length] ?? "#7F7F7F" }] }));
  const removeSeries = () => setData((d) => (d.series.length > 1 ? { ...d, series: d.series.slice(0, -1) } : d));
  const preview = useMemo(() => chartSvg({ ...data, width: 360, height: 220 }), [data]);
  const apply = () => {
    editor.chain().focus().insertChart(data).run();
    onClose();
  };
  return (
    <Dialog
      open
      title={editing ? "Verileri Düzenle" : "Grafik Ekle"}
      width={760}
      onClose={onClose}
      footer={
        <>
          <DialogButton primary onClick={apply}>
            Tamam
          </DialogButton>
          <DialogButton onClick={onClose}>İptal</DialogButton>
        </>
      }
    >
      <div className="bl-chart-dlg">
        <div className="bl-chart-types" role="radiogroup" aria-label="Grafik türü">
          {TYPES.map((t) => (
            <button key={t} type="button" role="radio" aria-checked={data.type === t} className={`bl-chart-type${data.type === t ? " is-on" : ""}`} onClick={() => set({ type: t })}>
              <span dangerouslySetInnerHTML={{ __html: chartSvg({ ...defaultChart(t, accents), title: "", width: 72, height: 52 }, true) }} />
              {CHART_TYPE_LABEL[t]}
            </button>
          ))}
        </div>
        <div className="bl-chart-main">
          <div className="bl-chart-left">
            <Field label="Grafik başlığı:" inline>
              <input value={data.title} onChange={(e) => set({ title: e.target.value })} aria-label="Grafik başlığı" />
            </Field>
            <div className="bl-chart-grid-wrap">
              <table className="bl-chart-grid" aria-label="Grafik verileri">
                <thead>
                  <tr>
                    <th />
                    {data.series.map((s, si) => (
                      <th key={si}>
                        <input value={s.name} aria-label={`Seri ${si + 1} adı`} onChange={(e) => setData((d) => ({ ...d, series: d.series.map((x, k) => (k === si ? { ...x, name: e.target.value } : x)) }))} />
                      </th>
                    ))}
                  </tr>
                </thead>
                <tbody>
                  {data.categories.map((c, ci) => (
                    <tr key={ci}>
                      <th>
                        <input value={c} aria-label={`Kategori ${ci + 1}`} onChange={(e) => setData((d) => ({ ...d, categories: d.categories.map((x, k) => (k === ci ? e.target.value : x)) }))} />
                      </th>
                      {data.series.map((s, si) => (
                        <td key={si}>
                          <input defaultValue={show(s.values[ci] ?? 0)} aria-label={`${s.name} ${c}`} onChange={(e) => setCell(si, ci, e.target.value)} />
                        </td>
                      ))}
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
            <div className="bl-chart-actions">
              <button type="button" className="oh-btn" onClick={addCategory}>
                Kategori Ekle
              </button>
              <button type="button" className="oh-btn" onClick={removeCategory}>
                Kategori Sil
              </button>
              <button type="button" className="oh-btn" onClick={addSeries}>
                Seri Ekle
              </button>
              <button type="button" className="oh-btn" onClick={removeSeries}>
                Seri Sil
              </button>
            </div>
            <div className="bl-dlg-grid">
              <Field label="Genişlik (cm):">
                <input type="number" step={0.5} min={3} value={+(data.width / 37.795).toFixed(1)} onChange={(e) => set({ width: Math.max(113, Math.round(num(e.target.value) * 37.795)) })} />
              </Field>
              <Field label="Yükseklik (cm):">
                <input type="number" step={0.5} min={2} value={+(data.height / 37.795).toFixed(1)} onChange={(e) => set({ height: Math.max(76, Math.round(num(e.target.value) * 37.795)) })} />
              </Field>
            </div>
          </div>
          <div className="bl-chart-preview" aria-label="Önizleme" dangerouslySetInnerHTML={{ __html: preview }} />
        </div>
      </div>
    </Dialog>
  );
}
