/** Durum Çözümlemesi iletişim kutuları: Senaryo Yöneticisi, Senaryo Ekle/Düzenle, Senaryo Değerleri, Senaryo Özeti, Veri Tablosu. */
import { useState } from "react";
import { addr } from "@officehere/formula";
import { Dialog, DialogButton, Field } from "@officehere/ui";
import type { Controller } from "../controller";
import { parseRangeLoose } from "../controller";
import type { CellPos, Scenario } from "../model/types";
import { displayText } from "../model/display";
import { cellsText, createDataTable, deleteScenario, parseCells, saveScenario, scenarioSummary, showScenario } from "../actions/whatIfOps";
import { Check, SimpleDialog, parseTrNumber } from "./common";

type P = { ctl: Controller; onClose: () => void; arg?: unknown };

const parse = (s: string) => parseRangeLoose(s.replace(/\$/g, ""));

/** Senaryo Yöneticisi: senaryolar listesi, Göster, Ekle, Sil, Düzenle, Özet. */
export function ScenarioManagerDialog({ ctl, onClose }: P) {
  const sheet = ctl.sheet;
  const list = sheet.scenarios;
  const [sel, setSel] = useState(list[0]?.name ?? "");
  const cur = list.find((s) => s.name === sel) ?? list[0];
  return (
    <Dialog
      open
      title="Senaryo Yöneticisi"
      onClose={onClose}
      width={460}
      footer={
        <>
          <DialogButton primary disabled={!cur} onClick={() => cur && showScenario(ctl, cur.name)}>
            Göster
          </DialogButton>
          <DialogButton onClick={onClose}>Kapat</DialogButton>
        </>
      }
    >
      <div className="tb-scn">
        <div className="tb-scn-list">
          <div className="tb-scn-label">Senaryolar:</div>
          <div className="tb-list" role="listbox" aria-label="Senaryolar" style={{ height: 150 }}>
            {list.length === 0 && <div className="tb-scn-empty">Tanımlı senaryo yok. Senaryo eklemek için Ekle'yi seçin.</div>}
            {list.map((s) => (
              <div key={s.name} role="option" aria-selected={s === cur} className={`tb-list-item${s === cur ? " is-active" : ""}`} onClick={() => setSel(s.name)} onDoubleClick={() => showScenario(ctl, s.name)}>
                {s.name}
              </div>
            ))}
          </div>
        </div>
        <div className="tb-scn-buttons">
          <DialogButton onClick={() => ctl.openDialog("scenarioEdit")}>Ekle...</DialogButton>
          <DialogButton disabled={!cur} onClick={() => cur && (deleteScenario(ctl, cur.name), setSel(list.find((s) => s !== cur)?.name ?? ""))}>
            Sil
          </DialogButton>
          <DialogButton disabled={!cur} onClick={() => cur && ctl.openDialog("scenarioEdit", { name: cur.name })}>
            Düzenle...
          </DialogButton>
          <DialogButton disabled={!list.length} onClick={() => (onClose(), ctl.openDialog("scenarioSummary"))}>
            Özet...
          </DialogButton>
        </div>
      </div>
      {cur && (
        <div className="tb-scn-info">
          <div>
            <span className="tb-scn-label">Değişen hücreler:</span> {cellsText(cur.cells)}
          </div>
          <div>
            <span className="tb-scn-label">Açıklama:</span> {cur.comment || "—"}
          </div>
        </div>
      )}
    </Dialog>
  );
}

/** Senaryo Ekle / Düzenle: ad, değişen hücreler, açıklama, koruma; Tamam değerler kutusunu açar. */
export function ScenarioEditDialog({ ctl, onClose, arg }: P) {
  const editing = (arg as { name?: string } | undefined)?.name;
  const old = editing ? ctl.sheet.scenarios.find((s) => s.name === editing) : undefined;
  const [name, setName] = useState(old?.name ?? "");
  const [cells, setCells] = useState(old ? cellsText(old.cells) : ctl.sel.ranges.map((r) => (r.r1 === r.r2 && r.c1 === r.c2 ? `${addr(r.r1, r.c1)}` : `${addr(r.r1, r.c1)}:${addr(r.r2, r.c2)}`)).join(";"));
  const [comment, setComment] = useState(old?.comment ?? `${new Date().toLocaleDateString("tr-TR")} tarihinde oluşturuldu`);
  const [locked, setLocked] = useState(old?.locked !== false);
  const [hidden, setHidden] = useState(!!old?.hidden);
  const ok = () => {
    const list = parseCells(cells, parse);
    if (!name.trim()) return ctl.alert("Senaryo adı boş bırakılamaz.");
    if (!list) return ctl.alert("Değişen hücreler başvurusu geçerli değil.");
    if (list.length > 32) return ctl.alert("Bir senaryoda en çok 32 değişen hücre olabilir.");
    if (ctl.sheet.scenarios.some((s) => s.name.toLocaleUpperCase("tr-TR") === name.trim().toLocaleUpperCase("tr-TR") && s.name !== editing)) return ctl.alert(`“${name.trim()}” adlı bir senaryo zaten var. Başka bir ad girin.`);
    onClose();
    ctl.openDialog("scenarioValues", { name: name.trim(), cells: list, comment, locked, hidden, oldName: editing });
  };
  return (
    <SimpleDialog title={editing ? "Senaryoyu Düzenle" : "Senaryo Ekle"} onClose={onClose} onOk={ok} width={420}>
      <Field label="Senaryo adı:">
        <input type="text" value={name} onChange={(e) => setName(e.target.value)} aria-label="Senaryo adı" data-autofocus="true" />
      </Field>
      <Field label="Değişen hücreler:">
        <input type="text" value={cells} onChange={(e) => setCells(e.target.value)} aria-label="Değişen hücreler" />
      </Field>
      <p className="tb-fc-note">Bitişik olmayan hücreleri noktalı virgülle ayırın (ör. B2;B4:B5).</p>
      <Field label="Açıklama:">
        <textarea value={comment} onChange={(e) => setComment(e.target.value)} rows={3} aria-label="Açıklama" />
      </Field>
      <fieldset className="oh-fieldset">
        <legend>Koruma</legend>
        <Check checked={locked} onChange={setLocked}>
          Değişiklikleri engelle
        </Check>
        <Check checked={hidden} onChange={setHidden}>
          Gizle
        </Check>
      </fieldset>
    </SimpleDialog>
  );
}

/** Senaryo Değerleri: her değişen hücre için değer. */
export function ScenarioValuesDialog({ ctl, onClose, arg }: P) {
  const a = arg as { name: string; cells: CellPos[]; comment?: string; locked: boolean; hidden: boolean; oldName?: string };
  const sheet = ctl.sheet;
  const old = a.oldName ? sheet.scenarios.find((s) => s.name === a.oldName) : undefined;
  const [vals, setVals] = useState(() =>
    a.cells.map((c) => {
      const v = old?.cells.find((x) => x.row === c.row && x.col === c.col)?.value;
      if (v !== undefined) return v === null ? "" : typeof v === "number" ? v.toLocaleString("tr-TR", { useGrouping: false, maximumFractionDigits: 10 }) : String(v);
      return displayText(ctl.wb, sheet, c.row, c.col);
    }),
  );
  const ok = () => {
    const cells = a.cells.map((c, i) => {
      const t = vals[i].trim();
      const n = parseTrNumber(t);
      return { row: c.row, col: c.col, value: t === "" ? null : Number.isFinite(n) ? n : t };
    });
    const sc: Scenario = { name: a.name, cells, ...(a.comment ? { comment: a.comment } : {}), ...(a.locked ? {} : { locked: false }), ...(a.hidden ? { hidden: true } : {}) };
    const err = saveScenario(ctl, sc, a.oldName);
    if (err) return ctl.alert(err);
    // Altta açık duran Senaryo Yöneticisi yeni listeyle güncellenir
    onClose();
  };
  return (
    <SimpleDialog title="Senaryo Değerleri" onClose={onClose} onOk={ok} width={360}>
      <p className="tb-fc-note">Değişen hücrelerin her biri için değer girin.</p>
      <div className="tb-scn-values">
        {a.cells.map((c, i) => (
          <label key={i} className="tb-scn-value">
            <span>
              {i + 1}: {addr(c.row, c.col)}
            </span>
            <input type="text" value={vals[i]} onChange={(e) => setVals(vals.map((v, j) => (j === i ? e.target.value : v)))} aria-label={`${addr(c.row, c.col)} değeri`} data-autofocus={i === 0 ? "true" : undefined} />
          </label>
        ))}
      </div>
    </SimpleDialog>
  );
}

/** Senaryo Özeti: sonuç hücreleri sorulur, özet yeni sayfaya yazılır. */
export function ScenarioSummaryDialog({ ctl, onClose }: P) {
  const [res, setRes] = useState(() => {
    const a = ctl.sel.active;
    return ctl.wb.getFormula(ctl.sheet, a.row, a.col) ? addr(a.row, a.col) : "";
  });
  const ok = () => {
    const list = res.trim() ? parseCells(res, parse) : [];
    if (list === null) return ctl.alert("Sonuç hücreleri başvurusu geçerli değil.");
    onClose();
    scenarioSummary(ctl, list);
  };
  return (
    <SimpleDialog title="Senaryo Özeti" onClose={onClose} onOk={ok} width={360}>
      <fieldset className="oh-fieldset">
        <legend>Rapor türü</legend>
        <label className="tb-radio">
          <input type="radio" checked readOnly />
          <span>Senaryo özeti</span>
        </label>
      </fieldset>
      <Field label="Sonuç hücreleri:">
        <input type="text" value={res} onChange={(e) => setRes(e.target.value)} aria-label="Sonuç hücreleri" data-autofocus="true" />
      </Field>
    </SimpleDialog>
  );
}

/** Veri Tablosu: satır ve/veya sütun giriş hücresi. */
export function DataTableDialog({ ctl, onClose }: P) {
  const [rowIn, setRowIn] = useState("");
  const [colIn, setColIn] = useState("");
  const cell = (s: string): CellPos | undefined | null => {
    if (!s.trim()) return undefined;
    const r = parse(s.trim().replace(/^.*!/, ""));
    return r && r.r1 === r.r2 && r.c1 === r.c2 ? { row: r.r1, col: r.c1 } : null;
  };
  const ok = () => {
    const ri = cell(rowIn);
    const ci = cell(colIn);
    if (ri === null || ci === null || (!ri && !ci)) return ctl.alert("Giriş hücresi başvurusu geçerli değil.");
    const err = createDataTable(ctl, ctl.range, ri, ci);
    if (err) return ctl.alert(err);
    onClose();
  };
  return (
    <SimpleDialog title="Veri Tablosu" onClose={onClose} onOk={ok} width={320}>
      <Field label="Satır giriş hücresi:">
        <input type="text" value={rowIn} onChange={(e) => setRowIn(e.target.value)} aria-label="Satır giriş hücresi" data-autofocus="true" />
      </Field>
      <Field label="Sütun giriş hücresi:">
        <input type="text" value={colIn} onChange={(e) => setColIn(e.target.value)} aria-label="Sütun giriş hücresi" />
      </Field>
    </SimpleDialog>
  );
}
