/** Veri araçları iletişim kutuları: Gruplandır (satır/sütun), Alt Toplam, Hedef Ara ve sonucu. */
import { useState } from "react";
import { Dialog, DialogButton, Field } from "@officehere/ui";
import { addr } from "@officehere/formula";
import type { Controller } from "../controller";
import { parseRangeLoose } from "../controller";
import { addSubtotals, groupSelection, removeSubtotals, ungroupSelection } from "../actions/outline";
import { applyGoalSeek, runGoalSeek } from "../actions/whatIf";
import { displayText } from "../model/display";
import { SUBTOTAL_FUNCS } from "../model/outline";
import { currentRegion } from "../model/selection";
import { Check, Radio, SimpleDialog, parseTrNumber } from "./common";

type P = { ctl: Controller; onClose: () => void; arg?: unknown };

/** Gruplandır / Grubu Çöz: Satırlar mı Sütunlar mı? */
export function GroupAxisDialog({ ctl, onClose, arg }: P) {
  const mode = (arg as { mode: "group" | "ungroup" }).mode;
  const [axis, setAxis] = useState<"row" | "col">("row");
  return (
    <SimpleDialog
      title={mode === "group" ? "Gruplandır" : "Grubu Çöz"}
      width={240}
      onClose={onClose}
      onOk={() => {
        onClose();
        if (mode === "group") groupSelection(ctl, axis);
        else ungroupSelection(ctl, axis);
      }}
    >
      <Radio name="ga" value="row" current={axis} onChange={setAxis}>
        Satırlar
      </Radio>
      <Radio name="ga" value="col" current={axis} onChange={setAxis}>
        Sütunlar
      </Radio>
    </SimpleDialog>
  );
}

/** Veri > Alt Toplam. */
export function SubtotalDialog({ ctl, onClose }: P) {
  const { wb, sheet } = ctl;
  const region = (() => {
    const r = ctl.range;
    return r.r1 === r.r2 && r.c1 === r.c2 ? currentRegion(wb, sheet, ctl.sel.active) : r;
  })();
  const headers = Array.from({ length: region.c2 - region.c1 + 1 }, (_, i) => displayText(wb, sheet, region.r1, region.c1 + i) || `Sütun ${addr(0, region.c1 + i).replace(/\d+$/, "")}`);
  const [groupCol, setGroupCol] = useState(region.c1);
  const [fn, setFn] = useState(9);
  const lastNumeric = (() => {
    for (let c = region.c2; c >= region.c1; c--) if (typeof wb.getValue(sheet, region.r1 + 1, c) === "number") return c;
    return region.c2;
  })();
  const [cols, setCols] = useState<number[]>([lastNumeric]);
  const [replace, setReplace] = useState(true);
  const ok = () => {
    if (region.r2 <= region.r1) {
      ctl.alert("Alt toplam için başlık satırı ve veri içeren bir aralık seçin.");
      return;
    }
    if (!cols.length) {
      ctl.alert("Alt toplam eklenecek en az bir sütun seçin.");
      return;
    }
    onClose();
    addSubtotals(ctl, { range: region, groupCol, fn, cols }, replace);
  };
  return (
    <Dialog
      open
      title="Alt Toplam"
      width={360}
      onClose={onClose}
      footer={
        <>
          <DialogButton onClick={() => (onClose(), removeSubtotals(ctl, region))}>Tümünü Kaldır</DialogButton>
          <DialogButton primary onClick={ok}>
            Tamam
          </DialogButton>
          <DialogButton onClick={onClose}>İptal</DialogButton>
        </>
      }
    >
      <Field label="Her değişiklikte:">
        <select value={groupCol} onChange={(e) => setGroupCol(Number(e.target.value))} aria-label="Her değişiklikte">
          {headers.map((h, i) => (
            <option key={i} value={region.c1 + i}>
              {h}
            </option>
          ))}
        </select>
      </Field>
      <Field label="Kullanılacak işlev:">
        <select value={fn} onChange={(e) => setFn(Number(e.target.value))} aria-label="Kullanılacak işlev">
          {SUBTOTAL_FUNCS.map((f) => (
            <option key={f.id} value={f.id}>
              {f.label}
            </option>
          ))}
        </select>
      </Field>
      <div style={{ margin: "6px 0 2px" }}>Alt toplam ekle:</div>
      <div className="tb-list" style={{ height: 120 }}>
        {headers.map((h, i) => (
          <div key={i} className="tb-list-item">
            <Check checked={cols.includes(region.c1 + i)} onChange={(v) => setCols(v ? [...cols, region.c1 + i] : cols.filter((c) => c !== region.c1 + i))}>
              {h}
            </Check>
          </div>
        ))}
      </div>
      <div style={{ marginTop: 8 }}>
        <Check checked={replace} onChange={setReplace}>
          Geçerli alt toplamları değiştir
        </Check>
        <Check checked onChange={() => undefined}>
          Verilerin altında özet
        </Check>
      </div>
    </Dialog>
  );
}

/** Durum Çözümlemesi > Hedef Ara. */
export function GoalSeekDialog({ ctl, onClose }: P) {
  const a = ctl.sel.active;
  const [setCell, setSetCell] = useState(addr(a.row, a.col));
  const [toValue, setToValue] = useState("");
  const [byCell, setByCell] = useState("");
  const ok = () => {
    const t = parseRangeLoose(setCell.replace(/\$/g, ""));
    const c = parseRangeLoose(byCell.replace(/\$/g, ""));
    const goal = parseTrNumber(toValue);
    if (!t || t.r1 !== t.r2 || t.c1 !== t.c2 || !ctl.wb.getFormula(ctl.sheet, t.r1, t.c1)) {
      ctl.alert("Ayarlanacak hücre bir formül içermelidir.");
      return;
    }
    if (!c || c.r1 !== c.r2 || c.c1 !== c.c2 || ctl.wb.getFormula(ctl.sheet, c.r1, c.c1)) {
      ctl.alert("Değişecek hücre, sabit bir değer içeren tek bir hücre olmalıdır.");
      return;
    }
    if (!toValue.trim() || !Number.isFinite(goal)) {
      ctl.alert("Sonuç değeri bir sayı olmalıdır.");
      return;
    }
    const target = { row: t.r1, col: t.c1 };
    const changing = { row: c.r1, col: c.c1 };
    const res = runGoalSeek(ctl, target, goal, changing);
    onClose();
    // Değer önce uygulanır; durum kutusunda İptal geri alır (Excel gibi)
    applyGoalSeek(ctl, changing, res.value);
    ctl.openDialog("goalSeekStatus", { cell: setCell.toUpperCase(), goal, result: res.result, found: res.found });
  };
  return (
    <SimpleDialog title="Hedef Ara" width={300} onClose={onClose} onOk={ok}>
      <Field label="Ayarlanacak hücre:">
        <input type="text" value={setCell} onChange={(e) => setSetCell(e.target.value)} aria-label="Ayarlanacak hücre" />
      </Field>
      <Field label="Sonuç:">
        <input type="text" value={toValue} onChange={(e) => setToValue(e.target.value)} aria-label="Sonuç" data-autofocus="true" />
      </Field>
      <Field label="Değişecek hücre:">
        <input type="text" value={byCell} onChange={(e) => setByCell(e.target.value)} aria-label="Değişecek hücre" />
      </Field>
    </SimpleDialog>
  );
}

export function GoalSeekStatusDialog({ ctl, onClose, arg }: P) {
  const a = arg as { cell: string; goal: number; result: number; found: boolean };
  const fmt = (n: number) => n.toLocaleString("tr-TR", { maximumFractionDigits: 6 });
  return (
    <Dialog
      open
      title="Hedef Arama Durumu"
      width={340}
      onClose={() => (onClose(), ctl.ops.undo())}
      footer={
        <>
          <DialogButton primary onClick={onClose}>
            Tamam
          </DialogButton>
          <DialogButton onClick={() => (onClose(), ctl.ops.undo())}>İptal</DialogButton>
        </>
      }
    >
      <p className="oh-dialog-message">
        {a.cell} hücresi için hedef arama {a.found ? "bir çözüm buldu." : "bir çözüm bulamadı."}
      </p>
      <p className="oh-dialog-message">Hedef değer: {fmt(a.goal)}</p>
      <p className="oh-dialog-message">Geçerli değer: {fmt(a.result)}</p>
    </Dialog>
  );
}
