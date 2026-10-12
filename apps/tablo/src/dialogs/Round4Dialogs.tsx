/** 4. tur iletişim kutuları: Gelişmiş Filtre, Çalışma Kitabını Koru / Korumayı Kaldır, Dilimleyici Ekle. */
import { useState } from "react";
import { Field } from "@officehere/ui";
import type { Controller } from "../controller";
import { parseRangeLoose } from "../controller";
import { formatRange } from "@officehere/formula";
import { currentRegion } from "../model/selection";
import { applyAdvancedFilter } from "../actions/advFilter";
import { checkProtectionPassword, hashPassword } from "../model/protection";
import { insertSlicers, slicerSourceAt } from "../actions/slicers";
import { Check, Radio, SimpleDialog } from "./common";

type P = { ctl: Controller; onClose: () => void; arg?: unknown };

const parse = (s: string) => (s.trim() ? parseRangeLoose(s.trim().replace(/\$/g, "").replace(/^.*!/, "")) : null);

export function AdvancedFilterDialog({ ctl, onClose }: P) {
  const sheet = ctl.sheet;
  const r = ctl.range;
  const single = r.r1 === r.r2 && r.c1 === r.c2;
  const initial = sheet.advFilter?.range ?? (single ? currentRegion(ctl.wb, sheet, ctl.sel.active) : r);
  const [mode, setMode] = useState<"inplace" | "copy">("inplace");
  const [list, setList] = useState(`$${formatRange(initial).replace(/([A-Z]+)(\d+)/g, "$1$$$2").replace(":", ":$")}`);
  const [crit, setCrit] = useState(sheet.advFilter?.criteria ? formatRange(sheet.advFilter.criteria) : "");
  const [dest, setDest] = useState("");
  const [unique, setUnique] = useState(!!sheet.advFilter?.unique);
  const ok = () => {
    const l = parse(list);
    if (!l) return ctl.alert("Liste aralığı başvurusu geçerli değil.");
    const c = crit.trim() ? parse(crit) : null;
    if (crit.trim() && !c) return ctl.alert("Ölçüt aralığı başvurusu geçerli değil.");
    let copyTo: { row: number; col: number } | undefined;
    if (mode === "copy") {
      const d = parse(dest);
      if (!d) return ctl.alert("Hedef başvurusu geçerli değil.");
      copyTo = { row: d.r1, col: d.c1 };
    }
    const err = applyAdvancedFilter(ctl, { list: l, criteria: c, copyTo, unique });
    if (err) return ctl.alert(err);
    onClose();
  };
  return (
    <SimpleDialog title="Gelişmiş Filtre" onClose={onClose} onOk={ok} width={380}>
      <fieldset className="oh-fieldset">
        <legend>Eylem</legend>
        <Radio name="af-mode" value="inplace" current={mode} onChange={setMode}>
          Listeyi yerinde filtrele
        </Radio>
        <Radio name="af-mode" value="copy" current={mode} onChange={setMode}>
          Başka bir yere kopyala
        </Radio>
      </fieldset>
      <Field label="Liste aralığı:" inline>
        <input type="text" value={list} onChange={(e) => setList(e.target.value)} aria-label="Liste aralığı" />
      </Field>
      <Field label="Ölçüt aralığı:" inline>
        <input type="text" value={crit} onChange={(e) => setCrit(e.target.value)} aria-label="Ölçüt aralığı" data-autofocus="true" />
      </Field>
      <Field label="Hedef:" inline>
        <input type="text" value={dest} disabled={mode !== "copy"} onChange={(e) => setDest(e.target.value)} aria-label="Hedef" />
      </Field>
      <Check checked={unique} onChange={setUnique}>
        Yalnızca benzersiz kayıtlar
      </Check>
    </SimpleDialog>
  );
}

/** Çalışma Kitabını Koru (yapı) ya da korumayı kaldır (parola sorulur). */
export function ProtectWorkbookDialog({ ctl, onClose }: P) {
  const wb = ctl.wb;
  const on = !wb.protection?.structure;
  const [pw, setPw] = useState("");
  const [pw2, setPw2] = useState("");
  const [step, setStep] = useState<1 | 2>(1);
  const [busy, setBusy] = useState(false);
  const done = () => {
    ctl.dirty = true;
    ctl.pristine = false;
    ctl.dataVersion++;
    ctl.emit();
    onClose();
  };
  const ok = async () => {
    if (busy) return;
    if (on) {
      if (pw && step === 1) return setStep(2);
      if (pw && pw !== pw2) {
        ctl.alert("Onay parolası aynı değil.");
        setStep(1);
        setPw2("");
        return;
      }
      setBusy(true);
      const hash = pw ? await hashPassword(pw) : undefined;
      wb.protection = { structure: true, ...(hash ? { hash } : {}) };
      return done();
    }
    setBusy(true);
    const good = await checkProtectionPassword(wb.protection!, pw);
    setBusy(false);
    if (!good) {
      ctl.alert("Girdiğiniz parola doğru değil. CAPS LOCK tuşunun kapalı olduğundan ve doğru büyük/küçük harf kullandığınızdan emin olun.");
      return;
    }
    wb.protection = null;
    done();
  };
  const needsPw = !on && !!(wb.protection?.hash || wb.protection?.legacy);
  if (!on && !needsPw) {
    queueMicrotask(() => {
      if (wb.protection) {
        wb.protection = null;
        done();
      }
    });
    return null;
  }
  return (
    <SimpleDialog title={on ? (step === 2 ? "Parolayı Onayla" : "Yapıyı ve Pencereleri Koru") : "Çalışma Kitabının Korumasını Kaldır"} onClose={onClose} onOk={() => void ok()} width={360} disabled={busy}>
      {step === 1 ? (
        <Field label={on ? "Parola (isteğe bağlı):" : "Parola:"}>
          <input type="password" value={pw} onChange={(e) => setPw(e.target.value)} aria-label="Parola" data-autofocus="true" />
        </Field>
      ) : (
        <Field label="Parolayı yeniden girin:">
          <input type="password" value={pw2} onChange={(e) => setPw2(e.target.value)} aria-label="Parolayı yeniden girin" data-autofocus="true" />
        </Field>
      )}
      {on && step === 1 && (
        <fieldset className="oh-fieldset">
          <legend>Koruma</legend>
          <Check checked onChange={() => undefined}>
            Yapı
          </Check>
        </fieldset>
      )}
      {on && <p className="tb-fc-note">Yapı korunurken sayfa eklenemez, silinemez, yeniden adlandırılamaz, taşınamaz, gizlenemez ya da gösterilemez. Koruma dosyaya kaydedilir.</p>}
      {busy && <p className="tb-fc-note">Parola denetleniyor...</p>}
    </SimpleDialog>
  );
}

/** Dilimleyici Ekle: etkin tablo ya da özet tablonun alanları onay kutularıyla. */
export function InsertSlicerDialog({ ctl, onClose }: P) {
  const src = slicerSourceAt(ctl);
  const fields = !src ? [] : src.kind === "table" ? src.table.columns : src.fields;
  const [checked, setChecked] = useState<Set<number>>(new Set());
  if (!src) {
    queueMicrotask(() => {
      onClose();
      ctl.alert("Dilimleyici eklemek için bir tablonun ya da özet tablonun içindeki bir hücreyi seçin.");
    });
    return null;
  }
  const ok = () => {
    const list = [...checked].sort((a, b) => a - b).map((i) => ({ name: fields[i], field: i }));
    if (!list.length) return onClose();
    onClose();
    insertSlicers(ctl, src.kind === "table" ? { kind: "table", tableId: src.table.id } : { kind: "pivot", pivotId: src.pivotId }, list);
  };
  return (
    <SimpleDialog title="Dilimleyici Ekle" onClose={onClose} onOk={ok} width={300}>
      <div className="tb-list" role="listbox" aria-label="Alanlar" style={{ height: 220, padding: 4 }}>
        {fields.map((f, i) => (
          <Check key={i} checked={checked.has(i)} onChange={(v) => setChecked((cur) => {
            const n = new Set(cur);
            if (v) n.add(i);
            else n.delete(i);
            return n;
          })}>
            {f || `Sütun ${i + 1}`}
          </Check>
        ))}
      </div>
    </SimpleDialog>
  );
}
