/** Mini grafik iletişim kutuları: Mini Grafikler Oluştur / Düzenle, Tek Mini Grafik Verisi, Gizli ve Boş Hücre Ayarları. */
import { useMemo, useState } from "react";
import { Field, FieldSet } from "@officehere/ui";
import type { Controller } from "../controller";
import { rectToA1 } from "../model/rect";
import type { SparklineGroup } from "../model/types";
import { Check, Radio, SimpleDialog } from "../dialogs/common";
import { sparkAt } from "./geometry";
import { createSparklines, editGroupData, editItemData, groupDataText, groupLocationText } from "./sparkOps";

type P = { ctl: Controller; onClose: () => void; arg?: unknown };

const abs = (a1: string) => a1.replace(/([A-Z]+)(\d+)/g, "$$$1$$$2");

/** Ekle > Mini Grafikler (arg.type) ya da Veri Düzenle (arg.edit = grup kimliği). */
export function SparklineDialog({ ctl, onClose, arg }: P) {
  const a = (arg ?? {}) as { type?: SparklineGroup["type"]; edit?: string };
  const sheet = ctl.sheet;
  const group = a.edit ? sheet.sparklines.find((g) => g.id === a.edit) : undefined;
  const init = useMemo(() => {
    if (group) return { data: groupDataText(ctl.wb, sheet, group), loc: abs(groupLocationText(group)) };
    const r = ctl.range;
    const single = r.r1 === r.r2 && r.c1 === r.c2;
    // Seçim veri içeriyorsa veri aralığı, tek boş hücreyse konum olarak önerilir
    if (single && ctl.wb.isEmpty(sheet, r.r1, r.c1)) return { data: "", loc: abs(rectToA1(r)) };
    return { data: rectToA1(r), loc: "" };
  }, []); // eslint-disable-line react-hooks/exhaustive-deps
  const [data, setData] = useState(init.data);
  const [loc, setLoc] = useState(init.loc);
  const ok = () => {
    const err = group ? editGroupData(ctl, group.id, data, loc) : createSparklines(ctl, a.type ?? "line", data, loc);
    if (err) {
      ctl.alert(err);
      return;
    }
    onClose();
  };
  const focusData = !init.data;
  return (
    <SimpleDialog title={group ? "Mini Grafikleri Düzenle" : "Mini Grafikler Oluştur"} onClose={onClose} onOk={ok} width={400}>
      <FieldSet legend="İstediğiniz verileri seçin">
        <Field label="Veri Aralığı:" inline>
          <input type="text" value={data} onChange={(e) => setData(e.target.value)} data-autofocus={focusData ? "true" : undefined} style={{ flex: 1 }} aria-label="Veri Aralığı" />
        </Field>
      </FieldSet>
      <FieldSet legend="Mini grafiklerin yerleştirilmesini istediğiniz konumu seçin">
        <Field label="Konum Aralığı:" inline>
          <input type="text" value={loc} onChange={(e) => setLoc(e.target.value)} data-autofocus={focusData ? undefined : "true"} style={{ flex: 1 }} aria-label="Konum Aralığı" />
        </Field>
      </FieldSet>
    </SimpleDialog>
  );
}

/** Tek Mini Grafiğin Verilerini Düzenle. */
export function SparklineItemDialog({ ctl, onClose }: P) {
  const { row, col } = ctl.sel.active;
  const s = sparkAt(ctl.sheet, row, col);
  const [data, setData] = useState(s?.item.range ?? "");
  if (!s) return null;
  const ok = () => {
    const err = editItemData(ctl, row, col, data);
    if (err) return ctl.alert(err);
    onClose();
  };
  return (
    <SimpleDialog title="Mini Grafik Verilerini Düzenle" onClose={onClose} onOk={ok} width={400}>
      <Field label={`${rectToA1({ r1: row, c1: col, r2: row, c2: col })} hücresindeki mini grafik için veri kaynağını seçin:`}>
        <input type="text" value={data} onChange={(e) => setData(e.target.value)} data-autofocus="true" aria-label="Veri Aralığı" />
      </Field>
    </SimpleDialog>
  );
}

/** Gizli ve Boş Hücre Ayarları. */
export function SparklineHiddenDialog({ ctl, onClose, arg }: P) {
  const g = ctl.sheet.sparklines.find((x) => x.id === (arg as { group: string }).group);
  const [empty, setEmpty] = useState<NonNullable<SparklineGroup["emptyAs"]>>(g?.emptyAs ?? "gap");
  const [hidden, setHidden] = useState(!!g?.showHidden);
  if (!g) return null;
  const ok = () => {
    const sheet = ctl.sheet;
    ctl.ops.sheetOp(sheet, "Gizli ve Boş Hücre Ayarları", () => {
      const live = sheet.sparklines.find((x) => x.id === g.id);
      if (!live) return;
      live.emptyAs = empty === "gap" ? undefined : empty;
      live.showHidden = hidden || undefined;
      if (!live.emptyAs) delete live.emptyAs;
      if (!live.showHidden) delete live.showHidden;
    });
    ctl.emit();
    onClose();
  };
  return (
    <SimpleDialog title="Gizli ve Boş Hücre Ayarları" onClose={onClose} onOk={ok} width={380}>
      <FieldSet legend="Boş hücreleri farklı göster:">
        <Radio name="spempty" value="gap" current={empty} onChange={setEmpty}>
          Boşluklar
        </Radio>
        <Radio name="spempty" value="zero" current={empty} onChange={setEmpty}>
          Sıfır
        </Radio>
        <Radio name="spempty" value="span" current={empty} onChange={setEmpty}>
          Veri noktalarını çizgiyle bağla
        </Radio>
      </FieldSet>
      <Check checked={hidden} onChange={setHidden}>
        Gizli satır ve sütunlardaki verileri göster
      </Check>
    </SimpleDialog>
  );
}

/** Dikey Eksen için Özel Değer (en küçük/en büyük). */
export function SparklineAxisDialog({ ctl, onClose, arg }: P) {
  const a = arg as { group: string; which: "min" | "max" };
  const g = ctl.sheet.sparklines.find((x) => x.id === a.group);
  const cur = a.which === "min" ? g?.minValue : g?.maxValue;
  const [text, setText] = useState(cur !== undefined ? String(cur).replace(".", ",") : "0,0");
  if (!g) return null;
  const ok = () => {
    const v = Number(text.trim().replace(/\./g, "").replace(",", "."));
    if (!Number.isFinite(v)) return ctl.alert("Geçerli bir sayı girin.");
    const sheet = ctl.sheet;
    ctl.ops.sheetOp(sheet, "Dikey Eksen Ayarları", () => {
      const live = sheet.sparklines.find((x) => x.id === g.id);
      if (!live) return;
      if (a.which === "min") {
        live.minAxis = "custom";
        live.minValue = v;
      } else {
        live.maxAxis = "custom";
        live.maxValue = v;
      }
    });
    ctl.emit();
    onClose();
  };
  return (
    <SimpleDialog title="Mini Grafik Dikey Ekseni Ayarları" onClose={onClose} onOk={ok} width={360}>
      <Field label={`Dikey eksen için ${a.which === "min" ? "en küçük" : "en büyük"} değeri girin:`}>
        <input type="text" value={text} onChange={(e) => setText(e.target.value)} data-autofocus="true" aria-label="Eksen değeri" />
      </Field>
    </SimpleDialog>
  );
}
