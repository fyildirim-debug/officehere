/** PivotTable iletişim kutuları: PivotTable Oluştur, Değer Alanı Ayarları, PivotTable'ı Taşı, Veri Kaynağını Değiştir, PivotTable Seçenekleri. */
import { useMemo, useState } from "react";
import { Field, FieldSet } from "@officehere/ui";
import { quoteSheetName } from "@officehere/formula";
import type { Controller } from "../controller";
import { parseRangeLoose } from "../controller";
import { rectToA1 } from "../model/rect";
import { currentRegion } from "../model/selection";
import type { PivotFn, Rect } from "../model/types";
import type { SheetModel } from "../model/workbook";
import { Check, Radio, SimpleDialog } from "../dialogs/common";
import { FN_LABEL, valueCaption } from "./compute";
import { applyPivot, createPivot, movePivot, readSource } from "./pivotOps";
import { NUMBER_FORMATS } from "../model/styles";

type P = { ctl: Controller; onClose: () => void; arg?: unknown };

function absRef(sheet: SheetModel, r: Rect): string {
  return `${quoteSheetName(sheet.name)}!${rectToA1(r).replace(/([A-Z]+)(\d+)/g, "$$$1$$$2")}`;
}

/** "Sayfa1!$A$1:$D$20" gibi metni sayfa + aralığa çevirir. */
export function parseSheetRange(ctl: Controller, text: string): { sheet: SheetModel; rect: Rect } | null {
  const t = text.trim().replace(/^=/, "");
  if (!t) return null;
  const bang = t.lastIndexOf("!");
  let sheet: SheetModel | undefined = ctl.sheet;
  let body = t;
  if (bang > 0) {
    let name = t.slice(0, bang);
    if (name.startsWith("'")) name = name.slice(1, -1).replace(/''/g, "'");
    sheet = ctl.wb.sheetByName(name);
    body = t.slice(bang + 1);
  }
  if (!sheet) return null;
  let r = parseRangeLoose(body);
  if (!r) {
    // Tablo adı
    const tb = ctl.wb.sheets.flatMap((s) => s.tables.map((x) => ({ s, x }))).find((o) => o.x.name.toLocaleUpperCase("tr-TR") === body.toLocaleUpperCase("tr-TR"));
    if (tb) {
      sheet = tb.s;
      r = tb.x.ref;
    }
  }
  return r ? { sheet, rect: r } : null;
}

export function CreatePivotDialog({ ctl, onClose }: P) {
  const initial = useMemo(() => {
    const r = ctl.range;
    const t = ctl.sheet.tableAt(ctl.sel.active.row, ctl.sel.active.col);
    if (t) return t.name;
    const region = r.r1 === r.r2 && r.c1 === r.c2 ? currentRegion(ctl.wb, ctl.sheet, ctl.sel.active) : r;
    return absRef(ctl.sheet, region);
  }, []); // eslint-disable-line react-hooks/exhaustive-deps
  const [src, setSrc] = useState(initial);
  const [where, setWhere] = useState<"new" | "existing">("new");
  const [loc, setLoc] = useState("");
  const ok = () => {
    const s = parseSheetRange(ctl, src);
    if (!s || (s.rect.r1 === s.rect.r2 && s.rect.c1 === s.rect.c2)) {
      ctl.alert("Başvuru geçerli değil.");
      return;
    }
    let target: { sheet: SheetModel; row: number; col: number } | null = null;
    if (where === "existing") {
      const l = parseSheetRange(ctl, loc);
      if (!l) {
        ctl.alert("Hedef başvuru geçerli değil.");
        return;
      }
      target = { sheet: l.sheet, row: l.rect.r1, col: l.rect.c1 };
    }
    onClose();
    createPivot(ctl, s.sheet, s.rect, target);
  };
  return (
    <SimpleDialog title="PivotTable Oluştur" onClose={onClose} onOk={ok} width={430}>
      <FieldSet legend="Çözümlemek istediğiniz verileri seçin">
        <Field label="Tablo/Aralık:" inline>
          <input type="text" value={src} onChange={(e) => setSrc(e.target.value)} data-autofocus="true" style={{ flex: 1 }} aria-label="Tablo/Aralık" />
        </Field>
      </FieldSet>
      <FieldSet legend="PivotTable raporunun yerleştirileceği yeri seçin">
        <Radio name="pvwhere" value="new" current={where} onChange={setWhere}>
          Yeni Çalışma Sayfası
        </Radio>
        <Radio name="pvwhere" value="existing" current={where} onChange={(v) => (setWhere(v), !loc && setLoc(absRef(ctl.sheet, { r1: ctl.sel.active.row, c1: ctl.sel.active.col, r2: ctl.sel.active.row, c2: ctl.sel.active.col })))}>
          Var Olan Çalışma Sayfası
        </Radio>
        <Field label="Konum:" inline>
          <input type="text" value={loc} disabled={where !== "existing"} onChange={(e) => setLoc(e.target.value)} style={{ flex: 1 }} aria-label="Konum" />
        </Field>
      </FieldSet>
    </SimpleDialog>
  );
}

/** Değer Alanı Ayarları: özetleme işlevi, özel ad, sayı biçimi. */
export function ValueFieldDialog({ ctl, onClose, arg }: P) {
  const a = arg as { pivotId: string; index: number };
  const sheet = ctl.sheet;
  const def = sheet.pivots.find((p) => p.id === a.pivotId);
  const src = def ? readSource(ctl.wb, def) : null;
  const vf = def?.values[a.index];
  const headers = src && typeof src !== "string" ? src.headers : [];
  const [fn, setFn] = useState<PivotFn>(vf?.fn ?? "sum");
  const [name, setName] = useState(vf ? valueCaption(vf, headers) : "");
  const [fmt, setFmt] = useState(vf?.numFmt ?? "");
  if (!def || !vf) return null;
  const ok = () => {
    const defaultOld = valueCaption({ ...vf, name: undefined }, headers);
    const custom = name.trim() && name.trim() !== defaultOld ? name.trim() : undefined;
    // İşlev değiştiyse ve ad varsayılansa yeni varsayılan ad kullanılır
    const finalName = custom && custom !== valueCaption({ ...vf, fn, name: undefined }, headers) ? custom : undefined;
    applyPivot(
      ctl,
      sheet,
      def.id,
      (d) => {
        d.values[a.index] = { ...d.values[a.index], fn, name: finalName, numFmt: fmt || undefined };
      },
      "Değer Alanı Ayarları",
    );
    onClose();
  };
  return (
    <SimpleDialog title="Değer Alanı Ayarları" onClose={onClose} onOk={ok} width={380}>
      <p className="tb-dlg-note">Kaynak Adı: {headers[vf.field]}</p>
      <Field label="Özel Ad:">
        <input type="text" value={name} onChange={(e) => setName(e.target.value)} data-autofocus="true" aria-label="Özel Ad" />
      </Field>
      <FieldSet legend="Değer alanını özetleme ölçütü">
        <div className="tb-list tb-pv-fnlist" role="listbox" aria-label="Özetleme işlevi">
          {(Object.keys(FN_LABEL) as PivotFn[]).map((k) => (
            <div key={k} role="option" aria-selected={k === fn} className={`tb-list-item${k === fn ? " is-active" : ""}`} onClick={() => setFn(k)} onDoubleClick={() => (setFn(k), setTimeout(ok))}>
              {FN_LABEL[k]}
            </div>
          ))}
        </div>
      </FieldSet>
      <Field label="Sayı Biçimi:" inline>
        <select value={fmt} onChange={(e) => setFmt(e.target.value)} aria-label="Sayı Biçimi">
          <option value="">Kaynak biçimi</option>
          {NUMBER_FORMATS.filter((f) => f.id !== "text").map((f) => (
            <option key={f.id} value={f.code === "General" ? "General" : f.code}>
              {f.label}
            </option>
          ))}
        </select>
      </Field>
    </SimpleDialog>
  );
}

export function MovePivotDialog({ ctl, onClose, arg }: P) {
  const a = arg as { pivotId: string };
  const sheet = ctl.sheet;
  const def = sheet.pivots.find((p) => p.id === a.pivotId);
  const [loc, setLoc] = useState(def ? absRef(sheet, { r1: def.anchor.row, c1: def.anchor.col, r2: def.anchor.row, c2: def.anchor.col }) : "");
  if (!def) return null;
  const ok = () => {
    const l = parseSheetRange(ctl, loc);
    if (!l) return ctl.alert("Başvuru geçerli değil.");
    if (l.sheet !== sheet) return ctl.alert("PivotTable yalnız aynı sayfa içinde taşınabilir.");
    onClose();
    movePivot(ctl, sheet, def.id, { row: l.rect.r1, col: l.rect.c1 });
  };
  return (
    <SimpleDialog title="PivotTable'ı Taşı" onClose={onClose} onOk={ok} width={360}>
      <Field label="Konum:">
        <input type="text" value={loc} onChange={(e) => setLoc(e.target.value)} data-autofocus="true" aria-label="Konum" />
      </Field>
    </SimpleDialog>
  );
}

export function ChangePivotSourceDialog({ ctl, onClose, arg }: P) {
  const a = arg as { pivotId: string };
  const sheet = ctl.sheet;
  const def = sheet.pivots.find((p) => p.id === a.pivotId);
  const srcSheet = def ? ctl.wb.sheetById(def.sourceSheet) : undefined;
  const [src, setSrc] = useState(def && srcSheet ? absRef(srcSheet, def.source) : "");
  if (!def) return null;
  const ok = () => {
    const s = parseSheetRange(ctl, src);
    if (!s) return ctl.alert("Başvuru geçerli değil.");
    onClose();
    applyPivot(
      ctl,
      sheet,
      def.id,
      (d) => {
        d.sourceSheet = s.sheet.id;
        d.source = s.rect;
      },
      "Veri Kaynağını Değiştir",
    );
  };
  return (
    <SimpleDialog title="PivotTable Veri Kaynağını Değiştir" onClose={onClose} onOk={ok} width={400}>
      <Field label="Tablo/Aralık:">
        <input type="text" value={src} onChange={(e) => setSrc(e.target.value)} data-autofocus="true" aria-label="Tablo/Aralık" />
      </Field>
    </SimpleDialog>
  );
}

export function PivotOptionsDialog({ ctl, onClose, arg }: P) {
  const a = arg as { pivotId: string };
  const sheet = ctl.sheet;
  const def = sheet.pivots.find((p) => p.id === a.pivotId);
  const [name, setName] = useState(def?.name ?? "");
  const [grandRows, setGrandRows] = useState(!!def?.grandRows);
  const [grandCols, setGrandCols] = useState(!!def?.grandCols);
  if (!def) return null;
  const ok = () => {
    const n = name.trim();
    if (!n) return ctl.alert("PivotTable adı boş olamaz.");
    const clash = ctl.wb.sheets.some((s) => s.pivots.some((p) => p.id !== def.id && p.name.toLocaleUpperCase("tr-TR") === n.toLocaleUpperCase("tr-TR")));
    if (clash) return ctl.alert("Bu ada sahip bir PivotTable zaten var.");
    onClose();
    applyPivot(
      ctl,
      sheet,
      def.id,
      (d) => {
        d.name = n;
        d.grandRows = grandRows;
        d.grandCols = grandCols;
      },
      "PivotTable Seçenekleri",
    );
  };
  return (
    <SimpleDialog title="PivotTable Seçenekleri" onClose={onClose} onOk={ok} width={380}>
      <Field label="PivotTable Adı:">
        <input type="text" value={name} onChange={(e) => setName(e.target.value)} data-autofocus="true" aria-label="PivotTable Adı" />
      </Field>
      <FieldSet legend="Genel Toplamlar">
        {/* Excel adlandırması: "satırlar için" = her satırın toplamı (sağdaki Genel Toplam sütunu) */}
        <Check checked={grandCols} onChange={setGrandCols}>
          Satırlar için genel toplamları göster
        </Check>
        <Check checked={grandRows} onChange={setGrandRows}>
          Sütunlar için genel toplamları göster
        </Check>
      </FieldSet>
    </SimpleDialog>
  );
}
