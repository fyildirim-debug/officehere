/** Otomatik Filtre açılır menüsü: sıralama, değer onay kutuları, arama, sayı/metin filtreleri. */
import { useMemo, useState } from "react";
import { ArrowSortDownLines20Regular, ArrowSortUpLines20Regular, FilterDismiss20Regular } from "@fluentui/react-icons";
import { DialogButton, Menu, MenuItem, MenuSeparator, Popover } from "@officehere/ui";
import type { Controller } from "../controller";
import * as D from "../actions/data";
import { textMatcher } from "@officehere/formula";
import { displayText } from "../model/display";
import type { FilterCriteria } from "../model/types";

/** Listede en çok bu kadar öğe gösterilir (Excel gibi; büyük sütunlarda menü hızlı açılır). */
const MAX_ITEMS = 10000;

export function FilterMenu({ ctl, row, col, x, y, onClose }: { ctl: Controller; row: number; col: number; x: number; y: number; onClose: () => void }) {
  const sheet = ctl.sheet;
  const f = D.filterAt(sheet, row, col);
  const { values, hasBlank } = useMemo(() => (f ? D.columnValues(ctl.wb, sheet, f, col) : { values: [], hasBlank: false }), [f, col]); // eslint-disable-line react-hooks/exhaustive-deps
  const cur = f?.columns[col - (f?.range.c1 ?? 0)];
  const [checked, setChecked] = useState<Set<string>>(() => {
    if (cur?.kind === "values") return new Set([...cur.values, ...(cur.blanks ? ["\u0000blank"] : [])]);
    return new Set([...values, "\u0000blank"]);
  });
  const [q, setQ] = useState("");
  if (!f) return null;
  const header = displayText(ctl.wb, sheet, row, col) || `Sütun ${col - f.range.c1 + 1}`;
  const numeric = (() => {
    let n = 0;
    let t = 0;
    for (let r = f.range.r1 + 1; r <= Math.min(f.range.r2, f.range.r1 + 200); r++) {
      const v = ctl.wb.getValue(sheet, r, col);
      if (typeof v === "number") n++;
      else if (typeof v === "string") t++;
    }
    return n >= t;
  })();
  const match = q ? textMatcher(q, "contains") : null;
  const shown = values.filter((v) => !match || match(v));
  const allChecked = shown.every((v) => checked.has(v)) && (!hasBlank || checked.has("\u0000blank"));
  const apply = () => {
    const vals = values.filter((v) => checked.has(v));
    const blanks = checked.has("\u0000blank");
    const crit: FilterCriteria | null = vals.length === values.length && (blanks || !hasBlank) && !q ? null : { kind: "values", values: q ? vals.filter((v) => shown.includes(v)) : vals, blanks: q ? false : blanks };
    D.setColumnFilter(ctl, f, col, crit);
    onClose();
  };
  const sortCol = (desc: boolean) => {
    const range = { ...f.range };
    D.sortRange(ctl, range, [{ col, desc }], true);
    f.sort = { col: col - f.range.c1, desc };
    onClose();
  };
  const custom = (op1: string) => {
    onClose();
    ctl.openDialog("customFilter", { col, op1, filterRow: row });
  };
  const sub = numeric ? (
    <Menu width={210}>
      <MenuItem label="Eşittir..." onClick={() => custom("eq")} />
      <MenuItem label="Eşit Değildir..." onClick={() => custom("ne")} />
      <MenuItem label="Büyüktür..." onClick={() => custom("gt")} />
      <MenuItem label="Büyük veya Eşittir..." onClick={() => custom("ge")} />
      <MenuItem label="Küçüktür..." onClick={() => custom("lt")} />
      <MenuItem label="Küçük veya Eşittir..." onClick={() => custom("le")} />
      <MenuItem label="Arasında..." onClick={() => custom("between")} />
      <MenuItem label="İlk 10..." onClick={() => (D.setColumnFilter(ctl, f, col, { kind: "top10", count: 10, bottom: false, percent: false }), onClose())} />
      <MenuItem label="Ortalamanın Üstünde" onClick={() => (D.setColumnFilter(ctl, f, col, { kind: "custom", op1: "gt", v1: String(avgOf(ctl, f, col)).replace(".", ",") }), onClose())} />
      <MenuItem label="Ortalamanın Altında" onClick={() => (D.setColumnFilter(ctl, f, col, { kind: "custom", op1: "lt", v1: String(avgOf(ctl, f, col)).replace(".", ",") }), onClose())} />
      <MenuSeparator />
      <MenuItem label="Özel Filtre..." onClick={() => custom("eq")} />
    </Menu>
  ) : (
    <Menu width={210}>
      <MenuItem label="Eşittir..." onClick={() => custom("eq")} />
      <MenuItem label="Eşit Değildir..." onClick={() => custom("ne")} />
      <MenuSeparator />
      <MenuItem label="İle Başlar..." onClick={() => custom("begins")} />
      <MenuItem label="İle Biter..." onClick={() => custom("ends")} />
      <MenuSeparator />
      <MenuItem label="İçerir..." onClick={() => custom("contains")} />
      <MenuItem label="İçermez..." onClick={() => custom("notContains")} />
      <MenuSeparator />
      <MenuItem label="Özel Filtre..." onClick={() => custom("eq")} />
    </Menu>
  );
  return (
    <Popover open onClose={onClose} point={{ x, y }}>
      <div className="oh-menu tb-filter-menu" onKeyDown={(e) => e.key === "Enter" && apply()}>
        <MenuItem icon={<ArrowSortDownLines20Regular />} label={numeric ? "Küçükten Büyüğe Sırala" : "A'dan Z'ye Sırala"} onClick={() => sortCol(false)} />
        <MenuItem icon={<ArrowSortUpLines20Regular />} label={numeric ? "Büyükten Küçüğe Sırala" : "Z'den A'ya Sırala"} onClick={() => sortCol(true)} />
        <MenuSeparator />
        <MenuItem icon={<FilterDismiss20Regular />} label={`"${header}" Sütunundan Filtreyi Temizle`} disabled={!cur} onClick={() => (D.setColumnFilter(ctl, f, col, null), onClose())} />
        <MenuItem label={numeric ? "Sayı Filtreleri" : "Metin Filtreleri"} submenu={sub} />
        <input className="tb-filter-search" placeholder="Ara" value={q} onChange={(e) => setQ(e.target.value)} />
        <div className="tb-filter-values">
          <label>
            <input
              type="checkbox"
              checked={allChecked}
              onChange={(e) => {
                const next = new Set(checked);
                for (const v of shown) e.target.checked ? next.add(v) : next.delete(v);
                if (hasBlank && !q) e.target.checked ? next.add("\u0000blank") : next.delete("\u0000blank");
                setChecked(next);
              }}
            />
            {q ? "(Tüm Arama Sonuçlarını Seç)" : "(Tümünü Seç)"}
          </label>
          {shown.slice(0, MAX_ITEMS).map((v) => (
            <label key={v}>
              <input
                type="checkbox"
                checked={checked.has(v)}
                onChange={(e) => {
                  const next = new Set(checked);
                  if (e.target.checked) next.add(v);
                  else next.delete(v);
                  setChecked(next);
                }}
              />
              {v}
            </label>
          ))}
          {hasBlank && !q && (
            <label>
              <input
                type="checkbox"
                checked={checked.has("\u0000blank")}
                onChange={(e) => {
                  const next = new Set(checked);
                  if (e.target.checked) next.add("\u0000blank");
                  else next.delete("\u0000blank");
                  setChecked(next);
                }}
              />
              (Boş)
            </label>
          )}
        </div>
        {shown.length > MAX_ITEMS && <div className="tb-filter-note">Tüm öğeler gösterilmiyor ({MAX_ITEMS.toLocaleString("tr-TR")} / {shown.length.toLocaleString("tr-TR")}). Aramayla daraltın.</div>}
        <div className="tb-filter-footer">
          <DialogButton primary onClick={apply}>
            Tamam
          </DialogButton>
          <DialogButton onClick={onClose}>İptal</DialogButton>
        </div>
      </div>
    </Popover>
  );
}

function avgOf(ctl: Controller, f: { range: { r1: number; r2: number } }, col: number): number {
  let s = 0;
  let n = 0;
  for (let r = f.range.r1 + 1; r <= f.range.r2; r++) {
    const v = ctl.wb.getValue(ctl.sheet, r, col);
    if (typeof v === "number") {
      s += v;
      n++;
    }
  }
  return n ? Math.round((s / n) * 1e6) / 1e6 : 0;
}
