/**
 * PivotTable Alanları bölmesi: alan listesi (onay kutuları, arama, öğe
 * filtresi) ve sürükle-bırak ile dört bölge: Filtreler, Sütunlar, Satırlar,
 * Σ Değerler. Düzen güncelleştirmesi ertelenebilir.
 */
import { useMemo, useRef, useState } from "react";
import { ArrowDown16Regular, ArrowUp16Regular, Dismiss16Regular, Filter16Regular, Search16Regular, TableSimple16Regular, Navigation16Regular } from "@fluentui/react-icons";
import { DialogButton, Menu, MenuItem, MenuProvider, MenuSeparator, Popover } from "@officehere/ui";
import type { Controller } from "../controller";
import type { PivotDef } from "../model/types";
import type { SheetModel } from "../model/workbook";
import { BLANK_ITEM, fieldItems, itemLabel, valueCaption, VALUES_FIELD, type PivotSource } from "./compute";
import { AREA_LABEL, isFieldUsed, moveWithin, placeField, removeAt, toggleField, type PivotArea } from "./fields";
import { applyPivot, readSource } from "./pivotOps";

interface DragData {
  field: number;
  from: PivotArea | "list";
  index: number;
}

const AREAS: PivotArea[] = ["filters", "cols", "rows", "values"];

export function PivotFieldsPane({ ctl, sheet, def }: { ctl: Controller; sheet: SheetModel; def: PivotDef }) {
  const src = useMemo(() => readSource(ctl.wb, def), [ctl.wb, def.sourceSheet, def.source, ctl.version]); // eslint-disable-line react-hooks/exhaustive-deps
  const [q, setQ] = useState("");
  const [defer, setDefer] = useState(false);
  const [draft, setDraft] = useState<PivotDef | null>(null);
  const [menu, setMenu] = useState<{ area: PivotArea; index: number; anchor: HTMLElement } | null>(null);
  const [filterFor, setFilterFor] = useState<{ field: number; anchor: HTMLElement } | null>(null);
  const drag = useRef<DragData | null>(null);
  const [dropAt, setDropAt] = useState<{ area: PivotArea; index: number } | null>(null);
  if (typeof src === "string") {
    return (
      <aside className="tb-pane tb-pv-pane" aria-label="Özet Tablo Alanları">
        <PaneHeader ctl={ctl} />
        <p className="tb-pv-error">{src}</p>
      </aside>
    );
  }
  const cur = draft ?? def;
  const change = (fn: (d: PivotDef) => void, label = "Özet Tablo Alanı") => {
    if (defer) {
      const d = structuredClone(cur);
      fn(d);
      setDraft(d);
      return;
    }
    applyPivot(ctl, sheet, def.id, fn, label);
  };
  const headers = src.headers;
  const shown = headers.map((h, i) => ({ h, i })).filter((x) => !q || x.h.toLocaleLowerCase("tr-TR").includes(q.toLocaleLowerCase("tr-TR")));

  const chipsOf = (area: PivotArea): { field: number; label: string; index: number }[] => {
    if (area === "values") return cur.values.map((v, i) => ({ field: v.field, label: valueCaption(v, headers), index: i }));
    if (area === "filters") return cur.filters.map((f, i) => ({ field: f.field, label: headers[f.field] ?? "?", index: i }));
    const list = (area === "rows" ? cur.rows : cur.cols).map((f, i) => ({ field: f, label: headers[f] ?? "?", index: i }));
    // Σ Değerler sanal alanı
    if (cur.values.length > 1 && (area === "rows") === !!cur.valuesOnRows) list.push({ field: VALUES_FIELD, label: "Σ Değerler", index: list.length });
    return list;
  };

  const onDrop = (area: PivotArea, index: number) => {
    const d = drag.current;
    drag.current = null;
    setDropAt(null);
    if (!d) return;
    if (d.from === area && area !== "values" && d.field !== VALUES_FIELD) {
      change((def2) => moveWithin(def2, area, d.index, index > d.index ? index - 1 : index), "Özet Tablo Alanını Taşı");
      return;
    }
    if (d.field === VALUES_FIELD && (area === "filters" || area === "values")) return;
    change((def2) => placeField(def2, src, d.field, area, index, d.from === "values" ? d.index : undefined), "Özet Tablo Alanını Taşı");
  };

  const dropProps = (area: PivotArea, index: number) => ({
    onDragOver: (e: React.DragEvent) => {
      if (!drag.current) return;
      e.preventDefault();
      e.stopPropagation();
      if (!dropAt || dropAt.area !== area || dropAt.index !== index) setDropAt({ area, index });
    },
    onDrop: (e: React.DragEvent) => {
      e.preventDefault();
      e.stopPropagation();
      onDrop(area, index);
    },
  });

  const startDrag = (e: React.DragEvent, data: DragData) => {
    drag.current = data;
    e.dataTransfer.effectAllowed = "move";
    e.dataTransfer.setData("text/plain", JSON.stringify(data));
  };

  return (
    <aside className="tb-pane tb-pv-pane" aria-label="Özet Tablo Alanları">
      <PaneHeader ctl={ctl} />
      <div className="tb-pv-sub">Rapora eklenecek alanları seçin:</div>
      <div className="tb-pv-search">
        <Search16Regular />
        <input type="text" placeholder="Ara" value={q} onChange={(e) => setQ(e.target.value)} aria-label="Alanlarda ara" />
      </div>
      <div className="tb-pv-fields" role="list">
        {shown.map(({ h, i }) => {
          const used = isFieldUsed(cur, i);
          const hasFilter = !!cur.hidden?.[i]?.length || !!cur.filters.find((f) => f.field === i && f.selected);
          return (
            <div key={i} role="listitem" className={`tb-pv-field${used ? " is-used" : ""}`} draggable onDragStart={(e) => startDrag(e, { field: i, from: "list", index: -1 })} onDragEnd={() => setDropAt(null)}>
              <label>
                <input type="checkbox" checked={used} onChange={() => change((d) => toggleField(d, src, i), used ? "Alanı Kaldır" : "Alan Ekle")} />
                <span className={used ? "is-bold" : ""}>{h}</span>
              </label>
              {hasFilter && <Filter16Regular className="tb-pv-funnel" />}
              <button type="button" className="tb-pv-more" aria-label={`${h} öğelerini filtrele`} onClick={(e) => setFilterFor({ field: i, anchor: e.currentTarget })}>
                <svg width="8" height="8" viewBox="0 0 8 8" aria-hidden="true">
                  <path d="M1 2.5l3 3 3-3" fill="none" stroke="currentColor" strokeWidth="1.1" />
                </svg>
              </button>
            </div>
          );
        })}
      </div>
      <div className="tb-pv-sub">Alanları aşağıdaki alanlar arasında sürükleyin:</div>
      <div className="tb-pv-areas">
        {AREAS.map((area) => {
          const chips = chipsOf(area);
          return (
            <div key={area} className={`tb-pv-area${dropAt?.area === area ? " is-over" : ""}`} {...dropProps(area, chips.length)} data-area={area}>
              <div className="tb-pv-area-title">
                {area === "filters" ? <Filter16Regular /> : area === "cols" ? <TableSimple16Regular /> : area === "rows" ? <Navigation16Regular /> : <span className="tb-pv-sigma">Σ</span>}
                {AREA_LABEL[area]}
              </div>
              <div className="tb-pv-chips">
                {chips.map((c) => (
                  <div
                    key={`${c.field}-${c.index}`}
                    className={`tb-pv-chip${dropAt?.area === area && dropAt.index === c.index ? " is-drop-before" : ""}`}
                    draggable
                    onDragStart={(e) => startDrag(e, { field: c.field, from: area, index: c.index })}
                    onDragEnd={() => setDropAt(null)}
                    {...dropProps(area, c.index)}
                    onClick={(e) => setMenu({ area, index: c.index, anchor: e.currentTarget })}
                    role="button"
                    tabIndex={0}
                    aria-label={`${AREA_LABEL[area]}: ${c.label}`}
                  >
                    <span>{c.label}</span>
                    <svg width="8" height="8" viewBox="0 0 8 8" aria-hidden="true">
                      <path d="M1 2.5l3 3 3-3" fill="none" stroke="currentColor" strokeWidth="1.1" />
                    </svg>
                  </div>
                ))}
              </div>
            </div>
          );
        })}
      </div>
      <div className="tb-pv-footer">
        <label>
          <input
            type="checkbox"
            checked={defer}
            onChange={(e) => {
              setDefer(e.target.checked);
              if (!e.target.checked && draft) {
                const d = draft;
                setDraft(null);
                applyPivot(ctl, sheet, def.id, (x) => Object.assign(x, d), "Özet Tablo Alanları");
              }
            }}
          />
          Düzen Güncelleştirmesini Ertele
        </label>
        <DialogButton
          disabled={!defer || !draft}
          onClick={() => {
            if (!draft) return;
            const d = draft;
            setDraft(null);
            applyPivot(ctl, sheet, def.id, (x) => Object.assign(x, d), "Özet Tablo Alanları");
          }}
        >
          Güncelleştir
        </DialogButton>
      </div>
      {menu && (
        <Popover open onClose={() => setMenu(null)} anchor={{ current: menu.anchor }}>
          <MenuProvider onClose={() => setMenu(null)}>
          <ChipMenu
            ctl={ctl}
            def={cur}
            area={menu.area}
            index={menu.index}
            chips={chipsOf(menu.area)}
            onClose={() => setMenu(null)}
            change={change}
            src={src}
          />
          </MenuProvider>
        </Popover>
      )}
      {filterFor && (
        <Popover open onClose={() => setFilterFor(null)} anchor={{ current: filterFor.anchor }}>
          <ItemFilter
            src={src}
            def={cur}
            field={filterFor.field}
            onApply={(sel) => {
              setFilterFor(null);
              change((d) => setItemSelection(d, filterFor.field, sel, src), "Özet Tablo Filtresi");
            }}
            onCancel={() => setFilterFor(null)}
          />
        </Popover>
      )}
    </aside>
  );
}

function PaneHeader({ ctl }: { ctl: Controller }) {
  return (
    <div className="tb-pane-head">
      <span>Özet Tablo Alanları</span>
      <button
        type="button"
        className="tb-pane-close"
        aria-label="Kapat"
        onClick={() => {
          ctl.showFieldList = false;
          ctl.emit();
          ctl.focusGrid();
        }}
      >
        <Dismiss16Regular />
      </button>
    </div>
  );
}

function ChipMenu({ ctl, def, area, index, chips, onClose, change, src }: { ctl: Controller; def: PivotDef; area: PivotArea; index: number; chips: { field: number; label: string; index: number }[]; onClose: () => void; change: (fn: (d: PivotDef) => void, label?: string) => void; src: PivotSource }) {
  const chip = chips.find((c) => c.index === index);
  if (!chip) return null;
  const isSigma = chip.field === VALUES_FIELD;
  const n = area === "values" ? def.values.length : area === "filters" ? def.filters.length : (area === "rows" ? def.rows : def.cols).length;
  const run = (fn: (d: PivotDef) => void, label: string) => () => {
    onClose();
    change(fn, label);
  };
  const moveTo = (to: PivotArea) => run((d) => placeField(d, src, chip.field, to, undefined, area === "values" ? index : undefined), "Özet Tablo Alanını Taşı");
  return (
    <Menu width={230}>
      <MenuItem icon={<ArrowUp16Regular />} label="Yukarı Taşı" disabled={isSigma || index === 0} onClick={run((d) => moveWithin(d, area, index, index - 1), "Özet Tablo Alanını Taşı")} />
      <MenuItem icon={<ArrowDown16Regular />} label="Aşağı Taşı" disabled={isSigma || index >= n - 1} onClick={run((d) => moveWithin(d, area, index, index + 1), "Özet Tablo Alanını Taşı")} />
      <MenuItem label="Başa Taşı" disabled={isSigma || index === 0} onClick={run((d) => moveWithin(d, area, index, 0), "Özet Tablo Alanını Taşı")} />
      <MenuItem label="Sona Taşı" disabled={isSigma || index >= n - 1} onClick={run((d) => moveWithin(d, area, index, n - 1), "Özet Tablo Alanını Taşı")} />
      <MenuSeparator />
      <MenuItem label="Rapor Filtresine Taşı" disabled={isSigma || area === "filters"} onClick={moveTo("filters")} />
      <MenuItem label="Satır Etiketlerine Taşı" disabled={area === "rows"} onClick={moveTo("rows")} />
      <MenuItem label="Sütun Etiketlerine Taşı" disabled={area === "cols"} onClick={moveTo("cols")} />
      <MenuItem label="Değerlere Taşı" disabled={isSigma || area === "values"} onClick={moveTo("values")} />
      <MenuSeparator />
      <MenuItem icon={<Dismiss16Regular />} label="Alanı Kaldır" disabled={isSigma} onClick={run((d) => removeAt(d, area, index), "Alanı Kaldır")} />
      <MenuSeparator />
      {area === "values" ? (
        <MenuItem
          label="Değer Alanı Ayarları..."
          onClick={() => {
            onClose();
            ctl.openDialog("pivotValueField", { pivotId: def.id, index });
          }}
        />
      ) : (
        <MenuItem label="Alan Ayarları..." disabled={isSigma} onClick={() => (onClose(), ctl.openDialog("pivotOptions", { pivotId: def.id }))} />
      )}
    </Menu>
  );
}

/** Öğe seçimi: satır/sütun alanında gizli öğeler, filtre alanında seçili öğeler. */
export function setItemSelection(d: PivotDef, field: number, selected: string[] | null, src: PivotSource): void {
  const all = fieldItems(src, field).map((x) => x.key);
  const pf = d.filters.find((f) => f.field === field);
  if (pf) {
    pf.selected = selected && selected.length < all.length ? selected : null;
    return;
  }
  const hidden = selected ? all.filter((k) => !selected.includes(k)) : [];
  d.hidden = { ...(d.hidden ?? {}), [field]: hidden };
  if (!hidden.length) delete d.hidden[field];
}

export function ItemFilter({ src, def, field, onApply, onCancel }: { src: PivotSource; def: PivotDef; field: number; onApply: (sel: string[] | null) => void; onCancel: () => void }) {
  const items = useMemo(() => fieldItems(src, field), [src, field]);
  const pf = def.filters.find((f) => f.field === field);
  const initial = pf ? (pf.selected ?? items.map((i) => i.key)) : items.map((i) => i.key).filter((k) => !def.hidden?.[field]?.includes(k));
  const [checked, setChecked] = useState(new Set(initial));
  const [q, setQ] = useState("");
  const shown = items.filter((i) => !q || String(itemLabel(i.value)).toLocaleLowerCase("tr-TR").includes(q.toLocaleLowerCase("tr-TR")));
  const all = shown.every((i) => checked.has(i.key));
  return (
    <div className="oh-menu tb-pv-itemfilter" role="dialog" aria-label="Öğe filtresi">
      <input type="text" placeholder="Ara" value={q} onChange={(e) => setQ(e.target.value)} aria-label="Öğelerde ara" />
      <div className="tb-pv-itemlist">
        <label>
          <input
            type="checkbox"
            checked={all}
            onChange={() => {
              const next = new Set(checked);
              for (const i of shown) {
                if (all) next.delete(i.key);
                else next.add(i.key);
              }
              setChecked(next);
            }}
          />
          (Tümünü Seç)
        </label>
        {shown.map((i) => (
          <label key={i.key}>
            <input
              type="checkbox"
              checked={checked.has(i.key)}
              onChange={() => {
                const next = new Set(checked);
                if (next.has(i.key)) next.delete(i.key);
                else next.add(i.key);
                setChecked(next);
              }}
            />
            {i.value === null || i.value === "" ? BLANK_ITEM : String(typeof i.value === "number" ? String(i.value).replace(".", ",") : itemLabel(i.value))}
          </label>
        ))}
      </div>
      <div className="tb-pv-itembtns">
        <DialogButton primary disabled={checked.size === 0} onClick={() => onApply(checked.size === items.length ? null : [...checked])}>
          Tamam
        </DialogButton>
        <DialogButton onClick={onCancel}>İptal</DialogButton>
      </div>
    </div>
  );
}

/** Izgaradaki PivotTable açılır düğmesi: sıralama + öğe filtresi. */
export function PivotButtonMenu({ ctl, sheet, pivotId, field, x, y, onClose }: { ctl: Controller; sheet: SheetModel; pivotId: string; field: number; x: number; y: number; onClose: () => void }) {
  const def = sheet.pivots.find((p) => p.id === pivotId);
  const src = def ? readSource(ctl.wb, def) : "";
  if (!def || typeof src === "string") return null;
  const isPage = def.filters.some((f) => f.field === field);
  const sort = (desc: boolean) => {
    onClose();
    applyPivot(ctl, sheet, def.id, (d) => (d.sort = { ...(d.sort ?? {}), [field]: desc }), "Sırala");
  };
  return (
    <Popover open onClose={onClose} point={{ x, y }}>
      <div className="oh-menu tb-pv-btnmenu">
        {!isPage && (
          <MenuProvider onClose={onClose}>
            <div className="tb-pv-btnfield">Alan: {src.headers[field]}</div>
            <MenuItem icon={<ArrowUp16Regular />} label="A'dan Z'ye Sırala" onClick={() => sort(false)} />
            <MenuItem icon={<ArrowDown16Regular />} label="Z'den A'ya Sırala" onClick={() => sort(true)} />
            <MenuSeparator />
          </MenuProvider>
        )}
        <ItemFilter
          src={src}
          def={def}
          field={field}
          onApply={(sel) => {
            onClose();
            applyPivot(ctl, sheet, def.id, (d) => setItemSelection(d, field, sel, src), "Özet Tablo Filtresi");
          }}
          onCancel={onClose}
        />
      </div>
    </Popover>
  );
}
