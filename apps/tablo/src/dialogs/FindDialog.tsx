/** Bul ve Değiştir iletişim kutusu. */
import { useMemo, useState } from "react";
import { Dialog, DialogButton, DialogTabs, Field } from "@officehere/ui";
import { addr } from "@officehere/formula";
import type { Controller } from "../controller";
import { findAll, findPattern, nextHitIndex, type FindHit, type FindOptions } from "../model/find";
import { Check } from "./common";

type P = { ctl: Controller; onClose: () => void; arg?: unknown };

export function FindDialog({ ctl, onClose, arg }: P) {
  const a = (arg ?? {}) as { query?: string; replace?: boolean };
  const [tab, setTab] = useState(a.replace ? "replace" : "find");
  const [query, setQuery] = useState(a.query ?? "");
  const [repl, setRepl] = useState("");
  const [more, setMore] = useState(false);
  const [scope, setScope] = useState<"sheet" | "book">("sheet");
  const [byColumns, setByColumns] = useState(false);
  const [lookIn, setLookIn] = useState<FindOptions["lookIn"]>("formulas");
  const [matchCase, setMatchCase] = useState(false);
  const [wholeCell, setWholeCell] = useState(false);
  const [list, setList] = useState<FindHit[] | null>(null);
  const [msg, setMsg] = useState("");
  const isReplace = tab === "replace";

  // Çok hücreli seçimde yalnız seçimde ara
  const within = useMemo(() => {
    const r = ctl.range;
    return ctl.sel.ranges.length > 1 || r.r1 !== r.r2 || r.c1 !== r.c2 ? ctl.sel.ranges.map((x) => ({ ...x })) : undefined;
  }, []); // eslint-disable-line react-hooks/exhaustive-deps

  const opts = (): FindOptions => ({ query, matchCase, wholeCell, lookIn: isReplace ? "formulas" : lookIn, byColumns, within: scope === "sheet" ? within : undefined });
  const sheets = () => (scope === "book" ? ctl.wb.sheets.filter((s) => !s.hidden) : [ctl.sheet]);
  const go = (h: FindHit) => {
    const idx = ctl.wb.sheets.indexOf(h.sheet);
    if (idx !== ctl.wb.activeIndex) ctl.activateSheet(idx);
    if (within && scope === "sheet") ctl.setSelection({ ranges: within, active: { row: h.row, col: h.col }, anchor: { row: h.row, col: h.col } });
    else ctl.selectCell(h.row, h.col);
  };
  const notFound = () => setMsg("Aradığınızı bulamıyoruz. Diğer arama seçenekleri için Seçenekler'i tıklatın.");

  const findNext = (backward = false) => {
    setMsg("");
    const hits = findAll(ctl.wb, sheets(), opts());
    if (!hits.length) return notFound();
    const { row, col } = ctl.sel.active;
    go(hits[nextHitIndex(hits, ctl.sheet, row, col, byColumns, backward)]);
  };
  const findAllClick = () => {
    const hits = findAll(ctl.wb, sheets(), opts());
    setList(hits);
    setMsg(hits.length ? `${hits.length} hücre bulundu` : "");
    if (!hits.length) notFound();
  };
  const replaceOne = () => {
    setMsg("");
    const o = opts();
    const { row, col } = ctl.sel.active;
    const hits = findAll(ctl.wb, sheets(), o);
    const here = hits.find((h) => h.sheet === ctl.sheet && h.row === row && h.col === col);
    if (here) {
      ctl.ops.replaceAll(ctl.sheet, [{ row, col }], wholeCell ? /^[\s\S]*$/u : findPattern(o, true), repl);
    }
    const after = findAll(ctl.wb, sheets(), o);
    if (!after.length) return here ? setMsg("Başka eşleşme yok.") : notFound();
    go(after[nextHitIndex(after, ctl.sheet, row, col, byColumns)]);
  };
  const replaceAllClick = () => {
    const o = opts();
    const hits = findAll(ctl.wb, sheets(), o);
    if (!hits.length) return notFound();
    let n = 0;
    for (const s of sheets()) {
      const cells = hits.filter((h) => h.sheet === s).map((h) => ({ row: h.row, col: h.col }));
      if (cells.length) n += ctl.ops.replaceAll(s, cells, wholeCell ? /^[\s\S]*$/u : findPattern(o, true), repl);
    }
    setList(null);
    ctl.alert(`Tümü tamamlandı. ${n.toLocaleString("tr-TR")} değiştirme yapıldı.`);
  };

  return (
    <Dialog
      open
      title="Bul ve Değiştir"
      onClose={onClose}
      width={more || list ? 580 : 520}
      onSubmit={() => findNext()}
      footer={
        <>
          {isReplace && <DialogButton onClick={replaceAllClick}>Tümünü Değiştir</DialogButton>}
          {isReplace && <DialogButton onClick={replaceOne}>Değiştir</DialogButton>}
          <DialogButton onClick={findAllClick}>Tümünü Bul</DialogButton>
          <DialogButton primary onClick={() => findNext()}>
            Sonrakini Bul
          </DialogButton>
          <DialogButton onClick={onClose}>Kapat</DialogButton>
        </>
      }
    >
      <DialogTabs tabs={[{ id: "find", label: "Bul" }, { id: "replace", label: "Değiştir" }]} active={tab} onChange={setTab} />
      <div className="tb-find">
        <Field label="Aranan:" inline>
          <input type="text" value={query} onChange={(e) => setQuery(e.target.value)} data-autofocus="true" style={{ flex: 1 }} />
        </Field>
        {isReplace && (
          <Field label="Yeni değer:" inline>
            <input type="text" value={repl} onChange={(e) => setRepl(e.target.value)} style={{ flex: 1 }} />
          </Field>
        )}
        {more && (
          <div className="tb-find-more">
            <div className="tb-find-col">
              <Field label="İçinde:" inline>
                <select value={scope} onChange={(e) => setScope(e.target.value as "sheet" | "book")}>
                  <option value="sheet">Sayfa</option>
                  <option value="book">Çalışma Kitabı</option>
                </select>
              </Field>
              <Field label="Ara:" inline>
                <select value={byColumns ? "cols" : "rows"} onChange={(e) => setByColumns(e.target.value === "cols")}>
                  <option value="rows">Satırlara Göre</option>
                  <option value="cols">Sütunlara Göre</option>
                </select>
              </Field>
              <Field label="Bakılacak yer:" inline>
                <select value={isReplace ? "formulas" : lookIn} disabled={isReplace} onChange={(e) => setLookIn(e.target.value as FindOptions["lookIn"])}>
                  <option value="formulas">Formüller</option>
                  <option value="values">Değerler</option>
                  <option value="notes">Notlar</option>
                </select>
              </Field>
            </div>
            <div className="tb-find-col">
              <Check checked={matchCase} onChange={setMatchCase}>
                Büyük/küçük harf eşleştir
              </Check>
              <Check checked={wholeCell} onChange={setWholeCell}>
                Tüm hücre içeriğini eşleştir
              </Check>
            </div>
          </div>
        )}
        <div className="tb-find-bar">
          <button type="button" className="tb-link-btn" onClick={() => setMore((m) => !m)}>
            {more ? "Seçenekler <<" : "Seçenekler >>"}
          </button>
          <span className="tb-find-msg">{msg}</span>
        </div>
        {list && list.length > 0 && (
          <div className="tb-find-results">
            <div className="tb-find-row tb-find-head">
              <span>Sayfa</span>
              <span>Hücre</span>
              <span>Değer</span>
            </div>
            <div className="tb-find-scroll">
              {list.slice(0, 2000).map((h, i) => (
                <div key={i} className="tb-find-row" onClick={() => go(h)}>
                  <span>{h.sheet.name}</span>
                  <span>{"$" + addr(h.row, h.col).replace(/(\d)/, "$$$1")}</span>
                  <span>{h.text}</span>
                </div>
              ))}
            </div>
          </div>
        )}
      </div>
    </Dialog>
  );
}
