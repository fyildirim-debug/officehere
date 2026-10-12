/** Tanımlı adlar: Ad Tanımla ve Ad Yöneticisi. */
import { useState } from "react";
import { Dialog, DialogButton, Field } from "@officehere/ui";
import { colToName, fromLocalFormula, quoteSheetName, toLocalFormula } from "@officehere/formula";
import type { Controller } from "../controller";
import { formatScalar } from "../model/display";
import { Check, SimpleDialog } from "./common";

type P = { ctl: Controller; onClose: () => void; arg?: unknown };

/** Ad geçerli mi? Hata metni ya da null. */
export function nameError(name: string): string | null {
  if (!name) return "Bir ad girin.";
  if (!/^[\p{L}_\\][\p{L}\p{N}_.]*$/u.test(name)) return "Girdiğiniz ad geçerli değil. Adın bir harf ya da alt çizgiyle başladığından ve boşluk içermediğinden emin olun.";
  if (/^[A-Za-z]{1,3}\d{1,7}$/.test(name) || /^[RrCc]$/.test(name)) return "Ad bir hücre başvurusuna benzediği için kullanılamaz.";
  if (name.length > 255) return "Ad çok uzun.";
  return null;
}

function selectionRef(ctl: Controller): string {
  const r = ctl.range;
  const cell = (row: number, col: number) => `$${colToName(col)}$${row + 1}`;
  const a = cell(r.r1, r.c1);
  const b = cell(r.r2, r.c2);
  return `=${quoteSheetName(ctl.sheet.name)}!${a === b ? a : `${a}:${b}`}`;
}

function valuePreview(ctl: Controller, formula: string): string {
  try {
    const m = ctl.wb.engine.evaluateArray(ctl.sheet.id, 0, 0, formula);
    if (m.length === 1 && m[0].length === 1) return formatScalar(m[0][0]);
    const cells = m.slice(0, 3).map((row) => row.slice(0, 4).map((v) => formatScalar(v)).join("\\"));
    return `{${cells.join(";")}${m.length > 3 ? "..." : ""}}`;
  } catch {
    return "#AD?";
  }
}

export function DefineNameDialog({ ctl, onClose, arg }: P) {
  const a = (arg ?? {}) as { edit?: { name: string; formula: string; sheet?: number }; onDone?: () => void };
  const { row, col } = ctl.sel.active;
  const guess = (() => {
    if (a.edit) return a.edit.name;
    const left = col > 0 ? ctl.wb.getValue(ctl.sheet, row, col - 1) : null;
    const up = row > 0 ? ctl.wb.getValue(ctl.sheet, row - 1, col) : null;
    const cand = typeof up === "string" ? up : typeof left === "string" ? left : "";
    const clean = cand.trim().replace(/\s+/g, "_");
    return nameError(clean) ? "" : clean;
  })();
  const [name, setName] = useState(guess);
  const [scope, setScope] = useState<string>(a.edit?.sheet !== undefined ? String(a.edit.sheet) : "book");
  const [ref, setRef] = useState(a.edit ? toLocalFormula(a.edit.formula, "tr") : selectionRef(ctl));
  const ok = () => {
    const err = nameError(name.trim());
    if (err) return ctl.alert(err);
    const sheetId = scope === "book" ? undefined : Number(scope);
    const exists = ctl.wb.engine.getNames().find((n) => n.name.toLocaleUpperCase("tr-TR") === name.trim().toLocaleUpperCase("tr-TR") && n.sheet === sheetId);
    if (exists && (!a.edit || a.edit.name.toLocaleUpperCase("tr-TR") !== name.trim().toLocaleUpperCase("tr-TR"))) return ctl.alert("Girdiğiniz ad zaten var. Benzersiz bir ad girin.");
    let canon: string;
    try {
      canon = fromLocalFormula(ref.startsWith("=") ? ref : "=" + ref, "tr");
    } catch {
      return ctl.alert("Girdiğiniz formülde bir sorun var.");
    }
    ctl.ops.namesOp(a.edit ? "Adı Düzenle" : "Ad Tanımla", () => {
      if (a.edit) ctl.wb.engine.removeName(a.edit.name, a.edit.sheet);
      ctl.wb.engine.defineName(name.trim(), canon, sheetId);
    });
    a.onDone?.();
    onClose();
  };
  return (
    <SimpleDialog title={a.edit ? "Adı Düzenle" : "Yeni Ad"} onClose={onClose} onOk={ok} width={420}>
      <Field label="Ad:">
        <input type="text" value={name} onChange={(e) => setName(e.target.value)} data-autofocus="true" />
      </Field>
      <Field label="Kapsam:">
        <select value={scope} onChange={(e) => setScope(e.target.value)} disabled={!!a.edit}>
          <option value="book">Çalışma Kitabı</option>
          {ctl.wb.sheets.map((s) => (
            <option key={s.id} value={String(s.id)}>
              {s.name}
            </option>
          ))}
        </select>
      </Field>
      <Field label="Başvuru yeri:">
        <input type="text" value={ref} onChange={(e) => setRef(e.target.value)} />
      </Field>
    </SimpleDialog>
  );
}

export function NameManagerDialog({ ctl, onClose }: P) {
  const [, force] = useState(0);
  const [sel, setSel] = useState<string | null>(null);
  const [sub, setSub] = useState<null | "new" | "edit">(null);
  const [filter, setFilter] = useState<"all" | "book" | "sheet">("all");
  const names = ctl.wb.engine
    .getNames()
    .filter((n) => filter === "all" || (filter === "book" ? n.sheet === undefined : n.sheet !== undefined))
    .sort((a, b) => a.name.localeCompare(b.name, "tr"));
  const keyOf = (n: { name: string; sheet?: number }) => `${n.sheet ?? ""}!${n.name}`;
  const cur = names.find((n) => keyOf(n) === sel);
  const scopeName = (id?: number) => (id === undefined ? "Çalışma Kitabı" : (ctl.wb.sheetById(id)?.name ?? "?"));
  const tables = ctl.wb.sheets.flatMap((s) => s.tables.map((t) => ({ t, s })));

  if (sub) return <DefineNameDialog ctl={ctl} onClose={() => (setSub(null), force((x) => x + 1))} arg={sub === "edit" && cur ? { edit: cur } : {}} />;

  return (
    <Dialog open title="Ad Yöneticisi" onClose={onClose} width={680} footer={<DialogButton onClick={onClose}>Kapat</DialogButton>}>
      <div className="tb-row-gap">
        <DialogButton onClick={() => setSub("new")}>Yeni...</DialogButton>
        <DialogButton disabled={!cur} onClick={() => setSub("edit")}>
          Düzenle...
        </DialogButton>
        <DialogButton
          disabled={!cur}
          onClick={() => {
            if (!cur) return;
            ctl.ops.namesOp("Adı Sil", () => ctl.wb.engine.removeName(cur.name, cur.sheet));
            setSel(null);
            force((x) => x + 1);
          }}
        >
          Sil
        </DialogButton>
        <span style={{ flex: 1 }} />
        <Field label="Filtre:" inline>
          <select value={filter} onChange={(e) => setFilter(e.target.value as typeof filter)}>
            <option value="all">Tümü</option>
            <option value="book">Çalışma Kitabı kapsamlı adlar</option>
            <option value="sheet">Çalışma Sayfası kapsamlı adlar</option>
          </select>
        </Field>
      </div>
      <div className="tb-names">
        <div className="tb-names-row tb-names-head">
          <span>Ad</span>
          <span>Değer</span>
          <span>Başvuru Yeri</span>
          <span>Kapsam</span>
        </div>
        <div className="tb-names-scroll">
          {names.map((n) => (
            <div key={keyOf(n)} className={`tb-names-row${keyOf(n) === sel ? " is-active" : ""}`} onClick={() => setSel(keyOf(n))} onDoubleClick={() => (setSel(keyOf(n)), setSub("edit"))}>
              <span>{n.name}</span>
              <span>{valuePreview(ctl, n.formula)}</span>
              <span>{toLocalFormula(n.formula, "tr")}</span>
              <span>{scopeName(n.sheet)}</span>
            </div>
          ))}
          {tables.map(({ t, s }) => (
            <div key={"t" + t.id} className="tb-names-row is-table">
              <span>{t.name}</span>
              <span>(Tablo)</span>
              <span>
                ={quoteSheetName(s.name)}!{rangeA1(t.ref)}
              </span>
              <span>Çalışma Kitabı</span>
            </div>
          ))}
          {names.length === 0 && tables.length === 0 && <div className="tb-names-empty">Tanımlı ad yok. Yeni bir ad eklemek için Yeni... düğmesini kullanın.</div>}
        </div>
      </div>
    </Dialog>
  );
}

function rangeA1(r: { r1: number; c1: number; r2: number; c2: number }): string {
  return `$${colToName(r.c1)}$${r.r1 + 1}:$${colToName(r.c2)}$${r.r2 + 1}`;
}

/** Seçimden Oluştur: üst satır / sol sütun etiketlerinden ad tanımlar. */
export function CreateNamesDialog({ ctl, onClose }: P) {
  const r = ctl.range;
  const wb = ctl.wb;
  const sheet = ctl.sheet;
  const isText = (row: number, col: number) => typeof wb.getValue(sheet, row, col) === "string";
  const [top, setTop] = useState(() => isText(r.r1, r.c2) && r.r2 > r.r1);
  const [left, setLeft] = useState(() => isText(r.r2, r.c1) && r.c2 > r.c1 && !(isText(r.r1, r.c2) && r.r2 > r.r1));
  const [bottom, setBottom] = useState(false);
  const [right, setRight] = useState(false);
  const ok = () => {
    const defs: { name: string; ref: string }[] = [];
    const clean = (v: unknown) => String(v ?? "").trim().replace(/[\s-]+/g, "_").replace(/^(\d)/, "_$1");
    const q = quoteSheetName(sheet.name);
    const add = (label: unknown, r1: number, c1: number, r2: number, c2: number) => {
      const n = clean(label);
      if (!n || nameError(n)) return;
      const a = `$${colToName(c1)}$${r1 + 1}`;
      const b = `$${colToName(c2)}$${r2 + 1}`;
      defs.push({ name: n, ref: `=${q}!${a === b ? a : `${a}:${b}`}` });
    };
    const r1 = r.r1 + (top ? 1 : 0);
    const r2 = r.r2 - (bottom ? 1 : 0);
    const c1 = r.c1 + (left ? 1 : 0);
    const c2 = r.c2 - (right ? 1 : 0);
    if (top) for (let c = c1; c <= c2; c++) add(wb.getValue(sheet, r.r1, c), r1, c, r2, c);
    if (bottom) for (let c = c1; c <= c2; c++) add(wb.getValue(sheet, r.r2, c), r1, c, r2, c);
    if (left) for (let row = r1; row <= r2; row++) add(wb.getValue(sheet, row, r.c1), row, c1, row, c2);
    if (right) for (let row = r1; row <= r2; row++) add(wb.getValue(sheet, row, r.c2), row, c1, row, c2);
    if (!defs.length) {
      ctl.alert("Etiketlerden geçerli ad oluşturulamadı.");
      return;
    }
    ctl.ops.namesOp("Seçimden Oluştur", () => {
      for (const d of defs) wb.engine.defineName(d.name, fromLocalFormula(d.ref, "tr"));
    });
    onClose();
  };
  return (
    <SimpleDialog title="Seçimden Ad Oluştur" onClose={onClose} onOk={ok} width={300}>
      <p className="tb-dlg-text">Şuradaki değerlerden ad oluştur:</p>
      <Check checked={top} onChange={setTop}>
        Üst satır
      </Check>
      <Check checked={left} onChange={setLeft}>
        Sol sütun
      </Check>
      <Check checked={bottom} onChange={setBottom}>
        Alt satır
      </Check>
      <Check checked={right} onChange={setRight}>
        Sağ sütun
      </Check>
    </SimpleDialog>
  );
}
