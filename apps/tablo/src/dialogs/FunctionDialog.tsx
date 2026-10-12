/** İşlev Ekle ve İşlev Bağımsız Değişkenleri iletişim kutuları. */
import { useMemo, useState } from "react";
import { Dialog, Field } from "@officehere/ui";
import { CATEGORY_LABELS, FUNCTIONS, argNames, colToName as colName, fromLocalFormula, type FnDef, type Scalar } from "@officehere/formula";
import type { Controller } from "../controller";
import { formatScalar } from "../model/display";
import { OkCancel } from "./common";

type P = { ctl: Controller; onClose: () => void; arg?: unknown };

const RECENT_KEY = "tablo.recentFunctions";
const DEFAULT_RECENT = ["SUM", "AVERAGE", "IF", "HYPERLINK", "COUNT", "MAX", "SIN", "SUMIF", "PMT", "STDEV"];

function loadRecent(): string[] {
  try {
    const s = localStorage.getItem(RECENT_KEY);
    if (s) return JSON.parse(s) as string[];
  } catch {
    /* depolama yok */
  }
  return DEFAULT_RECENT;
}

function pushRecent(name: string): void {
  const list = [name, ...loadRecent().filter((n) => n !== name)].slice(0, 10);
  try {
    localStorage.setItem(RECENT_KEY, JSON.stringify(list));
  } catch {
    /* depolama yok */
  }
}

/** Kullanıcıya gösterilen fonksiyonlar (benzersiz, Türkçe ada göre sıralı). */
export function allFunctions(): FnDef[] {
  const seen = new Set<FnDef>();
  for (const d of FUNCTIONS.values()) seen.add(d);
  return [...seen].filter((d) => !d.name.startsWith("_")).sort((a, b) => a.tr.localeCompare(b.tr, "tr"));
}

export function InsertFunctionDialog({ ctl, onClose, arg }: P) {
  const preset = (arg as { name?: string } | undefined)?.name;
  const all = useMemo(allFunctions, []);
  const [search, setSearch] = useState("");
  const [cat, setCat] = useState<string>("recent");
  const [selName, setSelName] = useState<string>(preset ?? "SUM");
  const [step, setStep] = useState<FnDef | null>(preset ? (FUNCTIONS.get(preset) ?? null) : null);
  const recent = useMemo(loadRecent, []);

  const list = useMemo(() => {
    if (search.trim()) {
      const q = search.trim().toLocaleUpperCase("tr-TR");
      const score = (d: FnDef) => (d.tr === q || d.name === q ? 0 : d.tr.startsWith(q) || d.name.startsWith(q) ? 1 : d.tr.includes(q) || d.name.includes(q) ? 2 : d.desc.toLocaleUpperCase("tr-TR").includes(q) ? 3 : 9);
      return all
        .map((d) => ({ d, s: score(d) }))
        .filter((x) => x.s < 9)
        .sort((a, b) => a.s - b.s || a.d.tr.localeCompare(b.d.tr, "tr"))
        .map((x) => x.d);
    }
    if (cat === "recent") return recent.map((n) => FUNCTIONS.get(n)).filter((d): d is FnDef => !!d);
    if (cat === "all") return all;
    return all.filter((d) => d.category === cat);
  }, [search, cat, all, recent]);
  const sel = list.find((d) => d.name === selName) ?? list[0];

  if (step) return <FunctionArgsDialog ctl={ctl} def={step} onClose={onClose} />;

  const ok = () => sel && setStep(sel);
  return (
    <Dialog open title="İşlev Ekle" onClose={onClose} width={520} onSubmit={ok} footer={<OkCancel onOk={ok} onCancel={onClose} disabled={!sel} />}>
      <Field label="Bir işlev arayın:">
        <div className="tb-fn-search">
          <input type="text" value={search} onChange={(e) => setSearch(e.target.value)} placeholder="Ne yapmak istediğinizi kısaca yazın" data-autofocus="true" />
        </div>
      </Field>
      <Field label="Bir kategori seçin:" inline>
        <select value={search ? "search" : cat} onChange={(e) => (setSearch(""), setCat(e.target.value))}>
          {search && <option value="search">Önerilen</option>}
          <option value="recent">En Son Kullanılan</option>
          <option value="all">Tümü</option>
          {Object.entries(CATEGORY_LABELS).map(([k, v]) => (
            <option key={k} value={k}>
              {v}
            </option>
          ))}
        </select>
      </Field>
      <Field label="Bir işlev seçin:">
        <div className="tb-list" style={{ height: 180 }}>
          {list.map((d) => (
            <div key={d.name} className={`tb-list-item${sel?.name === d.name ? " is-active" : ""}`} onClick={() => setSelName(d.name)} onDoubleClick={() => setStep(d)}>
              {d.tr}
            </div>
          ))}
          {list.length === 0 && <div className="tb-list-item">Eşleşen işlev yok.</div>}
        </div>
      </Field>
      {sel && (
        <div className="tb-fn-desc">
          <b>
            {sel.tr}({sel.args})
          </b>
          <p>{sel.desc}</p>
        </div>
      )}
    </Dialog>
  );
}

function evalLocal(ctl: Controller, text: string): string {
  if (!text.trim()) return "";
  try {
    const canon = fromLocalFormula("=" + text, "tr");
    const { row, col } = ctl.sel.active;
    const show = (v: Scalar) => (typeof v === "number" ? v.toLocaleString("tr-TR", { maximumFractionDigits: 10 }) : typeof v === "string" ? `"${v}"` : formatScalar(v));
    const m = ctl.wb.engine.evaluateArray(ctl.sheet.id, row, col, canon);
    if (m.length === 1 && m[0].length === 1) return show(m[0][0]);
    const rows = m.slice(0, 4).map((r) => r.slice(0, 5).map(show).join("\\"));
    return `{${rows.join(";")}${m.length > 4 || (m[0]?.length ?? 0) > 5 ? "..." : ""}}`;
  } catch {
    return "Geçersiz";
  }
}

/** Formülden mevcut bağımsız değişkenleri ayıklar (=AD(a;b) biçimi). */
function existingArgs(text: string, tr: string): string[] | null {
  const m = new RegExp(`^=${tr.replace(/[.]/g, "\\.")}\\((.*)\\)$`, "iu").exec(text.trim());
  if (!m) return null;
  const out: string[] = [];
  let depth = 0;
  let cur = "";
  let q = false;
  for (const ch of m[1]) {
    if (ch === '"') q = !q;
    if (!q && (ch === "(" || ch === "{")) depth++;
    if (!q && (ch === ")" || ch === "}")) depth--;
    if (!q && depth === 0 && ch === ";") {
      out.push(cur);
      cur = "";
    } else cur += ch;
  }
  out.push(cur);
  return out;
}

export function FunctionArgsDialog({ ctl, def, onClose }: { ctl: Controller; def: FnDef; onClose: () => void }) {
  const names = argNames(def);
  const repeat = names.some((n) => n.includes("..."));
  const fixed = names.filter((n) => !n.includes("..."));
  const initial = useMemo(() => {
    const { row, col } = ctl.sel.active;
    const ex = existingArgs(ctl.edit?.text ?? ctl.wb.editText(ctl.sheet, row, col), def.tr);
    if (ex) return ex;
    // TOPLA gibi işlevlerde üstteki sayıları öner
    if (["SUM", "AVERAGE", "COUNT", "MAX", "MIN"].includes(def.name)) {
      let r = row - 1;
      while (r >= 0 && typeof ctl.wb.getValue(ctl.sheet, r, col) === "number") r--;
      if (r < row - 1) return [`${colName(col)}${r + 2}:${colName(col)}${row}`];
    }
    return [];
  }, []); // eslint-disable-line react-hooks/exhaustive-deps
  const [vals, setVals] = useState<string[]>(() => {
    const n = Math.max(fixed.length, initial.length, Math.min(def.min, 3), 1);
    return Array.from({ length: n }, (_, i) => initial[i] ?? "");
  });
  const [focus, setFocus] = useState(0);
  const labelOf = (i: number) => {
    if (i < fixed.length) return fixed[i].replace(/^\[|\]$/g, "");
    const base = (fixed[fixed.length - 1] ?? "değer").replace(/^\[|\]$/g, "").replace(/\d+$/, "");
    return `${base}${i + 1}`;
  };
  const optional = (i: number) => i >= def.min;
  const formula = () => {
    const used = [...vals];
    while (used.length > def.min && !used[used.length - 1].trim()) used.pop();
    return `=${def.tr}(${used.join(";")})`;
  };
  const result = evalLocal(ctl, formula().slice(1));

  const ok = () => {
    if (vals.slice(0, def.min).some((v) => !v.trim())) {
      ctl.alert("Bu işlev için çok az bağımsız değişken girdiniz.");
      return;
    }
    const text = formula();
    pushRecent(def.name);
    const { row, col } = ctl.sel.active;
    if (ctl.edit) ctl.cancelEdit();
    ctl.ops.enterText(ctl.sheet, [{ row, col }], text, "İşlev Ekle");
    onClose();
  };
  const setAt = (i: number, v: string) => {
    const next = [...vals];
    next[i] = v;
    if (repeat && i === next.length - 1 && v.trim() && next.length < Math.min(def.max, 255)) next.push("");
    setVals(next);
  };

  return (
    <Dialog open title="İşlev Bağımsız Değişkenleri" onClose={onClose} width={540} onSubmit={ok} footer={<OkCancel onOk={ok} onCancel={onClose} />}>
      <div className="tb-fnargs">
        <div className="tb-fnargs-name">{def.tr}</div>
        <div className="tb-fnargs-list">
          {vals.map((v, i) => (
            <div key={i} className="tb-fnargs-row">
              <label className={optional(i) ? "" : "is-required"}>{labelOf(i)}</label>
              <input type="text" value={v} data-autofocus={i === 0 ? "true" : undefined} onFocus={() => setFocus(i)} onChange={(e) => setAt(i, e.target.value)} />
              <span className="tb-fnargs-val">= {evalLocal(ctl, v) || (optional(i) ? "isteğe bağlı" : "")}</span>
            </div>
          ))}
        </div>
        <div className="tb-fnargs-result">= {result}</div>
        <p className="tb-fn-desc">{def.desc}</p>
        <p className="tb-fn-desc">
          <b>{labelOf(focus)}:</b> {optional(focus) ? "İsteğe bağlı bağımsız değişken." : "Gerekli bağımsız değişken."} Başvuru, sayı, metin ya da formül girebilirsiniz.
        </p>
        <div className="tb-fnargs-foot">
          Formül sonucu = <b>{result}</b>
        </div>
      </div>
    </Dialog>
  );
}
