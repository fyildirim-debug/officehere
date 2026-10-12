/** Veri iletişim kutuları: Sırala, Yinelenenleri Kaldır, Metni Sütunlara Dönüştür, Veri Doğrulama, Özel Filtre, Seriler, Metin İçeri Aktar. */
import { useMemo, useState } from "react";
import { Dialog, DialogButton, DialogTabs, Field, FieldSet } from "@officehere/ui";
import { colToName, dateToSerial, fromLocalFormula, toLocalFormula, type Scalar } from "@officehere/formula";
import type { Controller } from "../controller";
import * as D from "../actions/data";
import { CUSTOM_LISTS, addDate } from "../model/fill";
import { displayText } from "../model/display";
import { csvToInputs, decodeText, parseCsv } from "../io/csv";
import type { FilterCriteria, Rect, ValidationRule } from "../model/types";
import { contains, subtractRects } from "../model/rect";
import { Check, OkCancel, Radio, SimpleDialog, parseTrNumber } from "./common";

type P = { ctl: Controller; onClose: () => void; arg?: unknown };

/* ------------------------------ Sırala ------------------------------ */

interface Level {
  col: number;
  by: "value" | "fill" | "color";
  order: "asc" | "desc" | number;
}

export function SortDialog({ ctl, onClose }: P) {
  const region = useMemo(() => D.sortRegion(ctl), []); // eslint-disable-line react-hooks/exhaustive-deps
  const [header, setHeader] = useState(() => D.guessHeader(ctl.wb, ctl.sheet, region) || !!ctl.sheet.tableAt(region.r1, region.c1)?.headerRow);
  const [levels, setLevels] = useState<Level[]>([{ col: Math.max(region.c1, Math.min(region.c2, ctl.sel.active.col)), by: "value", order: "asc" }]);
  const [sel, setSel] = useState(0);
  const [matchCase, setMatchCase] = useState(false);
  const [byCols, setByCols] = useState(false);
  const [opts, setOpts] = useState(false);
  void matchCase;
  const colLabel = (c: number) => {
    if (byCols) return `Satır ${c + 1}`;
    const h = header ? displayText(ctl.wb, ctl.sheet, region.r1, c) : "";
    return h || `Sütun ${colToName(c)}`;
  };
  const choices: number[] = [];
  if (byCols) for (let r = region.r1; r <= region.r2; r++) choices.push(r);
  else for (let c = region.c1; c <= region.c2; c++) choices.push(c);
  const upd = (i: number, p: Partial<Level>) => setLevels(levels.map((l, j) => (j === i ? { ...l, ...p } : l)));
  const ok = () => {
    D.sortRange(
      ctl,
      region,
      levels.map((l) => ({ col: l.col, desc: l.order === "desc", by: l.by, order: typeof l.order === "number" ? CUSTOM_LISTS[l.order] : undefined })),
      header,
      byCols,
    );
    ctl.selectRange(region, ctl.sel.active);
    onClose();
  };
  return (
    <Dialog open title="Sırala" onClose={onClose} width={640} onSubmit={ok} footer={<OkCancel onOk={ok} onCancel={onClose} />}>
      <div className="tb-sort-bar">
        <DialogButton onClick={() => (setLevels([...levels, { col: choices[0], by: "value", order: "asc" }]), setSel(levels.length))}>Düzey Ekle</DialogButton>
        <DialogButton disabled={levels.length < 2} onClick={() => (setLevels(levels.filter((_, j) => j !== sel)), setSel(Math.max(0, sel - 1)))}>
          Düzeyi Sil
        </DialogButton>
        <DialogButton onClick={() => setLevels([...levels.slice(0, sel + 1), { ...levels[sel] }, ...levels.slice(sel + 1)])}>Düzeyi Kopyala</DialogButton>
        <DialogButton disabled={sel === 0} onClick={() => (setLevels(swap(levels, sel, sel - 1)), setSel(sel - 1))}>
          ▲
        </DialogButton>
        <DialogButton disabled={sel >= levels.length - 1} onClick={() => (setLevels(swap(levels, sel, sel + 1)), setSel(sel + 1))}>
          ▼
        </DialogButton>
        <DialogButton onClick={() => setOpts(!opts)}>Seçenekler...</DialogButton>
        <Check checked={header} onChange={setHeader}>
          Verilerimde üst bilgi var
        </Check>
      </div>
      {opts && (
        <div className="tb-sort-opts">
          <Check checked={matchCase} onChange={setMatchCase}>
            Büyük/küçük harf duyarlı
          </Check>
          <Radio name="orient" value="rows" current={byCols ? "cols" : "rows"} onChange={() => setByCols(false)}>
            Yukarıdan aşağıya sırala
          </Radio>
          <Radio name="orient" value="cols" current={byCols ? "cols" : "rows"} onChange={() => (setByCols(true), setLevels([{ col: region.r1, by: "value", order: "asc" }]))}>
            Soldan sağa sırala
          </Radio>
        </div>
      )}
      <div className="tb-sort-grid">
        <div className="tb-sort-row tb-sort-head">
          <span />
          <span>{byCols ? "Satır" : "Sütun"}</span>
          <span>Sıralama Ölçütü</span>
          <span>Düzen</span>
        </div>
        {levels.map((l, i) => (
          <div key={i} className={`tb-sort-row${i === sel ? " is-active" : ""}`} onClick={() => setSel(i)}>
            <span>{i === 0 ? "Sıralama ölçütü" : "Sonra buna göre"}</span>
            <select value={l.col} onChange={(e) => upd(i, { col: Number(e.target.value) })}>
              {choices.map((c) => (
                <option key={c} value={c}>
                  {colLabel(c)}
                </option>
              ))}
            </select>
            <select value={l.by} onChange={(e) => upd(i, { by: e.target.value as Level["by"] })}>
              <option value="value">Hücre Değerleri</option>
              <option value="fill">Hücre Rengi</option>
              <option value="color">Yazı Tipi Rengi</option>
            </select>
            <select value={String(l.order)} onChange={(e) => upd(i, { order: e.target.value === "asc" || e.target.value === "desc" ? e.target.value : Number(e.target.value) })}>
              <option value="asc">{l.by === "value" ? "Küçükten Büyüğe / A'dan Z'ye" : "Üstte"}</option>
              <option value="desc">{l.by === "value" ? "Büyükten Küçüğe / Z'den A'ya" : "Altta"}</option>
              {l.by === "value" &&
                CUSTOM_LISTS.map((cl, k) => (
                  <option key={k} value={k}>
                    {cl.slice(0, 3).join(", ")}...
                  </option>
                ))}
            </select>
          </div>
        ))}
      </div>
    </Dialog>
  );
}

function swap<T>(a: T[], i: number, j: number): T[] {
  const b = [...a];
  [b[i], b[j]] = [b[j], b[i]];
  return b;
}

/* ------------------------- Yinelenenleri Kaldır ------------------------- */

export function RemoveDuplicatesDialog({ ctl, onClose }: P) {
  const region = useMemo(() => D.sortRegion(ctl), []); // eslint-disable-line react-hooks/exhaustive-deps
  const [header, setHeader] = useState(() => D.guessHeader(ctl.wb, ctl.sheet, region) || !!ctl.sheet.tableAt(region.r1, region.c1)?.headerRow);
  const cols: number[] = [];
  for (let c = region.c1; c <= region.c2; c++) cols.push(c);
  const [checked, setChecked] = useState<Set<number>>(new Set(cols));
  const ok = () => {
    if (!checked.size) return;
    ctl.selectRange(region);
    const { removed, kept } = D.removeDuplicates(ctl, region, cols.filter((c) => checked.has(c)), header);
    onClose();
    ctl.alert(removed ? `${removed.toLocaleString("tr-TR")} yinelenen değer bulundu ve kaldırıldı; ${kept.toLocaleString("tr-TR")} benzersiz değer kaldı.` : "Yinelenen değer bulunamadı.");
  };
  return (
    <SimpleDialog title="Yinelenenleri Kaldır" onClose={onClose} onOk={ok} width={420} disabled={!checked.size}>
      <p className="tb-dlg-text">Yinelenenleri silmek için yinelenen değerleri içeren bir veya daha fazla sütun seçin.</p>
      <div className="tb-row-gap">
        <DialogButton onClick={() => setChecked(new Set(cols))}>Tümünü Seç</DialogButton>
        <DialogButton onClick={() => setChecked(new Set())}>Tüm Seçimleri Kaldır</DialogButton>
        <Check checked={header} onChange={setHeader}>
          Verilerimde üst bilgi var
        </Check>
      </div>
      <div className="tb-list" style={{ height: 150 }}>
        {cols.map((c) => (
          <label key={c} className="tb-list-item tb-check">
            <input
              type="checkbox"
              checked={checked.has(c)}
              onChange={(e) => {
                const n = new Set(checked);
                if (e.target.checked) n.add(c);
                else n.delete(c);
                setChecked(n);
              }}
            />
            <span>{header ? displayText(ctl.wb, ctl.sheet, region.r1, c) || `Sütun ${colToName(c)}` : `Sütun ${colToName(c)}`}</span>
          </label>
        ))}
      </div>
    </SimpleDialog>
  );
}

/* ------------------------- Metni Sütunlara Dönüştür ------------------------- */

export function TextToColumnsDialog({ ctl, onClose }: P) {
  const r = ctl.range;
  const src: Rect = { r1: r.r1, c1: r.c1, r2: Math.min(r.r2, Math.max(r.r1, ctl.wb.dataExtent(ctl.sheet).rows - 1)), c2: r.c1 };
  const [tab, setTab] = useState<string[]>([";"]);
  const [other, setOther] = useState("");
  const [consecutive, setConsecutive] = useState(false);
  const [qual, setQual] = useState('"');
  const delims = [...tab, ...(other ? [other] : [])];
  const preview = useMemo(() => {
    const out: string[][] = [];
    for (let row = src.r1; row <= Math.min(src.r2, src.r1 + 6); row++) out.push(D.splitDelimited(ctl.wb.plainText(ctl.sheet, row, src.c1), delims, consecutive, qual));
    return out;
  }, [delims.join(""), consecutive, qual]); // eslint-disable-line react-hooks/exhaustive-deps
  const toggle = (d: string, on: boolean) => setTab(on ? [...tab, d] : tab.filter((x) => x !== d));
  const ok = () => {
    if (!delims.length) return;
    D.textToColumns(ctl, src, delims, consecutive, qual);
    onClose();
  };
  return (
    <SimpleDialog title="Metni Sütunlara Dönüştürme Sihirbazı" onClose={onClose} onOk={ok} width={560} okLabel="Son">
      <p className="tb-dlg-text">Verilerinizdeki alanları ayıran sınırlayıcıları seçin. Önizlemede metninizin nasıl etkilendiğini görebilirsiniz.</p>
      <FieldSet legend="Sınırlayıcılar">
        <div className="tb-row-gap">
          <Check checked={tab.includes("\t")} onChange={(v) => toggle("\t", v)}>
            Sekme
          </Check>
          <Check checked={tab.includes(";")} onChange={(v) => toggle(";", v)}>
            Noktalı virgül
          </Check>
          <Check checked={tab.includes(",")} onChange={(v) => toggle(",", v)}>
            Virgül
          </Check>
          <Check checked={tab.includes(" ")} onChange={(v) => toggle(" ", v)}>
            Boşluk
          </Check>
          <label className="tb-check">
            <span>Diğer:</span>
            <input type="text" maxLength={1} value={other} onChange={(e) => setOther(e.target.value)} style={{ width: 28 }} />
          </label>
        </div>
      </FieldSet>
      <div className="tb-row-gap">
        <Check checked={consecutive} onChange={setConsecutive}>
          Ardışık sınırlayıcıları tek say
        </Check>
        <Field label="Metin niteleyicisi:" inline>
          <select value={qual} onChange={(e) => setQual(e.target.value)}>
            <option value={'"'}>"</option>
            <option value="'">'</option>
            <option value="">{"{yok}"}</option>
          </select>
        </Field>
      </div>
      <PreviewTable rows={preview} />
    </SimpleDialog>
  );
}

function PreviewTable({ rows }: { rows: string[][] }) {
  const n = Math.max(1, ...rows.map((r) => r.length));
  return (
    <div className="tb-preview-table">
      <table>
        <tbody>
          {rows.map((r, i) => (
            <tr key={i}>
              {Array.from({ length: n }, (_, j) => (
                <td key={j}>{r[j] ?? ""}</td>
              ))}
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  );
}

/* ------------------------------ Veri Doğrulama ------------------------------ */

const OPS: { id: NonNullable<ValidationRule["operator"]>; label: string }[] = [
  { id: "between", label: "arasında" },
  { id: "notBetween", label: "arasında değil" },
  { id: "equal", label: "eşit" },
  { id: "notEqual", label: "eşit değil" },
  { id: "greaterThan", label: "büyük" },
  { id: "lessThan", label: "küçük" },
  { id: "greaterThanOrEqual", label: "büyük veya eşit" },
  { id: "lessThanOrEqual", label: "küçük veya eşit" },
];

/** Kanonik kural değerini kullanıcıya gösterilecek metne çevirir. */
function ruleToLocal(v: string | undefined, type: ValidationRule["type"]): string {
  if (!v) return "";
  if (v.startsWith("=")) return toLocalFormula(v, "tr");
  if (type === "date" && /^\d+(\.\d+)?$/.test(v)) {
    const d = new Date(Date.UTC(1899, 11, 30) + Number(v) * 86400000);
    return `${String(d.getUTCDate()).padStart(2, "0")}.${String(d.getUTCMonth() + 1).padStart(2, "0")}.${d.getUTCFullYear()}`;
  }
  if (type === "list") return v;
  return v.replace(".", ",");
}

function ruleFromLocal(v: string, type: ValidationRule["type"]): string | undefined {
  const t = v.trim();
  if (!t) return undefined;
  if (t.startsWith("=")) return fromLocalFormula(t, "tr");
  if (type === "list") return t;
  if (type === "date") {
    const m = /^(\d{1,2})\.(\d{1,2})\.(\d{4})$/.exec(t);
    if (m) return String(dateToSerial(+m[3], +m[2], +m[1]));
  }
  const n = parseTrNumber(t);
  return Number.isFinite(n) ? String(n) : t;
}

export function DataValidationDialog({ ctl, onClose }: P) {
  const sheet = ctl.sheet;
  const { row, col } = ctl.sel.active;
  const existing = sheet.validations.find((v) => v.ranges.some((r) => contains(r, row, col)));
  const init: ValidationRule = existing?.rule ?? { type: "any", allowBlank: true, showDropdown: true };
  const [tab, setTab] = useState("settings");
  const [rule, setRule] = useState<ValidationRule>({ ...init });
  const [f1, setF1] = useState(ruleToLocal(init.formula1, init.type));
  const [f2, setF2] = useState(ruleToLocal(init.formula2, init.type));
  const [showPrompt, setShowPrompt] = useState(!!(init.prompt || init.promptTitle) || !existing);
  const [showError, setShowError] = useState(true);
  const upd = (p: Partial<ValidationRule>) => setRule({ ...rule, ...p });
  const op = rule.operator ?? "between";
  const needs2 = op === "between" || op === "notBetween";
  const typed = rule.type !== "any" && rule.type !== "list" && rule.type !== "custom";

  const save = (clear: boolean) => {
    const ranges = ctl.sel.ranges.map((r) => ({ ...r }));
    ctl.ops.sheetOp(sheet, "Veri Doğrulama", () => {
      // Seçimle çakışan kurallar seçim dışına kırpılır (hücre başına tek kural)
      sheet.validations = sheet.validations.map((v) => ({ ...v, ranges: subtractRects(v.ranges, ranges) })).filter((v) => v.ranges.length);
      if (clear || rule.type === "any") {
        if (!clear && (showPrompt && (rule.prompt || rule.promptTitle))) sheet.validations.push({ id: ctl.wb.newId("dv"), ranges, rule: { type: "any", prompt: rule.prompt, promptTitle: rule.promptTitle } });
        return;
      }
      const r: ValidationRule = {
        ...rule,
        operator: typed ? op : undefined,
        formula1: ruleFromLocal(f1, rule.type),
        formula2: typed && needs2 ? ruleFromLocal(f2, rule.type) : undefined,
        prompt: showPrompt ? rule.prompt : undefined,
        promptTitle: showPrompt ? rule.promptTitle : undefined,
        error: showError ? rule.error : undefined,
        errorTitle: showError ? rule.errorTitle : undefined,
        errorStyle: showError ? (rule.errorStyle ?? "stop") : "information",
      };
      sheet.validations.push({ id: ctl.wb.newId("dv"), ranges, rule: r });
    });
    onClose();
  };
  const ok = () => {
    if (rule.type !== "any" && !f1.trim()) {
      ctl.alert(rule.type === "list" ? "Kaynak için bir değer girmelisiniz." : "Bir değer girmelisiniz.");
      return;
    }
    save(false);
  };

  return (
    <Dialog
      open
      title="Veri Doğrulama"
      onClose={onClose}
      width={460}
      onSubmit={ok}
      footer={
        <>
          <DialogButton onClick={() => save(true)}>Tümünü Temizle</DialogButton>
          <OkCancel onOk={ok} onCancel={onClose} />
        </>
      }
    >
      <DialogTabs
        tabs={[
          { id: "settings", label: "Ayarlar" },
          { id: "input", label: "Giriş İletisi" },
          { id: "error", label: "Hata Uyarısı" },
        ]}
        active={tab}
        onChange={setTab}
      />
      <div className="tb-dv-body">
        {tab === "settings" && (
          <>
            <FieldSet legend="Doğrulama ölçütleri">
              <Field label="İzin Verilen:">
                <select value={rule.type} onChange={(e) => upd({ type: e.target.value as ValidationRule["type"] })}>
                  <option value="any">Her değer</option>
                  <option value="whole">Tam sayı</option>
                  <option value="decimal">Ondalık</option>
                  <option value="list">Liste</option>
                  <option value="date">Tarih</option>
                  <option value="textLength">Metin uzunluğu</option>
                  <option value="custom">Özel</option>
                </select>
              </Field>
              {rule.type !== "any" && (
                <div className="tb-row-gap">
                  <Check checked={rule.allowBlank !== false} onChange={(v) => upd({ allowBlank: v })}>
                    Boşluğu yoksay
                  </Check>
                  {rule.type === "list" && (
                    <Check checked={rule.showDropdown !== false} onChange={(v) => upd({ showDropdown: v })}>
                      Hücre içi açılan liste
                    </Check>
                  )}
                </div>
              )}
              {typed && (
                <Field label="Veri:">
                  <select value={op} onChange={(e) => upd({ operator: e.target.value as ValidationRule["operator"] })}>
                    {OPS.map((o) => (
                      <option key={o.id} value={o.id}>
                        {o.label}
                      </option>
                    ))}
                  </select>
                </Field>
              )}
              {rule.type === "list" && (
                <Field label="Kaynak:">
                  <input type="text" value={f1} onChange={(e) => setF1(e.target.value)} placeholder="Evet;Hayır ya da =$A$1:$A$5" />
                </Field>
              )}
              {rule.type === "custom" && (
                <Field label="Formül:">
                  <input type="text" value={f1} onChange={(e) => setF1(e.target.value)} placeholder="=A1>0" />
                </Field>
              )}
              {typed && (
                <>
                  <Field label={needs2 ? "En az:" : op === "equal" || op === "notEqual" ? "Değer:" : op.startsWith("greater") ? "En az:" : "En çok:"}>
                    <input type="text" value={f1} onChange={(e) => setF1(e.target.value)} placeholder={rule.type === "date" ? "gg.aa.yyyy" : ""} />
                  </Field>
                  {needs2 && (
                    <Field label="En çok:">
                      <input type="text" value={f2} onChange={(e) => setF2(e.target.value)} placeholder={rule.type === "date" ? "gg.aa.yyyy" : ""} />
                    </Field>
                  )}
                </>
              )}
            </FieldSet>
          </>
        )}
        {tab === "input" && (
          <>
            <Check checked={showPrompt} onChange={setShowPrompt}>
              Hücre seçildiğinde giriş iletisini göster
            </Check>
            <Field label="Başlık:">
              <input type="text" value={rule.promptTitle ?? ""} onChange={(e) => upd({ promptTitle: e.target.value })} />
            </Field>
            <Field label="Giriş iletisi:">
              <textarea rows={5} value={rule.prompt ?? ""} onChange={(e) => upd({ prompt: e.target.value })} />
            </Field>
          </>
        )}
        {tab === "error" && (
          <>
            <Check checked={showError} onChange={setShowError}>
              Geçersiz veri girildikten sonra hata uyarısını göster
            </Check>
            <Field label="Stil:">
              <select value={rule.errorStyle ?? "stop"} onChange={(e) => upd({ errorStyle: e.target.value as ValidationRule["errorStyle"] })}>
                <option value="stop">Durdur</option>
                <option value="warning">Uyarı</option>
                <option value="information">Bilgi</option>
              </select>
            </Field>
            <Field label="Başlık:">
              <input type="text" value={rule.errorTitle ?? ""} onChange={(e) => upd({ errorTitle: e.target.value })} />
            </Field>
            <Field label="Hata iletisi:">
              <textarea rows={5} value={rule.error ?? ""} onChange={(e) => upd({ error: e.target.value })} />
            </Field>
          </>
        )}
      </div>
    </Dialog>
  );
}

/* ------------------------------ Özel Otomatik Filtre ------------------------------ */

const FILTER_OPS: { id: string; label: string }[] = [
  { id: "eq", label: "eşittir" },
  { id: "ne", label: "eşit değildir" },
  { id: "gt", label: "büyüktür" },
  { id: "ge", label: "büyük veya eşittir" },
  { id: "lt", label: "küçüktür" },
  { id: "le", label: "küçük veya eşittir" },
  { id: "begins", label: "ile başlar" },
  { id: "ends", label: "ile biter" },
  { id: "contains", label: "içerir" },
  { id: "notContains", label: "içermez" },
];

export function CustomFilterDialog({ ctl, onClose, arg }: P) {
  const a = arg as { col: number; op1: string; filterRow: number };
  const f = D.filterAt(ctl.sheet, a.filterRow, a.col);
  const cur = f?.columns[a.col - (f?.range.c1 ?? 0)];
  const prev = cur?.kind === "custom" ? cur : null;
  const between = a.op1 === "between";
  const [op1, setOp1] = useState(prev?.op1 ?? (between ? "ge" : a.op1));
  const [v1, setV1] = useState(prev?.v1 ?? "");
  const [join, setJoin] = useState<"and" | "or">(prev?.join ?? "and");
  const [op2, setOp2] = useState(prev?.op2 ?? (between ? "le" : ""));
  const [v2, setV2] = useState(prev?.v2 ?? "");
  const values = useMemo(() => (f ? D.columnValues(ctl.wb, ctl.sheet, f, a.col).values : []), []); // eslint-disable-line react-hooks/exhaustive-deps
  if (!f) return null;
  const ok = () => {
    const crit: FilterCriteria = { kind: "custom", op1, v1, ...(op2 ? { join, op2, v2 } : {}) };
    D.setColumnFilter(ctl, f, a.col, crit);
    onClose();
  };
  const opSelect = (v: string, set: (s: string) => void, allowEmpty: boolean) => (
    <select value={v} onChange={(e) => set(e.target.value)}>
      {allowEmpty && <option value="" />}
      {FILTER_OPS.map((o) => (
        <option key={o.id} value={o.id}>
          {o.label}
        </option>
      ))}
    </select>
  );
  return (
    <SimpleDialog title="Özel Otomatik Filtre" onClose={onClose} onOk={ok} width={460}>
      <p className="tb-dlg-text">Satırları göster: {displayText(ctl.wb, ctl.sheet, f.range.r1, a.col)}</p>
      <datalist id="tb-cf-values">
        {values.map((v) => (
          <option key={v} value={v} />
        ))}
      </datalist>
      <div className="tb-row-gap">
        {opSelect(op1, setOp1, false)}
        <input type="text" list="tb-cf-values" value={v1} onChange={(e) => setV1(e.target.value)} data-autofocus="true" />
      </div>
      <div className="tb-row-gap">
        <Radio name="join" value="and" current={join} onChange={setJoin}>
          Ve
        </Radio>
        <Radio name="join" value="or" current={join} onChange={setJoin}>
          Veya
        </Radio>
      </div>
      <div className="tb-row-gap">
        {opSelect(op2, setOp2, true)}
        <input type="text" list="tb-cf-values" value={v2} onChange={(e) => setV2(e.target.value)} />
      </div>
      <p className="tb-fc-note">Tek bir karakteri göstermek için ? kullanın. Herhangi bir karakter dizisini göstermek için * kullanın.</p>
    </SimpleDialog>
  );
}

/* ------------------------------ Seriler ------------------------------ */

export function SeriesDialog({ ctl, onClose }: P) {
  const r = ctl.range;
  const [dir, setDir] = useState<"rows" | "cols">(r.c2 - r.c1 > r.r2 - r.r1 ? "rows" : "cols");
  const [type, setType] = useState<"linear" | "growth" | "date" | "auto">("linear");
  const [unit, setUnit] = useState<"day" | "weekday" | "month" | "year">("day");
  const [step, setStep] = useState("1");
  const [stop, setStop] = useState("");
  const ok = () => {
    const sheet = ctl.sheet;
    if (type === "auto") {
      const src = dir === "cols" ? { ...r, r2: r.r1 } : { ...r, c2: r.c1 };
      ctl.ops.fill(sheet, src, r, "series");
      onClose();
      return;
    }
    const st = parseTrNumber(step);
    const stopV = stop.trim() ? parseTrNumber(stop) : NaN;
    const entries: { row: number; col: number; input: Scalar }[] = [];
    const lines = dir === "cols" ? r.c2 - r.c1 + 1 : r.r2 - r.r1 + 1;
    const len = dir === "cols" ? r.r2 - r.r1 + 1 : r.c2 - r.c1 + 1;
    for (let i = 0; i < lines; i++) {
      const row0 = dir === "cols" ? r.r1 : r.r1 + i;
      const col0 = dir === "cols" ? r.c1 + i : r.c1;
      const start = ctl.wb.getValue(sheet, row0, col0);
      if (typeof start !== "number") continue;
      let v = start;
      for (let k = 1; k < (Number.isFinite(stopV) && len === 1 ? 100000 : len); k++) {
        if (type === "linear") v = v + st;
        else if (type === "growth") v = v * st;
        else v = addDate(start, unit, st * k);
        if (Number.isFinite(stopV) && (st >= 0 ? v > stopV : v < stopV)) break;
        entries.push({ row: dir === "cols" ? row0 + k : row0, col: dir === "cols" ? col0 : col0 + k, input: v });
      }
    }
    if (entries.length) ctl.ops.setInputs(sheet, entries, "Seriler");
    onClose();
  };
  return (
    <SimpleDialog title="Seriler" onClose={onClose} onOk={ok} width={420}>
      <div className="tb-series">
        <FieldSet legend="Seriler">
          <Radio name="sdir" value="rows" current={dir} onChange={setDir}>
            Satırlar
          </Radio>
          <Radio name="sdir" value="cols" current={dir} onChange={setDir}>
            Sütunlar
          </Radio>
        </FieldSet>
        <FieldSet legend="Tür">
          <Radio name="stype" value="linear" current={type} onChange={setType}>
            Doğrusal
          </Radio>
          <Radio name="stype" value="growth" current={type} onChange={setType}>
            Büyüme
          </Radio>
          <Radio name="stype" value="date" current={type} onChange={setType}>
            Tarih
          </Radio>
          <Radio name="stype" value="auto" current={type} onChange={setType}>
            Otomatik Doldurma
          </Radio>
        </FieldSet>
        <FieldSet legend="Tarih birimi">
          {(
            [
              ["day", "Gün"],
              ["weekday", "Hafta içi"],
              ["month", "Ay"],
              ["year", "Yıl"],
            ] as const
          ).map(([id, label]) => (
            <label key={id} className="tb-radio">
              <input type="radio" name="sunit" disabled={type !== "date"} checked={unit === id} onChange={() => setUnit(id)} />
              <span>{label}</span>
            </label>
          ))}
        </FieldSet>
      </div>
      <div className="tb-row-gap">
        <Field label="Adım değeri:" inline>
          <input type="text" value={step} onChange={(e) => setStep(e.target.value)} style={{ width: 80 }} />
        </Field>
        <Field label="Durdurma değeri:" inline>
          <input type="text" value={stop} onChange={(e) => setStop(e.target.value)} style={{ width: 80 }} />
        </Field>
      </div>
    </SimpleDialog>
  );
}

/* ------------------------------ Metin/CSV İçeri Aktar ------------------------------ */

export function TextImportDialog({ ctl, onClose, arg }: P) {
  const a = arg as { name: string; data: ArrayBuffer };
  const text = useMemo(() => decodeText(a.data), [a.data]);
  const [delim, setDelim] = useState<string>(() => csvToInputs(text.slice(0, 20000)).delimiter);
  const [locale, setLocale] = useState<"tr" | "en">(delim === "," ? "en" : "tr");
  const [dest, setDest] = useState<"new" | "here">("new");
  const preview = useMemo(() => parseCsv(text.split(/\r?\n/).slice(0, 12).join("\n"), delim), [text, delim]);
  const ok = () => {
    const { rows } = csvToInputs(text, { delimiter: delim, locale });
    const wb = ctl.wb;
    const go = () => {
      let sheet = ctl.sheet;
      let r0 = ctl.sel.active.row;
      let c0 = ctl.sel.active.col;
      if (dest === "new") {
        const base = a.name.replace(/\.[^.]+$/, "").slice(0, 31);
        sheet = wb.addSheet(base && wb.isValidSheetName(base) === null ? base : wb.nextSheetName());
        r0 = 0;
        c0 = 0;
      }
      rows.forEach((rr, i) =>
        rr.forEach((c, j) => {
          if (c.input === null) return;
          wb.setInputRaw(sheet, r0 + i, c0 + j, c.input);
          if (c.numFmt) wb.setStyleRaw(sheet, r0 + i, c0 + j, wb.styles.patch(wb.styleIdAt(sheet, r0 + i, c0 + j), { numFmt: c.numFmt }));
        }),
      );
      if (dest === "new") wb.activeIndex = wb.sheets.indexOf(sheet);
    };
    ctl.ops.snapshotOp("Metinden/CSV'den", go);
    ctl.selectCell(ctl.sel.active.row, ctl.sel.active.col);
    onClose();
  };
  return (
    <SimpleDialog title={a.name} onClose={onClose} onOk={ok} width={640} okLabel="Yükle">
      <div className="tb-row-gap">
        <Field label="Sınırlayıcı:" inline>
          <select value={delim} onChange={(e) => setDelim(e.target.value)}>
            <option value=";">Noktalı virgül</option>
            <option value=",">Virgül</option>
            <option value={"\t"}>Sekme</option>
            <option value="|">Dikey çizgi</option>
            <option value=" ">Boşluk</option>
          </select>
        </Field>
        <Field label="Sayı biçimi:" inline>
          <select value={locale} onChange={(e) => setLocale(e.target.value as "tr" | "en")}>
            <option value="tr">Türkçe (1.234,56)</option>
            <option value="en">İngilizce (1,234.56)</option>
          </select>
        </Field>
        <Field label="Hedef:" inline>
          <select value={dest} onChange={(e) => setDest(e.target.value as "new" | "here")}>
            <option value="new">Yeni çalışma sayfası</option>
            <option value="here">Var olan sayfa (etkin hücre)</option>
          </select>
        </Field>
      </div>
      <PreviewTable rows={preview} />
    </SimpleDialog>
  );
}
