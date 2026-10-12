/** Koşullu biçimlendirme iletişim kutuları: hazır kurallar, Yeni Kural ve Kural Yöneticisi. */
import { useMemo, useState } from "react";
import { Dialog, DialogButton, Field } from "@officehere/ui";
import { formatRange, fromLocalFormula, parseRangeAddress, toLocalFormula } from "@officehere/formula";
import type { Controller } from "../controller";
import * as D from "../actions/data";
import { COLOR_SCALES, DATA_BAR_COLORS, HIGHLIGHT_PRESETS } from "../model/condfmt";
import type { CellStyle } from "../model/styles";
import type { CondRule } from "../model/types";
import { ColorDrop, OkCancel, SimpleDialog, parseTrNumber } from "./common";

type P = { ctl: Controller; onClose: () => void; arg?: unknown };

/** Kullanıcı girişini kanonik kural değerine çevirir. */
export function condValueFromLocal(s: string): string {
  const t = s.trim();
  if (t.startsWith("=")) return fromLocalFormula(t, "tr");
  const n = parseTrNumber(t);
  return t !== "" && Number.isFinite(n) ? String(n) : t;
}

export function condValueToLocal(s: string | undefined): string {
  if (!s) return "";
  if (s.startsWith("=")) return toLocalFormula(s, "tr");
  return /^-?\d+(\.\d+)?$/.test(s) ? s.replace(".", ",") : s;
}

const TIME_PERIODS: { id: string; label: string }[] = [
  { id: "yesterday", label: "Dün" },
  { id: "today", label: "Bugün" },
  { id: "tomorrow", label: "Yarın" },
  { id: "last7Days", label: "Son 7 gün içinde" },
  { id: "lastWeek", label: "Geçen hafta" },
  { id: "thisWeek", label: "Bu hafta" },
  { id: "nextWeek", label: "Gelecek hafta" },
  { id: "lastMonth", label: "Geçen ay" },
  { id: "thisMonth", label: "Bu ay" },
  { id: "nextMonth", label: "Gelecek ay" },
];

function StylePick({ value, onChange, custom, setCustom }: { value: string; onChange: (v: string) => void; custom: Partial<CellStyle>; setCustom: (s: Partial<CellStyle>) => void }) {
  return (
    <>
      <select value={value} onChange={(e) => onChange(e.target.value)}>
        {HIGHLIGHT_PRESETS.map((p) => (
          <option key={p.id} value={p.id}>
            {p.label}
          </option>
        ))}
        <option value="custom">Özel Biçim...</option>
      </select>
      {value === "custom" && <CustomStyle style={custom} onChange={setCustom} />}
    </>
  );
}

function CustomStyle({ style, onChange }: { style: Partial<CellStyle>; onChange: (s: Partial<CellStyle>) => void }) {
  return (
    <div className="tb-cf-custom">
      <label className="tb-check">
        <input type="checkbox" checked={!!style.bold} onChange={(e) => onChange({ ...style, bold: e.target.checked || undefined })} />
        <span>Kalın</span>
      </label>
      <label className="tb-check">
        <input type="checkbox" checked={!!style.italic} onChange={(e) => onChange({ ...style, italic: e.target.checked || undefined })} />
        <span>İtalik</span>
      </label>
      <Field label="Yazı rengi:" inline>
        <ColorDrop value={style.color} width={110} onChange={(c) => onChange({ ...style, color: c ?? undefined })} />
      </Field>
      <Field label="Dolgu:" inline>
        <ColorDrop value={style.fill} width={110} noneLabel="Renk Yok" onChange={(c) => onChange({ ...style, fill: c ?? undefined })} />
      </Field>
    </div>
  );
}

function styleFor(id: string, custom: Partial<CellStyle>): Partial<CellStyle> {
  return id === "custom" ? custom : (HIGHLIGHT_PRESETS.find((p) => p.id === id)?.style ?? HIGHLIGHT_PRESETS[0].style);
}

/** Giriş > Koşullu Biçimlendirme menüsündeki hazır kurallar. */
export function CondRuleDialog({ ctl, onClose, arg }: P) {
  const kind = (arg as { rule: string }).rule;
  if (kind === "new" || kind === "edit") return <NewRuleDialog ctl={ctl} onClose={onClose} arg={arg} />;
  return <PresetRuleDialog ctl={ctl} onClose={onClose} kind={kind} />;
}

function PresetRuleDialog({ ctl, onClose, kind }: { ctl: Controller; onClose: () => void; kind: string }) {
  const [v1, setV1] = useState(() => {
    if (kind.startsWith("top") || kind.startsWith("bottom")) return "10";
    // Varsayılan: seçimin ortalaması (yaklaşık)
    if (["greaterThan", "lessThan", "equal", "between"].includes(kind)) {
      const nums: number[] = [];
      const r = ctl.range;
      for (let row = r.r1; row <= Math.min(r.r2, r.r1 + 2000); row++)
        for (let col = r.c1; col <= Math.min(r.c2, r.c1 + 50); col++) {
          const v = ctl.wb.getValue(ctl.sheet, row, col);
          if (typeof v === "number") nums.push(v);
        }
      if (nums.length) {
        if (kind === "between") return String(Math.min(...nums)).replace(".", ",");
        const avg = nums.reduce((a, b) => a + b, 0) / nums.length;
        return String(Math.round(avg * 100) / 100).replace(".", ",");
      }
    }
    return "";
  });
  const [v2, setV2] = useState(() => {
    if (kind !== "between") return "";
    let max = -Infinity;
    const r = ctl.range;
    for (let row = r.r1; row <= Math.min(r.r2, r.r1 + 2000); row++)
      for (let col = r.c1; col <= Math.min(r.c2, r.c1 + 50); col++) {
        const v = ctl.wb.getValue(ctl.sheet, row, col);
        if (typeof v === "number" && v > max) max = v;
      }
    return Number.isFinite(max) ? String(max).replace(".", ",") : "";
  });
  const [period, setPeriod] = useState("yesterday");
  const [dup, setDup] = useState<"dup" | "uniq">("dup");
  const [styleId, setStyleId] = useState("lightRed");
  const [custom, setCustom] = useState<Partial<CellStyle>>({ fill: "#FFC7CE", color: "#9C0006" });

  const meta: Record<string, { title: string; text: string }> = {
    greaterThan: { title: "Büyüktür", text: "Şundan BÜYÜK olan hücreleri biçimlendir:" },
    lessThan: { title: "Küçüktür", text: "Şundan KÜÇÜK olan hücreleri biçimlendir:" },
    between: { title: "Arasında", text: "ARASINDA olan hücreleri biçimlendir:" },
    equal: { title: "Eşittir", text: "Şuna EŞİT olan hücreleri biçimlendir:" },
    containsText: { title: "İçeren Metin", text: "Şunu İÇEREN metin hücrelerini biçimlendir:" },
    timePeriod: { title: "Gerçekleşen Tarih", text: "Şu tarihi içeren hücreleri biçimlendir:" },
    duplicateValues: { title: "Yinelenen Değerler", text: "Şunu içeren hücreleri biçimlendir:" },
    top10: { title: "İlk 10 Öğe", text: "EN ÜSTTE yer alan hücreleri biçimlendir:" },
    top10Pct: { title: "İlk %10", text: "EN ÜSTTE yer alan hücreleri biçimlendir (%):" },
    bottom10: { title: "Son 10 Öğe", text: "EN ALTTA yer alan hücreleri biçimlendir:" },
    bottom10Pct: { title: "Son %10", text: "EN ALTTA yer alan hücreleri biçimlendir (%):" },
    aboveAverage: { title: "Ortalamanın Üstünde", text: "ORTALAMANIN ÜSTÜNDEKİ hücreleri biçimlendir:" },
    belowAverage: { title: "Ortalamanın Altında", text: "ORTALAMANIN ALTINDAKİ hücreleri biçimlendir:" },
  };
  const m = meta[kind] ?? meta.greaterThan;

  const ok = () => {
    const style = styleFor(styleId, custom);
    let rule: Omit<CondRule, "id" | "ranges">;
    switch (kind) {
      case "greaterThan":
      case "lessThan":
      case "equal":
        rule = { type: "cellIs", operator: kind, values: [condValueFromLocal(v1)], style };
        break;
      case "between":
        rule = { type: "cellIs", operator: "between", values: [condValueFromLocal(v1), condValueFromLocal(v2)], style };
        break;
      case "containsText":
        rule = { type: "containsText", text: v1, style };
        break;
      case "timePeriod":
        rule = { type: "timePeriod", timePeriod: period, style };
        break;
      case "duplicateValues":
        rule = { type: dup === "dup" ? "duplicateValues" : "uniqueValues", style };
        break;
      case "aboveAverage":
      case "belowAverage":
        rule = { type: "aboveAverage", below: kind === "belowAverage", style };
        break;
      default:
        rule = { type: "top10", rank: Math.max(1, Math.round(parseTrNumber(v1)) || 10), percent: kind.endsWith("Pct"), bottom: kind.startsWith("bottom"), style };
    }
    D.addCondRule(ctl, rule);
    onClose();
  };
  const plainInput = ["greaterThan", "lessThan", "equal", "containsText", "between"].includes(kind);
  const rankInput = kind.startsWith("top") || kind.startsWith("bottom");
  return (
    <SimpleDialog title={m.title} onClose={onClose} onOk={ok} width={kind === "between" ? 560 : 480}>
      <p className="tb-dlg-text">{m.text}</p>
      <div className="tb-row-gap tb-cf-row">
        {plainInput && <input type="text" value={v1} onChange={(e) => setV1(e.target.value)} data-autofocus="true" style={{ width: 150 }} />}
        {kind === "between" && (
          <>
            <span>ile</span>
            <input type="text" value={v2} onChange={(e) => setV2(e.target.value)} style={{ width: 150 }} />
          </>
        )}
        {rankInput && <input type="number" min={1} max={1000} value={v1} onChange={(e) => setV1(e.target.value)} style={{ width: 70 }} />}
        {kind === "timePeriod" && (
          <select value={period} onChange={(e) => setPeriod(e.target.value)}>
            {TIME_PERIODS.map((t) => (
              <option key={t.id} value={t.id}>
                {t.label}
              </option>
            ))}
          </select>
        )}
        {kind === "duplicateValues" && (
          <select value={dup} onChange={(e) => setDup(e.target.value as "dup" | "uniq")}>
            <option value="dup">Yinelenen</option>
            <option value="uniq">Benzersiz</option>
          </select>
        )}
        <span>{kind === "duplicateValues" ? "değerler ve" : "ile"}</span>
        <StylePick value={styleId} onChange={setStyleId} custom={custom} setCustom={setCustom} />
      </div>
    </SimpleDialog>
  );
}

/* ------------------------------ Yeni Biçimlendirme Kuralı ------------------------------ */

const RULE_KINDS = [
  { id: "scale", label: "Tüm hücreleri değerlerine göre biçimlendir" },
  { id: "cell", label: "Yalnızca şunu içeren hücreleri biçimlendir" },
  { id: "top", label: "Yalnızca en üst veya en alt sıradaki değerleri biçimlendir" },
  { id: "avg", label: "Yalnızca ortalamanın üstünde veya altında olan değerleri biçimlendir" },
  { id: "dup", label: "Yalnızca benzersiz veya yinelenen değerleri biçimlendir" },
  { id: "expr", label: "Biçimlendirilecek hücreleri belirlemek için formül kullan" },
];

function kindOfRule(r: CondRule): string {
  if (r.type === "colorScale" || r.type === "dataBar") return "scale";
  if (r.type === "top10") return "top";
  if (r.type === "aboveAverage") return "avg";
  if (r.type === "duplicateValues" || r.type === "uniqueValues") return "dup";
  if (r.type === "expression") return "expr";
  return "cell";
}

function NewRuleDialog({ ctl, onClose, arg }: P) {
  const a = arg as { rule: string; edit?: CondRule; onDone?: (r: CondRule) => void };
  const e = a.edit;
  const [kind, setKind] = useState(e ? kindOfRule(e) : "cell");
  const [scaleType, setScaleType] = useState<"colorScale2" | "colorScale3" | "dataBar">(e?.type === "dataBar" ? "dataBar" : e?.colors?.length === 2 ? "colorScale2" : "colorScale3");
  const [colors, setColors] = useState<string[]>(e?.colors ?? COLOR_SCALES[0].colors);
  const [barColor, setBarColor] = useState(e?.color ?? DATA_BAR_COLORS[0]);
  const [cellType, setCellType] = useState<string>(e && e.type !== "cellIs" && kindOfRule(e) === "cell" ? e.type : "cellIs");
  const [op, setOp] = useState<NonNullable<CondRule["operator"]>>(e?.operator ?? "between");
  const [v1, setV1] = useState(condValueToLocal(e?.values?.[0]) || e?.text || "");
  const [v2, setV2] = useState(condValueToLocal(e?.values?.[1]));
  const [bottom, setBottom] = useState(!!e?.bottom);
  const [rank, setRank] = useState(String(e?.rank ?? 10));
  const [percent, setPercent] = useState(!!e?.percent);
  const [below, setBelow] = useState(!!e?.below);
  const [uniq, setUniq] = useState(e?.type === "uniqueValues");
  const [expr, setExpr] = useState(condValueToLocal(e?.values?.[0]) || "=");
  const [style, setStyle] = useState<Partial<CellStyle>>(e?.style ?? { fill: "#FFC7CE", color: "#9C0006" });

  const build = (): Omit<CondRule, "id" | "ranges"> | null => {
    switch (kind) {
      case "scale":
        return scaleType === "dataBar" ? { type: "dataBar", color: barColor } : { type: "colorScale", colors: scaleType === "colorScale2" ? [colors[0], colors[colors.length - 1]] : colors.length === 3 ? colors : [colors[0], "#FFEB84", colors[colors.length - 1]] };
      case "cell":
        if (cellType === "cellIs") return { type: "cellIs", operator: op, values: op === "between" || op === "notBetween" ? [condValueFromLocal(v1), condValueFromLocal(v2)] : [condValueFromLocal(v1)], style };
        if (["containsText", "notContainsText", "beginsWith", "endsWith"].includes(cellType)) return { type: cellType as CondRule["type"], text: v1, style };
        return { type: cellType as CondRule["type"], style };
      case "top":
        return { type: "top10", rank: Math.max(1, Number(rank) || 10), percent, bottom, style };
      case "avg":
        return { type: "aboveAverage", below, style };
      case "dup":
        return { type: uniq ? "uniqueValues" : "duplicateValues", style };
      case "expr":
        if (!expr.trim() || expr.trim() === "=") return null;
        return { type: "expression", values: [condValueFromLocal(expr.startsWith("=") ? expr : "=" + expr)], style };
    }
    return null;
  };
  const ok = () => {
    const r = build();
    if (!r) {
      ctl.alert("Geçerli bir formül girin.");
      return;
    }
    if (a.onDone && e) a.onDone({ ...e, ...r, id: e.id, ranges: e.ranges });
    else D.addCondRule(ctl, r);
    onClose();
  };
  return (
    <SimpleDialog title={e ? "Biçimlendirme Kuralını Düzenle" : "Yeni Biçimlendirme Kuralı"} onClose={onClose} onOk={ok} width={560}>
      <Field label="Bir Kural Türü Seçin:">
        <div className="tb-list" style={{ height: 134 }}>
          {RULE_KINDS.map((k) => (
            <div key={k.id} className={`tb-list-item${kind === k.id ? " is-active" : ""}`} onClick={() => setKind(k.id)}>
              ► {k.label}
            </div>
          ))}
        </div>
      </Field>
      <div className="tb-cf-edit">
        <b>Kural Açıklamasını Düzenleyin:</b>
        {kind === "scale" && (
          <>
            <Field label="Biçim Stili:" inline>
              <select value={scaleType} onChange={(ev) => setScaleType(ev.target.value as typeof scaleType)}>
                <option value="colorScale2">2 Renk Ölçeği</option>
                <option value="colorScale3">3 Renk Ölçeği</option>
                <option value="dataBar">Veri Çubuğu</option>
              </select>
            </Field>
            {scaleType === "dataBar" ? (
              <Field label="Çubuk Rengi:" inline>
                <ColorDrop value={barColor} onChange={(c) => setBarColor(c ?? DATA_BAR_COLORS[0])} />
              </Field>
            ) : (
              <div className="tb-row-gap">
                <Field label="En Düşük:" inline>
                  <ColorDrop value={colors[0]} width={100} onChange={(c) => setColors([c ?? "#F8696B", ...colors.slice(1)])} />
                </Field>
                {scaleType === "colorScale3" && (
                  <Field label="Orta Nokta:" inline>
                    <ColorDrop value={colors.length === 3 ? colors[1] : "#FFEB84"} width={100} onChange={(c) => setColors([colors[0], c ?? "#FFEB84", colors[colors.length - 1]])} />
                  </Field>
                )}
                <Field label="En Yüksek:" inline>
                  <ColorDrop value={colors[colors.length - 1]} width={100} onChange={(c) => setColors([...colors.slice(0, -1), c ?? "#63BE7B"])} />
                </Field>
              </div>
            )}
          </>
        )}
        {kind === "cell" && (
          <div className="tb-row-gap">
            <select value={cellType} onChange={(ev) => setCellType(ev.target.value)}>
              <option value="cellIs">Hücre Değeri</option>
              <option value="containsText">Belirli Metin (içeren)</option>
              <option value="notContainsText">Belirli Metin (içermeyen)</option>
              <option value="beginsWith">Belirli Metin (ile başlayan)</option>
              <option value="endsWith">Belirli Metin (ile biten)</option>
              <option value="timePeriod">Gerçekleşen Tarihler</option>
              <option value="containsBlanks">Boşluklar</option>
              <option value="notContainsBlanks">Boşluk Olmayanlar</option>
              <option value="containsErrors">Hatalar</option>
            </select>
            {cellType === "cellIs" && (
              <>
                <select value={op} onChange={(ev) => setOp(ev.target.value as typeof op)}>
                  <option value="between">arasında</option>
                  <option value="notBetween">arasında değil</option>
                  <option value="equal">eşit</option>
                  <option value="notEqual">eşit değil</option>
                  <option value="greaterThan">büyük</option>
                  <option value="lessThan">küçük</option>
                  <option value="greaterThanOrEqual">büyük veya eşit</option>
                  <option value="lessThanOrEqual">küçük veya eşit</option>
                </select>
                <input type="text" value={v1} onChange={(ev) => setV1(ev.target.value)} style={{ width: 90 }} />
                {(op === "between" || op === "notBetween") && (
                  <>
                    <span>ve</span>
                    <input type="text" value={v2} onChange={(ev) => setV2(ev.target.value)} style={{ width: 90 }} />
                  </>
                )}
              </>
            )}
            {["containsText", "notContainsText", "beginsWith", "endsWith"].includes(cellType) && <input type="text" value={v1} onChange={(ev) => setV1(ev.target.value)} />}
          </div>
        )}
        {kind === "top" && (
          <div className="tb-row-gap">
            <select value={bottom ? "bottom" : "top"} onChange={(ev) => setBottom(ev.target.value === "bottom")}>
              <option value="top">İlk</option>
              <option value="bottom">Son</option>
            </select>
            <input type="number" value={rank} min={1} onChange={(ev) => setRank(ev.target.value)} style={{ width: 70 }} />
            <label className="tb-check">
              <input type="checkbox" checked={percent} onChange={(ev) => setPercent(ev.target.checked)} />
              <span>seçili aralığın yüzdesi</span>
            </label>
          </div>
        )}
        {kind === "avg" && (
          <select value={below ? "below" : "above"} onChange={(ev) => setBelow(ev.target.value === "below")}>
            <option value="above">üstünde</option>
            <option value="below">altında</option>
          </select>
        )}
        {kind === "dup" && (
          <select value={uniq ? "uniq" : "dup"} onChange={(ev) => setUniq(ev.target.value === "uniq")}>
            <option value="dup">yinelenen</option>
            <option value="uniq">benzersiz</option>
          </select>
        )}
        {kind === "expr" && (
          <Field label="Bu formül doğru olduğunda değerleri biçimlendir:">
            <input type="text" value={expr} onChange={(ev) => setExpr(ev.target.value)} placeholder="=$B2>100" />
          </Field>
        )}
        {kind !== "scale" && (
          <div className="tb-cf-preview-row">
            <span>Önizleme:</span>
            <span className="tb-cf-preview" style={{ background: style.fill, color: style.color, fontWeight: style.bold ? 700 : undefined, fontStyle: style.italic ? "italic" : undefined }}>
              AaBbÇçĞğ
            </span>
            <CustomStyle style={style} onChange={setStyle} />
          </div>
        )}
      </div>
    </SimpleDialog>
  );
}

/* ------------------------------ Kural Yöneticisi ------------------------------ */

function describe(r: CondRule): string {
  const v = (i: number) => condValueToLocal(r.values?.[i]);
  const ops: Record<string, string> = { between: "arasında", notBetween: "arasında değil", equal: "=", notEqual: "<>", greaterThan: ">", lessThan: "<", greaterThanOrEqual: ">=", lessThanOrEqual: "<=" };
  switch (r.type) {
    case "cellIs":
      return r.operator === "between" || r.operator === "notBetween" ? `Hücre Değeri ${v(0)} ile ${v(1)} ${ops[r.operator]}` : `Hücre Değeri ${ops[r.operator ?? "equal"]} ${v(0)}`;
    case "containsText":
      return `"${r.text}" metnini içeren hücre`;
    case "notContainsText":
      return `"${r.text}" metnini içermeyen hücre`;
    case "beginsWith":
      return `"${r.text}" ile başlayan`;
    case "endsWith":
      return `"${r.text}" ile biten`;
    case "duplicateValues":
      return "Yinelenen Değerler";
    case "uniqueValues":
      return "Benzersiz Değerler";
    case "top10":
      return `${r.bottom ? "Son" : "İlk"} ${r.rank ?? 10}${r.percent ? "%" : ""}`;
    case "aboveAverage":
      return r.below ? "Ortalamanın Altında" : "Ortalamanın Üstünde";
    case "dataBar":
      return "Veri Çubuğu";
    case "colorScale":
      return `${r.colors?.length ?? 3} Renk Ölçeği`;
    case "expression":
      return `Formül: ${v(0)}`;
    case "containsBlanks":
      return "Boş Hücreler";
    case "notContainsBlanks":
      return "Boş Olmayan Hücreler";
    case "containsErrors":
      return "Hatalar";
    case "timePeriod":
      return `Tarih: ${TIME_PERIODS.find((t) => t.id === r.timePeriod)?.label ?? r.timePeriod}`;
  }
  return r.type;
}

export function CondManagerDialog({ ctl, onClose }: P) {
  const sheet = ctl.sheet;
  const [scope, setScope] = useState<"selection" | "sheet">("selection");
  const [rules, setRules] = useState<CondRule[]>(() => structuredClone(sheet.condFormats));
  const [sel, setSel] = useState<string | null>(null);
  const [editing, setEditing] = useState<CondRule | null>(null);
  const [adding, setAdding] = useState(false);
  const selRanges = ctl.sel.ranges;
  const shown = useMemo(
    () => (scope === "sheet" ? rules : rules.filter((r) => r.ranges.some((a) => selRanges.some((s) => s.r1 <= a.r2 && a.r1 <= s.r2 && s.c1 <= a.c2 && a.c1 <= s.c2)))),
    [rules, scope, selRanges],
  );
  const move = (d: -1 | 1) => {
    const i = rules.findIndex((r) => r.id === sel);
    const j = i + d;
    if (i < 0 || j < 0 || j >= rules.length) return;
    const n = [...rules];
    [n[i], n[j]] = [n[j], n[i]];
    setRules(n);
  };
  const apply = () => {
    ctl.ops.sheetOp(sheet, "Kuralları Yönet", () => {
      sheet.condFormats = rules;
    });
  };
  const ok = () => {
    apply();
    onClose();
  };
  if (editing) {
    return (
      <NewRuleDialog
        ctl={ctl}
        onClose={() => setEditing(null)}
        arg={{ rule: "edit", edit: editing, onDone: (r: CondRule) => setRules(rules.map((x) => (x.id === r.id ? r : x))) }}
      />
    );
  }
  if (adding) {
    return (
      <NewRuleDialog
        ctl={ctl}
        onClose={() => setAdding(false)}
        arg={{
          rule: "edit",
          edit: { id: ctl.wb.newId("cf"), type: "cellIs", operator: "greaterThan", values: ["0"], ranges: selRanges.map((r) => ({ ...r })), style: { fill: "#FFC7CE", color: "#9C0006" } },
          onDone: (r: CondRule) => setRules([r, ...rules]),
        }}
      />
    );
  }
  const cur = rules.find((r) => r.id === sel);
  return (
    <Dialog
      open
      title="Koşullu Biçimlendirme Kuralları Yöneticisi"
      onClose={onClose}
      width={680}
      onSubmit={ok}
      footer={
        <>
          <OkCancel onOk={ok} onCancel={onClose} />
          <DialogButton onClick={apply}>Uygula</DialogButton>
        </>
      }
    >
      <Field label="Biçimlendirme kurallarını göster:" inline>
        <select value={scope} onChange={(e) => setScope(e.target.value as "selection" | "sheet")}>
          <option value="selection">Geçerli Seçim</option>
          <option value="sheet">Bu Çalışma Sayfası</option>
        </select>
      </Field>
      <div className="tb-row-gap">
        <DialogButton onClick={() => setAdding(true)}>Yeni Kural...</DialogButton>
        <DialogButton disabled={!cur} onClick={() => cur && setEditing(cur)}>
          Kuralı Düzenle...
        </DialogButton>
        <DialogButton disabled={!cur} onClick={() => (setRules(rules.filter((r) => r.id !== sel)), setSel(null))}>
          Kuralı Sil
        </DialogButton>
        <DialogButton disabled={!cur} onClick={() => move(-1)}>
          ▲
        </DialogButton>
        <DialogButton disabled={!cur} onClick={() => move(1)}>
          ▼
        </DialogButton>
      </div>
      <div className="tb-cfm">
        <div className="tb-cfm-row tb-cfm-head">
          <span>Kural (gösterilen sırada uygulanır)</span>
          <span>Biçim</span>
          <span>Uygulandığı yer</span>
          <span>Doğruysa Durdur</span>
        </div>
        {shown.length === 0 && <div className="tb-cfm-empty">Bu kapsamda kural yok.</div>}
        {shown.map((r) => (
          <div key={r.id} className={`tb-cfm-row${r.id === sel ? " is-active" : ""}`} onClick={() => setSel(r.id)} onDoubleClick={() => setEditing(r)}>
            <span>{describe(r)}</span>
            <span>
              <span
                className="tb-cf-preview"
                style={{
                  background: r.type === "dataBar" ? `linear-gradient(90deg, ${r.color} 60%, transparent 60%)` : r.type === "colorScale" ? `linear-gradient(90deg, ${(r.colors ?? []).join(",")})` : r.style?.fill,
                  color: r.style?.color,
                  fontWeight: r.style?.bold ? 700 : undefined,
                }}
              >
                AaBbÇç
              </span>
            </span>
            <span>
              <input
                type="text"
                defaultValue={"=" + r.ranges.map((x) => formatRange(x).replace(/([A-Z]+)(\d+)/g, "$$$1$$$2")).join(";")}
                onBlur={(e) => {
                  const parts = e.target.value.replace(/^=/, "").split(/[;,]/).map((s) => parseRangeAddress(s.replace(/\$/g, ""))).filter((x): x is NonNullable<typeof x> => !!x);
                  if (parts.length) setRules(rules.map((x) => (x.id === r.id ? { ...x, ranges: parts } : x)));
                }}
              />
            </span>
            <span>
              <input type="checkbox" checked={!!r.stopIfTrue} onChange={(e) => setRules(rules.map((x) => (x.id === r.id ? { ...x, stopIfTrue: e.target.checked } : x)))} />
            </span>
          </div>
        ))}
      </div>
    </Dialog>
  );
}
