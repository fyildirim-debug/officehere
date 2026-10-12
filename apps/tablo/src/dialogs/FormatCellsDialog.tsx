/** Hücreleri Biçimlendir iletişim kutusu: Sayı, Hizalama, Yazı Tipi, Kenarlık, Dolgu, Koruma. */
import { useMemo, useState } from "react";
import { Dialog, DialogTabs, Field, FieldSet } from "@officehere/ui";
import { FONT_LIST, FONT_SIZES, fontStack } from "@officehere/core";
import { formatFromLocal, formatToLocal, formatValue, looksLocalFormat } from "@officehere/formula";
import type { Controller } from "../controller";
import * as F from "../actions/format";
import { DEFAULT_FONT, DEFAULT_SIZE, type BorderSide, type BorderStyle, type CellStyle, type HAlign, type VAlign } from "../model/styles";
import { MAX_COLS, MAX_ROWS } from "@officehere/formula";
import { rect } from "../model/rect";
import { Check, ColorDrop, OkCancel } from "./common";
import { CATEGORIES, CURRENCY_SYMBOLS, CUSTOM_CODES, DATE_TYPES, FRACTION_TYPES, NEGATIVE_SAMPLES, SPECIAL_TYPES, TIME_TYPES, buildCode, detectCategory, type NumCategory, type NumOpts } from "./numberFormats";

type P = { ctl: Controller; onClose: () => void; arg?: unknown };
type Side = "top" | "bottom" | "left" | "right" | "insideH" | "insideV";
/** undefined: değişmedi, null: kaldır. */
type BorderDraft = Partial<Record<Side, BorderSide | null>>;

const TABS = [
  { id: "number", label: "Sayı" },
  { id: "align", label: "Hizalama" },
  { id: "font", label: "Yazı Tipi" },
  { id: "border", label: "Kenarlık" },
  { id: "fill", label: "Dolgu" },
  { id: "protect", label: "Koruma" },
];

const H_ALIGNS: { id: HAlign; label: string }[] = [
  { id: "general", label: "Genel" },
  { id: "left", label: "Sol (Girinti)" },
  { id: "center", label: "Orta" },
  { id: "right", label: "Sağ (Girinti)" },
  { id: "fill", label: "Dolgu" },
  { id: "justify", label: "İki Yana Yasla" },
  { id: "centerContinuous", label: "Seçim Arasında Ortala" },
  { id: "distributed", label: "Dağıtılmış (Girinti)" },
];
const V_ALIGNS: { id: VAlign; label: string }[] = [
  { id: "top", label: "Üst" },
  { id: "middle", label: "Orta" },
  { id: "bottom", label: "Alt" },
  { id: "justify", label: "İki Yana Yasla" },
  { id: "distributed", label: "Dağıtılmış" },
];
const LINE_STYLES: BorderStyle[] = ["hair", "dotted", "dashDotDot", "dashDot", "dashed", "thin", "slantDashDot", "mediumDashDot", "mediumDashed", "medium", "thick", "double"];
const PATTERNS: { id: string; label: string }[] = [
  { id: "", label: "Yok" },
  { id: "gray125", label: "%12,5 Gri" },
  { id: "lightGray", label: "%25 Gri" },
  { id: "mediumGray", label: "%50 Gri" },
  { id: "darkGray", label: "%75 Gri" },
  { id: "lightHorizontal", label: "İnce Yatay Şerit" },
  { id: "lightVertical", label: "İnce Dikey Şerit" },
  { id: "lightDown", label: "İnce Ters Çapraz Şerit" },
  { id: "lightUp", label: "İnce Çapraz Şerit" },
  { id: "lightGrid", label: "İnce Yatay Tarama" },
  { id: "darkHorizontal", label: "Yatay Şerit" },
  { id: "darkVertical", label: "Dikey Şerit" },
];

/** Kenarlık çizgisi örneği (SVG). */
export function LineSample({ style, color = "#000", width = 60 }: { style: BorderStyle; color?: string; width?: number }) {
  const w = style === "thick" ? 3 : style.startsWith("medium") ? 2 : 1;
  const dash: Record<string, string> = { hair: "1 1", dotted: "1 2", dashDotDot: "6 2 1 2 1 2", dashDot: "6 2 1 2", dashed: "4 2", mediumDashed: "6 3", mediumDashDot: "6 3 2 3", slantDashDot: "8 2 2 2" };
  if (style === "double")
    return (
      <svg width={width} height={8} aria-hidden="true">
        <path d={`M2 2.5H${width - 2}M2 5.5H${width - 2}`} stroke={color} strokeWidth={1} />
      </svg>
    );
  return (
    <svg width={width} height={8} aria-hidden="true">
      <path d={`M2 4H${width - 2}`} stroke={color} strokeWidth={w} strokeDasharray={dash[style]} />
    </svg>
  );
}

export function FormatCellsDialog({ ctl, onClose, arg }: P) {
  const init = useMemo(() => F.activeStyle(ctl), []); // eslint-disable-line react-hooks/exhaustive-deps
  const [tab, setTab] = useState((arg as { tab?: string } | undefined)?.tab ?? "number");
  const [draft, setDraft] = useState<CellStyle>({ ...init });
  const [touched, setTouched] = useState<Set<keyof CellStyle>>(new Set());
  const [borders, setBorders] = useState<BorderDraft>({});
  const r = ctl.range;
  const merged = !!ctl.sheet.mergeAt(r.r1, r.c1);
  const [merge, setMerge] = useState(merged);

  const set = <K extends keyof CellStyle>(k: K, v: CellStyle[K]) => {
    setDraft((d) => ({ ...d, [k]: v }));
    setTouched((t) => new Set(t).add(k));
  };

  const ok = () => {
    if (ctl.edit) ctl.commitEdit(null);
    const mark = ctl.ops.history.mark();
    const patch: Partial<CellStyle> = {};
    for (const k of touched) (patch as Record<string, unknown>)[k] = draft[k] ?? undefined;
    if (touched.size) F.applyStyle(ctl, patch, "Hücreleri Biçimlendir");
    if (Object.keys(borders).length) applyBorderDraft(ctl, borders);
    if (merge !== merged) F.mergeCells(ctl, merge ? "merge" : "unmerge");
    ctl.ops.history.squash(mark, "Hücreleri Biçimlendir");
    onClose();
  };

  return (
    <Dialog open title="Hücreleri Biçimlendir" onClose={onClose} width={600} onSubmit={ok} footer={<OkCancel onOk={ok} onCancel={onClose} />}>
      <DialogTabs tabs={TABS} active={tab} onChange={setTab} />
      <div className="tb-fc-body">
        {tab === "number" && <NumberTab ctl={ctl} code={draft.numFmt} onChange={(c) => set("numFmt", c === "General" ? undefined : c)} />}
        {tab === "align" && <AlignTab draft={draft} set={set} merge={merge} setMerge={setMerge} />}
        {tab === "font" && <FontTab draft={draft} set={set} />}
        {tab === "border" && <BorderTab init={init} multi={r.r1 !== r.r2 || r.c1 !== r.c2} draft={borders} setDraft={setBorders} />}
        {tab === "fill" && <FillTab draft={draft} set={set} />}
        {tab === "protect" && (
          <div className="tb-fc-protect">
            <Check checked={draft.locked !== false} onChange={(v) => set("locked", v ? undefined : false)}>
              Kilitli
            </Check>
            <Check checked={!!draft.hidden} onChange={(v) => set("hidden", v || undefined)}>
              Gizli
            </Check>
            <p className="tb-fc-note">Hücreleri kilitlemek ya da formülleri gizlemek, sayfa korunana kadar etkili olmaz (Gözden Geçir sekmesi, Koruma grubu, Sayfayı Koru düğmesi).</p>
          </div>
        )}
      </div>
    </Dialog>
  );
}

/* ------------------------------ Sayı ------------------------------ */

function NumberTab({ ctl, code, onChange }: { ctl: Controller; code: string | undefined; onChange: (c: string) => void }) {
  const detected = useMemo(() => detectCategory(code), []); // eslint-disable-line react-hooks/exhaustive-deps
  const [cat, setCat] = useState<NumCategory>(detected.cat);
  const [o, setO] = useState<NumOpts>(detected.opts);
  const [custom, setCustom] = useState(formatToLocal(code ?? "General"));
  const { row, col } = ctl.sel.active;
  const v = ctl.wb.getValue(ctl.sheet, row, col);
  const sample = v === null ? (cat === "date" || cat === "time" ? 45658.5 : 1234.1) : v;
  const cur = code ?? "General";
  const preview = formatValue(sample, cur);

  const pick = (c: NumCategory, opts = o, type?: string) => {
    setCat(c);
    const next = buildCode(c, opts, type ?? (c === "date" ? DATE_TYPES[0] : c === "time" ? TIME_TYPES[0] : c === "fraction" ? FRACTION_TYPES[0].code : c === "special" ? SPECIAL_TYPES[0].code : c === "custom" ? cur : undefined));
    if (c !== "custom") setCustom(formatToLocal(next));
    onChange(next);
  };
  const upd = (patch: Partial<NumOpts>) => {
    const n = { ...o, ...patch };
    setO(n);
    pick(cat, n);
  };
  const desc = CATEGORIES.find((c) => c.id === cat)?.desc;
  const typeList = (items: { code: string; label: string }[]) => (
    <Field label="Tür:">
      <div className="tb-list" style={{ height: 150 }}>
        {items.map((it) => (
          <div key={it.code} className={`tb-list-item${cur === it.code ? " is-active" : ""}`} onClick={() => pick(cat, o, it.code)}>
            {it.label}
          </div>
        ))}
      </div>
    </Field>
  );

  return (
    <div className="tb-fc-number">
      <Field label="Kategori:">
        <div className="tb-list" style={{ height: 260, width: 140 }}>
          {CATEGORIES.map((c) => (
            <div key={c.id} className={`tb-list-item${cat === c.id ? " is-active" : ""}`} onClick={() => pick(c.id)}>
              {c.label}
            </div>
          ))}
        </div>
      </Field>
      <div className="tb-fc-numright">
        <FieldSet legend="Örnek">
          <div className="tb-fc-sample" style={{ color: preview.color }}>
            {preview.text.replace(/\u0000./g, " ") || " "}
          </div>
        </FieldSet>
        {(cat === "number" || cat === "currency" || cat === "accounting" || cat === "percent" || cat === "scientific") && (
          <Field label="Ondalık basamak sayısı:" inline>
            <input type="number" min={0} max={30} value={o.decimals} onChange={(e) => upd({ decimals: Math.max(0, Math.min(30, Number(e.target.value) || 0)) })} style={{ width: 60 }} />
          </Field>
        )}
        {cat === "number" && (
          <Check checked={o.thousands} onChange={(t) => upd({ thousands: t })}>
            1000 Ayırıcısı (.) Kullan
          </Check>
        )}
        {(cat === "currency" || cat === "accounting") && (
          <Field label="Simge:" inline>
            <select value={o.symbol} onChange={(e) => upd({ symbol: e.target.value })}>
              {CURRENCY_SYMBOLS.map((s) => (
                <option key={s.id} value={s.id}>
                  {s.label}
                </option>
              ))}
            </select>
          </Field>
        )}
        {(cat === "number" || cat === "currency") && (
          <Field label="Negatif sayılar:">
            <div className="tb-list" style={{ height: 100 }}>
              {NEGATIVE_SAMPLES.map((s, i) => (
                <div key={s} className={`tb-list-item${o.negative === i ? " is-active" : ""}`} style={{ color: s.includes("kırmızı") ? "#C00000" : undefined }} onClick={() => upd({ negative: i })}>
                  {s.replace(" (kırmızı)", "")}
                </div>
              ))}
            </div>
          </Field>
        )}
        {cat === "date" && typeList(DATE_TYPES.map((c) => ({ code: c, label: formatValue(45658.5625, c).text })))}
        {cat === "time" && typeList(TIME_TYPES.map((c) => ({ code: c, label: formatValue(45658.5625, c).text })))}
        {cat === "fraction" && typeList(FRACTION_TYPES)}
        {cat === "special" && typeList(SPECIAL_TYPES)}
        {cat === "custom" && (
          <>
            <Field label="Tür:">
              <input
                type="text"
                value={custom}
                onChange={(e) => {
                  setCustom(e.target.value);
                  const t = e.target.value.trim();
                  if (t) onChange(looksLocalFormat(t) ? formatFromLocal(t) : t);
                }}
                style={{ width: "100%" }}
              />
            </Field>
            <div className="tb-list" style={{ height: 140 }}>
              {CUSTOM_CODES.map((c) => (
                <div
                  key={c}
                  className={`tb-list-item${cur === c ? " is-active" : ""}`}
                  onClick={() => {
                    setCustom(formatToLocal(c));
                    onChange(c);
                  }}
                >
                  {formatToLocal(c) === "General" ? "Genel" : formatToLocal(c)}
                </div>
              ))}
            </div>
          </>
        )}
        <p className="tb-fc-note">{desc}</p>
      </div>
    </div>
  );
}

/* ------------------------------ Hizalama ------------------------------ */

type Setter = <K extends keyof CellStyle>(k: K, v: CellStyle[K]) => void;

function AlignTab({ draft, set, merge, setMerge }: { draft: CellStyle; set: Setter; merge: boolean; setMerge: (v: boolean) => void }) {
  const vertical = draft.rotation === 255;
  const deg = vertical ? 0 : (draft.rotation ?? 0);
  return (
    <div className="tb-fc-align">
      <div>
        <FieldSet legend="Metin hizalaması">
          <Field label="Yatay:">
            <select value={draft.hAlign ?? "general"} onChange={(e) => set("hAlign", e.target.value === "general" ? undefined : (e.target.value as HAlign))}>
              {H_ALIGNS.map((a) => (
                <option key={a.id} value={a.id}>
                  {a.label}
                </option>
              ))}
            </select>
          </Field>
          <Field label="Girinti:" inline>
            <input type="number" min={0} max={15} value={draft.indent ?? 0} onChange={(e) => set("indent", Number(e.target.value) || undefined)} style={{ width: 56 }} />
          </Field>
          <Field label="Dikey:">
            <select value={draft.vAlign ?? "bottom"} onChange={(e) => set("vAlign", e.target.value === "bottom" ? undefined : (e.target.value as VAlign))}>
              {V_ALIGNS.map((a) => (
                <option key={a.id} value={a.id}>
                  {a.label}
                </option>
              ))}
            </select>
          </Field>
        </FieldSet>
        <FieldSet legend="Metin denetimi">
          <Check checked={!!draft.wrap} onChange={(v) => set("wrap", v || undefined)}>
            Metni kaydır
          </Check>
          <Check checked={!!draft.shrink} onChange={(v) => set("shrink", v || undefined)}>
            Uyacak şekilde daralt
          </Check>
          <Check checked={merge} onChange={setMerge}>
            Hücreleri birleştir
          </Check>
        </FieldSet>
      </div>
      <FieldSet legend="Yönlendirme">
        <div className="tb-fc-orient">
          <button type="button" className={`tb-fc-vtext${vertical ? " is-on" : ""}`} onClick={() => set("rotation", vertical ? undefined : 255)} aria-label="Dikey metin">
            {"Metin".split("").map((ch, i) => (
              <span key={i}>{ch}</span>
            ))}
          </button>
          <svg width="110" height="110" viewBox="-55 -55 110 110" className="tb-fc-dial" onClick={(e) => {
            const b = (e.currentTarget as SVGSVGElement).getBoundingClientRect();
            const x = e.clientX - b.left - 8;
            const y = e.clientY - b.top - b.height / 2;
            const a = Math.round((Math.atan2(-y, Math.max(1, x)) * 180) / Math.PI);
            set("rotation", Math.max(-90, Math.min(90, a)) || undefined);
          }}>
            <path d="M-47 -50 A50 50 0 0 1 -47 50" fill="none" stroke="#bbb" />
            {[-90, -45, 0, 45, 90].map((a) => (
              <circle key={a} cx={-47 + 50 * Math.cos((a * Math.PI) / 180)} cy={-50 * Math.sin((a * Math.PI) / 180)} r={2} fill="#666" />
            ))}
            <line x1={-47} y1={0} x2={-47 + 46 * Math.cos((deg * Math.PI) / 180)} y2={-46 * Math.sin((deg * Math.PI) / 180)} stroke="#C00000" strokeWidth={1.5} />
            <text x={-40} y={-4} fontSize={10} transform={`rotate(${-deg} -47 0)`}>Metin</text>
          </svg>
        </div>
        <Field label="Derece:" inline>
          <input type="number" min={-90} max={90} value={deg} disabled={vertical} onChange={(e) => set("rotation", Math.max(-90, Math.min(90, Number(e.target.value) || 0)) || undefined)} style={{ width: 56 }} />
        </Field>
      </FieldSet>
    </div>
  );
}

/* ------------------------------ Yazı tipi ------------------------------ */

function FontTab({ draft, set }: { draft: CellStyle; set: Setter }) {
  const font = draft.font ?? DEFAULT_FONT;
  const size = draft.size ?? DEFAULT_SIZE;
  const styleId = draft.bold && draft.italic ? "bi" : draft.bold ? "b" : draft.italic ? "i" : "n";
  const [fontText, setFontText] = useState(font);
  const [sizeText, setSizeText] = useState(String(size).replace(".", ","));
  return (
    <div className="tb-fc-font">
      <div className="tb-fc-fontcols">
        <Field label="Yazı tipi:">
          <input type="text" value={fontText} onChange={(e) => (setFontText(e.target.value), FONT_LIST.includes(e.target.value) && set("font", e.target.value))} />
          <div className="tb-list" style={{ height: 120 }}>
            {FONT_LIST.map((f) => (
              <div key={f} className={`tb-list-item${f === font ? " is-active" : ""}`} style={{ fontFamily: fontStack(f) }} onClick={() => (setFontText(f), set("font", f))}>
                {f}
              </div>
            ))}
          </div>
        </Field>
        <Field label="Yazı tipi stili:">
          <div className="tb-list" style={{ height: 146, width: 120 }}>
            {[
              ["n", "Normal"],
              ["i", "İtalik"],
              ["b", "Kalın"],
              ["bi", "Kalın İtalik"],
            ].map(([id, label]) => (
              <div
                key={id}
                className={`tb-list-item${styleId === id ? " is-active" : ""}`}
                onClick={() => {
                  set("bold", id.includes("b") || undefined);
                  set("italic", id.includes("i") || undefined);
                }}
              >
                {label}
              </div>
            ))}
          </div>
        </Field>
        <Field label="Boyut:">
          <input
            type="text"
            value={sizeText}
            style={{ width: 60 }}
            onChange={(e) => {
              setSizeText(e.target.value);
              const n = Number(e.target.value.replace(",", "."));
              if (n >= 1 && n <= 409) set("size", n);
            }}
          />
          <div className="tb-list" style={{ height: 120, width: 60 }}>
            {FONT_SIZES.map((s) => (
              <div key={s} className={`tb-list-item${s === size ? " is-active" : ""}`} onClick={() => (setSizeText(String(s).replace(".", ",")), set("size", s))}>
                {String(s).replace(".", ",")}
              </div>
            ))}
          </div>
        </Field>
      </div>
      <div className="tb-fc-fontrow">
        <Field label="Altı çizili:">
          <select value={draft.underline ?? ""} onChange={(e) => set("underline", (e.target.value || undefined) as CellStyle["underline"])}>
            <option value="">Yok</option>
            <option value="single">Tek</option>
            <option value="double">Çift</option>
          </select>
        </Field>
        <Field label="Renk:">
          <ColorDrop value={draft.color} onChange={(c) => set("color", c ?? undefined)} />
        </Field>
      </div>
      <div className="tb-fc-fontrow">
        <FieldSet legend="Efektler">
          <Check checked={!!draft.strike} onChange={(v) => set("strike", v || undefined)}>
            Üstü çizili
          </Check>
        </FieldSet>
        <FieldSet legend="Önizleme">
          <div
            className="tb-fc-preview"
            style={{
              fontFamily: fontStack(font),
              fontSize: `${Math.min(size, 28) * 1.33}px`,
              fontWeight: draft.bold ? 700 : 400,
              fontStyle: draft.italic ? "italic" : undefined,
              textDecoration: [draft.underline ? "underline" : "", draft.strike ? "line-through" : ""].join(" ").trim() || undefined,
              textDecorationStyle: draft.underline === "double" ? "double" : undefined,
              color: draft.color,
            }}
          >
            AaBbÇçĞğİıŞş
          </div>
        </FieldSet>
      </div>
    </div>
  );
}

/* ------------------------------ Kenarlık ------------------------------ */

function BorderTab({ init, multi, draft, setDraft }: { init: CellStyle; multi: boolean; draft: BorderDraft; setDraft: (d: BorderDraft) => void }) {
  const [line, setLine] = useState<BorderStyle>("thin");
  const [color, setColor] = useState<string | null>(null);
  const side = (): BorderSide => ({ style: line, color: color ?? "#000000" });
  const cur = (s: Side): BorderSide | null | undefined => {
    if (s in draft) return draft[s];
    if (s === "top") return init.bTop;
    if (s === "bottom") return init.bBottom;
    if (s === "left") return init.bLeft;
    if (s === "right") return init.bRight;
    return undefined;
  };
  const toggle = (s: Side) => {
    const c = cur(s);
    setDraft({ ...draft, [s]: c && c.style === line && (c.color ?? "#000000") === (color ?? "#000000") ? null : side() });
  };
  const preset = (p: "none" | "outline" | "inside") => {
    if (p === "none") setDraft({ top: null, bottom: null, left: null, right: null, insideH: null, insideV: null });
    else if (p === "outline") setDraft({ ...draft, top: side(), bottom: side(), left: side(), right: side() });
    else setDraft({ ...draft, insideH: side(), insideV: side() });
  };
  const ln = (s: Side, d: string) => {
    const b = cur(s);
    if (!b) return null;
    const w = b.style === "thick" ? 3 : b.style.startsWith("medium") ? 2 : 1;
    return <path d={d} stroke={b.color ?? "#000"} strokeWidth={w} strokeDasharray={b.style === "dashed" || b.style === "mediumDashed" ? "4 2" : b.style === "dotted" || b.style === "hair" ? "1 2" : undefined} />;
  };
  const btn = (s: Side, label: string, path: string) => (
    <button type="button" className={`tb-fc-bbtn${cur(s) ? " is-on" : ""}`} onClick={() => toggle(s)} aria-label={label} title={label}>
      <svg width="20" height="20" viewBox="0 0 20 20" aria-hidden="true">
        <rect x="3" y="3" width="14" height="14" fill="none" stroke="#bbb" strokeDasharray="1 1" />
        <path d={path} stroke="currentColor" strokeWidth="1.6" />
      </svg>
    </button>
  );
  return (
    <div className="tb-fc-border">
      <FieldSet legend="Çizgi">
        <div className="tb-fc-lines">
          {LINE_STYLES.map((s) => (
            <button key={s} type="button" className={`tb-fc-line${line === s ? " is-on" : ""}`} onClick={() => setLine(s)} aria-label={s}>
              <LineSample style={s} />
            </button>
          ))}
        </div>
        <Field label="Renk:">
          <ColorDrop value={color} onChange={setColor} />
        </Field>
      </FieldSet>
      <div className="tb-fc-bmain">
        <FieldSet legend="Hazır Ayarlar">
          <div className="tb-fc-presets">
            <button type="button" onClick={() => preset("none")}>Yok</button>
            <button type="button" onClick={() => preset("outline")}>Anahat</button>
            <button type="button" disabled={!multi} onClick={() => preset("inside")}>İç</button>
          </div>
        </FieldSet>
        <FieldSet legend="Kenarlık">
          <div className="tb-fc-bgrid">
            <div className="tb-fc-bcol">
              {btn("top", "Üst", "M3 3H17")}
              {multi && btn("insideH", "İç Yatay", "M3 10H17")}
              {btn("bottom", "Alt", "M3 17H17")}
            </div>
            <svg width="180" height="110" className="tb-fc-bpreview" aria-hidden="true">
              <rect x="10" y="10" width="160" height="90" fill="#fff" />
              {ln("top", "M10 10H170")}
              {ln("bottom", "M10 100H170")}
              {ln("left", "M10 10V100")}
              {ln("right", "M170 10V100")}
              {multi && ln("insideH", "M10 55H170")}
              {multi && ln("insideV", "M90 10V100")}
              <text x="90" y="59" fontSize="11" textAnchor="middle" fill="#666">Metin</text>
            </svg>
          </div>
          <div className="tb-fc-brow">
            {btn("left", "Sol", "M3 3V17")}
            {multi && btn("insideV", "İç Dikey", "M10 3V17")}
            {btn("right", "Sağ", "M17 3V17")}
          </div>
        </FieldSet>
        <p className="tb-fc-note">Seçili kenarlık stili, hazır ayarlar, önizleme diyagramı ya da yukarıdaki düğmeler tıklatılarak uygulanabilir.</p>
      </div>
    </div>
  );
}

function applyBorderDraft(ctl: Controller, d: BorderDraft): void {
  const sheet = ctl.sheet;
  const wb = ctl.wb;
  const rects = ctl.sel.ranges.map((r) => ({ r1: r.r1, c1: r.c1, r2: Math.min(r.r2, r.r1 + 5000), c2: Math.min(r.c2, r.c1 + 500) }));
  const touched = rects.map((r) => rect(Math.max(0, r.r1 - 1), Math.max(0, r.c1 - 1), Math.min(MAX_ROWS - 1, r.r2 + 1), Math.min(MAX_COLS - 1, r.c2 + 1)));
  const val = (s: BorderSide | null | undefined) => (s === null ? undefined : s);
  ctl.ops.cellsOp(sheet, touched, "Kenarlıklar", () => {
    for (const r of rects)
      for (let row = r.r1; row <= r.r2; row++)
        for (let col = r.c1; col <= r.c2; col++) {
          const p: Partial<CellStyle> = {};
          const top = row === r.r1 ? d.top : d.insideH;
          const bottom = row === r.r2 ? d.bottom : d.insideH;
          const left = col === r.c1 ? d.left : d.insideV;
          const right = col === r.c2 ? d.right : d.insideV;
          if (top !== undefined) p.bTop = val(top);
          if (bottom !== undefined) p.bBottom = val(bottom);
          if (left !== undefined) p.bLeft = val(left);
          if (right !== undefined) p.bRight = val(right);
          if (!Object.keys(p).length) continue;
          const id = wb.styleIdAt(sheet, row, col);
          wb.setStyleRaw(sheet, row, col, wb.styles.patch(id, p));
        }
  });
}

/* ------------------------------ Dolgu ------------------------------ */

function FillTab({ draft, set }: { draft: CellStyle; set: Setter }) {
  const [patColor, setPatColor] = useState<string | null>(draft.pattern?.color ?? "#000000");
  return (
    <div className="tb-fc-fill">
      <div>
        <Field label="Arka Plan Rengi:">
          <ColorDrop value={draft.fill} noneLabel="Renk Yok" width={170} onChange={(c) => set("fill", c ?? undefined)} />
        </Field>
        <button type="button" className="tb-link-btn" onClick={() => (set("fill", undefined), set("pattern", undefined))}>
          Dolgu Yok
        </button>
      </div>
      <div>
        <Field label="Desen Rengi:">
          <ColorDrop value={patColor} width={170} onChange={(c) => (setPatColor(c), draft.pattern && set("pattern", { ...draft.pattern, color: c ?? "#000000" }))} />
        </Field>
        <Field label="Desen Stili:">
          <select value={draft.pattern?.type ?? ""} onChange={(e) => set("pattern", e.target.value ? { type: e.target.value, color: patColor ?? "#000000" } : undefined)}>
            {PATTERNS.map((p) => (
              <option key={p.id} value={p.id}>
                {p.label}
              </option>
            ))}
          </select>
        </Field>
      </div>
      <FieldSet legend="Örnek">
        <div className="tb-fc-fillsample" style={{ background: draft.fill ?? "#fff" }} />
      </FieldSet>
    </div>
  );
}
