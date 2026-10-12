// Tablo iletişim kutuları: Sırala (tablo), Metni Tabloya Dönüştür, Tabloyu Metne Dönüştür, Tablo Seçenekleri
// (varsayılan hücre kenar boşlukları).
import { useMemo, useState } from "react";
import { Dialog, DialogButton, Field, FieldSet } from "@officehere/ui";
import { TableMap } from "@tiptap/pm/tables";
import { useBelge } from "../state/context";
import { SortDialog } from "./MiscDialogs";
import { DEFAULT_CELL_MARGINS, findTable, guessSeparator, selectedParagraphs, sortTable, tableToText, textRows, textToTable, type CellMargins, type SortType } from "../editor/tableTools";

const okCancel = (ok: () => void, close: () => void, disabled = false) => (
  <>
    <DialogButton primary onClick={ok} disabled={disabled}>
      Tamam
    </DialogButton>
    <DialogButton onClick={close}>İptal</DialogButton>
  </>
);

const SORT_TYPES: [SortType, string][] = [
  ["text", "Metin"],
  ["number", "Sayı"],
  ["date", "Tarih"],
];

export function TableSortDialog({ onClose }: { onClose: () => void }) {
  const { editor } = useBelge();
  const t = findTable(editor.state);
  const width = t ? TableMap.get(t.node).width : 0;
  const [header, setHeader] = useState(!!t?.node.attrs.headerRow);
  const [col, setCol] = useState(0);
  const [type, setType] = useState<SortType>("text");
  const [desc, setDesc] = useState(false);
  const [error, setError] = useState("");
  const names = Array.from({ length: width }, (_, i) => {
    const h = header && t ? (t.node.firstChild?.maybeChild(i)?.textContent.trim() ?? "") : "";
    return h || `Sütun ${i + 1}`;
  });
  const apply = () => {
    if (!sortTable(editor, { col, type, desc, header })) {
      setError("Dikey olarak birleştirilmiş hücreler içeren tablo sıralanamaz.");
      return;
    }
    onClose();
  };
  return (
    <Dialog open title="Sırala" width={440} onClose={onClose} onSubmit={apply} footer={okCancel(apply, onClose, !t)}>
      <FieldSet legend="Önce sıralama ölçütü">
        <div className="bl-dlg-grid">
          <Field label="Sütun:">
            <select aria-label="Sıralama sütunu" value={col} onChange={(e) => setCol(Number(e.target.value))}>
              {names.map((n, i) => (
                <option key={i} value={i}>
                  {n}
                </option>
              ))}
            </select>
          </Field>
          <Field label="Tür:">
            <select aria-label="Sıralama türü" value={type} onChange={(e) => setType(e.target.value as SortType)}>
              {SORT_TYPES.map(([v, l]) => (
                <option key={v} value={v}>
                  {l}
                </option>
              ))}
            </select>
          </Field>
        </div>
        <label className="bl-radio">
          <input type="radio" checked={!desc} onChange={() => setDesc(false)} /> Artan
        </label>
        <label className="bl-radio">
          <input type="radio" checked={desc} onChange={() => setDesc(true)} /> Azalan
        </label>
      </FieldSet>
      <FieldSet legend="Listemde">
        <label className="bl-radio">
          <input type="radio" checked={header} onChange={() => setHeader(true)} /> Üst bilgi satırı var
        </label>
        <label className="bl-radio">
          <input type="radio" checked={!header} onChange={() => setHeader(false)} /> Üst bilgi satırı yok
        </label>
      </FieldSet>
      {error && (
        <p className="bl-dlg-note is-error" role="alert">
          {error}
        </p>
      )}
    </Dialog>
  );
}

type SepChoice = "para" | "\t" | "," | ";" | "other";

function SeparatorField({ sep, setSep, other, setOther, paraLabel }: { sep: SepChoice; setSep: (s: SepChoice) => void; other: string; setOther: (s: string) => void; paraLabel: string }) {
  const opts: [SepChoice, string][] = [
    ["para", paraLabel],
    ["\t", "Sekmeler"],
    [",", "Virgüller"],
    [";", "Noktalı virgüller"],
  ];
  return (
    <div className="bl-sep-grid">
      {opts.map(([v, l]) => (
        <label key={l} className="bl-radio">
          <input type="radio" name="bl-sep" checked={sep === v} onChange={() => setSep(v)} /> {l}
        </label>
      ))}
      <label className="bl-radio">
        <input type="radio" name="bl-sep" checked={sep === "other"} onChange={() => setSep("other")} /> Diğer:
        <input type="text" aria-label="Diğer ayraç" maxLength={1} value={other} onFocus={() => setSep("other")} onChange={(e) => setOther(e.target.value)} />
      </label>
    </div>
  );
}

const sepValue = (sep: SepChoice, other: string) => (sep === "other" ? other || "-" : sep);

export function TextToTableDialog({ onClose }: { onClose: () => void }) {
  const { editor } = useBelge();
  const sel = useMemo(() => selectedParagraphs(editor.state), [editor]);
  const guess = sel ? guessSeparator(sel.nodes) : "para";
  const [sep, setSep] = useState<SepChoice>(guess as SepChoice);
  const [other, setOther] = useState("-");
  const s = sepValue(sep, other);
  const natural = sel && s !== "para" ? Math.max(1, ...textRows(sel.nodes, s, 1).map((r) => r.length)) : 1;
  const [colsIn, setColsIn] = useState<number | null>(null);
  const cols = colsIn ?? natural;
  const rows = sel ? textRows(sel.nodes, s, cols).length : 0;
  const apply = () => {
    if (sel) textToTable(editor, s, cols);
    onClose();
  };
  return (
    <Dialog open title="Metni Tabloya Dönüştür" width={400} onClose={onClose} onSubmit={apply} footer={okCancel(apply, onClose, !sel)}>
      {!sel && <p className="bl-dlg-note">Önce tabloya dönüştürülecek paragrafları seçin.</p>}
      <FieldSet legend="Tablo boyutu">
        <div className="bl-dlg-grid">
          <Field label="Sütun sayısı:">
            <input type="number" aria-label="Sütun sayısı" min={1} max={63} value={cols} onChange={(e) => setColsIn(Math.max(1, Math.min(63, Number(e.target.value) || 1)))} />
          </Field>
          <Field label="Satır sayısı:">
            <input type="number" aria-label="Satır sayısı" value={rows} disabled readOnly />
          </Field>
        </div>
      </FieldSet>
      <FieldSet legend="Metni ayırma yeri">
        <SeparatorField
          sep={sep}
          setSep={(v) => {
            setSep(v);
            setColsIn(null);
          }}
          other={other}
          setOther={setOther}
          paraLabel="Paragraflar"
        />
      </FieldSet>
    </Dialog>
  );
}

export function TableToTextDialog({ onClose }: { onClose: () => void }) {
  const { editor } = useBelge();
  const [sep, setSep] = useState<SepChoice>("\t");
  const [other, setOther] = useState("-");
  const apply = () => {
    tableToText(editor, sepValue(sep, other));
    onClose();
  };
  return (
    <Dialog open title="Tabloyu Metne Dönüştür" width={360} onClose={onClose} onSubmit={apply} footer={okCancel(apply, onClose, !findTable(editor.state))}>
      <FieldSet legend="Metni ayırma yeri">
        <SeparatorField sep={sep} setSep={setSep} other={other} setOther={setOther} paraLabel="Paragraf işaretleri" />
      </FieldSet>
    </Dialog>
  );
}

const CM_PER_PT = 2.54 / 72;
const toCm = (pt: number) => (pt * CM_PER_PT).toFixed(2).replace(".", ",");
const fromCm = (s: string) => Math.max(0, (parseFloat(s.replace(",", ".")) || 0) / CM_PER_PT);

export function CellMarginsDialog({ onClose }: { onClose: () => void }) {
  const { editor } = useBelge();
  const t = findTable(editor.state);
  const cur = (t?.node.attrs.cellMargins as CellMargins | null) ?? DEFAULT_CELL_MARGINS;
  const [v, setV] = useState({ top: toCm(cur.top), bottom: toCm(cur.bottom), left: toCm(cur.left), right: toCm(cur.right) });
  const apply = () => {
    const m: CellMargins = { top: fromCm(v.top), bottom: fromCm(v.bottom), left: fromCm(v.left), right: fromCm(v.right) };
    const round = (x: number) => Math.round(x * 10) / 10;
    const same = (Object.keys(m) as (keyof CellMargins)[]).every((k) => Math.abs(m[k] - DEFAULT_CELL_MARGINS[k]) < 0.05);
    editor
      .chain()
      .focus()
      .updateAttributes("table", { cellMargins: same ? null : { top: round(m.top), bottom: round(m.bottom), left: round(m.left), right: round(m.right) } })
      .run();
    onClose();
  };
  const input = (k: keyof typeof v, label: string) => (
    <Field label={`${label}:`}>
      <input type="text" aria-label={`${label} kenar boşluğu`} value={v[k]} onChange={(e) => setV((x) => ({ ...x, [k]: e.target.value }))} />
    </Field>
  );
  return (
    <Dialog open title="Tablo Seçenekleri" width={380} onClose={onClose} onSubmit={apply} footer={okCancel(apply, onClose, !t)}>
      <FieldSet legend="Varsayılan hücre kenar boşlukları (cm)">
        <div className="bl-dlg-grid">
          {input("top", "Üst")}
          {input("left", "Sol")}
          {input("bottom", "Alt")}
          {input("right", "Sağ")}
        </div>
      </FieldSet>
    </Dialog>
  );
}

/** Giriş › Sırala: imleç tablodaysa tablo satırlarını, değilse paragrafları sıralar. */
export function SortHost({ onClose }: { onClose: () => void }) {
  const { editor } = useBelge();
  return editor.isActive("table") ? <TableSortDialog onClose={onClose} /> : <SortDialog onClose={onClose} />;
}
