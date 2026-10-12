/** Çeşitli iletişim kutuları: uyarı, boyutlar, hücre ekle/sil, sayfa işlemleri, köprü, not, simge, yakınlaştırma, koruma, Özel Yapıştır. */
import { useEffect, useMemo, useState } from "react";
import { Dialog, DialogButton, Field } from "@officehere/ui";
import { addr, quoteSheetName } from "@officehere/formula";
import type { Controller } from "../controller";
import { parseRangeLoose, validationList } from "../controller";
import * as E from "../actions/edit";
import * as D from "../actions/data";
import { widthCharsToPx, widthPxToChars, ptToPx, pxToPt } from "../model/axis";
import { rect, rectToA1 } from "../model/rect";
import { currentRegion } from "../model/selection";
import { displayText } from "../model/display";
import { Check, OkCancel, Radio, SimpleDialog, fmtTr, parseTrNumber } from "./common";
import type { PasteOptions } from "../model/clipboard";
import { getUserName } from "@officehere/ui";
import { sizeSelection } from "../actions/format";

type P = { ctl: Controller; onClose: () => void; arg?: unknown };

export function AlertDialog({ onClose, arg }: P) {
  const a = arg as { message: string; title?: string };
  return (
    <Dialog open title={a.title ?? "OfficeHere Tablo"} onClose={onClose} width={420} onSubmit={onClose} footer={<DialogButton primary onClick={onClose}>Tamam</DialogButton>}>
      <p className="oh-dialog-message" style={{ whiteSpace: "pre-wrap" }}>
        {a.message}
      </p>
    </Dialog>
  );
}

export function ConfirmDialog({ onClose, arg }: P) {
  const a = arg as { message: string; title?: string; onOk: () => void; okLabel?: string };
  return (
    <SimpleDialog title={a.title ?? "OfficeHere Tablo"} onClose={onClose} onOk={() => (onClose(), a.onOk())} okLabel={a.okLabel} width={420}>
      <p className="oh-dialog-message" style={{ whiteSpace: "pre-wrap" }}>
        {a.message}
      </p>
    </SimpleDialog>
  );
}

export function SizeDialog({ ctl, onClose, arg }: P & { arg: "rowHeight" | "colWidth" | "defaultWidth" }) {
  const sheet = ctl.sheet;
  const r = ctl.range;
  const isRow = arg === "rowHeight";
  const initial = isRow ? pxToPt(sheet.rows.size(r.r1)) : arg === "colWidth" ? widthPxToChars(sheet.cols.size(r.c1)) : widthPxToChars(sheet.cols.defaultSize);
  const [v, setV] = useState(fmtTr(initial));
  const title = isRow ? "Satır Yüksekliği" : arg === "colWidth" ? "Sütun Genişliği" : "Standart Genişlik";
  const ok = () => {
    const n = parseTrNumber(v);
    const max = isRow ? 409 : 255;
    if (!(n >= 0 && n <= max)) {
      ctl.alert(`Girdiğiniz değer geçerli değil. 0 ile ${max} arasında bir sayı girin.`);
      return;
    }
    if (arg === "defaultWidth") {
      ctl.ops.sheetOp(sheet, "Standart Genişlik", () => sheet.cols.setDefault(widthCharsToPx(n)));
    } else {
      sizeSelection(ctl, isRow ? "row" : "col", isRow ? ptToPx(n) : widthCharsToPx(n), title);
    }
    onClose();
  };
  return (
    <SimpleDialog title={title} onClose={onClose} onOk={ok} width={300}>
      <Field label={isRow ? "Satır yüksekliği:" : arg === "colWidth" ? "Sütun genişliği:" : "Standart sütun genişliği:"} inline>
        <input type="text" value={v} onChange={(e) => setV(e.target.value)} data-autofocus="true" />
      </Field>
    </SimpleDialog>
  );
}

export function InsertDeleteCellsDialog({ ctl, onClose, arg }: P & { arg: "insert" | "delete" }) {
  const ins = arg === "insert";
  const [mode, setMode] = useState<"shift" | "shift2" | "row" | "col">("shift");
  const ok = () => {
    const r = ctl.range;
    if (mode === "row") ins ? E.insertRows(ctl) : E.deleteRows(ctl);
    else if (mode === "col") ins ? E.insertCols(ctl) : E.deleteCols(ctl);
    else ctl.ops.shiftCells(ctl.sheet, r, ins ? (mode === "shift" ? "right" : "down") : mode === "shift" ? "left" : "up");
    onClose();
  };
  return (
    <SimpleDialog title={ins ? "Ekle" : "Sil"} onClose={onClose} onOk={ok} width={260}>
      <div className="tb-radio-list">
        <Radio name="sh" value="shift" current={mode} onChange={setMode}>
          {ins ? "Hücreleri sağa kaydır" : "Hücreleri sola kaydır"}
        </Radio>
        <Radio name="sh" value="shift2" current={mode} onChange={setMode}>
          {ins ? "Hücreleri aşağı kaydır" : "Hücreleri yukarı kaydır"}
        </Radio>
        <Radio name="sh" value="row" current={mode} onChange={setMode}>
          Tüm satır
        </Radio>
        <Radio name="sh" value="col" current={mode} onChange={setMode}>
          Tüm sütun
        </Radio>
      </div>
    </SimpleDialog>
  );
}

export function MoveSheetDialog({ ctl, onClose, arg }: P) {
  const from = (arg as { index?: number } | undefined)?.index ?? ctl.wb.activeIndex;
  const [before, setBefore] = useState(Math.min(from + 1, ctl.wb.sheets.length));
  const [copy, setCopy] = useState(false);
  const sheets = ctl.wb.sheets;
  const ok = () => {
    if (copy) ctl.ops.copySheet(from, before);
    else ctl.ops.moveSheet(from, before > from ? before - 1 : before);
    onClose();
  };
  return (
    <SimpleDialog title="Taşı veya Kopyala" onClose={onClose} onOk={ok} width={340}>
      <div style={{ marginBottom: 6 }}>Seçili sayfaları taşı</div>
      <div style={{ marginBottom: 4 }}>Sonraki sayfadan önce:</div>
      <div className="tb-list" style={{ height: 150 }}>
        {[...sheets.map((s, i) => ({ label: s.name, i })), { label: "(sona taşı)", i: sheets.length }].map((it) => (
          <div key={it.i} className={`tb-list-item${before === it.i ? " is-active" : ""}`} onClick={() => setBefore(it.i)} onDoubleClick={ok}>
            {it.label}
          </div>
        ))}
      </div>
      <div style={{ marginTop: 8 }}>
        <Check checked={copy} onChange={setCopy}>
          Kopya oluştur
        </Check>
      </div>
    </SimpleDialog>
  );
}

/** Sayfayı sil: boş sayfa sorulmadan silinir, dolu sayfa için onay istenir (çizim dışında karar verilir). */
export function requestDeleteSheet(ctl: Controller, index = ctl.wb.activeIndex): void {
  const sheet = ctl.wb.sheets[index];
  if (!sheet || ctl.wb.sheets.filter((s) => !s.hidden).length <= 1) return;
  const hasData = ctl.wb.engine.cellCount(sheet.id) > 0 || sheet.styles.size > 0 || sheet.charts.length + sheet.images.length + sheet.shapes.length > 0;
  if (!hasData) ctl.ops.deleteSheet(index);
  else ctl.openDialog("deleteSheet", { index });
}

export function DeleteSheetDialog({ ctl, onClose, arg }: P) {
  const index = (arg as { index?: number } | undefined)?.index ?? ctl.wb.activeIndex;
  const sheet = ctl.wb.sheets[index];
  useEffect(() => {
    if (!sheet) onClose();
  }, [sheet, onClose]);
  if (!sheet) return null;
  return (
    <SimpleDialog title="OfficeHere Tablo" onClose={onClose} onOk={() => (ctl.ops.deleteSheet(index), onClose())} okLabel="Sil" width={420}>
      <p className="oh-dialog-message">Bu sayfa kalıcı olarak silinecek. Devam etmek istiyor musunuz?</p>
    </SimpleDialog>
  );
}

export function UnhideSheetDialog({ ctl, onClose }: P) {
  const hidden = ctl.wb.sheets.map((s, i) => ({ s, i })).filter((x) => x.s.hidden);
  const [sel, setSel] = useState(hidden[0]?.i ?? -1);
  return (
    <SimpleDialog title="Göster" onClose={onClose} onOk={() => (sel >= 0 && ctl.ops.setSheetHidden(sel, false), onClose())} width={300}>
      <div style={{ marginBottom: 4 }}>Sayfa göster:</div>
      <div className="tb-list" style={{ height: 140 }}>
        {hidden.map(({ s, i }) => (
          <div key={s.id} className={`tb-list-item${sel === i ? " is-active" : ""}`} onClick={() => setSel(i)}>
            {s.name}
          </div>
        ))}
      </div>
    </SimpleDialog>
  );
}

export function HyperlinkDialog({ ctl, onClose }: P) {
  const { row, col } = ctl.sel.active;
  const sheet = ctl.sheet;
  const existing = sheet.linkAt(row, col);
  const [kind, setKind] = useState<"web" | "place" | "mail">(existing?.startsWith("#") ? "place" : existing?.startsWith("mailto:") ? "mail" : "web");
  const [text, setText] = useState(displayText(ctl.wb, sheet, row, col));
  const [url, setUrl] = useState(existing && !existing.startsWith("#") && !existing.startsWith("mailto:") ? existing : "");
  const [mail, setMail] = useState(existing?.startsWith("mailto:") ? existing.slice(7) : "");
  const [target, setTarget] = useState(existing?.startsWith("#") ? existing.slice(1).replace(/^.*!/, "") : "A1");
  const [targetSheet, setTargetSheet] = useState(existing?.startsWith("#") && existing.includes("!") ? existing.slice(1).split("!")[0].replace(/^'|'$/g, "") : sheet.name);
  const ok = () => {
    let link = "";
    if (kind === "web") {
      link = url.trim();
      if (link && !/^[a-z]+:/i.test(link)) link = "https://" + link;
    } else if (kind === "mail") link = mail.trim() ? `mailto:${mail.trim()}` : "";
    else link = `#${quoteSheetName(targetSheet)}!${target.trim() || "A1"}`;
    if (!link) {
      onClose();
      return;
    }
    ctl.ops.setLink(sheet, row, col, link, text);
    onClose();
  };
  return (
    <Dialog
      open
      title="Köprü Ekle"
      onClose={onClose}
      width={520}
      onSubmit={ok}
      footer={
        <>
          {existing && <DialogButton onClick={() => (ctl.ops.setLink(sheet, row, col, null), onClose())}>Köprüyü Kaldır</DialogButton>}
          <OkCancel onOk={ok} onCancel={onClose} />
        </>
      }
    >
      <div className="tb-hl">
        <div className="tb-hl-side">
          {(
            [
              ["web", "Varolan Dosya veya Web Sayfası"],
              ["place", "Bu Belgeye Yerleştir"],
              ["mail", "E-posta Adresi"],
            ] as const
          ).map(([k, l]) => (
            <button key={k} type="button" className={`tb-hl-tab${kind === k ? " is-active" : ""}`} onClick={() => setKind(k)}>
              {l}
            </button>
          ))}
        </div>
        <div className="tb-hl-main">
          <Field label="Görüntülenecek metin:">
            <input type="text" value={text} onChange={(e) => setText(e.target.value)} />
          </Field>
          {kind === "web" && (
            <Field label="Adres:">
              <input type="text" value={url} onChange={(e) => setUrl(e.target.value)} placeholder="https://" data-autofocus="true" />
            </Field>
          )}
          {kind === "mail" && (
            <Field label="E-posta adresi:">
              <input type="text" value={mail} onChange={(e) => setMail(e.target.value)} data-autofocus="true" />
            </Field>
          )}
          {kind === "place" && (
            <>
              <Field label="Hücre başvurusunu yazın:">
                <input type="text" value={target} onChange={(e) => setTarget(e.target.value)} data-autofocus="true" />
              </Field>
              <Field label="Ya da bu belgede bir yer seçin:">
                <select value={targetSheet} onChange={(e) => setTargetSheet(e.target.value)} size={5}>
                  {ctl.wb.sheets.map((s) => (
                    <option key={s.id} value={s.name}>
                      {s.name}
                    </option>
                  ))}
                </select>
              </Field>
            </>
          )}
        </div>
      </div>
    </Dialog>
  );
}

export function NoteDialog({ ctl, onClose }: P) {
  const { row, col } = ctl.sel.active;
  const m = ctl.sheet.mergeAt(row, col);
  const r = m ? m.r1 : row;
  const c = m ? m.c1 : col;
  const existing = ctl.sheet.noteAt(r, c);
  const [text, setText] = useState(existing?.text ?? "");
  const author = existing?.author ?? getUserName();
  return (
    <Dialog
      open
      title={`Not — ${addr(r, c)}`}
      onClose={onClose}
      width={340}
      footer={
        <>
          {existing && <DialogButton onClick={() => (ctl.ops.setNote(ctl.sheet, r, c, null), onClose())}>Notu Sil</DialogButton>}
          <OkCancel onOk={() => (ctl.ops.setNote(ctl.sheet, r, c, text.trim() ? text : null, author), onClose())} onCancel={onClose} />
        </>
      }
    >
      <div style={{ fontWeight: 600, marginBottom: 4 }}>{author}:</div>
      <textarea className="tb-note-input" value={text} onChange={(e) => setText(e.target.value)} rows={6} data-autofocus="true" style={{ width: "100%", background: "#FFFFE1", border: "1px solid #c8c8a0", padding: 6 }} />
    </Dialog>
  );
}

const SYMBOL_SETS: { label: string; chars: string }[] = [
  { label: "Para birimi", chars: "₺$€£¥¢₹₽₩₪₫฿" },
  { label: "Matematik", chars: "±×÷≠≈≤≥∞√∑∏∫∂∆∇∈∉∩∪⊂⊃∀∃¬∧∨°′″‰‱µπΩ" },
  { label: "Oklar", chars: "←↑→↓↔↕⇐⇑⇒⇓⇔↖↗↘↙➔➜➤" },
  { label: "Noktalama", chars: "•·…–—‘’“”«»‹›§¶†‡©®™℠№" },
  { label: "Türkçe", chars: "ÇçĞğİıÖöŞşÜüÂâÎîÛû" },
  { label: "Yunan", chars: "ΑΒΓΔΕΖΗΘΙΚΛΜΝΞΟΠΡΣΤΥΦΧΨΩαβγδεζηθικλμνξοπρστυφχψω" },
  { label: "Şekiller", chars: "■□▲△▼▽◆◇○●◎★☆♠♣♥♦✓✔✗✘☐☑☒☀☁☂☎☺☻" },
  { label: "Kesirler", chars: "½⅓⅔¼¾⅕⅖⅗⅘⅙⅚⅛⅜⅝⅞¹²³⁴⁵⁶⁷⁸⁹⁰" },
];

export function SymbolDialog({ ctl, onClose }: P) {
  const [set, setSet] = useState(0);
  const [ch, setCh] = useState("");
  const insert = () => {
    if (!ch) return;
    const { row, col } = ctl.sel.active;
    if (ctl.edit) {
      const e = ctl.edit;
      ctl.updateEdit(e.text.slice(0, e.caret) + ch + e.text.slice(e.caret), e.caret + ch.length);
    } else {
      const cur = ctl.wb.editText(ctl.sheet, row, col);
      ctl.ops.enterText(ctl.sheet, [{ row, col }], cur + ch, "Simge");
    }
  };
  return (
    <Dialog open title="Simge" onClose={onClose} width={470} footer={<><DialogButton primary onClick={() => (insert(), onClose())} disabled={!ch}>Ekle</DialogButton><DialogButton onClick={onClose}>Kapat</DialogButton></>}>
      <Field label="Alt küme:" inline>
        <select value={set} onChange={(e) => setSet(Number(e.target.value))}>
          {SYMBOL_SETS.map((s, i) => (
            <option key={s.label} value={i}>
              {s.label}
            </option>
          ))}
        </select>
      </Field>
      <div className="tb-symbols">
        {[...SYMBOL_SETS[set].chars].map((c) => (
          <button key={c} type="button" className={`tb-symbol${c === ch ? " is-active" : ""}`} onClick={() => setCh(c)} onDoubleClick={() => (setCh(c), insert())}>
            {c}
          </button>
        ))}
      </div>
      <div style={{ marginTop: 6, color: "var(--oh-text-muted)" }}>{ch ? `Unicode: U+${ch.codePointAt(0)!.toString(16).toUpperCase().padStart(4, "0")}` : "Bir simge seçin."}</div>
    </Dialog>
  );
}

export function ZoomDialog({ ctl, onClose }: P) {
  const [mode, setMode] = useState(String(ctl.sheet.zoom));
  const [custom, setCustom] = useState(String(ctl.sheet.zoom));
  const ok = () => {
    let z = mode === "custom" ? Number(custom) : mode === "fit" ? -1 : Number(mode);
    if (z === -1) {
      const r = ctl.range;
      const w = ctl.sheet.cols.offset(r.c2 + 1) - ctl.sheet.cols.offset(r.c1);
      const h = ctl.sheet.rows.offset(r.r2 + 1) - ctl.sheet.rows.offset(r.r1);
      z = Math.floor(Math.min((window.innerWidth - 120) / w, (window.innerHeight - 300) / h) * 100);
    }
    if (!(z >= 10 && z <= 400)) {
      ctl.alert("Yakınlaştırma %10 ile %400 arasında olmalıdır.");
      return;
    }
    ctl.sheet.zoom = Math.round(z);
    ctl.emit();
    onClose();
  };
  return (
    <SimpleDialog title="Yakınlaştır" onClose={onClose} onOk={ok} width={240}>
      <div className="tb-radio-list">
        {["200", "100", "75", "50", "25"].map((z) => (
          <Radio key={z} name="z" value={z} current={mode} onChange={setMode}>
            %{z}
          </Radio>
        ))}
        <Radio name="z" value="fit" current={mode} onChange={setMode}>
          Seçime sığdır
        </Radio>
        <Radio name="z" value="custom" current={mode} onChange={setMode}>
          Özel: <input type="text" value={custom} onChange={(e) => (setCustom(e.target.value), setMode("custom"))} style={{ width: 50 }} /> %
        </Radio>
      </div>
    </SimpleDialog>
  );
}

export function CreateTableDialog({ ctl, onClose, arg }: P) {
  const style = (arg as { style?: string } | undefined)?.style;
  const initial = useMemo(() => {
    const r = ctl.range;
    return r.r1 === r.r2 && r.c1 === r.c2 ? currentRegion(ctl.wb, ctl.sheet, ctl.sel.active) : r;
  }, []); // eslint-disable-line react-hooks/exhaustive-deps
  const [ref, setRef] = useState("=" + rectToA1(initial).replace(/([A-Z]+)(\d+)/g, "$$$1$$$2"));
  const [header, setHeader] = useState(D.guessHeader(ctl.wb, ctl.sheet, initial));
  const ok = () => {
    const r = parseRangeLoose(ref.replace(/^=/, "").replace(/^.*!/, ""));
    if (!r) {
      ctl.alert("Başvuru geçerli değil.");
      return;
    }
    D.createTable(ctl, r, header, style);
    onClose();
  };
  return (
    <SimpleDialog title={style ? "Tablo Olarak Biçimlendir" : "Tablo Oluştur"} onClose={onClose} onOk={ok} width={320}>
      <Field label="Tablonuzun verileri nerede?">
        <input type="text" value={ref} onChange={(e) => setRef(e.target.value)} data-autofocus="true" />
      </Field>
      <Check checked={header} onChange={setHeader}>
        Tablomda üst bilgi var
      </Check>
    </SimpleDialog>
  );
}

export function ProtectDialog({ ctl, onClose }: P) {
  const sheet = ctl.sheet;
  const [pw, setPw] = useState("");
  const [pw2, setPw2] = useState("");
  const [step, setStep] = useState<1 | 2>(1);
  const on = !sheet.protected;
  const ok = () => {
    if (on && pw && step === 1) return setStep(2);
    if (on && pw && pw !== pw2) {
      ctl.alert("Onay parolası aynı değil.");
      setStep(1);
      setPw2("");
      return;
    }
    if (!on && sheet.protectPassword && pw !== sheet.protectPassword) {
      ctl.alert("Girdiğiniz parola doğru değil. CAPS LOCK tuşunun kapalı olduğundan ve doğru büyük/küçük harf kullandığınızdan emin olun.");
      return;
    }
    ctl.ops.sheetOp(sheet, on ? "Sayfayı Koru" : "Korumayı Kaldır", () => {
      sheet.protected = on;
      sheet.protectPassword = on && pw ? pw : undefined;
    });
    onClose();
  };
  if (!on && !sheet.protectPassword) {
    queueMicrotask(() => {
      ctl.ops.sheetOp(sheet, "Korumayı Kaldır", () => (sheet.protected = false));
      onClose();
    });
    return null;
  }
  return (
    <SimpleDialog title={on ? (step === 2 ? "Parolayı Onayla" : "Sayfayı Koru") : "Sayfanın Korumasını Kaldır"} onClose={onClose} onOk={ok} width={360}>
      {step === 1 ? (
        <Field label={on ? "Sayfa korumasını kaldırmak için parola (isteğe bağlı):" : "Parola:"}>
          <input type="password" value={pw} onChange={(e) => setPw(e.target.value)} data-autofocus="true" />
        </Field>
      ) : (
        <Field label="Parolayı yeniden girin:">
          <input type="password" value={pw2} onChange={(e) => setPw2(e.target.value)} data-autofocus="true" />
        </Field>
      )}
      {on && <p className="oh-bs-note">Korumalı sayfada kilitli hücreler düzenlenemez. Hücrelerin kilidini Hücreleri Biçimlendir &gt; Koruma sekmesinden açabilirsiniz. Parola yalnız bu oturumda geçerlidir ve dosyaya yazılmaz.</p>}
    </SimpleDialog>
  );
}

export function PasteSpecialDialog({ ctl, onClose }: P) {
  const [what, setWhat] = useState<PasteOptions["what"]>("all");
  const [op, setOp] = useState<NonNullable<PasteOptions["op"]>>("none");
  const [skip, setSkip] = useState(false);
  const [transpose, setTranspose] = useState(false);
  const ok = () => {
    onClose();
    E.pasteData(ctl, "", undefined, { what, op, skipBlanks: skip, transpose });
  };
  return (
    <SimpleDialog title="Özel Yapıştır" onClose={onClose} onOk={ok} width={440}>
      <fieldset className="oh-fieldset">
        <legend>Yapıştır</legend>
        <div className="tb-dialog-grid">
          {(
            [
              ["all", "Tümü"],
              ["allExceptBorders", "Kenarlıklar dışında tümü"],
              ["formulas", "Formüller"],
              ["widths", "Sütun genişlikleri"],
              ["values", "Değerler"],
              ["formulasAndFormats", "Formüller ve sayı biçimleri"],
              ["formats", "Biçimler"],
              ["valuesAndFormats", "Değerler ve sayı biçimleri"],
            ] as const
          ).map(([k, l]) => (
            <Radio key={k} name="pw" value={k} current={what} onChange={setWhat}>
              {l}
            </Radio>
          ))}
        </div>
      </fieldset>
      <fieldset className="oh-fieldset">
        <legend>İşlem</legend>
        <div className="tb-dialog-grid">
          {(
            [
              ["none", "Yok"],
              ["mul", "Çarp"],
              ["add", "Ekle"],
              ["div", "Böl"],
              ["sub", "Çıkar"],
            ] as const
          ).map(([k, l]) => (
            <Radio key={k} name="po" value={k} current={op} onChange={setOp}>
              {l}
            </Radio>
          ))}
        </div>
      </fieldset>
      <div style={{ display: "flex", gap: 24 }}>
        <Check checked={skip} onChange={setSkip}>
          Boşlukları atla
        </Check>
        <Check checked={transpose} onChange={setTranspose}>
          İşlem sırasını değiştir
        </Check>
      </div>
    </SimpleDialog>
  );
}

/** Veri doğrulama listesi ya da sütundaki değerlerden seçim. */
export function PickListDialog({ ctl, onClose, arg }: P) {
  const { row, col } = ctl.sel.active;
  const items = useMemo(() => {
    const a = arg as { items?: string[] } | undefined;
    if (a?.items) return a.items;
    const v = ctl.sheet.validations.find((x) => x.rule.type === "list" && x.ranges.some((r) => row >= r.r1 && row <= r.r2 && col >= r.c1 && col <= r.c2));
    if (v) return validationList(ctl.wb, ctl.sheet, v.rule.formula1 ?? "", row, col);
    const region = currentRegion(ctl.wb, ctl.sheet, { row, col });
    const set = new Set<string>();
    for (let r = region.r1; r <= region.r2; r++) {
      if (r === row) continue;
      const v2 = ctl.wb.getValue(ctl.sheet, r, col);
      if (typeof v2 === "string" && v2) set.add(v2);
    }
    return [...set].sort((a2, b) => a2.localeCompare(b, "tr"));
  }, []); // eslint-disable-line react-hooks/exhaustive-deps
  const [sel, setSel] = useState(0);
  const pick = (s: string) => {
    ctl.ops.enterText(ctl.sheet, [{ row, col }], s, "Listeden Seç");
    onClose();
  };
  return (
    <Dialog open title="Listeden Seç" onClose={onClose} width={260} onSubmit={() => items[sel] && pick(items[sel])} footer={<OkCancel onOk={() => items[sel] && pick(items[sel])} onCancel={onClose} />}>
      <div className="tb-list" style={{ height: 180 }} tabIndex={0} data-autofocus="true" onKeyDown={(e) => {
        if (e.key === "ArrowDown") setSel((s) => Math.min(items.length - 1, s + 1));
        if (e.key === "ArrowUp") setSel((s) => Math.max(0, s - 1));
      }}>
        {items.length === 0 && <div className="tb-list-item">(Liste boş)</div>}
        {items.map((s, i) => (
          <div key={s} className={`tb-list-item${i === sel ? " is-active" : ""}`} onClick={() => setSel(i)} onDoubleClick={() => pick(s)}>
            {s}
          </div>
        ))}
      </div>
    </Dialog>
  );
}

export function ShapeTextDialog({ ctl, onClose, arg }: P) {
  const id = (arg as { id: string }).id;
  const shape = ctl.sheet.shapes.find((s) => s.id === id);
  // Şekil seçiliyken yazılan karakter metnin sonuna eklenir
  const [text, setText] = useState((shape?.text ?? "") + ((arg as { text?: string }).text ?? ""));
  const [fill, setFill] = useState(shape?.fill ?? "#156082");
  return (
    <SimpleDialog
      title="Şekli Biçimlendir"
      onClose={onClose}
      onOk={() => {
        ctl.ops.sheetOp(ctl.sheet, "Şekil", () => {
          const s = ctl.sheet.shapes.find((x) => x.id === id);
          if (s) {
            s.text = text;
            s.fill = fill;
          }
        });
        onClose();
      }}
    >
      <Field label="Metin:">
        <textarea value={text} onChange={(e) => setText(e.target.value)} rows={3} data-autofocus="true" />
      </Field>
      <Field label="Dolgu rengi:" inline>
        <input type="color" value={/^#[0-9a-f]{6}$/i.test(fill) ? fill : "#ffffff"} onChange={(e) => setFill(e.target.value)} />
      </Field>
    </SimpleDialog>
  );
}

export function GoToDialog({ ctl, onClose }: P) {
  const names = ctl.wb.engine.getNames();
  const [ref, setRef] = useState("");
  const ok = () => {
    if (!ctl.goTo(ref)) {
      ctl.alert("Başvuru geçerli değil.");
      return;
    }
    onClose();
  };
  return (
    <SimpleDialog title="Git" onClose={onClose} onOk={ok} width={320}>
      <button type="button" className="oh-btn tb-goto-special" onClick={() => (onClose(), ctl.openDialog("goSpecial"))}>
        Özel...
      </button>
      <div style={{ marginBottom: 4 }}>Git:</div>
      <div className="tb-list" style={{ height: 130 }}>
        {names.map((n) => (
          <div key={n.name} className={`tb-list-item${ref === n.name ? " is-active" : ""}`} onClick={() => setRef(n.name)} onDoubleClick={() => (ctl.goTo(n.name), onClose())}>
            {n.name}
          </div>
        ))}
      </div>
      <Field label="Başvuru:">
        <input type="text" value={ref} onChange={(e) => setRef(e.target.value)} data-autofocus="true" />
      </Field>
    </SimpleDialog>
  );
}

export function rangeText(ctl: Controller): string {
  return `${quoteSheetName(ctl.sheet.name)}!${rectToA1(ctl.range).replace(/([A-Z]+)(\d+)/g, "$$$1$$$2")}`;
}

export { rect };
