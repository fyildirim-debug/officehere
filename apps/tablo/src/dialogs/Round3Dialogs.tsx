/** 3. tur iletişim kutuları: veri doğrulama hata kutusu, Özel Git. */
import { useState } from "react";
import { Dialog, DialogButton } from "@officehere/ui";
import type { Controller } from "../controller";
import { goSpecial, type GoSpecialKind } from "../model/goSpecial";
import { Check, Radio, SimpleDialog } from "./common";

type P = { ctl: Controller; onClose: () => void; arg?: unknown };

export interface ValidationErrorArg {
  title: string;
  message: string;
  style: "stop" | "warning" | "information";
  /** Yeniden Dene / Hayır: düzenlemeye dön. */
  onRetry: () => void;
  /** İptal: girişi geri al. */
  onCancel: () => void;
  /** Evet / Tamam: değeri kabul et. */
  onAccept: () => void;
}

function Glyph({ style }: { style: ValidationErrorArg["style"] }) {
  const color = style === "stop" ? "#C42B1C" : style === "warning" ? "#C27C0E" : "#0F6CBD";
  return (
    <svg width="32" height="32" viewBox="0 0 32 32" aria-hidden="true" style={{ flex: "none" }}>
      {style === "warning" ? <path d="M16 3 30 28H2Z" fill={color} /> : <circle cx="16" cy="16" r="14" fill={color} />}
      {style === "stop" && <path d="M10.5 10.5l11 11M21.5 10.5l-11 11" stroke="#fff" strokeWidth="3" strokeLinecap="round" />}
      {style === "warning" && <path d="M16 11v9M16 23.5v.5" stroke="#fff" strokeWidth="3" strokeLinecap="round" />}
      {style === "information" && <path d="M16 14v9M16 9.5v.5" stroke="#fff" strokeWidth="3" strokeLinecap="round" />}
    </svg>
  );
}

export function ValidationErrorDialog({ onClose, arg }: P) {
  const a = arg as ValidationErrorArg;
  const run = (fn: () => void) => () => {
    onClose();
    fn();
  };
  const footer =
    a.style === "stop" ? (
      <>
        <DialogButton primary onClick={run(a.onRetry)}>
          Yeniden Dene
        </DialogButton>
        <DialogButton onClick={run(a.onCancel)}>İptal</DialogButton>
      </>
    ) : a.style === "warning" ? (
      <>
        <DialogButton onClick={run(a.onAccept)}>Evet</DialogButton>
        <DialogButton primary onClick={run(a.onRetry)}>
          Hayır
        </DialogButton>
        <DialogButton onClick={run(a.onCancel)}>İptal</DialogButton>
      </>
    ) : (
      <>
        <DialogButton primary onClick={run(a.onAccept)}>
          Tamam
        </DialogButton>
        <DialogButton onClick={run(a.onCancel)}>İptal</DialogButton>
      </>
    );
  return (
    <Dialog open title={a.title} onClose={run(a.onCancel)} width={440} footer={footer}>
      <div style={{ display: "flex", gap: 14, alignItems: "flex-start" }}>
        <Glyph style={a.style} />
        <p className="oh-dialog-message" style={{ whiteSpace: "pre-wrap", margin: 0 }}>
          {a.message}
          {a.style === "warning" ? "\n\nDevam etmek istiyor musunuz?" : ""}
        </p>
      </div>
    </Dialog>
  );
}

const SPECIAL_LEFT: { k: GoSpecialKind; label: string }[] = [
  { k: "notes", label: "Notlar" },
  { k: "constants", label: "Sabitler" },
  { k: "formulas", label: "Formüller" },
  { k: "blanks", label: "Boşluklar" },
  { k: "currentRegion", label: "Geçerli bölge" },
  { k: "currentArray", label: "Geçerli dizi" },
  { k: "objects", label: "Nesneler" },
];
const SPECIAL_RIGHT: { k: GoSpecialKind; label: string }[] = [
  { k: "rowDiffs", label: "Satır farkları" },
  { k: "colDiffs", label: "Sütun farkları" },
  { k: "precedents", label: "Etkileyenler" },
  { k: "dependents", label: "Etkilenenler" },
  { k: "lastCell", label: "Son hücre" },
  { k: "visible", label: "Yalnızca görünür hücreler" },
  { k: "condFormats", label: "Koşullu biçimler" },
  { k: "validation", label: "Veri doğrulama" },
];

/** Git > Özel... iletişim kutusu. */
export function GoToSpecialDialog({ ctl, onClose }: P) {
  const [kind, setKind] = useState<GoSpecialKind>("blanks");
  const [types, setTypes] = useState({ numbers: true, text: true, logicals: true, errors: true });
  const [allLevels, setAllLevels] = useState(false);
  const [same, setSame] = useState(false);
  const ok = () => {
    const r = goSpecial(ctl.wb, ctl.sheet, ctl.sel.ranges, ctl.sel.active, kind, { ...types, allLevels, same });
    if (!r) {
      ctl.alert("Hiçbir hücre bulunamadı.");
      return;
    }
    onClose();
    if (r.object) {
      ctl.selObj = r.object;
      ctl.emit();
      return;
    }
    ctl.corner = { row: r.ranges[r.ranges.length - 1].r2, col: r.ranges[r.ranges.length - 1].c2 };
    ctl.setSelection({ ranges: r.ranges, active: r.active, anchor: r.active }, false);
    ctl.scrollIntoView(r.active.row, r.active.col);
  };
  const typed = kind === "constants" || kind === "formulas";
  const levels = kind === "precedents" || kind === "dependents";
  const sameOpt = kind === "condFormats" || kind === "validation";
  const radio = (x: { k: GoSpecialKind; label: string }) => (
    <Radio key={x.k} name="gospecial" value={x.k} current={kind} onChange={setKind}>
      {x.label}
    </Radio>
  );
  return (
    <SimpleDialog title="Özel Git" onClose={onClose} onOk={ok} width={440}>
      <div style={{ marginBottom: 6 }}>Seç</div>
      <div className="tb-gospecial">
        <div className="tb-gospecial-col">
          {SPECIAL_LEFT.slice(0, 3).map(radio)}
          <div className={`tb-gospecial-sub${typed ? "" : " is-disabled"}`}>
            {(["numbers", "text", "logicals", "errors"] as const).map((t) => (
              <Check key={t} checked={types[t]} onChange={(v) => typed && setTypes({ ...types, [t]: v })}>
                {{ numbers: "Sayılar", text: "Metin", logicals: "Mantıksal", errors: "Hatalar" }[t]}
              </Check>
            ))}
          </div>
          {SPECIAL_LEFT.slice(3).map(radio)}
        </div>
        <div className="tb-gospecial-col">
          {SPECIAL_RIGHT.slice(0, 4).map(radio)}
          <div className={`tb-gospecial-sub${levels ? "" : " is-disabled"}`}>
            <Radio name="levels" value="direct" current={allLevels ? "all" : "direct"} onChange={() => setAllLevels(false)}>
              Yalnızca doğrudan
            </Radio>
            <Radio name="levels" value="all" current={allLevels ? "all" : "direct"} onChange={() => setAllLevels(true)}>
              Tüm düzeyler
            </Radio>
          </div>
          {SPECIAL_RIGHT.slice(4).map(radio)}
          <div className={`tb-gospecial-sub${sameOpt ? "" : " is-disabled"}`}>
            <Radio name="same" value="all" current={same ? "same" : "all"} onChange={() => setSame(false)}>
              Tümü
            </Radio>
            <Radio name="same" value="same" current={same ? "same" : "all"} onChange={() => setSame(true)}>
              Aynı
            </Radio>
          </div>
        </div>
      </div>
    </SimpleDialog>
  );
}

/**
 * Yazım Denetimi (F7): sayfadaki metin hücrelerini sırayla gösterir. Sözcükler
 * sistemin Türkçe yazım denetleyicisiyle işaretlenir (altı kırmızı çizili);
 * düzeltilen metin "Değiştir" ile hücreye yazılır.
 */
export function SpellingDialog({ ctl, onClose }: P) {
  // Tarayıcının işaretlediği sözcükler betikle okunamadığından kendi (uydurma) denetimimiz yoktur:
  // denetim, düzenleyicide tarayıcının/işletim sisteminin Türkçe sözlüğüyle yapılır.
  const [on, setOn] = useState(ctl.spellcheck);
  return (
    <SimpleDialog
      title="Yazım Denetimi: Türkçe"
      onClose={onClose}
      onOk={() => {
        ctl.spellcheck = on;
        ctl.status = on ? "Yazım denetimi açık: hücre düzenlenirken hatalı sözcüklerin altı çizilir" : "";
        ctl.emit();
        onClose();
      }}
      width={460}
    >
      <p className="oh-dialog-message">
        Yazım denetimi tarayıcının Türkçe sözlüğüyle yapılır: bir hücreyi düzenlerken (F2 ya da çift tıklama) hatalı sözcüklerin altı kırmızı dalgalı çizgiyle çizilir. Öneriler için altı çizili sözcüğe sağ tıklayın.
      </p>
      <Check checked={on} onChange={setOn}>
        Yazarken yazım hatalarını işaretle
      </Check>
    </SimpleDialog>
  );
}
