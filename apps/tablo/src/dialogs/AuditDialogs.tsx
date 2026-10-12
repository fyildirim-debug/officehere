/** Formül Denetimi iletişim kutuları: Formülü Değerlendir, Hata Denetimi, Hata Denetimi Seçenekleri. */
import { useMemo, useRef, useState } from "react";
import { Dialog, DialogButton } from "@officehere/ui";
import { addr, FormulaStepper, toLocalFormula } from "@officehere/formula";
import type { Controller } from "../controller";
import { cellError, ignoreError, sheetErrors } from "../actions/audit";
import type { CellPos } from "../model/types";
import { Check } from "./common";

type P = { ctl: Controller; onClose: () => void; arg?: unknown };

function Marked({ text, mark, italic }: { text: string; mark: [number, number] | null; italic: boolean }) {
  if (!mark) return <span style={italic ? { fontStyle: "italic" } : undefined}>{text}</span>;
  return (
    <>
      {text.slice(0, mark[0])}
      <u className="tb-eval-mark">{text.slice(mark[0], mark[1])}</u>
      {text.slice(mark[1])}
    </>
  );
}

export function EvaluateFormulaDialog({ ctl, onClose }: P) {
  const { wb, sheet } = ctl;
  const pos = ctl.sel.active;
  const formula = wb.getFormula(sheet, pos.row, pos.col);
  const make = () => (formula ? [new FormulaStepper(wb.engine, sheet.id, pos.row, pos.col, formula)] : []);
  const stack = useRef<FormulaStepper[]>(make());
  const [, force] = useState(0);
  const refresh = () => {
    force((n) => n + 1);
    // Devre dışı kalan düğmeden düşen odak iletişim kutusunda kalsın (Esc/Enter çalışsın)
    requestAnimationFrame(() => document.querySelector<HTMLButtonElement>('.oh-dialog[aria-label="Formülü Değerlendir"] .oh-btn.is-primary')?.focus());
  };
  const cur = stack.current[stack.current.length - 1];
  if (!cur) {
    return (
      <Dialog open title="Formülü Değerlendir" onClose={onClose} width={480} footer={<DialogButton onClick={onClose}>Kapat</DialogButton>}>
        <p className="oh-dialog-message">Etkin hücre bir formül içermiyor. Değerlendirmek için formül içeren bir hücre seçin.</p>
      </Dialog>
    );
  }
  const root = stack.current[0];
  const evaluate = () => {
    if (root.done && stack.current.length === 1) {
      stack.current = make();
    } else cur.step();
    refresh();
  };
  const stepIn = () => {
    const t = cur.stepInTarget();
    if (!t) return;
    stack.current.push(new FormulaStepper(wb.engine, t.sheet, t.row, t.col, t.formula));
    refresh();
  };
  const stepOut = () => {
    if (stack.current.length < 2) return;
    stack.current.pop();
    stack.current[stack.current.length - 1].completeStep();
    refresh();
  };
  return (
    <Dialog
      open
      title="Formülü Değerlendir"
      onClose={onClose}
      width={560}
      footer={
        <>
          <DialogButton primary onClick={evaluate}>
            {root.done && stack.current.length === 1 ? "Yeniden Başlat" : "Değerlendir"}
          </DialogButton>
          <DialogButton onClick={stepIn} disabled={!cur.view().canStepIn || cur.done}>
            İçeri Adımla
          </DialogButton>
          <DialogButton onClick={stepOut} disabled={stack.current.length < 2}>
            Dışarı Adımla
          </DialogButton>
          <DialogButton onClick={onClose}>Kapat</DialogButton>
        </>
      }
    >
      <div className="tb-eval">
        <div className="tb-eval-head">
          <span>Başvuru:</span>
          <span>Değerlendirme:</span>
        </div>
        {stack.current.map((st, i) => {
          const v = st.view();
          return (
            <div key={i} className={`tb-eval-row${i === stack.current.length - 1 ? " is-current" : ""}`}>
              <span className="tb-eval-ref">{st.address()}</span>
              <span className="tb-eval-text" aria-label="Değerlendirme">
                <Marked text={v.text} mark={i === stack.current.length - 1 ? v.mark : null} italic={v.done} />
              </span>
            </div>
          );
        })}
        <p className="tb-eval-help">
          {cur.done ? "Formül sonucu italik olarak gösteriliyor. Baştan başlamak için Yeniden Başlat'ı tıklatın." : "Altı çizili ifadeyi değerlendirmek için Değerlendir'i tıklatın. En son sonuç italik olarak görünür."}
        </p>
      </div>
    </Dialog>
  );
}

export function ErrorCheckDialog({ ctl, onClose }: P) {
  const { wb, sheet } = ctl;
  const list = useMemo(() => sheetErrors(wb, sheet), [wb, sheet, ctl.version]); // eslint-disable-line react-hooks/exhaustive-deps
  const startIdx = Math.max(
    0,
    list.findIndex((c) => c.row > ctl.sel.active.row || (c.row === ctl.sel.active.row && c.col >= ctl.sel.active.col)),
  );
  const [idx, setIdx] = useState(startIdx);
  if (!list.length) {
    return (
      <Dialog open title="Hata Denetimi" onClose={onClose} width={420} onSubmit={onClose} footer={<DialogButton primary onClick={onClose}>Tamam</DialogButton>}>
        <p className="oh-dialog-message">Sayfanın tamamı için hata denetimi tamamlandı.</p>
      </Dialog>
    );
  }
  const i = Math.min(idx, list.length - 1);
  const c: CellPos = list[i];
  const err = cellError(wb, sheet, c.row, c.col);
  const go = (n: number) => {
    const j = (n + list.length) % list.length;
    setIdx(j);
    ctl.selectCell(list[j].row, list[j].col);
  };
  const f = wb.getFormula(sheet, c.row, c.col);
  return (
    <Dialog
      open
      title="Hata Denetimi"
      onClose={onClose}
      width={520}
      footer={
        <>
          <DialogButton onClick={() => go(i - 1)}>Önceki</DialogButton>
          <DialogButton primary onClick={() => go(i + 1)}>
            Sonraki
          </DialogButton>
        </>
      }
    >
      <div className="tb-errcheck">
        <div className="tb-errcheck-main">
          <p>
            <strong>{addr(c.row, c.col)} hücresinde hata</strong>
          </p>
          <p className="tb-errcheck-formula">{f ? toLocalFormula(f, "tr") : wb.editText(sheet, c.row, c.col)}</p>
          <p>
            <strong>{err?.title}</strong>
          </p>
          <p>{err?.help}</p>
        </div>
        <div className="tb-errcheck-actions">
          <DialogButton
            onClick={() => {
              ctl.selectCell(c.row, c.col);
              ctl.openDialog("evaluateFormula");
            }}
            disabled={!f}
          >
            Hesaplama Adımlarını Göster...
          </DialogButton>
          <DialogButton
            onClick={() => {
              ignoreError(ctl, c.row, c.col);
            }}
          >
            Hatayı Yoksay
          </DialogButton>
          <DialogButton
            onClick={() => {
              ctl.selectCell(c.row, c.col);
              onClose();
              ctl.startEdit(undefined, true);
            }}
          >
            Formül Çubuğunda Düzenle
          </DialogButton>
        </div>
      </div>
    </Dialog>
  );
}

export function ErrorOptionsDialog({ ctl, onClose }: P) {
  const [on, setOn] = useState(ctl.errorChecking);
  const ignored = ctl.wb.sheets.reduce((n, s) => n + s.ignoredErrors.size, 0);
  return (
    <Dialog
      open
      title="Hata Denetimi Seçenekleri"
      onClose={onClose}
      width={420}
      onSubmit={() => {
        ctl.errorChecking = on;
        ctl.emit();
        onClose();
      }}
      footer={
        <>
          <DialogButton
            primary
            onClick={() => {
              ctl.errorChecking = on;
              ctl.emit();
              onClose();
            }}
          >
            Tamam
          </DialogButton>
          <DialogButton onClick={onClose}>İptal</DialogButton>
        </>
      }
    >
      <Check checked={on} onChange={setOn}>
        Arka planda hata denetimini etkinleştir
      </Check>
      <p className="tb-dlg-note">
        Hata göstergesi rengi: yeşil. Denetlenen kurallar: hata veren formüller, metin olarak depolanmış sayılar, bölgedeki diğer formüllerle tutarsız formüller.
      </p>
      <DialogButton
        disabled={!ignored}
        onClick={() => {
          for (const s of ctl.wb.sheets) s.ignoredErrors.clear();
          ctl.emit();
        }}
      >
        Yoksayılan Hataları Sıfırla
      </DialogButton>
    </Dialog>
  );
}
