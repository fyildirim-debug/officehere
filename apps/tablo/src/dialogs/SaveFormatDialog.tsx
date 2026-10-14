/** Yerel olmayan biçimde kaydederken bir kez sorulan soru: biçimi koru ya da .xlsx. */
import { useEffect, useRef } from "react";
import { Dialog, DialogButton } from "@officehere/ui";
import type { Controller } from "../controller";
import type { SaveFormat } from "../io/formats";

export type FormatChoice = "same" | "xlsx" | null;

const NAMES: Record<SaveFormat, string> = { xlsx: ".xlsx", xlsm: ".xlsm", ods: "ODS", csv: "CSV", tsv: "TSV" };

function message(format: SaveFormat, sheet: string): string {
  if (format === "ods") return "Bu çalışma kitabı OpenDocument Hesap Tablosu (.ods) biçiminde. Bu biçimde bazı özellikler (grafikler, özet tablolar, koşullu biçimlendirme, veri doğrulama gibi) kaybolabilir.\n\nTüm özellikleri korumak için .xlsx olarak kaydedin.";
  const kind = format === "csv" ? "CSV" : "sekmeyle ayrılmış metin (.tsv)";
  return `Bu dosya ${kind} biçiminde. Bu biçimde bazı özellikler kaybolabilir: yalnız etkin sayfa («${sheet}») ve değerler kaydedilir; biçimler, formüller ve diğer sayfalar kaydedilmez.\n\nTüm özellikleri korumak için .xlsx olarak kaydedin.`;
}

export function SaveFormatDialog({ ctl, onClose, arg }: { ctl: Controller; onClose: () => void; arg?: unknown }) {
  const a = arg as { format: SaveFormat; onChoose: (c: FormatChoice) => void };
  const answered = useRef(false);
  const choose = (c: FormatChoice) => {
    if (!answered.current) {
      answered.current = true;
      a.onChoose(c);
    }
    onClose();
  };
  // Yanıtsız kapanırsa (Esc) İptal sayılır; geliştirme kipindeki çift bağlama yanıltmasın diye ertelenir
  const alive = useRef(false);
  useEffect(() => {
    alive.current = true;
    return () => {
      alive.current = false;
      setTimeout(() => {
        if (!alive.current && !answered.current) {
          answered.current = true;
          a.onChoose(null);
        }
      }, 0);
    };
  }, []); // eslint-disable-line react-hooks/exhaustive-deps
  return (
    <Dialog
      open
      title="OfficeHere Tablo"
      width={460}
      onClose={() => choose(null)}
      onSubmit={() => choose("same")}
      footer={
        <>
          <DialogButton primary onClick={() => choose("same")}>
            {NAMES[a.format]} olarak kaydet
          </DialogButton>
          <DialogButton onClick={() => choose("xlsx")}>.xlsx olarak kaydet</DialogButton>
          <DialogButton onClick={() => choose(null)}>İptal</DialogButton>
        </>
      }
    >
      <p className="oh-dialog-message" style={{ whiteSpace: "pre-wrap" }}>
        {message(a.format, ctl.sheet.name)}
      </p>
    </Dialog>
  );
}
