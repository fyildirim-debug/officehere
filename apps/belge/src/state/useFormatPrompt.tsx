// "Bu biçimde bazı özellikler kaybolabilir" sorusu: yerel olmayan bir biçimde ilk kayıtta bir kez sorulur.
import { useCallback, useRef, useState, type ReactNode } from "react";
import { Dialog, DialogButton } from "@officehere/ui";

export type FormatChoice = "keep" | "docx" | "cancel";

export function useFormatPrompt(): { ask: (label: string, ext: string, fileName: string) => Promise<FormatChoice>; dialog: ReactNode } {
  const [state, setState] = useState<{ label: string; ext: string; fileName: string } | null>(null);
  const resolver = useRef<((c: FormatChoice) => void) | null>(null);

  const ask = useCallback(
    (label: string, ext: string, fileName: string) =>
      new Promise<FormatChoice>((resolve) => {
        resolver.current?.("cancel");
        resolver.current = resolve;
        setState({ label, ext, fileName });
      }),
    [],
  );

  const choose = (c: FormatChoice) => {
    const r = resolver.current;
    resolver.current = null;
    setState(null);
    r?.(c);
  };

  const dialog = (
    <Dialog
      open={!!state}
      title="Biçimi koru"
      onClose={() => choose("cancel")}
      width={480}
      footer={
        <>
          <DialogButton primary onClick={() => choose("keep")}>
            {state ? `.${state.ext} olarak kaydet` : ""}
          </DialogButton>
          <DialogButton onClick={() => choose("docx")}>.docx olarak kaydet</DialogButton>
          <DialogButton onClick={() => choose("cancel")}>İptal</DialogButton>
        </>
      }
    >
      {state && (
        <p className="oh-dialog-message">
          <strong>"{state.fileName}"</strong> {state.label} biçiminde. Bu biçimde bazı özellikler (stiller, açıklamalar, değişiklik izleme, bazı yerleşim
          ayarları) kaybolabilir. Belgeyi bu biçimde mi kaydetmek istiyorsunuz, yoksa tüm özellikleri koruyan .docx biçiminde mi?
        </p>
      )}
    </Dialog>
  );

  return { ask, dialog };
}
