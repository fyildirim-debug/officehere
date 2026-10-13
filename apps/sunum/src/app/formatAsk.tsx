/** İlk kayıtta biçim sorusu: "Bu biçimde bazı özellikler kaybolabilir". */
import { useEffect, useState } from "react";
import { Dialog, DialogButton } from "@officehere/ui";

export type FormatChoice = "keep" | "native" | "cancel";

let asker: ((ext: string) => Promise<FormatChoice>) | null = null;

const LABELS: Record<string, string> = { odp: "OpenDocument Sunu" };

/** Soruyu gösterir; uygulama kökü bağlı değilse (testler) biçim korunur. */
export function askKeepFormat(ext: string): Promise<FormatChoice> {
  return asker ? asker(ext) : Promise.resolve("keep");
}

export function FormatAskHost() {
  const [req, setReq] = useState<{ ext: string; resolve: (c: FormatChoice) => void } | null>(null);
  useEffect(() => {
    asker = (ext) => new Promise((resolve) => setReq({ ext, resolve }));
    return () => {
      asker = null;
    };
  }, []);
  if (!req) return null;
  const done = (c: FormatChoice) => {
    req.resolve(c);
    setReq(null);
  };
  return (
    <Dialog
      open
      title="Biçimi Koru"
      onClose={() => done("cancel")}
      width={470}
      footer={
        <>
          <DialogButton primary onClick={() => done("keep")}>
            .{req.ext} olarak kaydet
          </DialogButton>
          <DialogButton onClick={() => done("native")}>.pptx olarak kaydet</DialogButton>
          <DialogButton onClick={() => done("cancel")}>İptal</DialogButton>
        </>
      }
    >
      <p className="oh-dialog-message">
        Bu sunu <strong>{LABELS[req.ext] ?? req.ext} (.{req.ext})</strong> biçiminde. Bu biçimde kaydederseniz bazı özellikler (animasyonlar, geçişler, grafikler ve bazı efektler) kaybolabilir. Tüm özellikleri korumak için .pptx biçimini seçin.
      </p>
    </Dialog>
  );
}
