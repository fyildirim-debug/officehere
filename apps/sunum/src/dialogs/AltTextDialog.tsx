/** Alternatif Metin: nesnenin ekran okuyucularda okunan açıklaması (descr). */
import { useState } from "react";
import { Dialog, DialogButton, Field } from "@officehere/ui";

export interface AltTextRequest {
  value: string;
  onOk: (text: string) => void;
}

export function AltTextDialog({ request, onClose }: { request: AltTextRequest | null; onClose: () => void }) {
  const [value, setValue] = useState(request?.value ?? "");
  const [last, setLast] = useState<AltTextRequest | null>(null);
  if (request !== last) {
    setLast(request);
    if (request) setValue(request.value);
  }
  if (!request) return null;
  const ok = () => {
    request.onOk(value.trim());
    onClose();
  };
  return (
    <Dialog
      open
      title="Alternatif Metin"
      onClose={onClose}
      width={380}
      footer={
        <>
          <DialogButton primary onClick={ok}>
            Tamam
          </DialogButton>
          <DialogButton onClick={onClose}>İptal</DialogButton>
        </>
      }
    >
      <p style={{ margin: "0 0 8px", color: "var(--oh-text-muted)" }}>Görme engelli kişiler için nesneyi bir iki cümleyle anlatın.</p>
      <Field label="Açıklama:">
        <textarea aria-label="Alternatif metin" rows={4} value={value} onChange={(e) => setValue(e.target.value)} style={{ width: "100%", resize: "vertical" }} />
      </Field>
    </Dialog>
  );
}
