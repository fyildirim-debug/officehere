/** Bölümü / Düzeni Yeniden Adlandır iletişim kutusu. */
import { useState } from "react";
import { Dialog, DialogButton, Field } from "@officehere/ui";

export interface RenameRequest {
  title: string;
  label: string;
  value: string;
  onOk: (name: string) => void;
}

export function RenameDialog({ request, onClose }: { request: RenameRequest | null; onClose: () => void }) {
  const [value, setValue] = useState(request?.value ?? "");
  const [last, setLast] = useState<RenameRequest | null>(null);
  if (request !== last) {
    setLast(request);
    if (request) setValue(request.value);
  }
  if (!request) return null;
  const ok = () => {
    if (value.trim()) request.onOk(value.trim());
    onClose();
  };
  return (
    <Dialog
      open
      title={request.title}
      onClose={onClose}
      onSubmit={ok}
      width={340}
      footer={
        <>
          <DialogButton primary onClick={ok}>
            Yeniden Adlandır
          </DialogButton>
          <DialogButton onClick={onClose}>İptal</DialogButton>
        </>
      }
    >
      <Field label={request.label}>
        <input type="text" aria-label={request.label.replace(/:$/, "")} value={value} onChange={(e) => setValue(e.target.value)} onFocus={(e) => e.currentTarget.select()} style={{ width: "100%", height: 26 }} />
      </Field>
    </Dialog>
  );
}
