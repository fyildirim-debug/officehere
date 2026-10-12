// Düzenlemeyi Kısıtla: izin verilen düzenleme türü, isteğe bağlı parola; korumayı başlat / durdur.
import { useState } from "react";
import { Dialog, DialogButton, Field, FieldSet, MessageBar } from "@officehere/ui";
import { useBelge } from "../state/context";
import type { DocProtection } from "../model/document";
import { PROTECTION_LABEL, checkPassword, makeProtection } from "../editor/protection";
import { Check } from "./FontDialog";

/** Koruma zorlanırken şeridin altındaki bilgi çubuğu. */
export function ProtectionBar({ mode, onStop }: { mode: DocProtection["mode"]; onStop: () => void }) {
  return (
    <MessageBar
      kind="info"
      actions={
        <button type="button" className="oh-btn" onClick={onStop}>
          Korumayı Durdur
        </button>
      }
    >
      <strong>Kısıtlanmış Düzenleme</strong> — bu belge korunuyor; izin verilen düzenleme: {PROTECTION_LABEL[mode]}.
    </MessageBar>
  );
}

export function ProtectDialog({ onClose }: { onClose: () => void }) {
  const { settings, updateSettings, notify } = useBelge();
  const cur = settings.protection ?? null;
  const enforced = !!cur?.enforced;
  const [limit, setLimit] = useState(true);
  const [mode, setMode] = useState<DocProtection["mode"]>(cur?.mode ?? "readOnly");
  const [pw, setPw] = useState("");
  const [pw2, setPw2] = useState("");
  const [msg, setMsg] = useState("");
  const [busy, setBusy] = useState(false);

  const start = async () => {
    if (!limit) return onClose();
    if (pw !== pw2) return setMsg("Parolalar eşleşmiyor.");
    setBusy(true);
    const p = await makeProtection(mode, pw);
    setBusy(false);
    updateSettings((s) => ({ ...s, protection: p }));
    notify(`Belge korunuyor: ${PROTECTION_LABEL[mode]}.`, "info");
    onClose();
  };

  const stop = async () => {
    if (!cur) return onClose();
    setBusy(true);
    const ok = await checkPassword(cur, pw);
    setBusy(false);
    if (!ok) return setMsg("Girdiğiniz parola yanlış.");
    // Zorlama kalkar (Word gibi kısıtlama ayarı korunur, parola silinir).
    updateSettings((s) => ({ ...s, protection: { mode: cur.mode, enforced: false } }));
    notify("Koruma durduruldu.", "success");
    onClose();
  };

  return (
    <Dialog
      open
      title="Düzenlemeyi Kısıtla"
      width={420}
      onClose={onClose}
      onSubmit={() => void (enforced ? stop() : start())}
      footer={
        <>
          <DialogButton primary disabled={busy} onClick={() => void (enforced ? stop() : start())}>
            {enforced ? "Korumayı Durdur" : "Evet, Korumayı Başlat"}
          </DialogButton>
          <DialogButton onClick={onClose}>İptal</DialogButton>
        </>
      }
    >
      {enforced ? (
        <>
          <p className="oh-dialog-message">
            Bu belge istenmeyen değişikliklere karşı korunuyor. İzin verilen düzenleme: <strong>{PROTECTION_LABEL[cur!.mode]}</strong>.
          </p>
          {cur?.hash && (
            <Field label="Parola:">
              <input type="password" value={pw} autoFocus onChange={(e) => setPw(e.target.value)} aria-label="Parola" />
            </Field>
          )}
        </>
      ) : (
        <>
          <FieldSet legend="Düzenleme kısıtlamaları">
            <Check label="Belgede yalnızca bu tür düzenlemelere izin ver:" v={limit} set={setLimit} />
            <select value={mode} disabled={!limit} onChange={(e) => setMode(e.target.value as DocProtection["mode"])} aria-label="İzin verilen düzenleme">
              {(Object.keys(PROTECTION_LABEL) as DocProtection["mode"][]).map((k) => (
                <option key={k} value={k}>
                  {PROTECTION_LABEL[k]}
                </option>
              ))}
            </select>
          </FieldSet>
          <FieldSet legend="Korumayı başlat (parola isteğe bağlı)">
            <Field label="Yeni parolayı girin:">
              <input type="password" value={pw} onChange={(e) => setPw(e.target.value)} aria-label="Yeni parola" />
            </Field>
            <Field label="Onaylamak için parolayı yeniden girin:">
              <input type="password" value={pw2} onChange={(e) => setPw2(e.target.value)} aria-label="Parola onayı" />
            </Field>
          </FieldSet>
        </>
      )}
      {msg && <div className="bl-dlg-msg is-error">{msg}</div>}
    </Dialog>
  );
}
