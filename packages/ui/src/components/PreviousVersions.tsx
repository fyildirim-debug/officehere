// Dosya › Bilgi › "Önceki sürümler" (masaüstü): OfficeHere'in bu dosyaya her pencere
// oturumundaki ilk yazmasından önceki içerik, kullanıcı veri dizininde en çok 3 sürüm.
// Geri yükleme dosyanın içeriğini o sürümle değiştirir (geçerli içerik de sürüm olarak
// saklanır, yani geri alınabilir) ve belgeyi yeniden açar.
import { useCallback, useEffect, useState } from "react";
import { History20Regular } from "@fluentui/react-icons";
import { getNative, type FileVersion } from "@officehere/core";
import { Dialog, DialogButton } from "./Dialog";
import { documentLifecycle } from "../hooks/useDocumentLifecycle";

function formatSize(n: number): string {
  if (n < 1024) return `${n} bayt`;
  if (n < 1024 * 1024) return `${(n / 1024).toLocaleString("tr-TR", { maximumFractionDigits: 0 })} KB`;
  return `${(n / 1024 / 1024).toLocaleString("tr-TR", { maximumFractionDigits: 1 })} MB`;
}

function formatWhen(at: number): string {
  return new Date(at).toLocaleString("tr-TR", { dateStyle: "medium", timeStyle: "short" });
}

/** Belgeyi aynı pencerede, diskten yeniden açar (sayfa yeniden yüklenir). */
function reopenFromDisk(path: string) {
  const url = new URL(window.location.href);
  url.searchParams.set("file", path);
  window.location.replace(url.toString());
}

export function PreviousVersions({ path, fileName }: { path?: string; fileName: string }) {
  const native = getNative();
  const [versions, setVersions] = useState<FileVersion[] | null>(null);
  const [asking, setAsking] = useState<FileVersion | null>(null);
  const [error, setError] = useState("");
  const [busy, setBusy] = useState(false);

  const load = useCallback(() => {
    if (!path || !native?.listVersions) return;
    native
      .listVersions(path)
      .then(setVersions)
      .catch(() => setVersions([]));
  }, [native, path]);

  useEffect(load, [load]);

  if (!path || !native?.listVersions || !native.restoreVersion) return null;

  const restore = async (v: FileVersion) => {
    setBusy(true);
    setError("");
    try {
      const r = await native.restoreVersion!(path, v.id);
      if (!r.ok) {
        setError(`Geri yüklenemedi: ${r.reason ?? r.message ?? "dosya yazılamadı"}.`);
        return;
      }
      // Bu pencerenin kurtarma taslağı silinir; belge diskten (geri yüklenen sürümle) yeniden açılır.
      await documentLifecycle.forgetDraft();
      reopenFromDisk(path);
    } catch (e) {
      setError(`Geri yüklenemedi: ${(e as Error).message}`);
    } finally {
      setBusy(false);
      setAsking(null);
    }
  };

  return (
    <section className="oh-versions" aria-label="Önceki sürümler">
      <h2 className="oh-bs-h2">Önceki sürümler</h2>
      <p className="oh-bs-note">
        OfficeHere bu dosyanın üzerine her oturumda ilk kez yazmadan önce önceki hâlini bilgisayarınızda saklar (en çok 3 sürüm). Otomatik kaydetmeden önceki özgün dosya ayrıca
        dosyanın yanında “.bak” uzantısıyla durur.
      </p>
      {versions === null ? null : versions.length === 0 ? (
        <p className="oh-versions-empty">Henüz önceki sürüm yok.</p>
      ) : (
        <ul className="oh-versions-list">
          {versions.map((v) => (
            <li key={v.id}>
              <History20Regular aria-hidden="true" />
              <span className="oh-versions-when">{formatWhen(v.savedAt)}</span>
              <span className="oh-versions-size">{formatSize(v.size)}</span>
              <button type="button" className="oh-btn" disabled={busy} onClick={() => setAsking(v)}>
                Geri yükle
              </button>
            </li>
          ))}
        </ul>
      )}
      {error && (
        <p className="oh-versions-error" role="alert">
          {error}
        </p>
      )}
      <Dialog
        open={!!asking}
        title="Önceki sürümü geri yükle"
        onClose={() => setAsking(null)}
        width={460}
        footer={
          <>
            <DialogButton primary disabled={busy} onClick={() => asking && void restore(asking)}>
              Geri yükle
            </DialogButton>
            <DialogButton onClick={() => setAsking(null)}>İptal</DialogButton>
          </>
        }
      >
        <p className="oh-dialog-message">
          <strong>“{fileName}”</strong> dosyası {asking ? formatWhen(asking.savedAt) : ""} tarihli sürümüyle değiştirilecek. Dosyanın şimdiki hâli de önceki sürümler arasında
          saklanır.
          {documentLifecycle.isDirty() ? " Bu penceredeki kaydedilmemiş değişiklikler kaybolacak." : ""}
        </p>
      </Dialog>
    </section>
  );
}
