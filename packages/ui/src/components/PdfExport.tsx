/**
 * PDF olarak dışa aktarma seçenekleri (masaüstü; Belge, Tablo ve Sunum aynı iletişim kutusunu kullanır):
 * biçim (PDF, PDF/A-2b, PDF/A-1b), sayfa aralığı, belge bilgileri, yer imleri, etiketler ve sonrası.
 */
import { useId, useState, type CSSProperties } from "react";
import { createRoot } from "react-dom/client";
import type { PdfAfterExport, PdfStandard } from "@officehere/core";
import { Dialog, DialogButton, Field, FieldSet } from "./Dialog";

export interface PdfExportDefaults {
  title?: string;
  author?: string;
  subject?: string;
  keywords?: string;
  /** Çıktının sayfa sayısı (biliniyorsa aralık denetlenir). */
  pageCount?: number;
  /** Önceden seçili sayfa aralığı ("2-3"). */
  pages?: string;
  /** "Başlıklardan yer imleri" seçeneği (başlığı olan belgeler). */
  headings?: boolean;
}

export interface PdfExportChoice {
  standard: PdfStandard;
  /** Chromium biçiminde sayfa aralığı ("1-3, 5"); tümü için yok. */
  pageRanges?: string;
  outline: boolean;
  tagged: boolean;
  after: PdfAfterExport;
  title?: string;
  author?: string;
  subject?: string;
  keywords?: string;
}

/**
 * "1-3, 5" biçimindeki aralığı denetleyip "1-3, 5" biçimine getirir. Sayfa sayısı biliniyorsa
 * "4-" (sona kadar) kabul edilir ve aralık belgeyle sınırlanır.
 */
export function parsePdfPageRange(text: string, count?: number): { ranges: string } | { error: string } {
  const parts = text
    .split(/[,;]/)
    .map((p) => p.trim())
    .filter(Boolean);
  if (!parts.length) return { error: "Sayfa aralığını girin (ör. 1-3, 5)." };
  const out: string[] = [];
  for (const part of parts) {
    const m = /^(\d+)\s*(?:[-–]\s*(\d*))?$/.exec(part);
    if (!m) return { error: `Geçersiz sayfa aralığı: "${part}"` };
    const a = Number(m[1]);
    let b = m[2] === undefined ? a : m[2] === "" ? (count ?? NaN) : Number(m[2]);
    if (Number.isNaN(b)) return { error: `"${part}" için son sayfayı da yazın.` };
    if (a < 1 || b < a) return { error: `Geçersiz sayfa aralığı: "${part}"` };
    if (count !== undefined) {
      if (a > count) return { error: `Belgede ${count} sayfa var; "${part}" belgenin dışında.` };
      b = Math.min(b, count);
    }
    out.push(a === b ? String(a) : `${a}-${b}`);
  }
  return { ranges: out.join(", ") };
}

const STANDARDS: { id: PdfStandard; label: string; hint: string }[] = [
  { id: "pdf", label: "PDF", hint: "Paylaşmak ve yazdırmak için." },
  { id: "pdfa-2b", label: "PDF/A-2b (arşiv)", hint: "Uzun süreli saklama: yazı tipleri ve renk profili dosyaya gömülür." },
  { id: "pdfa-1b", label: "PDF/A-1b (e-Yazışma üst yazısı)", hint: "Resmî yazışma ekleri için; gölge ve yarı saydam öğe içeremez." },
];

const AFTER: { id: PdfAfterExport; label: string }[] = [
  { id: "none", label: "Hiçbir şey yapma" },
  { id: "open", label: "Dosyayı aç" },
  { id: "folder", label: "Klasörünü göster" },
];

const STORE_KEY = "oh-pdf-secenekleri";

interface Remembered {
  standard?: PdfStandard;
  outline?: boolean;
  tagged?: boolean;
  after?: PdfAfterExport;
  author?: string;
}

function remembered(): Remembered {
  try {
    const v = JSON.parse(localStorage.getItem(STORE_KEY) ?? "{}") as Remembered;
    return v && typeof v === "object" ? v : {};
  } catch {
    return {};
  }
}

function remember(v: Remembered) {
  try {
    localStorage.setItem(STORE_KEY, JSON.stringify(v));
  } catch {
    // Depolama kapalı: seçimler yalnız bu oturumda geçerli.
  }
}

const row: CSSProperties = { display: "flex", alignItems: "flex-start", gap: 6, margin: "4px 0" };
const hint: CSSProperties = { color: "var(--oh-text-muted)", fontSize: 11 };

export interface PdfExportDialogProps {
  defaults: PdfExportDefaults;
  onDone: (choice: PdfExportChoice | null) => void;
}

export function PdfExportDialog({ defaults, onDone }: PdfExportDialogProps) {
  const id = useId();
  const [saved] = useState(remembered);
  const [standard, setStandard] = useState<PdfStandard>(saved.standard ?? "pdf");
  const [all, setAll] = useState(!defaults.pages);
  const [pages, setPages] = useState(defaults.pages ?? "");
  const [title, setTitle] = useState(defaults.title ?? "");
  const [author, setAuthor] = useState(defaults.author || saved.author || "");
  const [outline, setOutline] = useState(saved.outline ?? true);
  const [tagged, setTagged] = useState(saved.tagged ?? true);
  const [after, setAfter] = useState<PdfAfterExport>(saved.after ?? "none");
  const headings = !!defaults.headings;
  const parsed = all ? null : parsePdfPageRange(pages, defaults.pageCount);
  const error = parsed && "error" in parsed ? parsed.error : null;

  const submit = () => {
    if (error) return;
    const useOutline = headings && outline;
    remember({ standard, outline, tagged, after, author: author.trim() || saved.author });
    onDone({
      standard,
      pageRanges: parsed && "ranges" in parsed ? parsed.ranges : undefined,
      outline: useOutline,
      tagged: tagged || useOutline,
      after,
      title: title.trim() || undefined,
      author: author.trim() || undefined,
      subject: defaults.subject || undefined,
      keywords: defaults.keywords || undefined,
    });
  };

  return (
    <Dialog
      open
      title="PDF Olarak Dışa Aktar"
      onClose={() => onDone(null)}
      onSubmit={submit}
      width={480}
      footer={
        <>
          <DialogButton primary onClick={submit} disabled={!!error}>
            Dışa aktar
          </DialogButton>
          <DialogButton onClick={() => onDone(null)}>İptal</DialogButton>
        </>
      }
    >
      <FieldSet legend="Biçim">
        <div role="radiogroup" aria-label="Biçim">
          {STANDARDS.map((s) => (
            <label key={s.id} style={row}>
              <input type="radio" name={`${id}-std`} aria-label={s.label} aria-describedby={`${id}-${s.id}`} checked={standard === s.id} onChange={() => setStandard(s.id)} />
              <span>
                {s.label}
                <br />
                <span id={`${id}-${s.id}`} style={hint}>
                  {s.hint}
                </span>
              </span>
            </label>
          ))}
        </div>
      </FieldSet>
      <FieldSet legend="Sayfalar">
        <div role="radiogroup" aria-label="Sayfalar" style={{ display: "flex", alignItems: "center", gap: 14, flexWrap: "wrap" }}>
          <label style={{ ...row, alignItems: "center" }}>
            <input type="radio" name={`${id}-rng`} checked={all} onChange={() => setAll(true)} />
            Tümü{defaults.pageCount ? ` (${defaults.pageCount} sayfa)` : ""}
          </label>
          <label style={{ ...row, alignItems: "center" }}>
            <input type="radio" name={`${id}-rng`} checked={!all} onChange={() => setAll(false)} />
            Aralık:
          </label>
          <input
            type="text"
            aria-label="Sayfa aralığı"
            placeholder="ör. 1-3, 5"
            value={pages}
            onFocus={() => setAll(false)}
            onChange={(e) => {
              setPages(e.target.value);
              setAll(false);
            }}
            aria-invalid={!!error}
            style={{ width: 140, height: 26, padding: "0 6px", border: "1px solid var(--oh-border)", borderRadius: "var(--oh-radius)", background: "var(--oh-surface)", fontSize: 12 }}
          />
        </div>
        {error && (
          <div role="alert" style={{ background: "var(--oh-msg-error-bg)", borderRadius: 4, padding: "3px 6px", fontSize: 11, marginTop: 4 }}>
            {error}
          </div>
        )}
      </FieldSet>
      <FieldSet legend="Belge bilgileri">
        <Field label="Başlık" inline>
          <input type="text" value={title} onChange={(e) => setTitle(e.target.value)} style={{ flex: 1 }} />
        </Field>
        <Field label="Yazar" inline>
          <input type="text" value={author} onChange={(e) => setAuthor(e.target.value)} style={{ flex: 1 }} />
        </Field>
      </FieldSet>
      {headings && (
        <label style={row}>
          <input type="checkbox" checked={outline} onChange={(e) => setOutline(e.target.checked)} />
          Başlıklardan yer imleri oluştur
        </label>
      )}
      <label style={row}>
        <input
          type="checkbox"
          aria-label="Erişilebilirlik etiketleri ekle (etiketli PDF)"
          checked={tagged || (headings && outline)}
          disabled={headings && outline}
          onChange={(e) => setTagged(e.target.checked)}
        />
        <span>
          Erişilebilirlik etiketleri ekle (etiketli PDF)
          {headings && outline && (
            <>
              <br />
              <span style={hint}>Yer imleri için gereklidir.</span>
            </>
          )}
        </span>
      </label>
      <Field label="Dışa aktardıktan sonra" inline>
        <select value={after} onChange={(e) => setAfter(e.target.value as PdfAfterExport)}>
          {AFTER.map((a) => (
            <option key={a.id} value={a.id}>
              {a.label}
            </option>
          ))}
        </select>
      </Field>
      <details style={{ marginTop: 8 }}>
        <summary style={{ cursor: "pointer" }}>e-imza ile imzalamak</summary>
        <ol style={{ margin: "6px 0 0", paddingLeft: 20, lineHeight: 1.45 }}>
          <li>Belgeyi PDF/A biçiminde dışa aktarın (resmî yazışmalarda çoğunlukla PDF/A-1b istenir).</li>
          <li>Kaydedilen dosyayı e-imza uygulamanızla açın ve elektronik imza kartınız ya da mobil imzanızla PDF içine gömülü (PAdES) imza atın.</li>
          <li>İmzalanan dosyayı değiştirmeyin ve yeniden dışa aktarmayın: her değişiklik imzayı geçersiz kılar.</li>
        </ol>
        <p style={{ ...hint, margin: "6px 0 0" }}>OfficeHere belgeyi kendisi imzalamaz ve imza için ağa bağlanmaz.</p>
      </details>
    </Dialog>
  );
}

let asking = false;

/**
 * Seçenek iletişim kutusunu gösterir; iptalde null. Kutu kendi kökünde çizilir (uygulama kökü gerekmez)
 * ve söz çözülmeden önce belgeden kaldırılır, böylece yazdırma çıktısına girmez.
 */
export function askPdfExport(defaults: PdfExportDefaults): Promise<PdfExportChoice | null> {
  // Kutu açıkken gelen ikinci istek (ör. çift tıklama) yok sayılır.
  if (asking || typeof document === "undefined") return Promise.resolve(null);
  asking = true;
  return new Promise((resolve) => {
    const host = document.createElement("div");
    host.className = "oh-pdf-export-host";
    document.body.appendChild(host);
    const root = createRoot(host);
    let closed = false;
    const done = (choice: PdfExportChoice | null) => {
      if (closed) return;
      closed = true;
      setTimeout(() => {
        root.unmount();
        host.remove();
        asking = false;
        resolve(choice);
      }, 0);
    };
    root.render(<PdfExportDialog defaults={defaults} onDone={done} />);
  });
}
