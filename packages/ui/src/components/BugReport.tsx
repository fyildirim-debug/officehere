// "Hata bildir": sürüm, işletim sistemi ve son günlük satırlarını kullanıcıya GÖSTEREN,
// düzenlenebilir bir pencere. "GitHub'da bildir" önceden doldurulmuş bir sorun (issue)
// sayfası açar; kullanıcı orada göndermeden hiçbir şey paylaşılmaz (telemetri yok).
import { useEffect, useState } from "react";
import { Bug24Regular } from "@fluentui/react-icons";
import { getNative, type AppInfo, type AppKind } from "@officehere/core";
import { Dialog, DialogButton } from "./Dialog";
import { RibbonButton } from "./Ribbon";
import { APP_NAMES } from "./AppShell";
import { recentPageLog } from "../hooks/errorReporting";

export const ISSUES_URL = "https://github.com/fyildirim-debug/officehere/issues/new";
/** Tarayıcıların ve GitHub'ın kabul ettiği güvenli adres uzunluğu. */
const MAX_URL = 7000;

type ReportApp = AppKind | "baslangic";

function browserInfo(): string {
  if (typeof navigator === "undefined") return "bilinmiyor";
  const ua = navigator.userAgent;
  const os = /Windows NT [\d.]+/.exec(ua)?.[0] ?? /Mac OS X [\d_]+/.exec(ua)?.[0]?.replace(/_/g, ".") ?? (/Android [\d.]+/.exec(ua)?.[0] || (/Linux/.test(ua) ? "Linux" : "bilinmiyor"));
  const browser = /(Edg|Firefox|Chrome|Safari)\/[\d.]+/.exec(ua)?.[0] ?? "tarayıcı";
  return `${os} · ${browser}`;
}

/** Rapor metni (kullanıcı düzenleyebilir). */
export function buildReportBody(app: ReportApp, info: AppInfo | null, logLines: string[]): string {
  const env = info
    ? [`- Uygulama: ${APP_NAMES[app]} ${info.version} (masaüstü)`, `- İşletim sistemi: ${info.platform} ${info.osVersion} (${info.arch})`, info.electron ? `- Electron ${info.electron} · Chrome ${info.chrome ?? "?"}` : ""]
    : [`- Uygulama: ${APP_NAMES[app]} (web)`, `- Tarayıcı: ${browserInfo()}`];
  return [
    "## Ne oldu?",
    "(Sorunu ve nasıl yeniden oluşturulabileceğini kısaca yazın. Belgenizden kişisel bilgi eklemeyin.)",
    "",
    "## Ortam",
    ...env.filter(Boolean),
    "",
    "## Son günlük satırları",
    "```",
    ...(logLines.length ? logLines : ["(günlükte kayıt yok)"]),
    "```",
  ].join("\n");
}

/** Önceden doldurulmuş GitHub sorun adresi; uzun metinde günlük baştan kısaltılır. */
export function issueUrl(title: string, body: string): string {
  const make = (b: string) => `${ISSUES_URL}?title=${encodeURIComponent(title)}&body=${encodeURIComponent(b)}`;
  let url = make(body);
  if (url.length <= MAX_URL) return url;
  const note = "\n(Metin uzun olduğu için kısaltıldı; tamamını “Panoya kopyala” ile kopyalayıp buraya yapıştırabilirsiniz.)";
  let cut = body;
  while (cut.length > 200) {
    cut = cut.slice(0, Math.floor(cut.length * 0.85));
    url = make(cut + note);
    if (url.length <= MAX_URL) return url;
  }
  return make(note.trim());
}

async function copyText(text: string): Promise<boolean> {
  try {
    await navigator.clipboard.writeText(text);
    return true;
  } catch {
    // Eski yöntem (izin verilmeyen bağlamlar).
    try {
      const ta = document.createElement("textarea");
      ta.value = text;
      ta.style.position = "fixed";
      ta.style.opacity = "0";
      document.body.appendChild(ta);
      ta.select();
      const ok = document.execCommand("copy");
      ta.remove();
      return ok;
    } catch {
      return false;
    }
  }
}

export function BugReportDialog({ open, onClose, app }: { open: boolean; onClose: () => void; app: ReportApp }) {
  const [title, setTitle] = useState("");
  const [body, setBody] = useState("");
  const [note, setNote] = useState("");

  useEffect(() => {
    if (!open) return;
    let alive = true;
    setNote("");
    setTitle(`[${APP_NAMES[app].replace("OfficeHere ", "") || "OfficeHere"}] `);
    void (async () => {
      const native = getNative();
      let info: AppInfo | null = null;
      let lines: string[] = [];
      try {
        info = (await native?.getAppInfo?.()) ?? null;
        lines = (await native?.getLogTail?.(30)) ?? [];
      } catch {
        // masaüstü bilgisi alınamadı
      }
      if (!native?.getLogTail) lines = recentPageLog().slice(-30);
      if (alive) setBody(buildReportBody(app, info, lines));
    })();
    return () => {
      alive = false;
    };
  }, [open, app]);

  return (
    <Dialog
      open={open}
      title="Hata bildir"
      onClose={onClose}
      width={620}
      footer={
        <>
          <DialogButton
            primary
            onClick={() => {
              window.open(issueUrl(title.trim() || "Hata bildirimi", body), "_blank", "noopener");
              setNote("Tarayıcınızda GitHub sayfası açıldı. Bildirim, siz orada gönderdiğinizde paylaşılır.");
            }}
          >
            GitHub'da bildir
          </DialogButton>
          <DialogButton
            onClick={() => {
              void copyText(`${title}\n\n${body}`).then((ok) => setNote(ok ? "Rapor panoya kopyalandı." : "Panoya kopyalanamadı; metni seçip kendiniz kopyalayın."));
            }}
          >
            Panoya kopyala
          </DialogButton>
          <DialogButton onClick={onClose}>Kapat</DialogButton>
        </>
      }
    >
      <div className="oh-bugreport">
        <p className="oh-dialog-message">
          Aşağıdaki bilgiler yalnız siz göndermeyi seçerseniz paylaşılır; OfficeHere kendiliğinden hiçbir şey göndermez. Göndermeden önce metni düzenleyebilir, istemediğiniz
          satırları silebilirsiniz. Günlükte belge içeriği ve dosya adları bulunmaz.
        </p>
        <label className="oh-field">
          <span className="oh-field-label">Başlık</span>
          <input type="text" value={title} onChange={(e) => setTitle(e.target.value)} aria-label="Başlık" />
        </label>
        <label className="oh-field">
          <span className="oh-field-label">Rapor</span>
          <textarea className="oh-bugreport-body" value={body} onChange={(e) => setBody(e.target.value)} aria-label="Rapor" rows={14} spellCheck={false} />
        </label>
        {note && (
          <p className="oh-bugreport-note" role="status">
            {note}
          </p>
        )}
      </div>
    </Dialog>
  );
}

/** Yardım sekmesi için "Hata Bildir" düğmesi (iletişim kutusunu kendisi açar). */
export function ReportBugButton({ app, size = "large" }: { app: ReportApp; size?: "large" | "small" }) {
  const [open, setOpen] = useState(false);
  return (
    <>
      <RibbonButton size={size} icon={<Bug24Regular />} label="Hata Bildir" description="Sürüm, işletim sistemi ve son günlük satırlarıyla bir hata bildirimi hazırlar. Siz göndermeden hiçbir şey paylaşılmaz." onClick={() => setOpen(true)} />
      <BugReportDialog open={open} onClose={() => setOpen(false)} app={app} />
    </>
  );
}
