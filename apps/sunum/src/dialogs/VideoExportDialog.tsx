/** Dosya › Dışarı Aktar › Video Oluştur: kalite, zamanlamalar, slayt süresi ve ilerleme. */
import { useMemo, useRef, useState } from "react";
import { Dialog, DialogButton, Field } from "@officehere/ui";
import { downloadBlob } from "@officehere/core";
import { formatTime } from "../show/showKeys";
import type { Actions } from "../state/actions";
import { DEFAULT_VIDEO_OPTIONS, exportVideo, videoTimeline, VideoUnsupportedError, type VideoOptions } from "../app/videoExport";
import { documentPres } from "../state/masterActions";

const QUALITIES: { h: VideoOptions["height"]; label: string }[] = [
  { h: 1080, label: "Full HD (1080p)" },
  { h: 720, label: "HD (720p)" },
  { h: 480, label: "Standart (480p)" },
];

export function VideoExportDialog({ open, onClose, actions: a }: { open: boolean; onClose: () => void; actions: Actions }) {
  const pres = open ? documentPres(a) : null;
  const hasTimings = !!pres?.slides.some((s) => !s.hidden && s.transition?.advanceAfter !== undefined);
  const [opts, setOpts] = useState<VideoOptions>(DEFAULT_VIDEO_OPTIONS);
  const [secText, setSecText] = useState("5");
  const [progress, setProgress] = useState<number | null>(null);
  const abort = useRef<AbortController | null>(null);
  const eff: VideoOptions = { ...opts, useTimings: opts.useTimings && hasTimings };
  const timeline = useMemo(() => (pres ? videoTimeline(pres, eff) : { total: 0, segs: [] }), [pres, eff.height, eff.secondsPerSlide, eff.useTimings]); // eslint-disable-line react-hooks/exhaustive-deps
  const shown = pres?.slides.filter((s) => !s.hidden).length ?? 0;
  const busy = progress !== null;

  const close = () => {
    abort.current?.abort();
    abort.current = null;
    setProgress(null);
    onClose();
  };

  const run = async () => {
    if (!pres || busy) return;
    const ctl = new AbortController();
    abort.current = ctl;
    setProgress(0);
    try {
      const blob = await exportVideo(pres, eff, (p) => setProgress(p), ctl.signal);
      if (ctl.signal.aborted) return;
      const name = `${a.s.fileName}.webm`;
      downloadBlob(blob, name);
      a.message(`Video oluşturuldu: ${name}`, "success");
      abort.current = null;
      setProgress(null);
      onClose();
    } catch (e) {
      if ((e as Error).name === "AbortError") return;
      a.message(e instanceof VideoUnsupportedError ? `Video oluşturulamadı: ${e.message} Uygulamanın güncel sürümünü kullanın.` : `Video oluşturulamadı: ${(e as Error).message}`, "error");
      abort.current = null;
      setProgress(null);
      onClose();
    }
  };

  return (
    <Dialog
      open={open}
      title="Video Oluştur"
      onClose={close}
      width={440}
      footer={
        <>
          <DialogButton primary disabled={busy || !shown} onClick={() => void run()}>
            Video Oluştur
          </DialogButton>
          <DialogButton onClick={close}>İptal</DialogButton>
        </>
      }
    >
      <p style={{ marginTop: 0 }}>Sunuyu geçişler, animasyon adımları ve zamanlamalarla bir video dosyasına (.webm) dönüştürün.</p>
      <Field label="Kalite:">
        <select aria-label="Video kalitesi" disabled={busy} value={opts.height} onChange={(e) => setOpts({ ...opts, height: Number(e.target.value) as VideoOptions["height"] })} style={{ width: "100%", height: 26 }}>
          {QUALITIES.map((q) => (
            <option key={q.h} value={q.h}>
              {q.label}
            </option>
          ))}
        </select>
      </Field>
      <Field label="Zamanlamalar:">
        <select aria-label="Zamanlamalar" disabled={busy} value={eff.useTimings ? "1" : "0"} onChange={(e) => setOpts({ ...opts, useTimings: e.target.value === "1" })} style={{ width: "100%", height: 26 }}>
          <option value="1" disabled={!hasTimings}>
            Kaydedilmiş Zamanlamaları Kullan
          </option>
          <option value="0">Kaydedilmiş Zamanlamaları Kullanma</option>
        </select>
      </Field>
      <Field label="Her slaytta geçirilecek saniye:" inline>
        <input
          type="text"
          inputMode="decimal"
          aria-label="Her slaytta geçirilecek saniye"
          disabled={busy}
          value={secText}
          onChange={(e) => {
            setSecText(e.target.value);
            const v = Number(e.target.value.replace(",", "."));
            if (Number.isFinite(v) && v > 0) setOpts({ ...opts, secondsPerSlide: Math.min(600, v) });
          }}
          style={{ width: 70 }}
        />
      </Field>
      <p className="video-info" data-total={timeline.total}>
        {shown} slayt · Video süresi: {formatTime(timeline.total)}
      </p>
      {busy && (
        <div className="video-progress">
          <progress max={1} value={progress ?? 0} aria-label="Video oluşturuluyor" />
          <span>%{Math.round((progress ?? 0) * 100)}</span>
        </div>
      )}
    </Dialog>
  );
}
