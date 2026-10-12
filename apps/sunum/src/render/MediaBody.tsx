/** Video ve ses şekilleri: düzenleyicide kapak + oynat simgesi, gösteride HTML5 oynatıcı. */
import { useRef, useState } from "react";
import type { MediaInfo, PicShape } from "../model/types";

export function MediaBody({ shape, media, w, h, mode }: { shape: PicShape; media: MediaInfo; w: number; h: number; mode: "edit" | "show" }) {
  const audioRef = useRef<HTMLAudioElement>(null);
  const [playing, setPlaying] = useState(false);
  if (mode === "show") {
    if (media.kind === "video") {
      return (
        <video
          className="media-el"
          data-media="video"
          src={media.src}
          poster={shape.src || undefined}
          controls
          autoPlay={media.autoPlay}
          loop={media.loop}
          playsInline
          style={{ width: w, height: h }}
          onClick={(e) => e.stopPropagation()}
        />
      );
    }
    return (
      <div
        className={`media-audio${playing ? " is-playing" : ""}`}
        data-media="audio"
        role="button"
        aria-label={playing ? "Sesi duraklat" : "Sesi yürüt"}
        style={{ width: w, height: h }}
        onClick={(e) => {
          e.stopPropagation();
          const el = audioRef.current;
          if (!el) return;
          if (el.paused) void el.play().catch(() => undefined);
          else el.pause();
        }}
      >
        {shape.src ? <img src={shape.src} alt={shape.descr ?? "Ses"} draggable={false} /> : <SpeakerIcon />}
        <audio ref={audioRef} src={media.src} autoPlay={media.autoPlay} loop={media.loop} onPlay={() => setPlaying(true)} onPause={() => setPlaying(false)} onEnded={() => setPlaying(false)} />
      </div>
    );
  }
  return (
    <div className="media-edit" style={{ width: w, height: h }}>
      {shape.src ? <img src={shape.src} alt={shape.descr ?? ""} draggable={false} /> : media.kind === "audio" ? <SpeakerIcon /> : <div className="media-blank" />}
      {media.kind === "video" && (
        <div className="media-badge" aria-hidden="true">
          <svg viewBox="0 0 24 24" width="100%" height="100%">
            <circle cx="12" cy="12" r="11" fill="rgba(0,0,0,0.55)" />
            <path d="M9.5 7.5 L17 12 L9.5 16.5 Z" fill="#fff" />
          </svg>
        </div>
      )}
    </div>
  );
}

export function SpeakerIcon() {
  return (
    <svg viewBox="0 0 48 48" width="100%" height="100%" aria-hidden="true">
      <path d="M8 18h8l11-9v30l-11-9H8z" fill="#5b5b5b" />
      <path d="M32 17c3 2 4.5 4.5 4.5 7s-1.5 5-4.5 7M35.5 12c5 3.5 7.5 7.5 7.5 12s-2.5 8.5-7.5 12" fill="none" stroke="#5b5b5b" strokeWidth="2.6" strokeLinecap="round" />
    </svg>
  );
}

/** Ses simgesi kapağı (PPTX'e yazılan resim). */
export const SPEAKER_POSTER = `data:image/svg+xml;charset=utf-8,${encodeURIComponent(
  '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 48 48" width="96" height="96"><path d="M8 18h8l11-9v30l-11-9H8z" fill="#5b5b5b"/><path d="M32 17c3 2 4.5 4.5 4.5 7s-1.5 5-4.5 7M35.5 12c5 3.5 7.5 7.5 7.5 12s-2.5 8.5-7.5 12" fill="none" stroke="#5b5b5b" stroke-width="2.6" stroke-linecap="round"/></svg>',
)}`;
