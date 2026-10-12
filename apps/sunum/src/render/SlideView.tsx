/**
 * Slayt çizici: arka plan → asıl slayt süslemeleri → düzen süslemeleri →
 * slayt şekilleri. Düzenleyicide, küçük resimlerde, gösteride ve
 * yazdırmada aynı bileşen ölçeklenerek kullanılır.
 */
import { memo, useMemo, type CSSProperties, type ReactNode } from "react";
import { cssColor } from "../model/color";
import { effectiveBodyProps, isTextEmpty, levelResolver, promptText, type TextContext } from "../model/textStyle";
import type { GroupShape, Layout, Master, PicShape, Presentation, Shape, ShapeBase, Slide, SpShape, TableShape, Theme } from "../model/types";
import { customPathD, presetGeometry, type GeomResult } from "./geometry";
import { effectsFilter, fillToCss, lineStroke, PT, reflectionCss, svgPaint, type SvgPaint } from "./style";
import { cellFormat } from "./tableStyles";
import { paragraphsHtml, type TextRenderCtx } from "./textHtml";

export type RenderMode = "edit" | "thumb" | "show" | "print";

export interface EditorSlot {
  shapeId: string;
  cell?: [number, number];
}

export interface RenderCtx {
  theme: Theme;
  master: Master;
  layout?: Layout;
  mode: RenderMode;
  slideNumber: number;
  editing?: EditorSlot | null;
  renderEditor?: (shape: SpShape | TableShape, textCtx: TextRenderCtx, cell?: [number, number]) => ReactNode;
}

export type GeomOverride = Pick<ShapeBase, "x" | "y" | "w" | "h" | "rot"> & { flipH?: boolean; flipV?: boolean; crop?: { l: number; t: number; r: number; b: number } };

export interface SlideViewProps {
  pres: Presentation;
  slide: Slide;
  slideNumber: number;
  mode: RenderMode;
  overrides?: Map<string, GeomOverride>;
  editing?: EditorSlot | null;
  renderEditor?: RenderCtx["renderEditor"];
  /** Gösteride animasyonla gizlenen şekiller vb. */
  className?: string;
  style?: CSSProperties;
  /** Arka plan grafiklerini/şekilleri çizmeden yalnız arka plan (geçiş katmanları için gerekmez). */
  shapeClass?: (s: Shape) => string | undefined;
}

export function slideBackground(pres: Presentation, slide: Slide) {
  const layout = pres.layouts.find((l) => l.id === slide.layoutId);
  return slide.background ?? layout?.background ?? pres.master.background;
}

export const SlideView = memo(function SlideView({ pres, slide, slideNumber, mode, overrides, editing, renderEditor, className, style, shapeClass }: SlideViewProps) {
  const layout = pres.layouts.find((l) => l.id === slide.layoutId) ?? pres.layouts[0];
  const ctx = useMemo<RenderCtx>(
    () => ({ theme: pres.theme, master: pres.master, layout, mode, slideNumber, editing, renderEditor }),
    [pres.theme, pres.master, layout, mode, slideNumber, editing, renderEditor],
  );
  const bg = slide.background ?? layout?.background ?? pres.master.background;
  const showMaster = !slide.hideMasterShapes && (layout?.showMasterShapes ?? true);
  const decoCtx = useMemo<RenderCtx>(() => ({ ...ctx, editing: null, renderEditor: undefined }), [ctx]);
  return (
    <div
      className={`sv sv-${mode} ${className ?? ""}`}
      style={{ width: pres.width, height: pres.height, ...style }}
      data-slide-id={slide.id}
    >
      <div className="sv-bg" style={{ background: fillToCss(bg, pres.theme) }} />
      {showMaster && <DecoLayer shapes={pres.master.shapes} ctx={decoCtx} />}
      {!slide.hideMasterShapes && layout && <DecoLayer shapes={layout.shapes} ctx={decoCtx} />}
      {slide.shapes.map((s) => (
        <ShapeView key={s.id} shape={s} ctx={ctx} override={overrides?.get(s.id)} overrides={overrides} extraClass={shapeClass?.(s)} />
      ))}
    </div>
  );
});

const DecoLayer = memo(function DecoLayer({ shapes, ctx }: { shapes: Shape[]; ctx: RenderCtx }) {
  // Asıl slayt/düzen süslemeleri slayt sınırında kırpılır (düzenleme görünümünde taşmasın).
  return (
    <div className="sv-deco">
      {shapes
        .filter((s) => !(s.kind === "sp" && s.ph))
        .map((s) => (
          <ShapeView key={s.id} shape={s} ctx={ctx} deco />
        ))}
    </div>
  );
});

interface ShapeViewProps {
  shape: Shape;
  ctx: RenderCtx;
  override?: GeomOverride;
  overrides?: Map<string, GeomOverride>;
  /** Grup içindeyse grubun sol üst köşesi. */
  ox?: number;
  oy?: number;
  deco?: boolean;
  extraClass?: string;
}

export const ShapeView = memo(function ShapeView({ shape, ctx, override, overrides, ox = 0, oy = 0, deco, extraClass }: ShapeViewProps) {
  const g = override ?? shape;
  const style: CSSProperties = {
    left: g.x - ox,
    top: g.y - oy,
    width: Math.max(g.w, 0.01),
    height: Math.max(g.h, 0.01),
  };
  if (g.rot) style.transform = `rotate(${g.rot}deg)`;
  if (shape.hidden && ctx.mode !== "edit") return null;
  if (shape.hidden) style.opacity = 0.35;
  const fx = shape.kind === "sp" || shape.kind === "pic" ? shape.effects : undefined;
  const filter = effectsFilter(fx, ctx.theme);
  if (filter) style.filter = filter;
  const refl = reflectionCss(fx);
  if (refl) (style as Record<string, string>).WebkitBoxReflect = refl;
  const common = {
    className: `sh sh-${shape.kind}${deco ? " is-deco" : ""}${extraClass ? ` ${extraClass}` : ""}`,
    "data-shape-id": deco ? undefined : shape.id,
    style,
  };
  const flipH = override?.flipH ?? shape.flipH;
  const flipV = override?.flipV ?? shape.flipV;
  switch (shape.kind) {
    case "sp":
      return (
        <div {...common}>
          <SpBody shape={shape} w={g.w} h={g.h} ctx={ctx} flipH={flipH} flipV={flipV} deco={deco} />
        </div>
      );
    case "pic":
      return (
        <div {...common}>
          <PicBody shape={shape} w={g.w} h={g.h} ctx={ctx} flipH={flipH} flipV={flipV} crop={override?.crop} />
        </div>
      );
    case "table":
      return (
        <div {...common}>
          <TableBody shape={shape} w={g.w} ctx={ctx} />
        </div>
      );
    case "group":
      return (
        <div {...common} style={{ ...style, transform: `${style.transform ?? ""}${flipH ? " scaleX(-1)" : ""}${flipV ? " scaleY(-1)" : ""}` || undefined }}>
          <GroupBody shape={shape} g={g} ctx={ctx} overrides={overrides} deco={deco} />
        </div>
      );
  }
});

function GroupBody({ shape, g, ctx, overrides, deco }: { shape: GroupShape; g: GeomOverride; ctx: RenderCtx; overrides?: Map<string, GeomOverride>; deco?: boolean }) {
  // Grup kutusu değiştiyse (sürükleme önizlemesi) alt şekilleri ölçekle.
  const sx = shape.w ? g.w / shape.w : 1;
  const sy = shape.h ? g.h / shape.h : 1;
  return (
    <>
      {shape.children.map((c) => {
        const childOv = overrides?.get(c.id);
        const ov: GeomOverride | undefined =
          childOv ?? (sx !== 1 || sy !== 1 || g.x !== shape.x || g.y !== shape.y ? { x: g.x + (c.x - shape.x) * sx, y: g.y + (c.y - shape.y) * sy, w: c.w * sx, h: c.h * sy, rot: c.rot } : undefined);
        return <ShapeView key={c.id} shape={c} ctx={ctx} override={ov} overrides={overrides} ox={g.x} oy={g.y} deco={deco} />;
      })}
    </>
  );
}

let defSeq = 0;

function PaintDefs({ paints }: { paints: SvgPaint[] }) {
  const defs = paints.filter((p) => p.def);
  if (!defs.length) return null;
  return (
    <defs>
      {defs.map(({ def }) => {
        const d = def!;
        if (d.kind === "pattern")
          return (
            <pattern key={d.id} id={d.id} patternUnits="objectBoundingBox" width="1" height="1" patternContentUnits="objectBoundingBox">
              <image href={d.src} x="0" y="0" width="1" height="1" preserveAspectRatio="none" />
            </pattern>
          );
        const stops = d.stops!.map((s, i) => <stop key={i} offset={s.offset} stopColor={s.color} />);
        if (d.kind === "radial")
          return (
            <radialGradient key={d.id} id={d.id} cx="50%" cy="50%" r="70%">
              {stops}
            </radialGradient>
          );
        const a = ((d.angle ?? 0) * Math.PI) / 180;
        const x = Math.cos(a) / 2;
        const y = Math.sin(a) / 2;
        return (
          <linearGradient key={d.id} id={d.id} x1={0.5 - x} y1={0.5 - y} x2={0.5 + x} y2={0.5 + y}>
            {stops}
          </linearGradient>
        );
      })}
    </defs>
  );
}

function arrowMarker(id: string, type: string, color: string) {
  const common = { id, viewBox: "0 0 10 10", refX: 5, refY: 5, markerWidth: 4, markerHeight: 4, orient: "auto-start-reverse" as const, markerUnits: "strokeWidth" as const };
  switch (type) {
    case "oval":
      return (
        <marker key={id} {...common}>
          <circle cx="5" cy="5" r="4" fill={color} />
        </marker>
      );
    case "diamond":
      return (
        <marker key={id} {...common}>
          <path d="M5 0 L10 5 L5 10 L0 5 Z" fill={color} />
        </marker>
      );
    case "arrow":
      return (
        <marker key={id} {...common} refX={8}>
          <path d="M1 1 L9 5 L1 9" fill="none" stroke={color} strokeWidth="1.6" />
        </marker>
      );
    case "stealth":
      return (
        <marker key={id} {...common} refX={7}>
          <path d="M0 0 L10 5 L0 10 L3 5 Z" fill={color} />
        </marker>
      );
    default:
      return (
        <marker key={id} {...common} refX={7}>
          <path d="M0 0 L10 5 L0 10 Z" fill={color} />
        </marker>
      );
  }
}

export function shapeGeometry(shape: SpShape, w: number, h: number): GeomResult {
  if (shape.geom === "custom" && shape.paths?.length) {
    return {
      paths: shape.paths.map((p) => ({ d: customPathD(p, w, h), fill: p.fill === false ? "none" : "norm", stroke: p.stroke !== false })),
      text: { x: 0, y: 0, w, h },
      // Dolgusuz ve kapanmayan yollar (serbest çizim, mürekkep) açık çizgi gibi davranır.
      open: shape.paths.length === 1 && shape.paths[0].fill === false && !shape.paths[0].cmds.some((c) => c.c === "Z") ? true : undefined,
    };
  }
  return presetGeometry(shape.geom, w, h, shape.adj);
}

function SpBody({ shape, w, h, ctx, flipH, flipV, deco }: { shape: SpShape; w: number; h: number; ctx: RenderCtx; flipH?: boolean; flipV?: boolean; deco?: boolean }) {
  const geo = shapeGeometry(shape, w, h);
  const id = useMemo(() => `g${(defSeq += 1)}`, []);
  const ln = lineStroke(shape.line, ctx.theme);
  const paints = geo.paths.map((p) => svgPaint(shape.fill, ctx.theme, id, geo.open ? "none" : p.fill));
  const markers: ReactNode[] = [];
  if (ln.stroke !== "none") {
    if (shape.line.head && shape.line.head !== "none") markers.push(arrowMarker(`${id}-h`, shape.line.head, ln.stroke));
    if (shape.line.tail && shape.line.tail !== "none") markers.push(arrowMarker(`${id}-t`, shape.line.tail, ln.stroke));
  }
  const isTextual = !!shape.txBox || !!shape.ph;
  const flip = flipH || flipV ? `scale(${flipH ? -1 : 1},${flipV ? -1 : 1})` : undefined;
  const editMode = ctx.mode === "edit";
  const hasFill = shape.fill.type !== "none";
  const showText = !!shape.text;
  return (
    <>
      <svg className="sh-geom" width={Math.max(w, 1)} height={Math.max(h, 1)} viewBox={`0 0 ${Math.max(w, 1)} ${Math.max(h, 1)}`} overflow="visible" style={flip ? { transform: flip } : undefined}>
        <PaintDefs paints={paints} />
        {markers.length > 0 && <defs>{markers}</defs>}
        {isTextual && !deco && <rect className="sh-hit" x={0} y={0} width={Math.max(w, 1)} height={Math.max(h, 1)} fill="transparent" />}
        {geo.paths.map((p, i) => (
          <path
            key={i}
            className={hasFill && p.fill !== "none" && !geo.open ? "sh-path is-filled" : "sh-path"}
            d={p.d}
            fill={paints[i].paint}
            fillRule={p.evenodd ? "evenodd" : undefined}
            stroke={p.stroke ? ln.stroke : "none"}
            strokeWidth={p.stroke ? ln.width : 0}
            strokeDasharray={p.stroke ? ln.dash : undefined}
            strokeLinejoin="round"
            strokeLinecap={shape.line.cap === "rnd" ? "round" : shape.line.cap === "sq" ? "square" : geo.open ? "round" : undefined}
            markerStart={geo.open && shape.line.head && shape.line.head !== "none" ? `url(#${id}-h)` : undefined}
            markerEnd={geo.open && shape.line.tail && shape.line.tail !== "none" ? `url(#${id}-t)` : undefined}
          />
        ))}
        {geo.open && !deco && <path className="sh-hitline" d={geo.paths[0].d} fill="none" stroke="transparent" strokeWidth={Math.max(10, ln.width + 8)} />}
        {editMode && !deco && (shape.ph || shape.txBox) && isTextEmpty(shape.text) && ctx.editing?.shapeId !== shape.id && (
          <rect className="ph-outline" x={0} y={0} width={Math.max(w, 1)} height={Math.max(h, 1)} fill="none" vectorEffect="non-scaling-stroke" />
        )}
      </svg>
      {showText && <SpText shape={shape} geo={geo} ctx={ctx} />}
      {editMode && shape.ph && !deco && isTextEmpty(shape.text) && (shape.ph.type === "obj" || shape.ph.type === "pic" || shape.ph.type === "tbl") && <PlaceholderIcons type={shape.ph.type} w={w} h={h} />}
    </>
  );
}

function PlaceholderIcons({ type, w, h }: { type: string; w: number; h: number }) {
  const size = Math.max(28, Math.min(56, Math.min(w, h) / 5));
  const items = type === "pic" ? ["picture"] : type === "tbl" ? ["table"] : ["table", "picture"];
  return (
    <div className="ph-icons" style={{ gap: size / 3 }}>
      {items.map((it) => (
        <span key={it} role="button" className="ph-icon" data-ph-action={it} title={it === "table" ? "Tablo Ekle" : "Resimler"} aria-label={it === "table" ? "Tablo Ekle" : "Resimler"} style={{ width: size, height: size }}>
          {it === "table" ? (
            <svg viewBox="0 0 24 24" width="70%" height="70%">
              <rect x="3" y="4" width="18" height="16" rx="1.5" fill="none" stroke="currentColor" strokeWidth="1.4" />
              <path d="M3 9h18M3 14.5h18M9 4v16M15 4v16" stroke="currentColor" strokeWidth="1.2" />
            </svg>
          ) : (
            <svg viewBox="0 0 24 24" width="70%" height="70%">
              <rect x="3" y="5" width="18" height="14" rx="1.5" fill="none" stroke="currentColor" strokeWidth="1.4" />
              <circle cx="9" cy="10" r="1.8" fill="none" stroke="currentColor" strokeWidth="1.2" />
              <path d="M4 17l5-5 4 4 3-3 4 4" fill="none" stroke="currentColor" strokeWidth="1.3" />
            </svg>
          )}
        </span>
      ))}
    </div>
  );
}

function SpText({ shape, geo, ctx }: { shape: SpShape; geo: GeomResult; ctx: RenderCtx }) {
  const tctx: TextContext = { master: ctx.master, layout: ctx.layout };
  const body = shape.text!;
  const bp = effectiveBodyProps(shape, tctx);
  const lvl = useMemo(() => levelResolver(shape, tctx), [shape, ctx.master, ctx.layout]); // eslint-disable-line react-hooks/exhaustive-deps
  const editing = ctx.editing?.shapeId === shape.id && !ctx.editing.cell;
  const prompt = ctx.mode === "edit" && shape.ph ? promptText(shape, tctx) : undefined;
  const trc: TextRenderCtx = {
    theme: ctx.theme,
    lvl,
    fontScale: bp.autofit === "shrink" ? body.fontScale : undefined,
    lnReduction: bp.autofit === "shrink" ? body.lnSpcReduction : undefined,
    slideNumber: ctx.slideNumber,
    editing,
    prompt,
  };
  const r = geo.text;
  const ins = bp.insets;
  const vertical = bp.vert === "vert" || bp.vert === "vert270" || bp.vert === "stacked";
  const style: CSSProperties = {
    left: r.x + ins.l,
    top: r.y + ins.t,
    width: Math.max(0, r.w - ins.l - ins.r),
    height: Math.max(0, r.h - ins.t - ins.b),
    justifyContent: bp.anchor === "ctr" ? "center" : bp.anchor === "b" ? "flex-end" : "flex-start",
    whiteSpace: bp.wrap ? "pre-wrap" : "pre",
  };
  if (vertical) {
    style.writingMode = "vertical-rl";
    if (bp.vert === "vert270") style.transform = "rotate(180deg)";
    if (bp.vert === "stacked") style.textOrientation = "upright";
    style.alignItems = bp.anchor === "ctr" ? "center" : bp.anchor === "b" ? "flex-start" : "flex-end";
    style.justifyContent = "flex-start";
  }
  if (shape.flipV && !vertical) style.transform = "rotate(180deg)";
  const empty = isTextEmpty(body);
  const clickable = !empty || !!shape.ph || !!shape.txBox;
  return (
    <div className={`tb${clickable ? " is-hit" : ""}${empty && prompt && !editing ? " is-empty" : ""}`} style={style}>
      {editing && ctx.renderEditor ? ctx.renderEditor(shape, trc) : <div className="tb-c" dangerouslySetInnerHTML={{ __html: paragraphsHtml(body, trc) }} />}
    </div>
  );
}

function PicBody({ shape, w, h, ctx, flipH, flipV, crop }: { shape: PicShape; w: number; h: number; ctx: RenderCtx; flipH?: boolean; flipV?: boolean; crop?: PicShape["crop"] }) {
  const c = crop ?? shape.crop ?? { l: 0, t: 0, r: 0, b: 0 };
  const iw = w / Math.max(0.01, 1 - c.l - c.r);
  const ih = h / Math.max(0.01, 1 - c.t - c.b);
  const ln = lineStroke(shape.line, ctx.theme);
  const filters: string[] = [];
  const adj = shape.adjust;
  if (adj?.brightness) filters.push(`brightness(${(1 + adj.brightness).toFixed(2)})`);
  if (adj?.contrast) filters.push(`contrast(${(1 + adj.contrast).toFixed(2)})`);
  if (adj?.grayscale) filters.push("grayscale(1)");
  const geo = shape.geom && shape.geom !== "rect" ? presetGeometry(shape.geom, w, h) : null;
  const clip = geo ? `path("${geo.paths[0].d}")` : undefined;
  return (
    <>
      <div className="pic-clip" style={{ clipPath: clip, transform: flipH || flipV ? `scale(${flipH ? -1 : 1},${flipV ? -1 : 1})` : undefined }}>
        {shape.src ? (
          <img
            src={shape.src}
            alt={shape.descr ?? ""}
            draggable={false}
            style={{ left: -c.l * iw, top: -c.t * ih, width: iw, height: ih, filter: filters.join(" ") || undefined, opacity: adj?.alpha }}
          />
        ) : (
          <div className="pic-missing" />
        )}
      </div>
      {ln.stroke !== "none" && (
        <svg className="sh-geom" width={w} height={h} overflow="visible">
          <path d={geo ? geo.paths[0].d : `M0 0 H${w} V${h} H0 Z`} fill="none" stroke={ln.stroke} strokeWidth={ln.width} strokeDasharray={ln.dash} />
        </svg>
      )}
    </>
  );
}

const CELL_INS = { l: 9.6, r: 9.6, t: 4.8, b: 4.8 };

function TableBody({ shape, w, ctx }: { shape: TableShape; w: number; ctx: RenderCtx }) {
  const tctx: TextContext = { master: ctx.master, layout: ctx.layout };
  const scale = shape.w ? w / shape.w : 1;
  const base = useMemo(() => levelResolver({ kind: "sp" }, tctx, "other"), [ctx.master, ctx.layout]); // eslint-disable-line react-hooks/exhaustive-deps
  const border = (b: { color: import("../model/types").Color; width: number } | null) => (b ? `${Math.max(0.5, b.width * PT).toFixed(2)}px solid ${cssColor(b.color, ctx.theme)}` : "none");
  return (
    <table className="tbl" style={{ width: w }}>
      <colgroup>
        {shape.cols.map((c, i) => (
          <col key={i} style={{ width: c * scale }} />
        ))}
      </colgroup>
      <tbody>
        {shape.rows.map((row, r) => (
          <tr key={r} style={{ height: row.h }}>
            {row.cells.map((cell, c) => {
              if (cell.hMerge || cell.vMerge) return null;
              const fmt = cellFormat(shape, r, c);
              const lvl = (n: number) => {
                const b = base(n);
                return { para: b.para, run: { ...b.run, color: fmt.text ?? b.run.color, bold: fmt.bold || b.run.bold } };
              };
              const editing = ctx.editing?.shapeId === shape.id && ctx.editing.cell?.[0] === r && ctx.editing.cell?.[1] === c;
              const trc: TextRenderCtx = { theme: ctx.theme, lvl, slideNumber: ctx.slideNumber, editing };
              return (
                <td
                  key={c}
                  colSpan={cell.gridSpan && cell.gridSpan > 1 ? cell.gridSpan : undefined}
                  rowSpan={cell.rowSpan && cell.rowSpan > 1 ? cell.rowSpan : undefined}
                  data-cell={`${r},${c}`}
                  style={{
                    background: fillToCss(fmt.fill, ctx.theme),
                    borderTop: border(fmt.borders.t),
                    borderBottom: border(fmt.borders.b),
                    borderLeft: border(fmt.borders.l),
                    borderRight: border(fmt.borders.r),
                    verticalAlign: cell.anchor === "ctr" ? "middle" : cell.anchor === "b" ? "bottom" : "top",
                    padding: `${CELL_INS.t}px ${CELL_INS.r}px ${CELL_INS.b}px ${CELL_INS.l}px`,
                  }}
                >
                  {editing && ctx.renderEditor ? ctx.renderEditor(shape, trc, [r, c]) : <div className="tb-c" dangerouslySetInnerHTML={{ __html: paragraphsHtml(cell.text, trc) }} />}
                </td>
              );
            })}
          </tr>
        ))}
      </tbody>
    </table>
  );
}

/** Slaytı verilen kutuya sığdırarak çizer (küçük resim, gösteri, yazdırma). */
export function ScaledSlide({ pres, slide, slideNumber, width, mode = "thumb", className, shapeClass }: { pres: Presentation; slide: Slide; slideNumber: number; width: number; mode?: RenderMode; className?: string; shapeClass?: (s: Shape) => string | undefined }) {
  const scale = width / pres.width;
  return (
    <div className={`sv-scaled ${className ?? ""}`} style={{ width, height: pres.height * scale }}>
      <SlideView pres={pres} slide={slide} slideNumber={slideNumber} mode={mode} style={{ transform: `scale(${scale})`, transformOrigin: "0 0" }} shapeClass={shapeClass} />
    </div>
  );
}
