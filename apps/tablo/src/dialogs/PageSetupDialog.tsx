/** Sayfa Yapısı iletişim kutusu: Sayfa, Kenar Boşlukları, Üst Bilgi/Alt Bilgi, Sayfa (yazdırma) sekmeleri. */
import { useState } from "react";
import { Dialog, DialogButton, DialogTabs, Field, FieldSet } from "@officehere/ui";
import type { Controller } from "../controller";
import { parseRangeLoose } from "../controller";
import { rectToA1 } from "../model/rect";
import type { PageSetup } from "../model/types";
import { Check, OkCancel, Radio, fmtTr, parseTrNumber } from "./common";
import { colToName, nameToCol } from "@officehere/formula";

type P = { ctl: Controller; onClose: () => void; arg?: unknown };

export const PAPERS: { id: PageSetup["paper"]; label: string; w: number; h: number }[] = [
  { id: "A4", label: "A4 (21 x 29,7 cm)", w: 21, h: 29.7 },
  { id: "A3", label: "A3 (29,7 x 42 cm)", w: 29.7, h: 42 },
  { id: "A5", label: "A5 (14,8 x 21 cm)", w: 14.8, h: 21 },
  { id: "Letter", label: "Letter (21,59 x 27,94 cm)", w: 21.59, h: 27.94 },
  { id: "Legal", label: "Legal (21,59 x 35,56 cm)", w: 21.59, h: 35.56 },
];

export const MARGIN_PRESETS: { id: string; label: string; m: PageSetup["margins"] }[] = [
  { id: "normal", label: "Normal", m: { top: 0.75, bottom: 0.75, left: 0.7, right: 0.7, header: 0.3, footer: 0.3 } },
  { id: "wide", label: "Geniş", m: { top: 1, bottom: 1, left: 1, right: 1, header: 0.5, footer: 0.5 } },
  { id: "narrow", label: "Dar", m: { top: 0.75, bottom: 0.75, left: 0.25, right: 0.25, header: 0.3, footer: 0.3 } },
];

/** Sayfa düzeni değişikliği (geri alınabilir). */
export function setPage(ctl: Controller, patch: Partial<PageSetup>, label = "Sayfa Yapısı"): void {
  const sheet = ctl.sheet;
  ctl.ops.sheetOp(sheet, label, () => {
    sheet.page = { ...sheet.page, ...patch };
  });
}

const cm = (inch: number) => fmtTr(inch * 2.54, 2);
const inch = (s: string) => Math.max(0, parseTrNumber(s) / 2.54);
/** Metin değişmediyse özgün değeri korur (yuvarlama kaymasını önler). */
const keep = (s: string, orig: number) => (s === cm(orig) ? orig : inch(s));

export function PageSetupDialog({ ctl, onClose, arg }: P) {
  const p = ctl.sheet.page;
  const [tab, setTab] = useState((arg as { tab?: string } | undefined)?.tab ?? "page");
  const [orientation, setOrientation] = useState(p.orientation);
  const [paper, setPaper] = useState(p.paper);
  const [fit, setFit] = useState<"no" | "sheet" | "cols" | "rows">(p.fitMode ?? (p.fitToWidth ? "cols" : "no"));
  const [m, setM] = useState(() => Object.fromEntries(Object.entries(p.margins).map(([k, v]) => [k, cm(v)])) as Record<keyof PageSetup["margins"], string>);
  const [header, setHeader] = useState(p.header ?? "");
  const [footer, setFooter] = useState(p.footer ?? "");
  const [area, setArea] = useState(p.printArea ? rectToA1(p.printArea) : "");
  const [grid, setGrid] = useState(p.printGridlines);
  const [heads, setHeads] = useState(p.printHeadings);
  const [tRows, setTRows] = useState(p.titleRows ? `$${p.titleRows.r1 + 1}:$${p.titleRows.r2 + 1}` : "");
  const [tCols, setTCols] = useState(p.titleCols ? `$${colToName(p.titleCols.c1)}:$${colToName(p.titleCols.c2)}` : "");
  const [bw, setBw] = useState(!!p.blackAndWhite);
  const [draft, setDraft] = useState(!!p.draft);
  const [order, setOrder] = useState<NonNullable<PageSetup["pageOrder"]>>(p.pageOrder ?? "downThenOver");
  const build = (): Partial<PageSetup> | null => {
    let printArea = undefined;
    if (area.trim()) {
      const r = parseRangeLoose(area.replace(/^=/, "").replace(/^.*!/, ""));
      if (!r) {
        ctl.alert("Yazdırma alanı başvurusu geçerli değil.");
        return null;
      }
      printArea = r;
    }
    // Yazdırma başlıkları: $1:$2 (satırlar) ve $A:$B (sütunlar); tek satır/sütun da yazılabilir
    const rowsM = /^\$?(\d+)(?::\$?(\d+))?$/.exec(tRows.trim().replace(/^.*!/, ""));
    const colsM = /^\$?([A-Za-z]{1,3})(?::\$?([A-Za-z]{1,3}))?$/.exec(tCols.trim().replace(/^.*!/, ""));
    if ((tRows.trim() && !rowsM) || (tCols.trim() && !colsM)) {
      ctl.alert("Yazdırma başlığı başvurusu geçerli değil. Satırlar için $1:$2, sütunlar için $A:$B biçimini kullanın.");
      return null;
    }
    const titleRows = rowsM ? { r1: Math.min(+rowsM[1], +(rowsM[2] ?? rowsM[1])) - 1, r2: Math.max(+rowsM[1], +(rowsM[2] ?? rowsM[1])) - 1 } : undefined;
    const c1 = colsM ? nameToCol(colsM[1].toUpperCase()) : 0;
    const c2 = colsM ? nameToCol((colsM[2] ?? colsM[1]).toUpperCase()) : 0;
    const titleCols = colsM ? { c1: Math.min(c1, c2), c2: Math.max(c1, c2) } : undefined;
    return {
      orientation,
      paper,
      fitToWidth: undefined,
      fitMode: fit === "no" ? undefined : fit,
      margins: {
        top: keep(m.top, p.margins.top),
        bottom: keep(m.bottom, p.margins.bottom),
        left: keep(m.left, p.margins.left),
        right: keep(m.right, p.margins.right),
        header: keep(m.header, p.margins.header),
        footer: keep(m.footer, p.margins.footer),
      },
      header: header || undefined,
      footer: footer || undefined,
      printArea,
      printGridlines: grid,
      printHeadings: heads,
      titleRows,
      titleCols,
      blackAndWhite: bw || undefined,
      draft: draft || undefined,
      pageOrder: order === "overThenDown" ? order : undefined,
    };
  };
  const ok = () => {
    const b = build();
    if (!b) return;
    setPage(ctl, b);
    onClose();
  };
  const print = () => {
    const b = build();
    if (!b) return;
    setPage(ctl, b);
    onClose();
    ctl.openDialog("printPreview");
  };
  const mField = (k: keyof PageSetup["margins"], label: string) => (
    <Field label={label}>
      <input type="text" value={m[k]} onChange={(e) => setM({ ...m, [k]: e.target.value })} style={{ width: 70 }} />
    </Field>
  );
  return (
    <Dialog
      open
      title="Sayfa Yapısı"
      onClose={onClose}
      width={520}
      onSubmit={ok}
      footer={
        <>
          <DialogButton onClick={print}>Yazdır...</DialogButton>
          <OkCancel onOk={ok} onCancel={onClose} />
        </>
      }
    >
      <DialogTabs
        tabs={[
          { id: "page", label: "Sayfa" },
          { id: "margins", label: "Kenar Boşlukları" },
          { id: "hf", label: "Üst Bilgi/Alt Bilgi" },
          { id: "sheet", label: "Çalışma Sayfası" },
        ]}
        active={tab}
        onChange={setTab}
      />
      <div className="tb-ps-body">
        {tab === "page" && (
          <>
            <FieldSet legend="Yönlendirme">
              <div className="tb-row-gap">
                <Radio name="or" value="portrait" current={orientation} onChange={setOrientation}>
                  Dikey
                </Radio>
                <Radio name="or" value="landscape" current={orientation} onChange={setOrientation}>
                  Yatay
                </Radio>
              </div>
            </FieldSet>
            <FieldSet legend="Ölçek">
              <Radio name="fit" value="no" current={fit} onChange={setFit}>
                Normal boyutun %100'ü
              </Radio>
              <Radio name="fit" value="sheet" current={fit} onChange={setFit}>
                Sayfayı tek sayfaya sığdır
              </Radio>
              <Radio name="fit" value="cols" current={fit} onChange={setFit}>
                Tüm sütunları bir sayfaya sığdır
              </Radio>
              <Radio name="fit" value="rows" current={fit} onChange={setFit}>
                Tüm satırları bir sayfaya sığdır
              </Radio>
            </FieldSet>
            <Field label="Kağıt boyutu:">
              <select value={paper} onChange={(e) => setPaper(e.target.value as PageSetup["paper"])}>
                {PAPERS.map((x) => (
                  <option key={x.id} value={x.id}>
                    {x.label}
                  </option>
                ))}
              </select>
            </Field>
          </>
        )}
        {tab === "margins" && (
          <div className="tb-ps-margins">
            <div />
            {mField("top", "Üst (cm):")}
            {mField("header", "Üst Bilgi (cm):")}
            {mField("left", "Sol (cm):")}
            <div className="tb-ps-paper" style={{ aspectRatio: orientation === "portrait" ? "0.7" : "1.4" }} />
            {mField("right", "Sağ (cm):")}
            <div />
            {mField("bottom", "Alt (cm):")}
            {mField("footer", "Alt Bilgi (cm):")}
          </div>
        )}
        {tab === "hf" && (
          <>
            <Field label="Üst bilgi:">
              <input type="text" value={header} onChange={(e) => setHeader(e.target.value)} placeholder="Ör. &[Sayfa] / &[Sayfalar] ya da serbest metin" />
            </Field>
            <Field label="Alt bilgi:">
              <input type="text" value={footer} onChange={(e) => setFooter(e.target.value)} placeholder="Sayfa &[Sayfa]" />
            </Field>
            <div className="tb-row-gap">
              <DialogButton onClick={() => setFooter("Sayfa &[Sayfa]")}>Sayfa 1</DialogButton>
              <DialogButton onClick={() => setFooter("Sayfa &[Sayfa] / &[Sayfalar]")}>Sayfa 1 / ?</DialogButton>
              <DialogButton onClick={() => setHeader("&[Sekme]")}>Sayfa adı</DialogButton>
              <DialogButton onClick={() => setHeader("&[Dosya]")}>Dosya adı</DialogButton>
            </div>
            <p className="tb-fc-note">Kodlar: &amp;[Sayfa] sayfa numarası, &amp;[Sayfalar] toplam sayfa, &amp;[Sekme] sayfa adı, &amp;[Dosya] dosya adı, &amp;[Tarih] bugünün tarihi.</p>
          </>
        )}
        {tab === "sheet" && (
          <>
            <Field label="Yazdırma alanı:">
              <input type="text" value={area} onChange={(e) => setArea(e.target.value)} placeholder="Ör. A1:F40" />
            </Field>
            <FieldSet legend="Yazdırma başlıkları">
              <Field label="Üstte yinelenecek satırlar:" inline>
                <input type="text" value={tRows} onChange={(e) => setTRows(e.target.value)} placeholder="$1:$1" aria-label="Üstte yinelenecek satırlar" />
              </Field>
              <Field label="Solda yinelenecek sütunlar:" inline>
                <input type="text" value={tCols} onChange={(e) => setTCols(e.target.value)} placeholder="$A:$A" aria-label="Solda yinelenecek sütunlar" />
              </Field>
            </FieldSet>
            <FieldSet legend="Yazdır">
              <div className="tb-ps-checks">
                <Check checked={grid} onChange={setGrid}>
                  Kılavuz çizgileri
                </Check>
                <Check checked={bw} onChange={setBw}>
                  Siyah beyaz
                </Check>
                <Check checked={draft} onChange={setDraft}>
                  Taslak kalitesi
                </Check>
                <Check checked={heads} onChange={setHeads}>
                  Satır ve sütun başlıkları
                </Check>
              </div>
            </FieldSet>
            <FieldSet legend="Sayfa sırası">
              <Radio name="order" value="downThenOver" current={order} onChange={setOrder}>
                Önce aşağı, sonra yana
              </Radio>
              <Radio name="order" value="overThenDown" current={order} onChange={setOrder}>
                Önce yana, sonra aşağı
              </Radio>
            </FieldSet>
          </>
        )}
      </div>
    </Dialog>
  );
}
