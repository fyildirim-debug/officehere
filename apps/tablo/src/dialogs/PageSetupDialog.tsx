/** Sayfa Yapısı iletişim kutusu: Sayfa, Kenar Boşlukları, Üst Bilgi/Alt Bilgi, Sayfa (yazdırma) sekmeleri. */
import { useState } from "react";
import { Dialog, DialogButton, DialogTabs, Field, FieldSet } from "@officehere/ui";
import type { Controller } from "../controller";
import { parseRangeLoose } from "../controller";
import { rectToA1 } from "../model/rect";
import type { PageSetup } from "../model/types";
import { Check, OkCancel, Radio, fmtTr, parseTrNumber } from "./common";

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
  const [fit, setFit] = useState(!!p.fitToWidth);
  const [m, setM] = useState(() => Object.fromEntries(Object.entries(p.margins).map(([k, v]) => [k, cm(v)])) as Record<keyof PageSetup["margins"], string>);
  const [header, setHeader] = useState(p.header ?? "");
  const [footer, setFooter] = useState(p.footer ?? "");
  const [area, setArea] = useState(p.printArea ? rectToA1(p.printArea) : "");
  const [grid, setGrid] = useState(p.printGridlines);
  const [heads, setHeads] = useState(p.printHeadings);
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
    return {
      orientation,
      paper,
      fitToWidth: fit || undefined,
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
          { id: "sheet", label: "Sayfa Ayarları" },
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
              <Radio name="fit" value="no" current={fit ? "yes" : "no"} onChange={() => setFit(false)}>
                Normal boyutun %100'ü
              </Radio>
              <Radio name="fit" value="yes" current={fit ? "yes" : "no"} onChange={() => setFit(true)}>
                Tüm sütunları bir sayfaya sığdır
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
            <FieldSet legend="Yazdır">
              <Check checked={grid} onChange={setGrid}>
                Kılavuz çizgileri
              </Check>
              <Check checked={heads} onChange={setHeads}>
                Satır ve sütun başlıkları
              </Check>
            </FieldSet>
          </>
        )}
      </div>
    </Dialog>
  );
}
