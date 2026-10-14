import { describe, expect, it } from "vitest";
import { getExt, orderViolation, PRESENTATION_ORDER, SETTINGS_ORDER, topLevelNames, upsertExt, WORKSHEET_ORDER, XmlPart } from "../../src/opc/xmlPart";

describe("XmlPart (core)", () => {
  it("settings.xml yaması: sahip öğeler değişir, bilinmeyenler ve uzantılar yerinde kalır, şema sırası korunur", () => {
    const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:settings xmlns:w="w" xmlns:m="m" xmlns:w14="w14"><w:zoom w:percent="120"/><w:defaultTabStop w:val="720"/><w:compat><w:compatSetting w:name="compatibilityMode" w:val="15"/></w:compat><w:rsids><w:rsidRoot w:val="00A1"/></w:rsids><m:mathPr><m:mathFont m:val="Cambria Math"/></m:mathPr><w14:docId w14:val="1A2B"/></w:settings>`;
    const p = new XmlPart(xml, SETTINGS_ORDER);
    p.set("defaultTabStop", `<w:defaultTabStop w:val="708"/>`);
    p.set("evenAndOddHeaders", `<w:evenAndOddHeaders/>`);
    p.set("zoom", null);
    const out = p.toString();
    expect(out.startsWith(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:settings`)).toBe(true);
    expect(topLevelNames(out)).toEqual(["defaultTabStop", "evenAndOddHeaders", "compat", "rsids", "mathPr", "docId"]);
    expect(orderViolation(topLevelNames(out), SETTINGS_ORDER)).toBeNull();
    expect(out).toContain(`<w:compatSetting w:name="compatibilityMode" w:val="15"/>`);
    expect(out).toContain(`w:val="708"`);
  });

  it("presentation.xml: embeddedFontLst sıraya göre yerleşir, extLst girdileri URI ile korunur", () => {
    const xml = `<p:presentation xmlns:p="p"><p:sldMasterIdLst/><p:sldIdLst/><p:sldSz cx="1" cy="2"/><p:notesSz cx="1" cy="2"/><p:defaultTextStyle/><p:extLst><p:ext uri="{EFAFB233-063F-42B5-8137-9DF3F51BA10A}"><p15:sldGuideLst/></p:ext></p:extLst></p:presentation>`;
    const p = new XmlPart(xml, PRESENTATION_ORDER);
    p.set("embeddedFontLst", `<p:embeddedFontLst/>`);
    expect(topLevelNames(p.toString())).toEqual(["sldMasterIdLst", "sldIdLst", "sldSz", "notesSz", "embeddedFontLst", "defaultTextStyle", "extLst"]);
    expect(getExt(p, "{efafb233-063f-42b5-8137-9df3f51ba10a}")).toContain("sldGuideLst");
    upsertExt(p, "{X}", `<p:ext uri="{X}"/>`, []);
    expect(p.toString()).toContain(`<p:ext uri="{X}"/>`);
    expect(p.toString()).toContain("sldGuideLst");
  });

  it("çalışma sayfası sırası (Tablo'dan taşındı, davranış aynı)", () => {
    const p = new XmlPart(`<worksheet><sheetData/><pageMargins/></worksheet>`, WORKSHEET_ORDER);
    p.set("mergeCells", `<mergeCells count="1"><mergeCell ref="A1:B2"/></mergeCells>`);
    p.set("dimension", `<dimension ref="A1"/>`);
    expect(topLevelNames(p.toString())).toEqual(["dimension", "sheetData", "mergeCells", "pageMargins"]);
  });
});
