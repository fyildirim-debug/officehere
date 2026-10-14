#!/usr/bin/env python3
"""LibreOffice'i başvuru hesap motoru olarak kullanır: her çalışma kitabını açar, calculateAll()
yapar ve tüm formül hücrelerinin sonucunu JSON verir ({anahtar: {sayfa: {"satır,sütun": değer}}}). Tur, tablo motorunun sonuçlarını bununla
karşılaştırır ("formül(LO) 423/425").

Ön koşul: LibreOffice UNO dinleyicisi (LO_PORT, varsayılan 2002).
Kullanım: python3 scripts/derlem/gercek_lo.py <anahtar=dosya ...> > gercek_lo.json
  (anahtar verilmezse dosya yolu anahtar olur)
"""
import json
import os
import sys

import uno
from com.sun.star.beans import PropertyValue

ERR = {532: "#DIV/0!", 32767: "#N/A", 503: "#NUM!", 519: "#VALUE!", 524: "#REF!", 525: "#NAME?", 502: "#NUM!", 504: "#VALUE!", 511: "#VALUE!"}
PORT = int(os.environ.get("LO_PORT", "2002"))
local = uno.getComponentContext()
ctx = local.ServiceManager.createInstanceWithContext("com.sun.star.bridge.UnoUrlResolver", local).resolve(f"uno:socket,host=127.0.0.1,port={PORT};urp;StarOffice.ComponentContext")
desktop = ctx.ServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)


def P(n, v):
    p = PropertyValue()
    p.Name = n
    p.Value = v
    return p


out = {}
for arg in sys.argv[1:]:
    key, _, path = arg.partition("=") if "=" in arg else (arg, "", arg)
    try:
        doc = desktop.loadComponentFromURL(uno.systemPathToFileUrl(os.path.abspath(path)), "_blank", 0, (P("Hidden", True), P("ReadOnly", True)))
        if doc is None:
            out[key] = {"__hata__": "LO açamadı"}
            continue
        doc.calculateAll()
        res = {}
        for si in range(doc.Sheets.getCount()):
            sh = doc.Sheets.getByIndex(si)
            cells = {}
            ranges = sh.queryContentCells(16)  # FORMULA
            for ra in ranges.getRangeAddresses():
                for r in range(ra.StartRow, ra.EndRow + 1):
                    for c in range(ra.StartColumn, ra.EndColumn + 1):
                        cell = sh.getCellByPosition(c, r)
                        if cell.getType().value != "FORMULA":
                            continue
                        e = cell.getError()
                        if e:
                            v = ERR.get(e, f"#ERR{e}")
                        else:
                            v = cell.getString() if cell.FormulaResultType2 == 2 else cell.getValue()
                        cells[f"{r},{c}"] = v
            res[sh.Name] = cells
        out[key] = res
        doc.close(True)
    except Exception as ex:  # noqa: BLE001 - LO hatası rapora yazılır
        out[key] = {"__hata__": str(ex)}
print(json.dumps(out, ensure_ascii=False, indent=0, sort_keys=True))
