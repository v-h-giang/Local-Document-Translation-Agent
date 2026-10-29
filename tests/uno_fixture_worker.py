"""Create and inspect synthetic workbooks using native LibreOffice UNO."""
import json
import sys
import traceback
from pathlib import Path

import uno
from com.sun.star.beans import PropertyValue


def prop(name, value):
    p = PropertyValue()
    p.Name, p.Value = name, value
    return p


def connect(port):
    local = uno.getComponentContext()
    resolver = local.ServiceManager.createInstanceWithContext(
        "com.sun.star.bridge.UnoUrlResolver", local)
    uri = f"uno:socket,host=127.0.0.1,port={port};urp;StarOffice.ComponentContext"
    import time
    for _ in range(100):
        try:
            remote = resolver.resolve(uri)
            return remote.ServiceManager.createInstanceWithContext(
                "com.sun.star.frame.Desktop", remote)
        except Exception:
            time.sleep(.15)
    raise RuntimeError("UNO fixture worker could not connect")


def set_charstyle(cursor, style):
    if style.get("bold"):
        cursor.CharWeight = 150.0
    if style.get("italic"):
        cursor.CharPosture = uno.Enum("com.sun.star.awt.FontSlant", "ITALIC")
    if style.get("strike"):
        cursor.CharStrikeout = 1
    if style.get("color") is not None:
        cursor.CharColor = int(style["color"])
    if style.get("height") is not None:
        cursor.CharHeight = float(style["height"])


def set_rich(target, spec):
    runs = spec.get("runs")
    text = "".join(x["text"] for x in runs) if runs else spec.get("text", "")
    target.setString(text)
    cursor = target.createTextCursor()
    cursor.gotoStart(False)
    if runs:
        for run in runs:
            n = len(run["text"])
            if n:
                cursor.goRight(n, True)
                set_charstyle(cursor, run)
                cursor.collapseToEnd()
    elif spec.get("style"):
        cursor.gotoEnd(True)
        set_charstyle(cursor, spec["style"])


def make_shape(doc, sheet, spec):
    shp = doc.createInstance("com.sun.star.drawing.TextShape")
    shp.Name = spec["name"]
    pos = uno.createUnoStruct("com.sun.star.awt.Point")
    pos.X, pos.Y = 1200, 1200
    shp.setPosition(pos)
    size = uno.createUnoStruct("com.sun.star.awt.Size")
    size.Width, size.Height = int(spec.get("width", 15000)), int(spec.get("height", 5000))
    shp.setSize(size)
    sheet.getDrawPage().add(shp)
    set_rich(shp, spec)


def make_case(desktop, spec, folder):
    doc = desktop.loadComponentFromURL("private:factory/scalc", "_blank", 0, (prop("Hidden", True),))
    try:
        first = doc.Sheets.getByIndex(0)
        first.Name = spec["sheets"][0]["name"]
        for index, info in enumerate(spec["sheets"]):
            if index:
                doc.Sheets.insertNewByName(info["name"], index)
            sheet = doc.Sheets.getByName(info["name"])
            for addr, value in info.get("cells", {}).items():
                c = sheet.getCellRangeByName(addr)
                if "formula" in value:
                    c.setFormula(value["formula"])
                elif "number" in value:
                    c.setValue(float(value["number"]))
                else:
                    set_rich(c, value)
            for merge in info.get("merged", []):
                sheet.getCellRangeByName(merge).merge(True)
            for obj in info.get("shapes", []):
                make_shape(doc, sheet, obj)
        destination = (folder / (spec["id"] + ".xlsx")).resolve()
        doc.storeAsURL(uno.systemPathToFileUrl(str(destination)), (
            prop("FilterName", "Calc MS Excel 2007 XML"), prop("Overwrite", True)))
    finally:
        doc.close(True)


def simplify_style(fragment):
    def get(obj, name):
        try:
            return getattr(obj, name)
        except Exception:
            return None
    posture = get(fragment, "CharPosture")
    return {
        "bold": bool((get(fragment, "CharWeight") or 0) > 120),
        "strike": bool(get(fragment, "CharStrikeout")),
        "italic": bool(getattr(posture, "value", "") in ("ITALIC", "OBLIQUE")
                       or str(posture).endswith("ITALIC")),
        "color": get(fragment, "CharColor"),
        "height": get(fragment, "CharHeight"),
    }


def get_runs(target):
    results = []
    try:
        paragraphs = target.createEnumeration()
        while paragraphs.hasMoreElements():
            para = paragraphs.nextElement()
            parts = para.createEnumeration()
            while parts.hasMoreElements():
                fragment = parts.nextElement()
                text = fragment.getString()
                if text:
                    results.append({"text": text, "style": simplify_style(fragment)})
    except Exception as exc:
        return [{"error": str(exc)}]
    return results


def inspect_case(desktop, spec, folder):
    path = (folder / (spec["id"] + ".xlsx")).resolve()
    if not path.exists():
        return {"error": "missing output file"}
    doc = desktop.loadComponentFromURL(uno.systemPathToFileUrl(str(path)), "_blank", 0, (
        prop("Hidden", True), prop("ReadOnly", True),
        prop("UpdateDocMode", 0), prop("MacroExecutionMode", 0)))
    if not doc:
        return {"error": "could not open workbook"}
    try:
        result = {"cells": {}, "shapes": {}, "merged": {}}
        for si in spec["sheets"]:
            sn = si["name"]
            sh = doc.Sheets.getByName(sn)
            result["merged"][sn] = {r: bool(sh.getCellRangeByName(r).IsMerged)
                                     for r in si.get("merged", [])}
            for addr in si.get("cells", {}):
                target = sh.getCellRangeByName(addr)
                result["cells"][f"{sn}!{addr}"] = {
                    "text": target.getString(), "formula": target.getFormula(),
                    "number": target.getValue(), "runs": get_runs(target),
                }
            page = sh.getDrawPage()
            for i in range(page.getCount()):
                shp = page.getByIndex(i)
                name = str(getattr(shp, "Name", ""))
                size = shp.getSize()
                result["shapes"][f"{sn}!{name}"] = {
                    "text": shp.getString(), "width": int(size.Width),
                    "height": int(size.Height), "runs": get_runs(shp),
                }
        return result
    finally:
        doc.close(True)


def main():
    port = int(sys.argv[1])
    action = sys.argv[2]
    manifest = Path(sys.argv[3])
    root = Path(sys.argv[4])
    cases = json.loads(manifest.read_text(encoding="utf-8"))
    root.mkdir(parents=True, exist_ok=True)
    desktop = connect(port)
    if action == "create":
        for spec in cases:
            make_case(desktop, spec, root)
        print(json.dumps({"created": len(cases)}, ensure_ascii=True))
    elif action == "inspect":
        print(json.dumps({x["id"]: inspect_case(desktop, x, root)
                          for x in cases}, ensure_ascii=True))
    else:
        raise ValueError("Unknown action " + action)


if __name__ == "__main__":
    try:
        main()
    except Exception:
        traceback.print_exc()
        sys.exit(1)
