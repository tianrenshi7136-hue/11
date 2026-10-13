"""Read OOXML directly, retaining local names, cached errors and print definitions."""
import json
import posixpath
import re
from pathlib import Path
from xml.etree import ElementTree as E
from zipfile import ZipFile

NS = {"s": "http://schemas.openxmlformats.org/spreadsheetml/2006/main", "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships"}
REL = "http://schemas.openxmlformats.org/package/2006/relationships"


def audit(path):
    with ZipFile(path) as z:
        workbook = E.fromstring(z.read("xl/workbook.xml"))
        relationships = {r.get("Id"): r.get("Target") for r in E.fromstring(z.read("xl/_rels/workbook.xml.rels"))}
        strings = []
        if "xl/sharedStrings.xml" in z.namelist():
            strings = ["".join(t.text or "" for t in n.iter("{" + NS["s"] + "}t")) for n in E.fromstring(z.read("xl/sharedStrings.xml"))]
        names = [{"name": n.get("name"), "scope": n.get("localSheetId"), "value": n.text} for n in workbook.findall("s:definedNames/s:definedName", NS)]
        sheets = []
        for s in workbook.findall("s:sheets/s:sheet", NS):
            target = relationships[s.get("{" + NS["r"] + "}id")]
            target = target.lstrip("/") if target.startswith("/") else posixpath.normpath("xl/" + target)
            sheet = E.fromstring(z.read(target))
            cells = []
            for cell in sheet.findall("s:sheetData/s:row/s:c", NS):
                formula = cell.find("s:f", NS)
                val = cell.findtext("s:v", default="", namespaces=NS)
                if cell.get("t") == "s" and val:
                    val = strings[int(val)]
                elif cell.get("t") == "inlineStr":
                    val = "".join(cell.find("s:is", NS).itertext())
                if formula is not None or val:
                    cells.append({"cell": cell.get("r"), "type": cell.get("t", "n"), "value": val, "formula": formula.text if formula is not None else None, "formula_attributes": formula.attrib if formula is not None else {}})
            sheets.append({"name": s.get("name"), "state": s.get("state", "visible"), "cells": cells, "formula_count": sum(bool(c["formula"] is not None or c["formula_attributes"]) for c in cells), "cached_errors": [c for c in cells if c["type"] == "e"], "broken_formulas": [c for c in cells if "#REF!" in (c["formula"] or "")], "external_formulas": [c for c in cells if re.search(r"\[\d+\]", c["formula"] or "")]})
        links = []
        malformed_parts = []
        for part in z.namelist():
            if part.startswith("xl/externalLinks/_rels/") and part.endswith(".rels"):
                try:
                    rels = E.fromstring(z.read(part))
                except E.ParseError as exc:
                    malformed_parts.append({"part": part, "error": str(exc)})
                    continue
                links.extend(r.get("Target") for r in rels if r.get("TargetMode") == "External" and r.get("Type", "").endswith("/externalLinkPath"))
        return {"file": path.name, "names": names, "external_targets": links, "malformed_parts": malformed_parts, "sheets": sheets}


def main():
    root = Path(__file__).resolve().parent.parent
    reports = [audit(p) for p in root.glob("*.xlsx")]
    (root / "outputs").mkdir(exist_ok=True)
    (root / "outputs" / "template-audit.json").write_text(json.dumps(reports, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps([{"file": r["file"], "external_target_count": len(r["external_targets"]), "broken_names": sum("#REF!" in (n["value"] or "") for n in r["names"]), "sheets": [{k: s[k] for k in ("name", "state", "formula_count")} | {"cached_errors": len(s["cached_errors"]), "broken_formulas": len(s["broken_formulas"]), "external_formulas": len(s["external_formulas"])} for s in r["sheets"]]} for r in reports], ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
