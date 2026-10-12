"""Read and write .xlsx with the standard library only (no openpyxl / pandas needed).

read_workbook(path)          -> {sheet name: [row, row, ...]}, a row is a list of str / int / float / None
write_workbook(path, sheets) -> sheets is a list of dicts:
    name, rows, header_rows (frozen + styled), freeze_cols, widths, autofilter (bool),
    validations [(cell range, "yes,no")]
A cell is a plain value, or S(value, style) with style one of STYLES.
"""
import math
import re
import zipfile
import xml.etree.ElementTree as ET
from xml.sax.saxutils import escape

M = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
STYLES = {"header": 1, "bold": 2, "good": 3, "bad": 4, "warn": 5, "text": 6, "input": 7, "title": 8, "wrap": 9,
          "dim": 10}
_BAD_XML = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f]")


def S(value, style):
    return (value, style)


def txt(v):
    """Cell value -> clean text ('' for empty, 235.0 -> '235')."""
    if v is None:
        return ""
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    return str(v).strip()


def fnum(v):
    """Cell value -> float or None."""
    try:
        f = float(v)
        return None if math.isnan(f) else f
    except (TypeError, ValueError):
        return None


# ---------------------------------------------------------------- reading
def _col_idx(ref):
    n = 0
    for ch in ref:
        if not ch.isalpha():
            break
        n = n * 26 + ord(ch.upper()) - 64
    return n - 1


def read_workbook(path):
    with zipfile.ZipFile(path) as z:
        names = set(z.namelist())
        shared = []
        if "xl/sharedStrings.xml" in names:
            for si in ET.fromstring(z.read("xl/sharedStrings.xml")).findall(f"{{{M}}}si"):
                # skip phonetic runs (rPh), keep plain and rich text
                parts = [t.text or "" for t in si.findall(f"{{{M}}}t")]
                parts += [t.text or "" for r in si.findall(f"{{{M}}}r") for t in r.findall(f"{{{M}}}t")]
                shared.append("".join(parts))
        rels = {r.get("Id"): r.get("Target") for r in ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))}
        out = {}
        for s in ET.fromstring(z.read("xl/workbook.xml")).find(f"{{{M}}}sheets"):
            target = rels[s.get(f"{{{R}}}id")].lstrip("/")
            if not target.startswith("xl/"):
                target = "xl/" + target
            rows = []
            for row in ET.fromstring(z.read(target)).iter(f"{{{M}}}row"):
                cells = []
                for c in row.findall(f"{{{M}}}c"):
                    t = c.get("t")
                    if t == "inlineStr":
                        v = "".join(x.text or "" for x in c.iter(f"{{{M}}}t"))
                    else:
                        e = c.find(f"{{{M}}}v")
                        v = None if e is None else e.text
                        if v is not None:
                            if t == "s":
                                v = shared[int(v)]
                            elif t in (None, "n"):
                                f = fnum(v)
                                v = v if f is None else int(f) if f.is_integer() else f
                            elif t == "e":
                                v = None
                    if isinstance(v, str) and v == "":
                        v = None
                    i = _col_idx(c.get("r")) if c.get("r") else len(cells)
                    cells.extend([None] * (i + 1 - len(cells)))
                    cells[i] = v
                r = int(row.get("r")) if row.get("r") else len(rows) + 1
                rows.extend([] for _ in range(r - len(rows)))
                rows[r - 1] = cells
            out[s.get("name")] = rows
        return out


def cell(row, i):
    return row[i] if i is not None and 0 <= i < len(row) else None


# ---------------------------------------------------------------- writing
def col_letter(i):
    s = ""
    i += 1
    while i:
        i, r = divmod(i - 1, 26)
        s = chr(65 + r) + s
    return s


_STYLES_XML = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    f'<styleSheet xmlns="{M}">'
    '<fonts count="4"><font><sz val="11"/><name val="Calibri"/></font>'
    '<font><b/><sz val="11"/><name val="Calibri"/></font>'
    '<font><b/><sz val="14"/><name val="Calibri"/></font>'
    '<font><sz val="11"/><color rgb="FF808080"/><name val="Calibri"/></font></fonts>'
    '<fills count="7"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill>'
    + "".join(f'<fill><patternFill patternType="solid"><fgColor rgb="FF{c}"/><bgColor indexed="64"/></patternFill></fill>'
              for c in ("D9D9D9", "C6EFCE", "FFC7CE", "FFEB9C", "DDEBF7")) +
    '</fills>'
    '<borders count="2"><border/><border><left style="thin"><color rgb="FF9BC2E6"/></left>'
    '<right style="thin"><color rgb="FF9BC2E6"/></right><top style="thin"><color rgb="FF9BC2E6"/></top>'
    '<bottom style="thin"><color rgb="FF9BC2E6"/></bottom></border></borders>'
    '<cellStyleXfs count="1"><xf/></cellStyleXfs>'
    '<cellXfs count="11"><xf/>'
    '<xf fontId="1" fillId="2" applyFont="1" applyFill="1" applyAlignment="1"><alignment wrapText="1" vertical="top"/></xf>'
    '<xf fontId="1" applyFont="1"/>'
    '<xf fillId="3" applyFill="1"/><xf fillId="4" applyFill="1"/><xf fillId="5" applyFill="1"/>'
    '<xf numFmtId="49" applyNumberFormat="1"/>'
    '<xf numFmtId="49" fillId="6" borderId="1" applyNumberFormat="1" applyFill="1" applyBorder="1"/>'
    '<xf fontId="2" applyFont="1"/>'
    '<xf applyAlignment="1"><alignment wrapText="1" vertical="top"/></xf>'
    '<xf fontId="3" applyFont="1"/>'
    '</cellXfs></styleSheet>')


def _sheet_xml(sh):
    rows, hdr, fc = sh["rows"], sh.get("header_rows", 0), sh.get("freeze_cols", 0)
    out = [f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="{M}">']
    if hdr or fc:
        pane = (f'xSplit="{fc}" ' if fc else "") + (f'ySplit="{hdr}" ' if hdr else "")
        active = "bottomRight" if hdr and fc else "bottomLeft" if hdr else "topRight"
        out.append(f'<sheetViews><sheetView workbookViewId="0"><pane {pane}topLeftCell="{col_letter(fc)}{hdr + 1}" '
                   f'activePane="{active}" state="frozen"/></sheetView></sheetViews>')
    if sh.get("widths"):
        out.append("<cols>" + "".join(f'<col min="{i + 1}" max="{i + 1}" width="{w}" customWidth="1"/>'
                                      for i, w in enumerate(sh["widths"])) + "</cols>")
    out.append("<sheetData>")
    width = 0
    for r, row in enumerate(rows, 1):
        cells = []
        width = max(width, len(row))
        for c, v in enumerate(row):
            style = STYLES["header"] if r <= hdr else 0
            if isinstance(v, tuple):
                v, style = v[0], STYLES[v[1]]
            s = f' s="{style}"' if style else ""
            ref = f"{col_letter(c)}{r}"
            if isinstance(v, bool):
                v = "yes" if v else "no"
            if v is None or v == "" or (isinstance(v, float) and (math.isnan(v) or math.isinf(v))):
                if style:
                    cells.append(f'<c r="{ref}"{s}/>')
            elif isinstance(v, (int, float)):
                cells.append(f'<c r="{ref}"{s}><v>{v}</v></c>')
            else:
                cells.append(f'<c r="{ref}" t="inlineStr"{s}><is><t xml:space="preserve">'
                             f'{escape(_BAD_XML.sub("", str(v)))}</t></is></c>')
        out.append(f'<row r="{r}">' + "".join(cells) + "</row>")
    out.append("</sheetData>")
    if sh.get("autofilter") and hdr and width:
        out.append(f'<autoFilter ref="A{hdr}:{col_letter(width - 1)}{max(len(rows), hdr)}"/>')
    if sh.get("validations"):
        out.append(f'<dataValidations count="{len(sh["validations"])}">' + "".join(
            f'<dataValidation type="list" allowBlank="1" showErrorMessage="1" sqref="{ref}">'
            f'<formula1>"{escape(opts)}"</formula1></dataValidation>' for ref, opts in sh["validations"])
            + "</dataValidations>")
    out.append("</worksheet>")
    return "".join(out)


def write_workbook(path, sheets):
    ct = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
          '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
          '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
          '<Default Extension="xml" ContentType="application/xml"/>'
          '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
          '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>']
    wb = [f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="{M}" xmlns:r="{R}"><sheets>']
    rels = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">']
    used = set()
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        for i, sh in enumerate(sheets, 1):
            name = re.sub(r"[\[\]:*?/\\]", "_", sh["name"])[:31]
            while name.lower() in used:
                name = name[:28] + f"_{i}"
            used.add(name.lower())
            z.writestr(f"xl/worksheets/sheet{i}.xml", _sheet_xml(sh))
            ct.append(f'<Override PartName="/xl/worksheets/sheet{i}.xml" '
                      'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>')
            wb.append(f'<sheet name="{escape(name)}" sheetId="{i}" r:id="rId{i}"/>')
            rels.append(f'<Relationship Id="rId{i}" Type="{R}/worksheet" Target="worksheets/sheet{i}.xml"/>')
        rels.append(f'<Relationship Id="rId{len(sheets) + 1}" Type="{R}/styles" Target="styles.xml"/></Relationships>')
        z.writestr("[Content_Types].xml", "".join(ct) + "</Types>")
        z.writestr("_rels/.rels", '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                   '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                   f'<Relationship Id="rId1" Type="{R}/officeDocument" Target="xl/workbook.xml"/></Relationships>')
        z.writestr("xl/workbook.xml", "".join(wb) + "</sheets></workbook>")
        z.writestr("xl/_rels/workbook.xml.rels", "".join(rels))
        z.writestr("xl/styles.xml", _STYLES_XML)
