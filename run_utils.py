#!/usr/bin/env python3
import datetime as dt
import html
import json
import sys
import zipfile
from pathlib import Path
from typing import Any


class StepLogger:
    def __init__(self, log_path: Path | None = None) -> None:
        self.log_path = log_path
        if self.log_path:
            self.log_path.write_text("", encoding="utf-8")

    def emit(self, stage: str, status: str, word: str = "", detail: str = "") -> None:
        timestamp = dt.datetime.now().isoformat(timespec="seconds")
        event = {
            "timestamp": timestamp,
            "word": word,
            "stage": stage,
            "status": status,
            "detail": detail,
        }
        line = f"[{timestamp}] {stage} {status}"
        if word:
            line += f" word={word}"
        if detail:
            line += f" detail={detail}"
        print(line, file=sys.stderr, flush=True)
        if self.log_path:
            with self.log_path.open("a", encoding="utf-8") as handle:
                handle.write(json.dumps(event, ensure_ascii=False) + "\n")


def xlsx_col_name(index: int) -> str:
    name = ""
    while index:
        index, remainder = divmod(index - 1, 26)
        name = chr(65 + remainder) + name
    return name


def write_xlsx(path: Path, sheet_name: str, headers: list[str], rows: list[dict[str, Any]]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    matrix = [headers]
    for row in rows:
        matrix.append([str(row.get(header, "")) for header in headers])

    sheet_rows = []
    for row_index, row in enumerate(matrix, start=1):
        cells = []
        for col_index, value in enumerate(row, start=1):
            cell = f"{xlsx_col_name(col_index)}{row_index}"
            cells.append(f'<c r="{cell}" t="inlineStr"><is><t>{html.escape(value)}</t></is></c>')
        sheet_rows.append(f'<row r="{row_index}">{"".join(cells)}</row>')

    dimension = f"A1:{xlsx_col_name(len(headers))}{max(len(matrix), 1)}"
    widths = "".join(
        f'<col min="{i}" max="{i}" width="{min(max(len(header) + 4, 12), 60)}" customWidth="1"/>'
        for i, header in enumerate(headers, start=1)
    )
    worksheet = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <dimension ref="{dimension}"/>
  <cols>{widths}</cols>
  <sheetData>{"".join(sheet_rows)}</sheetData>
</worksheet>'''
    workbook = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="{html.escape(sheet_name)}" sheetId="1" r:id="rId1"/></sheets>
</workbook>'''
    workbook_rels = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>'''
    root_rels = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>'''
    content_types = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
</Types>'''

    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as archive:
        archive.writestr("[Content_Types].xml", content_types)
        archive.writestr("_rels/.rels", root_rels)
        archive.writestr("xl/workbook.xml", workbook)
        archive.writestr("xl/_rels/workbook.xml.rels", workbook_rels)
        archive.writestr("xl/worksheets/sheet1.xml", worksheet)


def write_text(path: Path, lines: list[str]) -> None:
    path.write_text("\n".join(lines) + ("\n" if lines else ""), encoding="utf-8")
