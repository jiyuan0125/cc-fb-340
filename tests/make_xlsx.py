"""Build minimal xlsx workbooks for archive.py boundary tests.

Reuses the container parts of the real workbook and swaps sheet1.xml;
cells use inline strings and numeric values exactly like the sample.
"""
import zipfile
from pathlib import Path

TEMPLATE = Path("/workspace/data/tickets.xlsx")

HEADERS = ["tid", "type", "closed_at", "reopen_cnt", "last_close_at", "attach_cnt"]
COL_LETTERS = ["A", "B", "C", "D", "E", "F"]


def _esc(text: str) -> str:
    return (
        text.replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
    )


def _str_cell(ref: str, text: str) -> str:
    return f'<c r="{ref}" t="inlineStr"><is><t>{_esc(text)}</t></is></c>'


def _num_cell(ref: str, value: int) -> str:
    return f'<c r="{ref}" t="n"><v>{value}</v></c>'


def build_xlsx(
    out_path: Path,
    job_date: str,
    retention: dict[str, int],
    tickets: list[dict],
) -> None:
    ret_text = "retention " + " ".join(
        f"{name} {days}" for name, days in retention.items()
    )
    rows_xml = []
    # row 1: headers
    rows_xml.append(
        '<row r="1">'
        + "".join(_str_cell(f"{c}1", h) for c, h in zip(COL_LETTERS, HEADERS))
        + "</row>"
    )
    # row 2: instructions
    rows_xml.append(
        '<row r="2">'
        + _str_cell("A2", f"job date {job_date}")
        + _str_cell("B2", ret_text)
        + "</row>"
    )
    # data rows
    for idx, ticket in enumerate(tickets, start=3):
        cells = [
            _str_cell(f"A{idx}", ticket["tid"]),
            _str_cell(f"B{idx}", ticket["type"]),
            _str_cell(f"C{idx}", ticket["closed_at"]),
            _num_cell(f"D{idx}", ticket.get("reopen_cnt", 0)),
        ]
        if ticket.get("last_close_at") is not None:
            cells.append(_str_cell(f"E{idx}", ticket["last_close_at"]))
        cells.append(_num_cell(f"F{idx}", ticket.get("attach_cnt", 0)))
        rows_xml.append(f'<row r="{idx}">' + "".join(cells) + "</row>")

    sheet_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        '<sheetData>' + "".join(rows_xml) + "</sheetData></worksheet>"
    )

    with zipfile.ZipFile(TEMPLATE) as src, zipfile.ZipFile(
        out_path, "w", zipfile.ZIP_DEFLATED
    ) as dst:
        for item in src.infolist():
            if item.filename == "xl/worksheets/sheet1.xml":
                dst.writestr(item, sheet_xml)
            else:
                dst.writestr(item, src.read(item.filename))
