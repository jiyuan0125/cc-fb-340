#!/usr/bin/env python3
"""Archive closed tickets whose retention period has elapsed.

Reads the ticket master sheet (tickets.xlsx), moves every *closed* ticket
past its type-specific retention period into a SQLite cold store (cold.db,
table ``archived``), and writes a migration report (report.json).

Design notes
------------
* Retention days are counted as calendar days including the closing day
  ("含头当天"): a ticket closed on day C with retention N is covered on
  days C .. C+(N-1), so ``due_date = C + (N-1)`` and it becomes eligible
  once ``job_date > due_date``. Calendar arithmetic handles month/year
  boundaries and leap days naturally.
* Archiving is idempotent by tid: tids already present in cold.db are
  skipped, so a re-run only inserts newly expired tickets and never
  double-counts attachments (attach_cnt is frozen at first insert).
* Reopened tickets whose last_close_at is empty are still open: they are
  never archived (even when past due) and are reported under open_reopened.

Only the Python standard library is used; the xlsx is read directly as
zip + XML so no third-party package is required.
"""

from __future__ import annotations

import argparse
import json
import re
import sqlite3
import zipfile
from datetime import date, timedelta
from pathlib import Path
from xml.etree import ElementTree as ET

XLSX_NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
TID_RE = re.compile(r"^T\d+$")
KNOWN_TYPES = ("incident", "change", "request", "problem")
COLS = {
    "A": "tid",
    "B": "type",
    "C": "closed_at",
    "D": "reopen_cnt",
    "E": "last_close_at",
    "F": "attach_cnt",
}


# --------------------------------------------------------------------------
# xlsx reading
# --------------------------------------------------------------------------
def _cell_text(cell: ET.Element) -> str | None:
    """Return the raw text value of a worksheet cell.

    Handles inline strings (this workbook has no sharedStrings.xml),
    plain numbers and booleans. Numeric date serials are also supported
    defensively; the sample data stores dates as inline ISO strings.
    """
    if cell.get("t") == "inlineStr":
        inline = cell.find(f"{XLSX_NS}is")
        if inline is None:
            return None
        return "".join(node.text or "" for node in inline.iter(f"{XLSX_NS}t"))
    value = cell.find(f"{XLSX_NS}v")
    return value.text if value is not None else None


def read_sheet_rows(xlsx_path: Path) -> list[dict[str, str]]:
    """Parse sheet1 into a list of {column-letter: raw-text} dicts."""
    with zipfile.ZipFile(xlsx_path) as zf:
        xml_bytes = zf.read("xl/worksheets/sheet1.xml")
    root = ET.fromstring(xml_bytes)
    rows: list[dict[str, str]] = []
    for row in root.iter(f"{XLSX_NS}row"):
        values: dict[str, str] = {}
        for cell in row.findall(f"{XLSX_NS}c"):
            column = re.match(r"[A-Z]+", cell.get("r", "")).group()
            text = _cell_text(cell)
            if text is not None:
                values[column] = text
        rows.append(values)
    return rows


def parse_instructions(rows: list[dict[str, str]]) -> tuple[date, dict[str, int]]:
    """Extract job date and per-type retention days from the header row 2.

    Expected text:
      A2: "job date 2028-03-15"
      B2: "retention incident 365 change 180 request 90 problem 60"
    """
    instruction_row = rows[1]
    job_match = re.search(r"(\d{4}-\d{2}-\d{2})", instruction_row["A"])
    if not job_match:
        raise ValueError(f"cannot find job date in header row: {instruction_row!r}")
    job_date = date.fromisoformat(job_match.group(1))

    retention: dict[str, int] = {}
    pairs = re.findall(r"([A-Za-z]+)\s+(\d+)", instruction_row["B"])
    for type_name, days in pairs:
        type_name = type_name.lower()
        if type_name in KNOWN_TYPES:
            retention[type_name] = int(days)
    missing = set(KNOWN_TYPES) - retention.keys()
    if missing:
        raise ValueError(f"retention for types missing from header row: {sorted(missing)}")
    return job_date, retention


def parse_tickets(rows: list[dict[str, str]]) -> list[dict]:
    """Turn data rows into normalized ticket records.

    Rows before the data (header, instructions) and summary rows such as
    the trailing TOTAL row are skipped.
    """
    tickets: list[dict] = []
    for raw in rows[2:]:
        tid = raw.get("A", "")
        if not TID_RE.match(tid):
            continue
        ticket_type = raw.get("B")
        if ticket_type not in KNOWN_TYPES:
            continue
        tickets.append(
            {
                "tid": tid,
                "type": ticket_type,
                "closed_at": raw["C"],
                "reopen_cnt": int(raw.get("D", "0")),
                # Empty cell means never reopened -> last_close_at stays NULL.
                "last_close_at": raw.get("E"),
                "attach_cnt": int(raw.get("F", "0")),
            }
        )
    return tickets


# --------------------------------------------------------------------------
# Archiving
# --------------------------------------------------------------------------
def open_cold_store(db_path: Path) -> sqlite3.Connection:
    conn = sqlite3.connect(db_path)
    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS archived (
            tid           TEXT PRIMARY KEY,
            type          TEXT NOT NULL,
            closed_at     TEXT NOT NULL,
            reopen_cnt    INTEGER NOT NULL,
            last_close_at TEXT,
            attach_cnt    INTEGER NOT NULL,
            due_date      TEXT NOT NULL
        )
        """
    )
    # due_date must remain YYYY-MM-DD text, never a numeric affinity column.
    conn.commit()
    return conn


def due_date_for(closed_at: str, retention_days: int) -> str:
    """Calendar-day due date; the closing day counts as day 1."""
    closed = date.fromisoformat(closed_at)
    due = closed + timedelta(days=retention_days - 1)
    return due.isoformat()


def archive_tickets(
    conn: sqlite3.Connection,
    tickets: list[dict],
    retention: dict[str, int],
    job_date: date,
) -> list[dict]:
    """Insert closed, past-retention tickets not already in the cold store.

    Returns the list of newly archived records (this run only).
    """
    existing = {row[0] for row in conn.execute("SELECT tid FROM archived")}
    to_insert: list[dict] = []
    for ticket in tickets:
        if ticket["tid"] in existing:
            continue  # idempotent: frozen record, attach_cnt never re-added
        due = due_date_for(ticket["closed_at"], retention[ticket["type"]])
        # Only closed tickets are eligible. A reopened ticket with no
        # last_close_at is still open and must stay out of the cold store.
        is_closed = ticket["reopen_cnt"] == 0 or ticket["last_close_at"] is not None
        past_retention = job_date > date.fromisoformat(due)
        if is_closed and past_retention:
            record = {**ticket, "due_date": due}
            to_insert.append(record)

    conn.executemany(
        """
        INSERT INTO archived
            (tid, type, closed_at, reopen_cnt, last_close_at, attach_cnt, due_date)
        VALUES
            (:tid, :type, :closed_at, :reopen_cnt, :last_close_at, :attach_cnt, :due_date)
        """,
        to_insert,
    )
    conn.commit()
    return to_insert


# --------------------------------------------------------------------------
# Report
# --------------------------------------------------------------------------
def build_report(
    conn: sqlite3.Connection,
    tickets: list[dict],
    job_date: date,
) -> dict:
    # Migrated counts come from the cold store itself, so they stay correct
    # across repeated runs even if the workbook later changes.
    migrated_counts = dict(
        conn.execute("SELECT type, COUNT(*) FROM archived GROUP BY type")
    )
    total_by_type: dict[str, int] = {}
    for ticket in tickets:
        total_by_type[ticket["type"]] = total_by_type.get(ticket["type"], 0) + 1

    migrated: dict[str, int] = {}
    remaining: dict[str, int] = {}
    for ticket_type in KNOWN_TYPES:
        migrated_n = int(migrated_counts.get(ticket_type, 0))
        migrated[ticket_type] = migrated_n
        remaining[ticket_type] = total_by_type.get(ticket_type, 0) - migrated_n

    open_reopened_tids = sorted(
        ticket["tid"]
        for ticket in tickets
        if ticket["reopen_cnt"] > 0 and ticket["last_close_at"] is None
    )

    return {
        "job_date": job_date.isoformat(),
        "migrated": migrated,
        "remaining": remaining,
        "open_reopened": {
            "count": len(open_reopened_tids),
            "tids": open_reopened_tids,
        },
    }


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--xlsx",
        type=Path,
        default=Path("tickets.xlsx"),
        help="path to the ticket master workbook",
    )
    parser.add_argument(
        "--db",
        type=Path,
        default=Path("cold.db"),
        help="path to the SQLite cold store",
    )
    parser.add_argument(
        "--report",
        type=Path,
        default=Path("report.json"),
        help="path to write the migration report",
    )
    args = parser.parse_args()

    rows = read_sheet_rows(args.xlsx)
    job_date, retention = parse_instructions(rows)
    tickets = parse_tickets(rows)

    conn = open_cold_store(args.db)
    try:
        inserted = archive_tickets(conn, tickets, retention, job_date)
        report = build_report(conn, tickets, job_date)
    finally:
        conn.close()

    args.report.write_text(
        json.dumps(report, ensure_ascii=False, indent=2) + "\n", encoding="utf-8"
    )

    print(f"job_date: {report['job_date']}")
    print(f"newly archived this run: {len(inserted)}")
    print(f"cold store total: {sum(report['migrated'].values())}")
    print(f"remaining: {sum(report['remaining'].values())}")
    print(
        "open_reopened: "
        f"{report['open_reopened']['count']} {report['open_reopened']['tids']}"
    )


if __name__ == "__main__":
    main()
