"""Boundary / idempotency tests for archive.py using synthetic workbooks."""
import json
import sqlite3
import sys
import tempfile
from datetime import date, timedelta
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))
sys.path.insert(0, str(Path(__file__).parent.parent / "src"))

import archive  # noqa: E402
from make_xlsx import build_xlsx  # noqa: E402

RET = {"incident": 365, "change": 180, "request": 90, "problem": 60}


def run_case(tmp: Path, xlsx_name: str, job_date: str, tickets, db_name: str):
    xlsx = tmp / xlsx_name
    db = tmp / db_name
    report = tmp / (db_name + ".json")
    build_xlsx(xlsx, job_date, RET, tickets)
    old_argv = sys.argv
    sys.argv = [
        "archive.py",
        "--xlsx", str(xlsx),
        "--db", str(db),
        "--report", str(report),
    ]
    try:
        archive.main()
    finally:
        sys.argv = old_argv
    conn = sqlite3.connect(db)
    rows = {r[0]: r for r in conn.execute("SELECT * FROM archived")}
    conn.close()
    return rows, json.loads(report.read_text())


def main() -> None:
    with tempfile.TemporaryDirectory() as td:
        tmp = Path(td)
        tickets = [
            # due exactly on job date -> must NOT be archived
            # problem 60: closed 2028-01-16 + 59 = due 2028-03-15 (2028 is leap)
            {"tid": "T1", "type": "problem", "closed_at": "2028-01-16",
             "reopen_cnt": 0, "last_close_at": None, "attach_cnt": 1},
            # one day past due -> archived (head included)
            {"tid": "T2", "type": "problem", "closed_at": "2028-01-15",
             "reopen_cnt": 0, "last_close_at": None, "attach_cnt": 2},
            # request 90 over year boundary: 2027-12-16 + 89 = 2028-03-14 -> due < job
            {"tid": "T3", "type": "request", "closed_at": "2027-12-16",
             "reopen_cnt": 0, "last_close_at": None, "attach_cnt": 3},
            # leap-year window: closed 2024-02-28 (leap), problem 60 -> due 2024-04-27
            {"tid": "T4", "type": "problem", "closed_at": "2024-02-28",
             "reopen_cnt": 0, "last_close_at": None, "attach_cnt": 4},
            # reopened but closed again, past due -> archived, last_close_at kept
            {"tid": "T5", "type": "change", "closed_at": "2027-08-01",
             "reopen_cnt": 2, "last_close_at": "2027-09-01", "attach_cnt": 5},
            # reopened, still open, past due -> excluded, listed in open_reopened
            {"tid": "T6", "type": "change", "closed_at": "2027-01-01",
             "reopen_cnt": 1, "last_close_at": None, "attach_cnt": 6},
        ]
        rows, report = run_case(tmp, "a.xlsx", "2028-03-15", tickets, "cold1.db")

        assert set(rows) == {"T2", "T3", "T4", "T5"}, set(rows)
        assert "T1" not in rows, "ticket due on job date must stay out"
        assert "T6" not in rows, "open reopened ticket must stay out"

        # head-inclusive calendar math
        assert rows["T2"][6] == "2028-03-14", rows["T2"]
        assert rows["T3"][6] == "2028-03-14", rows["T3"]
        # 2024 leap day included in the 60-day window
        assert rows["T4"][6] == "2024-04-27", rows["T4"]
        # reopened-closed: last_close_at preserved, reopen_cnt preserved
        t5 = rows["T5"]
        assert t5[3] == 2 and t5[4] == "2027-09-01", t5
        # due_date stored as TEXT
        assert all(isinstance(r[6], str) for r in rows.values())

        assert report["job_date"] == "2028-03-15"
        assert report["migrated"] == {
            "incident": 0, "change": 1, "request": 1, "problem": 2
        }
        assert report["remaining"] == {
            "incident": 0, "change": 1, "request": 0, "problem": 1
        }
        assert report["open_reopened"] == {"count": 1, "tids": ["T6"]}

        # ---- rerun: nothing new, then advance job date so T1 becomes due ----
        rows2, report2 = run_case(tmp, "a.xlsx", "2028-03-15", tickets, "cold1.db")
        assert set(rows2) == {"T2", "T3", "T4", "T5"}
        assert rows2["T2"][5] == 2, "attach_cnt must stay frozen"

        # same db, job date one day later; T1 now exactly one day past due
        rows3, report3 = run_case(tmp, "b.xlsx", "2028-03-16", tickets, "cold1.db")
        assert "T1" in rows3 and len(rows3) == 5
        assert rows3["T1"][6] == "2028-03-15"  # 2028-01-16 + 59
        # migrated count read from cold store: problem now 3
        assert report3["migrated"]["problem"] == 3
        assert report3["remaining"]["problem"] == 0

        # mutate T1 attach_cnt in a new workbook and rerun: frozen value kept
        tickets2 = [dict(t, attach_cnt=99) if t["tid"] == "T1" else t for t in tickets]
        rows4, _ = run_case(tmp, "c.xlsx", "2028-03-16", tickets2, "cold1.db")
        assert rows4["T1"][5] == 1, "re-migration must not accumulate/freeze-break attach_cnt"
        assert rows4["T2"][5] == 2

    print("ALL BOUNDARY TESTS PASSED")


if __name__ == "__main__":
    main()
