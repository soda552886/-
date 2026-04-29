import json
import os
import sqlite3
from contextlib import closing
from datetime import datetime
from typing import Iterable, List, Optional


DB_PATH = os.getenv("DB_PATH", "financial_reports.db")


def get_connection() -> sqlite3.Connection:
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn


def init_db() -> None:
    with closing(get_connection()) as conn:
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS import_batches (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                source_type TEXT NOT NULL,
                file_name TEXT NOT NULL,
                imported_at TEXT NOT NULL,
                row_count INTEGER NOT NULL
            )
            """
        )
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS payroll_records (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                batch_id INTEGER NOT NULL,
                source_type TEXT NOT NULL,
                sheet_name TEXT,
                employee_name TEXT,
                company_name TEXT,
                project_name TEXT,
                roc_year INTEGER,
                salary REAL,
                bonus REAL,
                welfare REAL,
                total_income REAL,
                note TEXT,
                raw_payload TEXT NOT NULL,
                FOREIGN KEY(batch_id) REFERENCES import_batches(id)
            )
            """
        )
        conn.commit()


def save_import_records(
    source_type: str,
    file_name: str,
    records: Iterable[dict],
) -> int:
    rows = list(records)
    if not rows:
        return 0

    with closing(get_connection()) as conn:
        cursor = conn.execute(
            """
            INSERT INTO import_batches (source_type, file_name, imported_at, row_count)
            VALUES (?, ?, ?, ?)
            """,
            (source_type, file_name, datetime.now().isoformat(timespec="seconds"), len(rows)),
        )
        batch_id = int(cursor.lastrowid)

        conn.executemany(
            """
            INSERT INTO payroll_records (
                batch_id,
                source_type,
                sheet_name,
                employee_name,
                company_name,
                project_name,
                roc_year,
                salary,
                bonus,
                welfare,
                total_income,
                note,
                raw_payload
            )
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            [
                (
                    batch_id,
                    source_type,
                    row.get("sheet_name"),
                    row.get("employee_name"),
                    row.get("company_name"),
                    row.get("project_name"),
                    row.get("roc_year"),
                    row.get("salary"),
                    row.get("bonus"),
                    row.get("welfare"),
                    row.get("total_income"),
                    row.get("note"),
                    json.dumps(row, ensure_ascii=False),
                )
                for row in rows
            ],
        )
        conn.commit()
    return len(rows)


def list_payroll_records(
    keyword: str = "",
    source_type: str = "全部",
    roc_year: Optional[int] = None,
    limit: int = 500,
) -> List[sqlite3.Row]:
    query = "SELECT * FROM payroll_records WHERE 1=1"
    params: List[object] = []

    if keyword.strip():
        query += " AND (employee_name LIKE ? OR company_name LIKE ? OR project_name LIKE ?)"
        like_kw = f"%{keyword.strip()}%"
        params.extend([like_kw, like_kw, like_kw])
    if source_type != "全部":
        query += " AND source_type = ?"
        params.append(source_type)
    if roc_year is not None:
        query += " AND roc_year = ?"
        params.append(roc_year)

    query += " ORDER BY id DESC LIMIT ?"
    params.append(limit)

    with closing(get_connection()) as conn:
        return conn.execute(query, params).fetchall()


def list_batches(limit: int = 100) -> List[sqlite3.Row]:
    with closing(get_connection()) as conn:
        return conn.execute(
            "SELECT * FROM import_batches ORDER BY id DESC LIMIT ?",
            (limit,),
        ).fetchall()


def delete_records_by_source(source_type: str) -> int:
    with closing(get_connection()) as conn:
        cursor = conn.execute(
            "DELETE FROM payroll_records WHERE source_type = ?",
            (source_type,),
        )
        conn.execute(
            "DELETE FROM import_batches WHERE source_type = ?",
            (source_type,),
        )
        conn.commit()
    return int(cursor.rowcount or 0)


def update_payroll_record(record_id: int, data: dict) -> None:
    with closing(get_connection()) as conn:
        conn.execute(
            """
            UPDATE payroll_records
            SET
                sheet_name = ?,
                employee_name = ?,
                company_name = ?,
                project_name = ?,
                roc_year = ?,
                salary = ?,
                bonus = ?,
                welfare = ?,
                total_income = ?,
                note = ?
            WHERE id = ?
            """,
            (
                data.get("sheet_name"),
                data.get("employee_name"),
                data.get("company_name"),
                data.get("project_name"),
                data.get("roc_year"),
                data.get("salary"),
                data.get("bonus"),
                data.get("welfare"),
                data.get("total_income"),
                data.get("note"),
                record_id,
            ),
        )
        conn.commit()


def delete_payroll_records(record_ids: List[int]) -> int:
    ids = [int(i) for i in record_ids if str(i).isdigit()]
    if not ids:
        return 0
    placeholders = ",".join(["?"] * len(ids))
    with closing(get_connection()) as conn:
        cursor = conn.execute(f"DELETE FROM payroll_records WHERE id IN ({placeholders})", ids)
        conn.commit()
    return int(cursor.rowcount or 0)
