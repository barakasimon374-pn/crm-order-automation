"""
db.py — SQLite persistence layer for the CRM Order Automation system.

Tables
------
runs
    One row per automation run (one Excel upload → one CRM order session).

excel_order_lines
    Every usable Excel line loaded during a run (after filtering).

posting_report_rows
    Every row written to the CRM Posting Report for a run.
"""

from __future__ import annotations

import sqlite3
import time
from pathlib import Path
from typing import Any

DB_PATH = Path(__file__).resolve().parent / "crm_automation.db"


# ============================================================
# CONNECTION
# ============================================================

def get_connection() -> sqlite3.Connection:
    conn = sqlite3.connect(str(DB_PATH))
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA journal_mode=WAL")
    conn.execute("PRAGMA foreign_keys=ON")
    return conn


# ============================================================
# SCHEMA
# ============================================================

def init_db() -> None:
    """Create tables if they do not already exist."""
    with get_connection() as conn:
        conn.executescript("""
            CREATE TABLE IF NOT EXISTS runs (
                id              INTEGER PRIMARY KEY AUTOINCREMENT,
                started_at      TEXT    NOT NULL,
                finished_at     TEXT,
                excel_filename  TEXT,
                status          TEXT    NOT NULL DEFAULT 'running',
                -- 'running' | 'success' | 'error'
                error_message   TEXT,
                total_documents INTEGER,
                total_lines     INTEGER,
                report_path     TEXT,
                customer_name   TEXT,
                created_at      TEXT    NOT NULL DEFAULT (datetime('now'))
            );

            CREATE TABLE IF NOT EXISTS excel_order_lines (
                id              INTEGER PRIMARY KEY AUTOINCREMENT,
                run_id          INTEGER NOT NULL REFERENCES runs(id),
                document_no     TEXT    NOT NULL,
                customer_no     TEXT    NOT NULL,
                item_no         TEXT    NOT NULL,
                description     TEXT    NOT NULL,
                quantity        REAL    NOT NULL,
                unit_price      REAL    NOT NULL,
                amount          REAL,
                unit_of_measure TEXT,
                location_code   TEXT,
                crm_search_text TEXT
            );

            CREATE TABLE IF NOT EXISTS posting_report_rows (
                id               INTEGER PRIMARY KEY AUTOINCREMENT,
                run_id           INTEGER NOT NULL REFERENCES runs(id),
                posting_date     TEXT,
                customer_no      TEXT,
                document_no      TEXT,
                excel_item_no    TEXT,
                excel_description TEXT,
                excel_qty        REAL,
                excel_unit_price REAL,
                crm_item_no      TEXT,
                crm_product      TEXT,
                crm_inventory    REAL,
                allocated_qty    REAL,
                crm_unit_price   REAL,
                inventory_status TEXT
            );

            CREATE INDEX IF NOT EXISTS idx_excel_order_lines_run
                ON excel_order_lines(run_id);

            CREATE INDEX IF NOT EXISTS idx_posting_report_rows_run
                ON posting_report_rows(run_id);

            CREATE INDEX IF NOT EXISTS idx_runs_started_at
                ON runs(started_at);
        """)


# ============================================================
# RUNS
# ============================================================

def create_run(
    excel_filename: str,
    customer_name: str,
    started_at: str | None = None,
) -> int:
    """Insert a new run row and return its id."""
    if started_at is None:
        started_at = time.strftime("%Y-%m-%d %H:%M:%S")

    with get_connection() as conn:
        cursor = conn.execute(
            """
            INSERT INTO runs (started_at, excel_filename, status, customer_name)
            VALUES (?, ?, 'running', ?)
            """,
            (started_at, excel_filename, customer_name),
        )
        return cursor.lastrowid


def finish_run(
    run_id: int,
    status: str,
    finished_at: str | None = None,
    error_message: str | None = None,
    total_documents: int | None = None,
    total_lines: int | None = None,
    report_path: str | None = None,
) -> None:
    """Update the run row after the automation completes or fails."""
    if finished_at is None:
        finished_at = time.strftime("%Y-%m-%d %H:%M:%S")

    with get_connection() as conn:
        conn.execute(
            """
            UPDATE runs
            SET finished_at     = ?,
                status          = ?,
                error_message   = ?,
                total_documents = ?,
                total_lines     = ?,
                report_path     = ?
            WHERE id = ?
            """,
            (
                finished_at,
                status,
                error_message,
                total_documents,
                total_lines,
                report_path,
                run_id,
            ),
        )


def get_all_runs(limit: int = 100) -> list[dict]:
    """Return the most recent runs, newest first."""
    with get_connection() as conn:
        rows = conn.execute(
            """
            SELECT id, started_at, finished_at, excel_filename,
                   status, error_message, total_documents,
                   total_lines, report_path, customer_name
            FROM runs
            ORDER BY id DESC
            LIMIT ?
            """,
            (limit,),
        ).fetchall()
        return [dict(row) for row in rows]


def get_run(run_id: int) -> dict | None:
    """Return a single run by id."""
    with get_connection() as conn:
        row = conn.execute(
            "SELECT * FROM runs WHERE id = ?", (run_id,)
        ).fetchone()
        return dict(row) if row else None


# ============================================================
# EXCEL ORDER LINES
# ============================================================

def save_excel_order_lines(run_id: int, excel_orders: dict) -> None:
    """
    Persist every Excel order line for a run.

    excel_orders is the dict returned by load_excel_orders():
        { document_no: { "document_no", "customer_no", "lines": [...] } }
    """
    rows = []
    for order in excel_orders.values():
        for line in order["lines"]:
            rows.append((
                run_id,
                line.get("document_no", ""),
                line.get("customer_no", ""),
                line.get("item_no", ""),
                line.get("description", ""),
                float(line.get("quantity", 0)),
                float(line.get("unit_price_excel", 0)),
                float(line["amount_excel"]) if line.get("amount_excel") is not None else None,
                line.get("unit_of_measure", ""),
                line.get("location_code", ""),
                line.get("crm_search_text", ""),
            ))

    with get_connection() as conn:
        conn.executemany(
            """
            INSERT INTO excel_order_lines
                (run_id, document_no, customer_no, item_no, description,
                 quantity, unit_price, amount, unit_of_measure,
                 location_code, crm_search_text)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            rows,
        )


def get_excel_order_lines(run_id: int) -> list[dict]:
    """Return all Excel order lines for a run."""
    with get_connection() as conn:
        rows = conn.execute(
            "SELECT * FROM excel_order_lines WHERE run_id = ? ORDER BY id",
            (run_id,),
        ).fetchall()
        return [dict(row) for row in rows]


# ============================================================
# POSTING REPORT ROWS
# ============================================================

def save_posting_report_rows(run_id: int, rows: list[dict]) -> None:
    """Persist all posting report rows collected during a run."""
    data = [
        (
            run_id,
            row.get("posting_date", ""),
            row.get("customer_no", ""),
            row.get("document_no", ""),
            row.get("excel_item_no", ""),
            row.get("excel_description", ""),
            float(row["excel_qty"]) if row.get("excel_qty") is not None else None,
            float(row["excel_unit_price"]) if row.get("excel_unit_price") is not None else None,
            row.get("crm_item_no", ""),
            row.get("crm_product", ""),
            float(row["crm_inventory"]) if row.get("crm_inventory") is not None else None,
            float(row["allocated_qty"]) if row.get("allocated_qty") is not None else None,
            float(row["crm_unit_price"]) if row.get("crm_unit_price") is not None else None,
            row.get("inventory_status", ""),
        )
        for row in rows
    ]

    with get_connection() as conn:
        conn.executemany(
            """
            INSERT INTO posting_report_rows
                (run_id, posting_date, customer_no, document_no,
                 excel_item_no, excel_description, excel_qty,
                 excel_unit_price, crm_item_no, crm_product,
                 crm_inventory, allocated_qty, crm_unit_price,
                 inventory_status)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            data,
        )


def get_posting_report_rows(run_id: int) -> list[dict]:
    """Return all posting report rows for a run."""
    with get_connection() as conn:
        rows = conn.execute(
            """
            SELECT * FROM posting_report_rows
            WHERE run_id = ?
            ORDER BY id
            """,
            (run_id,),
        ).fetchall()
        return [dict(row) for row in rows]


def get_posting_report_summary(run_id: int) -> dict:
    """Return inventory status counts for a run."""
    with get_connection() as conn:
        row = conn.execute(
            """
            SELECT
                COUNT(*)                                            AS total,
                SUM(CASE WHEN inventory_status = 'ENOUGH INVENTORY'        THEN 1 ELSE 0 END) AS enough,
                SUM(CASE WHEN inventory_status = 'INSUFFICIENT INVENTORY'  THEN 1 ELSE 0 END) AS insufficient,
                SUM(CASE WHEN inventory_status = 'ZERO INVENTORY'          THEN 1 ELSE 0 END) AS zero
            FROM posting_report_rows
            WHERE run_id = ?
            """,
            (run_id,),
        ).fetchone()
        return dict(row) if row else {}


# ============================================================
# INITIALISE ON IMPORT
# ============================================================

init_db()
