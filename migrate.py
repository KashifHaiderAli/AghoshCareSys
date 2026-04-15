#!/usr/bin/env python3
"""
SQL Server to SQLite Migration Script
Migrates data from Microsoft SQL Server 2005 (Aghosh) to an existing SQLite database.
Handles schema differences: renamed columns, missing columns, type conversions.
"""

import argparse
import sqlite3
import sys
import time
from datetime import datetime, date
from decimal import Decimal

import pyodbc

# ---------------------------------------------------------------------------
# Configuration: per-table column name mappings (SQL Server name → SQLite name)
# ---------------------------------------------------------------------------
COLUMN_MAPPINGS = {
    "tblAcademicProgressReports": {
        "APRID": "ReportID",
    },
    "tblChildren": {
        "GuardianAddress": "Guardianaddress",
    },
}

# Columns to skip during migration (exist in SQL Server but must be ignored)
SKIP_COLUMNS = {
    "tblAcademicProgressReports": {"Percentage", "CreatedDate"},
}

# Extra columns in SQLite that do not exist in SQL Server (column → default value)
EXTRA_SQLITE_COLUMNS = {
    "tblChildren": {"MonthlyAmount": None},
    "tblDonors": {"MonthlyCommitment": None},
}

# Per-table column type overrides: { table: { column: forced_source_type } }
# Used when SQLite expects a different type than what SQL Server reports
TYPE_OVERRIDES = {
    "tblSponsorships": {
        "Percentage": "int_to_text",
    },
}

# Tables in SQL Server that should NOT be migrated
SKIP_TABLES = {"tblPerformance"}

BATCH_SIZE = 500

# ---------------------------------------------------------------------------
# Connection helpers
# ---------------------------------------------------------------------------

def connect_sql_server(server: str, database: str, username: str, password: str) -> pyodbc.Connection:
    """Establish connection to SQL Server using pyodbc."""
    drivers = [d for d in pyodbc.drivers() if "SQL Server" in d or "FreeTDS" in d or "ODBC" in d]
    if not drivers:
        drivers = pyodbc.drivers()
    if not drivers:
        print("ERROR: No ODBC drivers found. Install FreeTDS or SQL Server ODBC driver.")
        sys.exit(1)

    driver = drivers[0]
    conn_str = (
        f"DRIVER={{{driver}}};"
        f"SERVER={server};"
        f"DATABASE={database};"
        f"UID={username};"
        f"PWD={password};"
        f"TDS_Version=8.0;"
    )
    try:
        conn = pyodbc.connect(conn_str, timeout=30)
        print(f"Connected to SQL Server: {server}/{database} (driver: {driver})")
        return conn
    except pyodbc.Error as e:
        print(f"ERROR connecting to SQL Server: {e}")
        sys.exit(1)


def connect_sqlite(db_path: str) -> sqlite3.Connection:
    """Open connection to existing SQLite database."""
    try:
        conn = sqlite3.connect(db_path)
        conn.execute("PRAGMA journal_mode=WAL")
        conn.execute("PRAGMA synchronous=NORMAL")
        print(f"Connected to SQLite: {db_path}")
        return conn
    except sqlite3.Error as e:
        print(f"ERROR connecting to SQLite: {e}")
        sys.exit(1)


# ---------------------------------------------------------------------------
# Schema introspection
# ---------------------------------------------------------------------------

def get_sql_server_tables(conn: pyodbc.Connection) -> list[str]:
    """Retrieve list of user tables from SQL Server."""
    cursor = conn.cursor()
    cursor.execute(
        "SELECT TABLE_NAME FROM INFORMATION_SCHEMA.TABLES "
        "WHERE TABLE_TYPE = 'BASE TABLE' ORDER BY TABLE_NAME"
    )
    tables = [row.TABLE_NAME for row in cursor.fetchall()]
    cursor.close()
    return tables


def get_sql_server_columns(conn: pyodbc.Connection, table: str) -> list[dict]:
    """Get column metadata from SQL Server for a given table."""
    cursor = conn.cursor()
    cursor.execute(
        "SELECT COLUMN_NAME, DATA_TYPE, CHARACTER_MAXIMUM_LENGTH, "
        "IS_NULLABLE, NUMERIC_PRECISION, NUMERIC_SCALE "
        "FROM INFORMATION_SCHEMA.COLUMNS "
        "WHERE TABLE_NAME = ? ORDER BY ORDINAL_POSITION",
        (table,),
    )
    columns = []
    for row in cursor.fetchall():
        columns.append({
            "name": row.COLUMN_NAME,
            "type": row.DATA_TYPE.lower(),
            "max_length": row.CHARACTER_MAXIMUM_LENGTH,
            "nullable": row.IS_NULLABLE == "YES",
            "precision": row.NUMERIC_PRECISION,
            "scale": row.NUMERIC_SCALE,
        })
    cursor.close()
    return columns


def get_sqlite_columns(conn: sqlite3.Connection, table: str) -> list[dict]:
    """Get column metadata from SQLite for a given table."""
    cursor = conn.cursor()
    cursor.execute(f'PRAGMA table_info("{table}")')
    columns = []
    for row in cursor.fetchall():
        columns.append({
            "name": row[1],
            "type": row[2].upper() if row[2] else "TEXT",
            "notnull": bool(row[3]),
            "default": row[4],
            "pk": bool(row[5]),
        })
    cursor.close()
    return columns


# ---------------------------------------------------------------------------
# Column mapping
# ---------------------------------------------------------------------------

def build_column_mapping(
    table: str,
    sql_columns: list[dict],
    sqlite_columns: list[dict],
) -> dict:
    """
    Build mapping: { sqlite_col_name: sql_server_col_name_or_None }
    Returns dict keyed by SQLite column name.
    If value is None the column gets a default value (not in SQL Server).
    """
    table_mappings = COLUMN_MAPPINGS.get(table, {})
    table_skips = SKIP_COLUMNS.get(table, set())
    table_extras = EXTRA_SQLITE_COLUMNS.get(table, {})
    table_type_overrides = TYPE_OVERRIDES.get(table, {})

    reverse_map = {v.lower(): k for k, v in table_mappings.items()}

    sql_col_names = {}
    sql_col_meta = {}
    for c in sql_columns:
        if c["name"] in table_skips:
            continue
        mapped_name = table_mappings.get(c["name"], c["name"])
        sql_col_names[mapped_name.lower()] = c["name"]
        sql_col_meta[c["name"]] = c

    sqlite_col_names = {c["name"].lower(): c for c in sqlite_columns}

    mapping = {}
    for sqlite_col in sqlite_columns:
        col_lower = sqlite_col["name"].lower()
        if col_lower in sql_col_names:
            source_type = sql_col_meta[sql_col_names[col_lower]]["type"]
            if sqlite_col["name"] in table_type_overrides:
                source_type = table_type_overrides[sqlite_col["name"]]
            mapping[sqlite_col["name"]] = {
                "source": sql_col_names[col_lower],
                "source_type": source_type,
                "target_type": sqlite_col["type"],
            }
        elif sqlite_col["name"] in table_extras:
            mapping[sqlite_col["name"]] = {
                "source": None,
                "default": table_extras[sqlite_col["name"]],
                "target_type": sqlite_col["type"],
            }
        else:
            mapping[sqlite_col["name"]] = {
                "source": None,
                "default": None,
                "target_type": sqlite_col["type"],
            }

    return mapping


# ---------------------------------------------------------------------------
# Type conversion
# ---------------------------------------------------------------------------

def convert_value(value, source_type: str, target_type: str):
    """Convert a single value from SQL Server type to SQLite-compatible type."""
    if value is None:
        return None

    target_upper = target_type.upper()

    # Explicit int-to-text override (e.g., tblSponsorships.Percentage)
    if source_type == "int_to_text":
        return str(int(value)) if value is not None else None

    # bit → INTEGER (0/1)
    if source_type in ("bit",):
        return 1 if value else 0

    # datetime → TEXT (ISO format)
    if source_type in ("datetime", "smalldatetime", "datetime2", "date"):
        if isinstance(value, (datetime, date)):
            return value.strftime("%Y-%m-%d %H:%M:%S")
        return str(value)

    # decimal/numeric → REAL or TEXT
    if source_type in ("decimal", "numeric", "money", "smallmoney"):
        if target_upper == "TEXT":
            return str(value)
        return float(value)

    # int/bigint → int or TEXT
    if source_type in ("int", "bigint", "smallint", "tinyint"):
        if target_upper == "TEXT":
            return str(value)
        return int(value)

    # float/real → REAL or TEXT
    if source_type in ("float", "real"):
        if target_upper == "TEXT":
            return str(value)
        return float(value)

    # string types
    if source_type in ("nvarchar", "varchar", "nchar", "char", "text", "ntext"):
        return str(value) if value is not None else None

    # Decimal objects from pyodbc
    if isinstance(value, Decimal):
        if target_upper == "TEXT":
            return str(value)
        return float(value)

    if isinstance(value, (datetime, date)):
        return value.strftime("%Y-%m-%d %H:%M:%S")

    return value


# ---------------------------------------------------------------------------
# Row transformation
# ---------------------------------------------------------------------------

def transform_row(mapping: dict, source_row: dict) -> tuple:
    """
    Transform a single source row (dict) into a tuple ordered by SQLite columns.
    """
    values = []
    for sqlite_col, info in mapping.items():
        if info["source"] is None:
            values.append(info.get("default"))
        else:
            raw = source_row.get(info["source"])
            values.append(convert_value(raw, info["source_type"], info["target_type"]))
    return tuple(values)


# ---------------------------------------------------------------------------
# Table migration
# ---------------------------------------------------------------------------

def migrate_table(
    table: str,
    sql_conn: pyodbc.Connection,
    sqlite_conn: sqlite3.Connection,
) -> int:
    """
    Migrate a single table from SQL Server to SQLite.
    Returns number of rows inserted.
    """
    sql_columns = get_sql_server_columns(sql_conn, table)
    sqlite_columns = get_sqlite_columns(sqlite_conn, table)

    if not sql_columns:
        raise ValueError(f"Table '{table}' has no columns in SQL Server")
    if not sqlite_columns:
        raise ValueError(f"Table '{table}' has no columns in SQLite")

    mapping = build_column_mapping(table, sql_columns, sqlite_columns)

    source_cols = [
        info["source"] for info in mapping.values() if info["source"] is not None
    ]
    if not source_cols:
        raise ValueError(f"No matching columns found for table '{table}'")

    select_sql = f'SELECT [{"], [".join(source_cols)}] FROM [{table}]'
    sqlite_col_names = list(mapping.keys())
    placeholders = ", ".join(["?"] * len(sqlite_col_names))
    insert_sql = (
        f'INSERT INTO "{table}" ('
        + ", ".join(f'"{c}"' for c in sqlite_col_names)
        + f") VALUES ({placeholders})"
    )

    sql_cursor = sql_conn.cursor()
    sql_cursor.execute(select_sql)

    col_names = [desc[0] for desc in sql_cursor.description]

    total_rows = 0
    batch = []

    for row in sql_cursor:
        row_dict = dict(zip(col_names, row))
        transformed = transform_row(mapping, row_dict)
        batch.append(transformed)

        if len(batch) >= BATCH_SIZE:
            sqlite_conn.executemany(insert_sql, batch)
            total_rows += len(batch)
            batch = []

    if batch:
        sqlite_conn.executemany(insert_sql, batch)
        total_rows += len(batch)

    sqlite_conn.commit()
    sql_cursor.close()
    return total_rows


# ---------------------------------------------------------------------------
# Main entry point
# ---------------------------------------------------------------------------

def main():
    parser = argparse.ArgumentParser(
        description="Migrate data from SQL Server (Aghosh) to SQLite"
    )
    parser.add_argument("--server", required=True, help="SQL Server hostname or IP")
    parser.add_argument("--database", default="Aghosh", help="SQL Server database name")
    parser.add_argument("--username", required=True, help="SQL Server username")
    parser.add_argument("--password", required=True, help="SQL Server password")
    parser.add_argument("--sqlite", required=True, help="Path to existing SQLite database")
    parser.add_argument("--tables", nargs="*", help="Specific tables to migrate (default: all)")
    args = parser.parse_args()

    print("=" * 60)
    print("SQL Server → SQLite Migration")
    print("=" * 60)
    start_time = time.time()

    sql_conn = connect_sql_server(args.server, args.database, args.username, args.password)
    sqlite_conn = connect_sqlite(args.sqlite)

    # Disable foreign keys during import
    sqlite_conn.execute("PRAGMA foreign_keys = OFF")

    # Get list of tables to migrate
    sql_tables = get_sql_server_tables(sql_conn)
    print(f"\nFound {len(sql_tables)} tables in SQL Server")

    # Get list of SQLite tables for filtering
    sqlite_cursor = sqlite_conn.cursor()
    sqlite_cursor.execute("SELECT name FROM sqlite_master WHERE type='table' ORDER BY name")
    sqlite_tables = {row[0] for row in sqlite_cursor.fetchall()}
    sqlite_cursor.close()
    print(f"Found {len(sqlite_tables)} tables in SQLite")

    if args.tables:
        tables_to_migrate = [t for t in args.tables if t in sql_tables]
    else:
        tables_to_migrate = sql_tables

    results = {"success": [], "failed": [], "skipped": [], "row_counts": {}}

    print(f"\n{'─' * 60}")
    print(f"{'Table':<35} {'Status':<12} {'Rows':>8}")
    print(f"{'─' * 60}")

    for table in tables_to_migrate:
        # Skip tables not in SQLite
        if table in SKIP_TABLES:
            results["skipped"].append(table)
            print(f"{table:<35} {'SKIPPED':<12} {'N/A':>8}")
            continue

        if table not in sqlite_tables:
            results["skipped"].append(table)
            print(f"{table:<35} {'SKIPPED':<12} {'N/A':>8}")
            continue

        try:
            row_count = migrate_table(table, sql_conn, sqlite_conn)
            results["success"].append(table)
            results["row_counts"][table] = row_count
            print(f"{table:<35} {'OK':<12} {row_count:>8}")
        except Exception as e:
            results["failed"].append(table)
            results["row_counts"][table] = 0
            print(f"{table:<35} {'FAILED':<12} {'0':>8}")
            print(f"  └─ Error: {e}")
            # Rollback any partial inserts for this table
            try:
                sqlite_conn.rollback()
            except Exception:
                pass

    # Re-enable foreign keys
    sqlite_conn.execute("PRAGMA foreign_keys = ON")

    elapsed = time.time() - start_time

    # Print summary
    print(f"\n{'=' * 60}")
    print("MIGRATION SUMMARY")
    print(f"{'=' * 60}")
    total_processed = len(results["success"]) + len(results["failed"])
    total_rows = sum(results["row_counts"].values())
    print(f"  Tables processed : {total_processed}")
    print(f"  Successful       : {len(results['success'])}")
    print(f"  Failed           : {len(results['failed'])}")
    print(f"  Skipped          : {len(results['skipped'])}")
    print(f"  Total rows       : {total_rows:,}")
    print(f"  Elapsed time     : {elapsed:.2f}s")

    if results["failed"]:
        print(f"\n  Failed tables: {', '.join(results['failed'])}")

    print(f"{'=' * 60}")

    sql_conn.close()
    sqlite_conn.close()

    sys.exit(1 if results["failed"] else 0)


if __name__ == "__main__":
    main()
