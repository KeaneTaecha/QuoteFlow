"""Rename the tb_modifier column to base_modifier in the products table."""

import argparse
import sqlite3
from pathlib import Path


def column_exists(cursor, table, column):
    cursor.execute(f"PRAGMA table_info({table})")
    return any(row[1] == column for row in cursor.fetchall())


def rename_column(db_path: Path):
    if not db_path.exists():
        raise FileNotFoundError(f"Database not found at {db_path}")

    conn = sqlite3.connect(db_path)
    cursor = conn.cursor()

    try:
        has_tb = column_exists(cursor, "products", "tb_modifier")
        has_base = column_exists(cursor, "products", "base_modifier")

        if not has_tb:
            print("tb_modifier column not found; nothing to rename.")
            return

        if has_base:
            print("base_modifier column already exists; no action taken.")
            return

        cursor.execute("ALTER TABLE products RENAME COLUMN tb_modifier TO base_modifier")
        conn.commit()
        print("Successfully renamed tb_modifier to base_modifier.")
    finally:
        conn.close()


def main():
    parser = argparse.ArgumentParser(description="Rename tb_modifier column to base_modifier.")
    parser.add_argument(
        "--db",
        default="/Users/kunanonttaechaaukarkaul/Documents/University/Komfortflow/QuoteFlow/prices.db",
        help="Path to the SQLite database (default: %(default)s)",
    )
    args = parser.parse_args()
    rename_column(Path(args.db))


if __name__ == "__main__":
    main()

