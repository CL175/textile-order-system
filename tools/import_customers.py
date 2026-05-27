# -*- coding: utf-8 -*-
"""Import customers from 图例\客户.txt into the database."""
import sys
import os
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from pypinyin import pinyin, Style


def get_prefix(name):
    """Generate prefix from pinyin initials. 新怡 → XY, 宏润 → HR"""
    initials = pinyin(name, style=Style.FIRST_LETTER)
    return "".join(i[0].upper() for i in initials if i[0].isalpha())


def import_customers(filepath):
    from db.database import init_database, get_connection

    init_database()

    # Read customer names
    names = []
    with open(filepath, "r", encoding="gbk") as f:
        for line in f:
            line = line.strip()
            if not line:
                continue
            # Split on space: first part = customer name, rest = contact/notes
            parts = line.split(None, 1)
            name = parts[0].strip()
            # Remove parenthetical notes from name
            if "(" in name:
                name = name.split("(")[0].strip()
            if "(" in name:
                name = name.split("(")[0].strip()
            if name and name not in names:
                names.append(name)

    # Insert into database
    conn = get_connection()
    added = 0
    skipped = 0

    try:
        for name in names:
            # Check if already exists
            existing = conn.execute(
                "SELECT COUNT(*) FROM customer WHERE name=?", (name,)
            ).fetchone()[0]
            if existing:
                skipped += 1
                continue

            prefix = get_prefix(name)
            conn.execute(
                "INSERT INTO customer (name, prefix, has_model_number, has_item_code)"
                " VALUES (?,?,?,?)",
                (name, prefix, 0, 0))
            print(u"  + {}  [{}]".format(name, prefix))
            added += 1

        total = conn.execute("SELECT COUNT(*) FROM customer").fetchone()[0]
        conn.commit()
    finally:
        conn.close()

    print(u"\nDone: {} added, {} skipped (already exist)".format(added, skipped))
    print(u"Total customers in DB: {}".format(total))


if __name__ == "__main__":
    path = sys.argv[1] if len(sys.argv) > 1 else r"D:\xxm\图例\客户.txt"
    import_customers(path)
