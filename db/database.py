# -*- coding: utf-8 -*-
"""SQLite database connection manager and schema initialization."""
import sqlite3
import os
import sys
import shutil


DB_FILENAME = "textile.db"


def get_data_dir():
    if getattr(sys, 'frozen', False):
        base = os.path.dirname(sys.executable)
    else:
        base = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    data_dir = os.path.join(base, "data")
    os.makedirs(data_dir, exist_ok=True)
    return data_dir


def get_db_path():
    return os.path.join(get_data_dir(), DB_FILENAME)


def get_connection():
    """Get a SQLite connection with WAL mode and foreign keys enabled."""
    db_path = get_db_path()
    conn = sqlite3.connect(db_path)
    conn.execute("PRAGMA journal_mode=WAL")
    conn.execute("PRAGMA foreign_keys=ON")
    conn.row_factory = sqlite3.Row
    return conn


def backup_database():
    """Create a timestamped backup of the database file."""
    db_path = get_db_path()
    if not os.path.exists(db_path):
        return None
    import time
    ts = time.strftime("%Y%m%d_%H%M%S")
    bak_path = db_path + ".{}.bak".format(ts)
    shutil.copy2(db_path, bak_path)
    return bak_path


def init_database():
    """Create all tables and seed default settings. Safe to call multiple times."""
    conn = get_connection()
    try:
        conn.executescript(SCHEMA)
        conn.commit()
    finally:
        conn.close()
    run_migrations()


def run_migrations():
    """Apply schema migrations safely (idempotent)."""
    conn = get_connection()
    try:
        # v1: add export_count to delivery_note
        try:
            conn.execute(
                "ALTER TABLE delivery_note ADD COLUMN"
                " export_count INTEGER NOT NULL DEFAULT 0")
            conn.commit()
        except sqlite3.OperationalError:
            pass

        # v2: remove UNIQUE from customer.prefix
        try:
            conn.execute("SAVEPOINT mig_prefix")
            # Check if any unique constraint exists on customer
            cu = conn.execute(
                "PRAGMA index_list(customer)")
            has_unique = any(r["unique"] and r["origin"] == "u"
                           for r in cu.fetchall())
            if has_unique:
                conn.executescript("""
                    PRAGMA foreign_keys = OFF;
                    CREATE TABLE customer_new (
                        id               INTEGER PRIMARY KEY AUTOINCREMENT,
                        name             TEXT    NOT NULL,
                        prefix           TEXT    NOT NULL,
                        contact_person   TEXT    DEFAULT '',
                        phone            TEXT    DEFAULT '',
                        email            TEXT    DEFAULT '',
                        address          TEXT    DEFAULT '',
                        notes            TEXT    DEFAULT '',
                        has_model_number INTEGER NOT NULL DEFAULT 1,
                        has_item_code    INTEGER NOT NULL DEFAULT 1,
                        excel_file_path  TEXT    DEFAULT '',
                        print_template   TEXT    DEFAULT '',
                        dn_headers       TEXT    DEFAULT '订单号,,款号',
                        created_at       TEXT    NOT NULL DEFAULT (datetime('now','localtime')),
                        updated_at       TEXT    NOT NULL DEFAULT (datetime('now','localtime'))
                    );
                    INSERT INTO customer_new SELECT
                        id, name, prefix, contact_person, phone, email,
                        address, notes, has_model_number, has_item_code,
                        excel_file_path, print_template,
                        '订单号,,款号',
                        created_at, updated_at
                    FROM customer;
                    DROP TABLE customer;
                    ALTER TABLE customer_new RENAME TO customer;
                    PRAGMA foreign_keys = ON;
                """)
                conn.commit()
        except Exception:
            conn.execute("ROLLBACK TO mig_prefix")

        # v3: add dn_headers column to customer
        try:
            conn.execute("SAVEPOINT mig_dn_headers")
            conn.execute(
                "ALTER TABLE customer ADD COLUMN dn_headers TEXT "
                "DEFAULT '订单号,,款号'")
            conn.commit()
        except Exception:
            conn.execute("ROLLBACK TO mig_dn_headers")

        # v4: multi-order delivery note support
        # Add customer_id to delivery_note, create junction table
        try:
            conn.execute("SAVEPOINT mig_multi_order")
            # Check if customer_id column already exists
            cols = [c[1] for c in conn.execute(
                "PRAGMA table_info(delivery_note)").fetchall()]
            if "customer_id" not in cols:
                conn.execute(
                    "ALTER TABLE delivery_note ADD COLUMN"
                    " customer_id INTEGER REFERENCES customer(id)")
                # Backfill customer_id from orders
                conn.execute(
                    "UPDATE delivery_note SET customer_id = ("
                    " SELECT o.customer_id FROM orders o"
                    " WHERE o.id = delivery_note.order_id)")
            # Create junction table (idempotent)
            conn.execute(
                "CREATE TABLE IF NOT EXISTS delivery_note_order ("
                " id INTEGER PRIMARY KEY AUTOINCREMENT,"
                " delivery_note_id INTEGER NOT NULL"
                "  REFERENCES delivery_note(id) ON DELETE CASCADE,"
                " order_id INTEGER NOT NULL"
                "  REFERENCES orders(id) ON DELETE CASCADE,"
                " UNIQUE(delivery_note_id, order_id))")
            conn.execute(
                "CREATE INDEX IF NOT EXISTS idx_dn_customer"
                " ON delivery_note(customer_id)")
            conn.execute(
                "CREATE INDEX IF NOT EXISTS idx_dn_order_dn"
                " ON delivery_note_order(delivery_note_id)")
            conn.execute(
                "CREATE INDEX IF NOT EXISTS idx_dn_order_ord"
                " ON delivery_note_order(order_id)")
            # Insert existing order relationships (idempotent)
            conn.execute(
                "INSERT OR IGNORE INTO delivery_note_order"
                " (delivery_note_id, order_id)"
                " SELECT id, order_id FROM delivery_note"
                " WHERE order_id IS NOT NULL")
            conn.commit()
        except Exception:
            conn.execute("ROLLBACK TO mig_multi_order")

        # v5: add order_number / mfg_number to order_item and delivery_note_item
        try:
            conn.execute("SAVEPOINT mig_order_number")
            oi_cols = [c[1] for c in conn.execute(
                "PRAGMA table_info(order_item)").fetchall()]
            if "order_number" not in oi_cols:
                conn.execute(
                    "ALTER TABLE order_item ADD COLUMN"
                    " order_number TEXT DEFAULT ''")
            if "mfg_number" not in oi_cols:
                conn.execute(
                    "ALTER TABLE order_item ADD COLUMN"
                    " mfg_number TEXT DEFAULT ''")
            dni_cols = [c[1] for c in conn.execute(
                "PRAGMA table_info(delivery_note_item)").fetchall()]
            if "mfg_number" not in dni_cols:
                conn.execute(
                    "ALTER TABLE delivery_note_item ADD COLUMN"
                    " mfg_number TEXT DEFAULT ''")
            conn.commit()
        except Exception:
            conn.execute("ROLLBACK TO mig_order_number")

        # v6: add customer_code to order_item and delivery_note_item
        try:
            conn.execute("SAVEPOINT mig_customer_code")
            oi_cols = [c[1] for c in conn.execute(
                "PRAGMA table_info(order_item)").fetchall()]
            if "customer_code" not in oi_cols:
                conn.execute(
                    "ALTER TABLE order_item ADD COLUMN"
                    " customer_code TEXT DEFAULT ''")
            dni_cols = [c[1] for c in conn.execute(
                "PRAGMA table_info(delivery_note_item)").fetchall()]
            if "customer_code" not in dni_cols:
                conn.execute(
                    "ALTER TABLE delivery_note_item ADD COLUMN"
                    " customer_code TEXT DEFAULT ''")
            conn.commit()
        except Exception:
            conn.execute("ROLLBACK TO mig_customer_code")
    finally:
        conn.close()


# ============================================================
# SCHEMA
# ============================================================
SCHEMA = """
CREATE TABLE IF NOT EXISTS customer (
    id               INTEGER PRIMARY KEY AUTOINCREMENT,
    name             TEXT    NOT NULL,
    prefix           TEXT    NOT NULL,
    contact_person   TEXT    DEFAULT '',
    phone            TEXT    DEFAULT '',
    email            TEXT    DEFAULT '',
    address          TEXT    DEFAULT '',
    notes            TEXT    DEFAULT '',
    has_model_number INTEGER NOT NULL DEFAULT 1,
    has_item_code    INTEGER NOT NULL DEFAULT 1,
    excel_file_path  TEXT    DEFAULT '',
    print_template   TEXT    DEFAULT '',
    dn_headers       TEXT    DEFAULT ',订单号,,款号',
    created_at       TEXT    NOT NULL DEFAULT (datetime('now','localtime')),
    updated_at       TEXT    NOT NULL DEFAULT (datetime('now','localtime'))
);

CREATE TABLE IF NOT EXISTS product (
    id            INTEGER PRIMARY KEY AUTOINCREMENT,
    model_number  TEXT    DEFAULT '',
    item_code     TEXT    DEFAULT '',
    product_name  TEXT    NOT NULL,
    specification TEXT    DEFAULT '',
    created_at    TEXT    NOT NULL DEFAULT (datetime('now','localtime'))
);
CREATE INDEX IF NOT EXISTS idx_product_name ON product(product_name);
CREATE INDEX IF NOT EXISTS idx_product_model ON product(model_number);
CREATE INDEX IF NOT EXISTS idx_product_code  ON product(item_code);

CREATE TABLE IF NOT EXISTS color (
    id   INTEGER PRIMARY KEY AUTOINCREMENT,
    name TEXT    NOT NULL UNIQUE
);

CREATE TABLE IF NOT EXISTS orders (
    id              INTEGER PRIMARY KEY AUTOINCREMENT,
    customer_id     INTEGER NOT NULL REFERENCES customer(id),
    display_name    TEXT    NOT NULL,
    order_date      TEXT    NOT NULL,
    delivery_number TEXT    DEFAULT '',
    delivery_date   TEXT    DEFAULT '',
    status          TEXT    NOT NULL DEFAULT 'draft',
    notes           TEXT    DEFAULT '',
    created_at      TEXT    NOT NULL DEFAULT (datetime('now','localtime')),
    updated_at      TEXT    NOT NULL DEFAULT (datetime('now','localtime'))
);
CREATE INDEX IF NOT EXISTS idx_orders_customer ON orders(customer_id);
CREATE INDEX IF NOT EXISTS idx_orders_status   ON orders(status);
CREATE INDEX IF NOT EXISTS idx_orders_date     ON orders(order_date);
CREATE INDEX IF NOT EXISTS idx_orders_dispname ON orders(display_name);

CREATE TABLE IF NOT EXISTS order_item (
    id               INTEGER PRIMARY KEY AUTOINCREMENT,
    order_id         INTEGER NOT NULL REFERENCES orders(id) ON DELETE CASCADE,
    product_id       INTEGER NOT NULL REFERENCES product(id),
    color_id         INTEGER NOT NULL REFERENCES color(id),
    model_number     TEXT    DEFAULT '',
    item_code        TEXT    DEFAULT '',
    customer_code    TEXT    DEFAULT '',
    product_name     TEXT    NOT NULL,
    color_name       TEXT    NOT NULL,
    quantity_formula TEXT    DEFAULT '',
    quantity         REAL    NOT NULL DEFAULT 0,
    print_count      INTEGER NOT NULL DEFAULT 1,
    is_printed       INTEGER NOT NULL DEFAULT 0,
    push_count       INTEGER NOT NULL DEFAULT 0,
    unit_price       TEXT    DEFAULT '',
    amount           REAL    DEFAULT 0,
    notes            TEXT    DEFAULT '',
    sort_order       INTEGER NOT NULL DEFAULT 0,
    created_at       TEXT    NOT NULL DEFAULT (datetime('now','localtime')),
    updated_at       TEXT    NOT NULL DEFAULT (datetime('now','localtime'))
);
CREATE INDEX IF NOT EXISTS idx_order_item_order ON order_item(order_id);

CREATE TABLE IF NOT EXISTS settings (
    key   TEXT PRIMARY KEY,
    value TEXT NOT NULL
);

CREATE TABLE IF NOT EXISTS delivery_note (
    id              INTEGER PRIMARY KEY AUTOINCREMENT,
    order_id        INTEGER REFERENCES orders(id) ON DELETE SET NULL,
    customer_id     INTEGER REFERENCES customer(id),
    delivery_number TEXT    NOT NULL UNIQUE,
    delivery_date   TEXT    DEFAULT '',
    status          TEXT    NOT NULL DEFAULT 'draft',
    notes           TEXT    DEFAULT '',
    export_count    INTEGER NOT NULL DEFAULT 0,
    created_at      TEXT    NOT NULL DEFAULT (datetime('now','localtime')),
    updated_at      TEXT    NOT NULL DEFAULT (datetime('now','localtime'))
);
CREATE INDEX IF NOT EXISTS idx_dn_order ON delivery_note(order_id);
CREATE INDEX IF NOT EXISTS idx_dn_number ON delivery_note(delivery_number);

CREATE TABLE IF NOT EXISTS delivery_note_order (
    id                INTEGER PRIMARY KEY AUTOINCREMENT,
    delivery_note_id  INTEGER NOT NULL REFERENCES delivery_note(id) ON DELETE CASCADE,
    order_id          INTEGER NOT NULL REFERENCES orders(id) ON DELETE CASCADE,
    UNIQUE(delivery_note_id, order_id)
);
CREATE INDEX IF NOT EXISTS idx_dn_order_dn ON delivery_note_order(delivery_note_id);
CREATE INDEX IF NOT EXISTS idx_dn_order_ord ON delivery_note_order(order_id);

CREATE TABLE IF NOT EXISTS delivery_note_item (
    id                INTEGER PRIMARY KEY AUTOINCREMENT,
    delivery_note_id  INTEGER NOT NULL REFERENCES delivery_note(id) ON DELETE CASCADE,
    order_item_id     INTEGER REFERENCES order_item(id) ON DELETE SET NULL,
    model_number      TEXT    DEFAULT '',
    item_code         TEXT    DEFAULT '',
    customer_code     TEXT    DEFAULT '',
    product_name      TEXT    NOT NULL,
    color_name        TEXT    NOT NULL,
    quantity_formula  TEXT    DEFAULT '',
    quantity          REAL    NOT NULL DEFAULT 0,
    unit_price        TEXT    DEFAULT '',
    amount            REAL    DEFAULT 0,
    notes             TEXT    DEFAULT '',
    sort_order        INTEGER NOT NULL DEFAULT 0,
    created_at        TEXT    NOT NULL DEFAULT (datetime('now','localtime')),
    updated_at        TEXT    NOT NULL DEFAULT (datetime('now','localtime'))
);
CREATE INDEX IF NOT EXISTS idx_dn_item_dn ON delivery_note_item(delivery_note_id);

INSERT OR IGNORE INTO settings (key, value) VALUES
    ('company_name', '佛山市和诚内衣辅料有限公司'),
    ('company_address', '地址：盐步联安路97号F座    电话：18988539567        传真：0757-81101784'),
    ('company_email', '邮箱：ZWQ197819@126.COM'),
    ('footer_text', '客户须知：如对产品质量有异议，请于收到货后三天内与本公司联系解决，逾期则视为合格。谢谢合作！'),
    ('maker_name', '制表人13829133080'),
    ('default_excel_folder', 'D:/xxm/送货单'),
    ('app_version', '1.0.0');
"""
