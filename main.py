# -*- coding: utf-8 -*-
"""Sponge Factory Order-to-Delivery System - Entry Point."""
import sys
import os
import shutil
import tempfile

# ==========================================
# 暴力清理历史 comtypes 缓存，防止 Win7 报 WinError 123
# ==========================================
_cache_dir = os.path.join(tempfile.gettempdir(), 'comtypes_cache')
if os.path.exists(_cache_dir):
    try:
        shutil.rmtree(_cache_dir)
    except Exception:
        pass

_APP_DIR = os.path.dirname(os.path.abspath(__file__))
if _APP_DIR not in sys.path:
    sys.path.insert(0, _APP_DIR)

try:
    import Tkinter as tk
    import ttk
except ImportError:
    import tkinter as tk
    from tkinter import ttk


def main():
    from db.database import init_database, backup_database

    # Enable DPI awareness for crisp rendering on high-DPI displays
    try:
        import ctypes
        ctypes.windll.shcore.SetProcessDpiAwareness(2)  # Per-monitor DPI
    except Exception:
        pass

    # Install error logger (before everything else)
    from logic.logger import install_hook, install_thread_hook, log_info
    install_hook()
    install_thread_hook()
    log_info("Application started")

    # Init DB
    init_database()

    # Backup on startup
    try:
        backup_database()
    except Exception:
        pass

    # Seed default data if empty
    _seed_defaults()

    # Build UI
    root = tk.Tk()
    root.withdraw()  # hide until main window is ready

    from ui.main_window import MainWindow
    app = MainWindow(root)

    # Center window
    root.update_idletasks()
    w = root.winfo_width()
    h = root.winfo_height()
    sw = root.winfo_screenwidth()
    sh = root.winfo_screenheight()
    x = (sw - w) // 2
    y = (sh - h) // 2
    root.geometry("+{}+{}".format(
        max(0, x), max(0, y)))
    root.deiconify()

    root.mainloop()


def _seed_defaults():
    """Add default data if database is empty."""
    from db.database import get_connection
    conn = get_connection()

    row = conn.execute("SELECT COUNT(*) as cnt FROM color").fetchone()
    if row["cnt"] == 0:
        for c in [u"黑色", u"白色", u"灰色", u"肤色", u"卡其色",
                  u"浅灰", u"深灰", u"米白", u"浅紫", u"深紫"]:
            conn.execute("INSERT OR IGNORE INTO color (name) VALUES (?)", (c,))

    row = conn.execute("SELECT COUNT(*) as cnt FROM customer").fetchone()
    if row["cnt"] == 0:
        conn.execute(
            "INSERT INTO customer (name, prefix, contact_person, phone,"
            " has_model_number, has_item_code, excel_file_path)"
            " VALUES (?,?,?,?,?,?,?)",
            (u"宏润", "HR", u"张伟强", "18988539567", 1, 0,
             "D:/xxm/送货单/宏润.xlsx"))
        conn.execute(
            "INSERT INTO customer (name, prefix, has_model_number, has_item_code,"
            " excel_file_path)"
            " VALUES (?,?,?,?,?)",
            (u"雅兰", "YL", 0, 0, "D:/xxm/送货单/雅兰.xlsx"))

    row = conn.execute("SELECT COUNT(*) as cnt FROM product").fetchone()
    if row["cnt"] == 0:
        for mn, ic, pn in [
            ("8604", "", u"雪纺顺纹三经双层细斜纹织带"),
            ("", "HR8246", u"D4570/7MM双面30D高弹高密织带"),
            ("8136", "", u"DX3030/6MM锦纶双面细纹织带"),
            ("W2503", "", u"40支纯棉色织纱"),
            ("", "", u"3MM双面弹力带"),
        ]:
            conn.execute(
                "INSERT INTO product (model_number, item_code, product_name)"
                " VALUES (?,?,?)", (mn, ic, pn))

    conn.commit()
    conn.close()


def test():
    """Command-line test suite."""
    from db.database import init_database
    init_database()
    from db.database import get_connection

    print("=" * 60)
    print("Running Database Tests")
    print("=" * 60)

    conn = get_connection()
    conn.execute("DELETE FROM order_item")
    conn.execute("DELETE FROM orders")
    conn.execute("DELETE FROM product")
    conn.execute("DELETE FROM color")
    conn.execute("DELETE FROM customer")
    conn.commit()
    conn.close()

    from db.models import Customer, Product, Color, Order, OrderItem, Settings
    from logic.order_number import generate_delivery_number
    from logic.formula import parse_formula

    # Customer
    cid = Customer.create(name=u"宏润", prefix="HR")
    print("[OK] Customer created")

    products = Product.get_all()
    print("[OK] Products: {}".format(len(products)))

    colors = Color.get_all()
    print("[OK] Colors: {}".format(len(colors)))

    # Order
    oid = Order.create(cid, u"宏润5.4", "2026-05-04")
    print("[OK] Order created: 宏润5.4")

    # Formula
    for expr in ["=30*2", "=10+23+123", "=10*5+10", "=9*9+1000000"]:
        r = parse_formula(expr)
        print("[OK] {} => qty={}, notes={}".format(expr, r["quantity"], r["notes"]))

    # Delivery number
    dn = generate_delivery_number(cid, "2026-05-15")
    print("[OK] Delivery number: {}".format(dn))

    print("=" * 60)
    print("ALL TESTS PASSED")
    print("=" * 60)


if __name__ == "__main__":
    # If --test flag passed, run tests; otherwise run UI
    if "--test" in sys.argv:
        test()
    else:
        main()
