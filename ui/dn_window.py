# -*- coding: utf-8 -*-
"""Delivery note window with Excel-style grid."""
try:
    import Tkinter as tk
    import ttk
    import tkMessageBox as messagebox
except ImportError:
    import tkinter as tk
    from tkinter import ttk
    from tkinter import messagebox

import json
from db.models import (Customer, Product, Color, Order, OrderItem,
                       DeliveryNote, DNItem, Settings)
from logic.formula import parse_formula
from logic.order_number import generate_delivery_number
from ui.excel_grid import ExcelGrid
from ui.widgets import DatePicker


# ---------------------------------------------------------------------------
# WPS COM helper – direct print (no pywinauto keystroke simulation)
# ---------------------------------------------------------------------------
def _wps_print_last_sheet(filepath):
    """
    【动态克隆版】WPS 打印核心
    不写死任何排版！读取 Sheet1(客户原始模板) 的排版属性，在内存中临时克隆到最后一页。
    打印后不保存，绝对不影响你导出的 Excel 源文件格式！
    """
    import os
    import time
    import win32com.client

    abs_path = os.path.abspath(filepath)
    app = None
    wb = None
    created = False

    for prog_id in ("ET.Application", "Ket.Application"):
        try:
            app = win32com.client.GetActiveObject(prog_id)
            break
        except Exception:
            pass
        try:
            # 使用 DispatchEx 强制启动独立新进程，更安全
            app = win32com.client.DispatchEx(prog_id)
            created = True
            break
        except Exception:
            continue

    if app is None:
        raise Exception(u"无法连接 WPS COM 组件，请确认 WPS 已安装或修复环境。")

    try:
        if created:
            app.Visible = False
            # 绝对禁止 WPS 弹任何确认框卡死进程
            app.DisplayAlerts = False

        # 【重点保障】：以只读模式打开，防止文件被锁死或被意外修改
        wb = app.Workbooks.Open(abs_path, ReadOnly=True)
        ws = wb.Worksheets(wb.Worksheets.Count)

        # ========================================================
        # 核心修复区：在内存中临时克隆 Sheet1 (客户模板) 的排版给当前页
        # 放弃会报错的 PageSetup.Copy，改用逐项赋值，WPS 完美支持这种写法
        # ========================================================
        try:
            src_ws = wb.Worksheets(1)
            if src_ws is not ws:
                src_ps = src_ws.PageSetup
                dst_ps = ws.PageSetup

                # 1. 克隆物理属性 (方向、纸张大小)
                dst_ps.Orientation = src_ps.Orientation
                dst_ps.PaperSize = src_ps.PaperSize

                # 2. 克隆边距
                dst_ps.LeftMargin = src_ps.LeftMargin
                dst_ps.RightMargin = src_ps.RightMargin
                dst_ps.TopMargin = src_ps.TopMargin
                dst_ps.BottomMargin = src_ps.BottomMargin

                # 3. 克隆缩放逻辑 (最关键：判断客户模板是按百分比缩放，还是强制缩放页宽)
                if src_ps.Zoom:
                    dst_ps.Zoom = src_ps.Zoom
                else:
                    dst_ps.Zoom = False
                    dst_ps.FitToPagesWide = src_ps.FitToPagesWide
                    dst_ps.FitToPagesTall = src_ps.FitToPagesTall

                # 顺手把页眉页脚清空，防止干扰打印机纸张边缘感应
                dst_ps.LeftHeader = ""
                dst_ps.CenterHeader = ""
                dst_ps.RightHeader = ""
                dst_ps.LeftFooter = ""
                dst_ps.CenterFooter = ""
                dst_ps.RightFooter = ""
        except Exception as e:
            print("克隆排版失败: " + str(e))
        # ========================================================

        ws.Activate()
        time.sleep(0.5)

        # 触发打印：极其保守的无参数调用
        try:
            ws.PrintOut(Copies=1)
        except Exception:
            try:
                ws.PrintOut()
            except Exception as e:
                raise Exception(u"WPS 拒绝执行底层打印: " + str(e))

        # 给网络打印机留足 4.5 秒的 Spooler 传输时间
        time.sleep(4.5)

    finally:
        # 【关键清理】：SaveChanges=False 保证内存里的排版修改随着关闭一起销毁，绝不污染你原本的导出文件
        if wb:
            try:
                wb.Close(SaveChanges=False)
            except Exception:
                pass
        if created and app:
            try:
                app.Quit()
            except Exception:
                pass


# Column indices
C_CUST_CODE, C_ORDER_NO, C_MFG_NO, C_MODEL, C_NAME, C_COLOR = range(6)
C_QTY, C_PRICE, C_AMOUNT, C_NOTES = range(6, 10)


class DeliveryNoteWindow(tk.Toplevel):

    COLUMNS = [u"客户号", u"订单号", u"制单号", u"款号",
               u"品名", u"颜色",
               u"数量", u"单价", u"金额", u"备注"]
    COL_WIDTHS = [80, 110, 110, 80, 260, 90, 90, 90, 90, 130]
    SEARCH_COLS = {C_NAME: "product", C_COLOR: "color"}

    def __init__(self, parent, order_id=None, dn_id=None,
                 on_save_callback=None, read_only=False):
        tk.Toplevel.__init__(self, parent)
        self.order_id = order_id
        self.dn_id = dn_id
        self.on_save_callback = on_save_callback
        self.read_only = read_only
        self._dirty = False

        if self.dn_id:
            dn = DeliveryNote.get_by_id(dn_id)
            if dn:
                self._customer = (Customer.get_by_id(dn["customer_id"])
                                  if dn.get("customer_id") else None)
                self._source_orders = dn.get("source_orders", [])
                # Backward compat: order_id might be None for multi-order DNs
                if dn.get("order_id"):
                    self._order = Order.get_by_id(dn["order_id"])
                elif self._source_orders:
                    self._order = Order.get_by_id(
                        self._source_orders[0]["id"])
                else:
                    self._order = None
            else:
                self._order = None
                self._customer = None
                self._source_orders = []
        elif self.order_id:
            self._order = Order.get_by_id(order_id)
            self._customer = (Customer.get_by_id(
                self._order["customer_id"]) if self._order else None)
            self._source_orders = []
        else:
            self._order = None
            self._customer = None
            self._source_orders = []

        self._customers = Customer.get_all()

        self.title(u"送货单")
        sw = self.winfo_screenwidth()
        sh = self.winfo_screenheight()
        w = int(sw * 0.85)
        h = int(sh * 0.7)
        self.geometry("{}x{}".format(w, h))
        self.minsize(920, 520)
        self.transient(parent)

        self._build()
        if self.dn_id:
            self.after(100, self._load_dn)
            self.after(200, lambda: setattr(self, '_dirty', False))
        elif self.order_id:
            self.after(100, self._load_from_order)
            self.after(200, lambda: setattr(self, '_dirty', False))

        self.bind("<Control-s>", lambda e: self._save())
        self.bind("<Control-S>", lambda e: self._save())
        self.protocol("WM_DELETE_WINDOW", self._on_close)

    def _safe_ui(self, fn):
        def _run():
            try:
                fn()
            except tk.TclError:
                pass
        self.after(0, _run)

    def _build(self):
        # ---- Big Title ----
        title_bar = tk.Frame(self, bg="#1a5276", height=36)
        title_bar.pack(fill=tk.X)
        title_bar.pack_propagate(False)
        st = u"（只读）" if self.read_only else u"（编辑模式）"
        tk.Label(title_bar, text=u"  送货单明细  {}".format(st),
                 font=("Microsoft YaHei", 14, "bold"),
                 bg="#1a5276", fg="white", anchor=tk.W).pack(
                     side=tk.LEFT, fill=tk.BOTH, expand=True)

        # ---- Header ----
        hdr = tk.Frame(self)
        hdr.pack(fill=tk.X, padx=10, pady=(8, 2))
        r1 = tk.Frame(hdr)
        r1.pack(fill=tk.X)

        tk.Label(r1, text=u"客户:", font=("Microsoft YaHei", 9)).pack(side=tk.LEFT)
        self.cust_lbl = tk.Label(r1, text="", width=14, anchor=tk.W,
                                 font=("Microsoft YaHei", 9, "bold"))
        self.cust_lbl.pack(side=tk.LEFT, padx=(2, 14))

        tk.Label(r1, text=u"送货单号:", font=("Microsoft YaHei", 9)).pack(side=tk.LEFT)
        self.dn_var = tk.StringVar()
        self.dn_e = ttk.Entry(r1, textvariable=self.dn_var, width=16)
        if self.read_only:
            self.dn_e.config(state="readonly")
        self.dn_e.pack(side=tk.LEFT, padx=(2, 14))

        tk.Label(r1, text=u"送货日期:", font=("Microsoft YaHei", 9)).pack(side=tk.LEFT)
        self.dlv_var = tk.StringVar(value=self._today())
        self.dlv_var.trace_add("write", lambda *a: self._on_date_var_changed())
        df = tk.Frame(r1)
        df.pack(side=tk.LEFT, padx=(2, 0))
        self.dlv_e = ttk.Entry(df, textvariable=self.dlv_var, width=12)
        if self.read_only:
            self.dlv_e.config(state="readonly")
        self.dlv_e.pack(side=tk.LEFT)
        if not self.read_only:
            ttk.Button(df, text=u"📅", width=2,
                       command=lambda: DatePicker(self, self.dlv_e)).pack(side=tk.LEFT)

        r2 = tk.Frame(hdr)
        r2.pack(fill=tk.X)
        self.info_lbl = tk.Label(r2, text="", fg="gray",
                                 font=("Microsoft YaHei", 8))
        self.info_lbl.pack(side=tk.LEFT)

        # ---- Grid ----
        # Apply customer-specific column names for the first 4 columns
        columns = list(self.COLUMNS)
        if self._customer and self._customer.get("dn_headers"):
            custom = self._customer["dn_headers"].split(",")
            if len(custom) < 4:
                custom = [""] + custom  # old 3-field → prepend 客户号
            for i in range(min(4, len(custom))):
                if custom[i].strip():
                    columns[i] = custom[i].strip()
        widths = list(self.COL_WIDTHS)
        saved = self._load_grid_widths()
        if saved and len(saved) == len(widths):
            widths = saved
        self.grid = ExcelGrid(
            self, columns=columns, col_widths=widths,
            search_cols={
                C_NAME: lambda r, c: self._search_product(r, c),
                C_COLOR: lambda r, c: self._search_color(r, c),
            })
        self.grid.on_cell_changed = self._on_grid_cell_changed
        self.grid.on_resize_done = self._save_grid_widths
        if self.read_only:
            self.grid.read_only = True
            self.grid.on_edit_blocked = self._on_readonly_edit
        self.grid.pack(fill=tk.BOTH, expand=True, padx=10, pady=2)

        # ---- Footer (always create, show/hide) ----
        self._ft = tk.Frame(self)
        self._ft.pack(fill=tk.X, padx=10, pady=(2, 8))

        self._btn_add = ttk.Button(self._ft, text=u"添加行",
                                   command=lambda: self.grid.add_row())
        self._btn_del = ttk.Button(self._ft, text=u"删除行",
                                   command=self._del_rows)
        self.done_var = tk.IntVar(value=0)
        self._chk_done = ttk.Checkbutton(
            self._ft, text=u"完单", variable=self.done_var,
            command=self._on_done_toggle)
        self._btn_cancel = ttk.Button(self._ft, text=u"取消",
                                      command=self._on_close)
        self._btn_save = ttk.Button(self._ft, text=u"保存",
                                    command=self._save)
        self._btn_export = ttk.Button(self._ft, text=u"导出",
                                      command=self._export)
        self._btn_print = ttk.Button(self._ft, text=u"打印",
                                     command=self._print)
        self._btn_export_ro = ttk.Button(self._ft, text=u"导出",
                                         command=self._export)

        self._set_footer_mode(self.read_only)

    # ============================================================
    # Helpers
    # ============================================================
    def _today(self):
        import datetime
        return datetime.date.today().strftime("%Y-%m-%d")

    def _set_footer_mode(self, read_only):
        """Show/hide footer buttons based on mode."""
        if read_only:
            self._btn_add.pack_forget()
            self._btn_del.pack_forget()
            self._btn_save.pack_forget()
            self._btn_export.pack_forget()
            self._btn_print.pack_forget()
            self._chk_done.pack(side=tk.LEFT, padx=(16, 2))
            self._btn_cancel.pack(side=tk.RIGHT, padx=2)
            self._btn_export_ro.pack(side=tk.RIGHT, padx=4)
        else:
            self._btn_export_ro.pack_forget()
            self._btn_add.pack(side=tk.LEFT, padx=2)
            self._btn_del.pack(side=tk.LEFT, padx=2)
            self._chk_done.pack(side=tk.LEFT, padx=(16, 2))
            self._btn_cancel.pack(side=tk.RIGHT, padx=2)
            self._btn_print.pack(side=tk.RIGHT, padx=2)
            self._btn_save.pack(side=tk.RIGHT, padx=2)
            self._btn_export.pack(side=tk.RIGHT, padx=4)

    def _del_rows(self):
        if self.grid.row_count() > 0:
            self.grid.delete_rows([self.grid.row_count() - 1])
            self._recalc()

    # ============================================================
    # Search dialogs
    # ============================================================
    def _search_product(self, row, col):
        cur = self.grid.get_cell(row, col)
        dlg = tk.Toplevel(self)
        dlg.title(u"搜索产品")
        dlg.geometry("520x380")
        dlg.transient(self)
        dlg.grab_set()
        tk.Label(dlg, text=u"输入关键字:", font=("Microsoft YaHei", 9)).pack(
            padx=10, pady=(10, 2), anchor=tk.W)
        entry = ttk.Entry(dlg, width=50)
        entry.pack(padx=10, fill=tk.X)
        entry.insert(0, cur if cur else "")
        lb_frame = tk.Frame(dlg)
        lb_frame.pack(padx=10, pady=4, fill=tk.BOTH, expand=True)
        lb = tk.Listbox(lb_frame, height=14, font=("Microsoft YaHei", 10))
        lb.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        sb = ttk.Scrollbar(lb_frame, orient=tk.VERTICAL, command=lb.yview)
        lb.configure(yscrollcommand=sb.set)
        sb.pack(side=tk.RIGHT, fill=tk.Y)
        results = []

        def refresh(kw=""):
            lb.delete(0, tk.END)
            results[:] = Product.search(kw) if kw else Product.get_all()
            for p in results:
                d = p["product_name"]
                if p["model_number"]:
                    d = u"[{}] {}".format(p["model_number"], d)
                lb.insert(tk.END, d)

        refresh(cur)
        entry.bind("<KeyRelease>", lambda e: refresh(entry.get().strip()))

        def confirm(e=None):
            sel = lb.curselection()
            if sel and sel[0] < len(results):
                p = results[sel[0]]
                self.grid.set_cell(row, C_MODEL, p.get("model_number", ""))
                self.grid.set_cell(row, C_NAME, p.get("product_name", ""))
                self.grid.draw()
            dlg.destroy()

        lb.bind("<Double-1>", confirm)
        lb.bind("<Return>", confirm)
        btn = tk.Frame(dlg)
        btn.pack(padx=10, pady=(4, 10), fill=tk.X)
        tk.Button(btn, text=u"新增产品",
                  command=lambda: (
                      Product.create(product_name=entry.get().strip()),
                      refresh())).pack(side=tk.LEFT, padx=2)
        tk.Button(btn, text=u"确定", command=confirm).pack(side=tk.RIGHT, padx=2)
        entry.focus_set()
        dlg.wait_window()

    def _search_color(self, row, col):
        cur = self.grid.get_cell(row, col)
        dlg = tk.Toplevel(self)
        dlg.title(u"选择颜色")
        dlg.geometry("350x340")
        dlg.transient(self)
        dlg.grab_set()
        tk.Label(dlg, text=u"选择或输入:", font=("Microsoft YaHei", 9)).pack(
            padx=10, pady=(10, 2), anchor=tk.W)
        entry = ttk.Entry(dlg, width=30)
        entry.pack(padx=10, fill=tk.X)
        entry.insert(0, cur if cur else "")
        lb = tk.Listbox(dlg, height=12, font=("Microsoft YaHei", 10))
        lb.pack(padx=10, pady=4, fill=tk.BOTH, expand=True)
        for c in Color.get_all():
            lb.insert(tk.END, c["name"])
        lb.bind("<ButtonRelease-1>",
                lambda e: (entry.delete(0, tk.END),
                          entry.insert(0, lb.get(lb.curselection()[0])))
                if lb.curselection() else None)

        def confirm():
            v = entry.get().strip()
            if v:
                self.grid.set_cell(row, col, v)
                self.grid.draw()
            dlg.destroy()

        lb.bind("<Double-1>", lambda e: confirm())
        btn = tk.Frame(dlg)
        btn.pack(padx=10, pady=(4, 10), fill=tk.X)
        tk.Button(btn, text=u"确定", command=confirm).pack(side=tk.RIGHT, padx=2)
        entry.focus_set()
        dlg.wait_window()

    # ============================================================
    # Realtime calc
    # ============================================================
    def _on_date_var_changed(self):
        """Mark dirty when the delivery date changes."""
        if not getattr(self, '_loading', False):
            self._dirty = True

    def _on_grid_cell_changed(self, row, col, new_val):
        self._dirty = True
        if col == C_QTY:
            self._parse_qty(row)
        if col in (C_QTY, C_PRICE):
            self._update_amt(row)
        self._update_summary()

    def _parse_qty(self, row):
        raw = self.grid.get_cell(row, C_QTY)
        if not raw:
            self.grid.set_cell(row, C_NOTES, "")
            self.grid.set_display_override(row, C_QTY, "")
            return
        s = raw.strip()
        if s.startswith("="):
            s = s[1:]
        r = parse_formula(s)
        if not r["error"] and r["piece_count"] > 0:
            self.grid.set_cell(row, C_QTY, s)
            self.grid.set_display_override(row, C_QTY, str(r["quantity"]))
            self.grid.set_cell(row, C_NOTES, r.get("notes", ""))
        elif s.replace(".", "").replace("-", "").isdigit():
            self.grid.set_cell(row, C_QTY, s)
            self.grid.set_display_override(row, C_QTY, "")
            self.grid.set_cell(row, C_NOTES, u"{}（1个）".format(s))
        else:
            self.grid.set_cell(row, C_QTY, s)
            self.grid.set_display_override(row, C_QTY, "")
            self.grid.set_cell(row, C_NOTES, "")

    def _update_amt(self, row):
        disp = self.grid._display_overrides.get((row, C_QTY))
        qs = disp if disp else self.grid.get_cell(row, C_QTY)
        ps = self.grid.get_cell(row, C_PRICE)
        try:
            q = float(qs) if qs else 0
        except ValueError:
            q = 0
        try:
            p = float(ps) if ps else 0
        except ValueError:
            self.grid.set_cell(row, C_AMOUNT, "")
            return
        a = q * p
        self.grid.set_cell(row, C_AMOUNT, "{:.2f}".format(a) if a else "")

    def _update_summary(self):
        data = self.grid.get_data()
        pieces = 0
        total = 0
        for ri, row in enumerate(data):
            if not row[C_NAME]:
                continue
            qs = row[C_QTY] if len(row) > C_QTY else ""
            if qs:
                s = qs.strip()
                if s.startswith("="):
                    s = s[1:]
                r = parse_formula(s)
                if not r["error"] and r["piece_count"] > 0:
                    pieces += r["piece_count"]
                elif s.replace(".", "").replace("-", "").isdigit():
                    pieces += 1
            amt = row[C_AMOUNT] if len(row) > C_AMOUNT else ""
            try:
                total += float(amt) if amt else 0
            except ValueError:
                pass
        if pieces > 0 or total > 0:
            self.grid.summary_text = u"合计个数: {}    合计金额: {:.2f}".format(
                pieces, total)
        else:
            self.grid.summary_text = ""
        self.grid.draw()

    # ============================================================
    # Recalc (full)
    # ============================================================
    def _recalc(self):
        self.grid.commit_edit()
        data = self.grid.get_data()
        self.grid.clear_overrides()
        for ri, row in enumerate(data):
            if not row[C_NAME]:
                continue
            qs = row[C_QTY] if len(row) > C_QTY else ""
            if qs:
                s = qs.strip()
                if s.startswith("="):
                    s = s[1:]
                r = parse_formula(s)
                if not r["error"] and r["piece_count"] > 0:
                    row[C_QTY] = s
                    self.grid.set_display_override(ri, C_QTY, str(r["quantity"]))
                    row[C_NOTES] = r.get("notes", "")
                elif s.replace(".", "").replace("-", "").isdigit():
                    row[C_NOTES] = u"{}（1个）".format(s)
            disp = self.grid._display_overrides.get((ri, C_QTY))
            qc = disp if disp else (row[C_QTY] if len(row) > C_QTY else "")
            ps = row[C_PRICE] if len(row) > C_PRICE else ""
            try:
                q = float(qc) if qc else 0
            except ValueError:
                q = 0
            try:
                p = float(ps) if ps else 0
            except ValueError:
                p = None
            if p is not None and q > 0 and p > 0:
                row[C_AMOUNT] = "{:.2f}".format(q * p)
            else:
                row[C_AMOUNT] = ""
        self.grid.set_data(data)
        self._update_summary()

    # ============================================================
    # Save / Load
    # ============================================================
    def _save(self, silent=False, callback=None):
        import threading, datetime
        dn = self.dn_var.get().strip()
        dd = self.dlv_var.get().strip()
        self.grid.commit_edit()
        self._recalc()
        data = self.grid.get_data()
        order_id = self.order_id
        dn_id_val = self.dn_id
        customer = self._customer
        on_cb = self.on_save_callback
        done = getattr(self, 'done_var', None)
        done_val = done.get() if done else 0
        overrides = dict(self.grid._display_overrides)
        self._btn_save.config(state="disabled")

        def do_save():
            from db.database import get_connection
            conn = get_connection()
            try:
                conn.execute("BEGIN IMMEDIATE")
                now = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                dn_str = dn
                if dn_id_val:
                    conn.execute(
                        "UPDATE delivery_note SET delivery_number=?,"
                        " delivery_date=?, updated_at=? WHERE id=?",
                        (dn, dd, now, dn_id_val))
                    conn.execute(
                        "DELETE FROM delivery_note_item"
                        " WHERE delivery_note_id=?", (dn_id_val,))
                    dnid = dn_id_val
                else:
                    if not dn_str and customer:
                        dn_str = generate_delivery_number(
                            customer["id"], dd or None)
                    if not dn_str:
                        if not silent:
                            self.after(0, lambda: messagebox.showwarning(
                                u"提示", u"请输入送货单号。"))
                        self._safe_ui(lambda: self._btn_save.config(state="normal"))
                        return
                    cust_id = customer["id"] if customer else None
                    cur = conn.execute(
                        "INSERT INTO delivery_note (order_id, customer_id,"
                        " delivery_number, delivery_date, status,"
                        " created_at, updated_at)"
                        " VALUES (?,?,?,?,'draft',?,?)",
                        (order_id, cust_id, dn_str, dd, now, now))
                    dnid = cur.lastrowid
                    if order_id:
                        conn.execute(
                            "INSERT OR IGNORE INTO delivery_note_order"
                            " (delivery_note_id, order_id) VALUES (?,?)",
                            (dnid, order_id))
                    self.after(0, lambda: setattr(self, 'dn_id', dnid))

                for i, row in enumerate(data):
                    if not row[C_NAME]:
                        continue
                    mn, pn, cn = row[C_MODEL] or "", row[C_NAME] or "", row[C_COLOR] or ""
                    # find-or-create product
                    conn.execute(
                        "INSERT OR IGNORE INTO product"
                        " (model_number, item_code, product_name)"
                        " VALUES (?,?,?)", (mn, "", pn))
                    pr = conn.execute(
                        "SELECT id, model_number FROM product"
                        " WHERE product_name=? AND model_number=?",
                        (pn, mn)).fetchone()
                    # find-or-create color
                    conn.execute(
                        "INSERT OR IGNORE INTO color (name) VALUES (?)", (cn,))
                    cr = conn.execute(
                        "SELECT id FROM color WHERE name=?", (cn,)).fetchone()

                    formula = row[C_QTY] if len(row) > C_QTY else ""
                    disp = overrides.get((i, C_QTY))
                    qs = disp if disp else formula
                    try:
                        q = float(qs) if qs else 0
                    except ValueError:
                        q = 0
                    try:
                        amt = float(row[C_AMOUNT]) if len(row) > C_AMOUNT and row[C_AMOUNT] else 0
                    except ValueError:
                        amt = 0
                    order_no = row[C_ORDER_NO] if len(row) > C_ORDER_NO else ""
                    mfg_no = row[C_MFG_NO] if len(row) > C_MFG_NO else ""
                    cust_code = row[C_CUST_CODE] if len(row) > C_CUST_CODE else ""
                    conn.execute(
                        "INSERT INTO delivery_note_item (delivery_note_id,"
                        " model_number, item_code, customer_code,"
                        " product_name, color_name,"
                        " quantity_formula, quantity, unit_price, amount,"
                        " notes, mfg_number, sort_order, created_at, updated_at)"
                        " VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)",
                        (dnid, pr["model_number"] if pr else mn, order_no,
                         cust_code,
                         pn, cn, formula, q,
                         row[C_PRICE] if len(row) > C_PRICE else "",
                         amt, row[C_NOTES] if len(row) > C_NOTES else "",
                         mfg_no, i, now, now))

                if done_val:
                    conn.execute(
                        "UPDATE delivery_note SET status='completed',"
                        " updated_at=? WHERE id=?", (now, dnid))

                conn.commit()
                self._dirty = False
                if not silent:
                    self.after(0, lambda: messagebox.showinfo(
                        u"保存成功", u"送货单已保存。\n单号: {}".format(dn_str)))
                if callback:
                    self._safe_ui(callback)
                elif on_cb:
                    self._safe_ui(on_cb)
            except Exception as e:
                try:
                    conn.rollback()
                except:
                    pass
                if not silent:
                    self.after(0, lambda: messagebox.showerror(
                        u"保存失败", str(e)))
            finally:
                conn.close()
                self._safe_ui(lambda: self._btn_save.config(state="normal"))

        t = threading.Thread(target=do_save)
        t.daemon = True
        t.start()

    def _load_from_order(self):
        if self._customer:
            self.cust_lbl.config(text=self._customer["name"])
        self.info_lbl.config(
            text=u"订单: {}    ".format(self._order.get("display_name", "")))
        items = OrderItem.get_by_order(self.order_id)
        rows = []
        for item in items:
            formula = item.get("quantity_formula") or ""
            q = item["quantity"] if item["quantity"] else ""
            qty_cell = formula if formula else (str(q) if q else "")
            rows.append([item.get("order_number", ""),
                         item.get("mfg_number", ""),
                         item["model_number"],
                         item.get("customer_code", ""),
                         item["product_name"],
                         item["color_name"], qty_cell,
                         "", "", item["notes"] if item["notes"] else ""])
        self.grid.set_data(rows)
        self._recalc()

    def _load_dn(self):
        dn = DeliveryNote.get_by_id(self.dn_id)
        if not dn:
            return
        self._loading = True
        self.dn_var.set(dn["delivery_number"])
        self.dlv_var.set(dn["delivery_date"] or self._today())
        if dn.get("customer_name"):
            self.cust_lbl.config(text=dn["customer_name"])
        if dn["status"] == "completed":
            if hasattr(self, 'done_var'):
                self.done_var.set(1)
        source_orders = dn.get("source_orders", [])
        if len(source_orders) > 1:
            names = u",".join(o["display_name"] for o in source_orders)
            self.info_lbl.config(
                text=u"来源订单: {}    ".format(names))
        else:
            self.info_lbl.config(
                text=u"订单: {}    ".format(dn.get("order_display_name", "")))
        items = DNItem.get_by_dn(self.dn_id)
        rows = []
        for item in items:
            # Use formula if saved, otherwise plain quantity
            formula = item.get("quantity_formula") or ""
            if formula and (formula.startswith("=") or "*" in formula
                          or "+" in formula):
                qty_cell = formula
            elif formula:
                qty_cell = str(formula)
            else:
                q = item["quantity"] if item["quantity"] else ""
                qty_cell = str(q) if q else ""
            rows.append([item.get("item_code", ""), item.get("mfg_number", ""),
                         item["model_number"], item.get("customer_code", ""),
                         item["product_name"],
                         item["color_name"], qty_cell,
                         item["unit_price"] if item["unit_price"] else "",
                         "{:.2f}".format(item["amount"]) if item["amount"] else "",
                         item["notes"] if item["notes"] else ""])
        self.grid.set_data(rows)
        self._recalc()
        self._loading = False
        self._dirty = False

    def _on_done_toggle(self):
        """Handle 完单 checkbox toggle."""
        if not self.done_var.get():
            # User unchecked → cancel 完单
            if self.read_only:
                if messagebox.askyesno(
                        u"取消完单",
                        u"确定要取消完单状态吗？\n取消后可以继续编辑。"):
                    if self.dn_id:
                        DeliveryNote.update(self.dn_id, status="draft")
                    self.grid.read_only = False
                    self.read_only = False
                    self._set_footer_mode(False)
                    self.title(self.title().replace(u"（只读）", u"（编辑模式）"))
                else:
                    self.done_var.set(1)  # revert
            # If not read_only, unchecking just means don't mark complete on save

    def _on_readonly_edit(self, row, col):
        if self._ask_yesno(
                u"已完单",
                u"此送货单已标记为完单，不能编辑。"
                u"\n\n要解除完单状态吗？"):
            if self.dn_id:
                DeliveryNote.update(self.dn_id, status="draft")
            self.grid.read_only = False
            self.read_only = False
            self.done_var.set(0)
            self._set_footer_mode(False)
            self.title(self.title().replace(u"（只读）", u"（编辑模式）"))

    def _ask_yesno(self, title, msg):
        dlg = tk.Toplevel(self)
        dlg.title(title)
        dlg.geometry("380x160")
        dlg.resizable(False, False)
        dlg.transient(self)
        dlg.grab_set()
        result = [False]
        tk.Label(dlg, text=msg, font=("Microsoft YaHei", 10),
                 wraplength=340, justify=tk.LEFT).pack(padx=20, pady=(20, 10))
        btn_frame = tk.Frame(dlg)
        btn_frame.pack()
        ttk.Button(btn_frame, text=u"是", command=lambda: (result.__setitem__(0, True), dlg.destroy()),
                   width=10).pack(side=tk.LEFT, padx=4)
        ttk.Button(btn_frame, text=u"否", command=lambda: (result.__setitem__(0, False), dlg.destroy()),
                   width=10).pack(side=tk.LEFT, padx=4)
        dlg.protocol("WM_DELETE_WINDOW", dlg.destroy)
        dlg.wait_window()
        return result[0]

    def _save_grid_widths(self):
        try:
            Settings.set("dn_grid_widths",
                         json.dumps(self.grid.col_widths))
        except Exception:
            pass

    def _load_grid_widths(self):
        try:
            raw = Settings.get_all().get("dn_grid_widths", "")
            if raw:
                return json.loads(raw)
        except Exception:
            pass
        return None

    def _on_close(self):
        self._save_grid_widths()
        if self._dirty and not self.read_only:
            answer = messagebox.askyesnocancel(
                u"未保存", u"有未保存的修改，要保存吗？\n\n"
                u"  是(Y) = 保存并退出\n  否(N) = 不保存退出\n  取消 = 继续编辑")
            if answer is None:
                return
            if answer:
                self._save(silent=True, callback=self.destroy)
                return
        self.destroy()

    def _export(self):
        if not self.dn_id or self._dirty:
            self._save(silent=False, callback=self._do_export)
            return
        self._do_export()

    def _do_export(self, silent=False, mark_completed=True):
        if not self.dn_id:
            return None
        dn = DeliveryNote.get_by_id(self.dn_id)
        if not dn:
            return None
        export_cnt = dn.get("export_count", 0)
        if export_cnt >= 1 and not silent:
            msg = u"此送货单已导出过 {} 次，是否再次导出？".format(export_cnt)
            if not messagebox.askyesno(u"再次导出确认", msg):
                return None
        items = DNItem.get_by_dn(self.dn_id)
        if len(items) > 12 and not silent:
            if not messagebox.askyesno(
                    u"超出模板限制",
                    u"物料行数（{}行）超过打印模板12行限制！\n\n"
                    u"是否继续导出？".format(len(items))):
                return None
        try:
            from logic.excel_export import generate_delivery_note
            ok, result = generate_delivery_note(self.dn_id)
            if ok:
                if mark_completed:
                    DeliveryNote.update(self.dn_id, export_count=export_cnt + 1,
                                        status="completed")
                    # Update local state
                    self.done_var.set(1)
                    self.grid.read_only = True
                    self.read_only = True
                    self._set_footer_mode(True)
                else:
                    DeliveryNote.update(self.dn_id, export_count=export_cnt + 1)
                if not silent:
                    msg = u"送货单已导出到：\n{}".format(result)
                    if mark_completed:
                        msg += u"\n已自动标记为完单。"
                    messagebox.showinfo(u"导出成功", msg)
                return result
            else:
                if not silent:
                    messagebox.showerror(u"导出失败", str(result))
                return None
        except Exception as e:
            if not silent:
                messagebox.showerror(u"导出失败", str(e))
            return None

    def _print(self):
        """Export then print directly to default printer."""
        # Save first if dirty
        if not self.dn_id or self._dirty:
            self._save(silent=False, callback=self._do_print)
            return
        self._do_print()

    def _do_print(self):
        if not self.dn_id:
            return

        # 先静默导出数据到 Excel 文件
        filepath = self._do_export(silent=True, mark_completed=False)
        if not filepath:
            return

        import os
        import threading

        # 禁用打印按钮，防止手抖狂点导致发了几十个打印任务
        self._btn_print.config(state="disabled")

        def _print_worker():
            import pythoncom
            # 【第一道保险】在后台线程调用 COM 必须先初始化！
            pythoncom.CoInitialize()

            try:
                # 尝试用我们修复好的 WPS COM 打印
                _wps_print_last_sheet(filepath)
                self.after(0, lambda: messagebox.showinfo(
                    u"打印完成", u"送货单已成功发送至打印机。\n\n文件保存于：{}".format(filepath)))

            except Exception as e:
                # 【第四道保险】COM 彻底崩溃时的系统级兜底方案
                try:
                    os.startfile(filepath, "print")
                    self.after(0, lambda: messagebox.showinfo(
                        u"打印提示", u"使用了系统备用打印通道发送完成。"))
                except Exception as fallback_e:
                    self.after(0, lambda: messagebox.showerror(
                        u"打印彻底失败", u"WPS接口及系统打印均失败：\n{}".format(str(e))))

            finally:
                # 【第一道保险】释放线程的 COM 资源
                pythoncom.CoUninitialize()
                # 恢复打印按钮
                self._safe_ui(lambda: self._btn_print.config(state="normal"))

        # 启动后台守护线程执行打印
        t = threading.Thread(target=_print_worker)
        t.daemon = True
        t.start()
