# -*- coding: utf-8 -*-
"""Order detail window with Excel-style grid."""
try:
    import Tkinter as tk
    import ttk
    import tkMessageBox as messagebox
except ImportError:
    import tkinter as tk
    from tkinter import ttk
    from tkinter import messagebox

import json
from db.models import (Customer, Product, Color, Order, OrderItem, Settings,
                       deduplicate_dn_items, get_column_visibility)
from logic.order_number import generate_delivery_number
from ui.excel_grid import ExcelGrid
from ui.widgets import DatePicker


def _safe_btn_config(btn, **kwargs):
    """Safely configure a button that may have been destroyed."""
    try:
        if btn.winfo_exists():
            btn.config(**kwargs)
    except Exception:
        pass


class OrderWindow(tk.Toplevel):

    COLUMNS = [u"打印", u"推送", u"客户号", u"订单号", u"制单号", u"款号",
               u"品名", u"颜色", u"数量", u"打印张数", u"状态"]
    COL_WIDTHS = [46, 46, 80, 100, 100, 80, 250, 90, 80, 75, 50]
    CHECKBOX_COL = 0
    SEARCH_COLS = {6: "product", 7: "color"}

    def __init__(self, parent, order_id=None, on_save_callback=None,
                 read_only=False):
        tk.Toplevel.__init__(self, parent)
        self.order_id = order_id
        self.on_save_callback = on_save_callback
        self.read_only = read_only

        self._customers = Customer.get_all()
        self._dirty = False
        self._last_cust_name = ""
        self._full_rows = []          # master copy of all rows
        self._filter_index_map = []    # filtered-row → full-row mapping
        self._filter_entries = {}      # col → (StringVar, Entry)

        if self.order_id:
            self.title(u"编辑订单")
        else:
            self.title(u"新建订单")

        sw = self.winfo_screenwidth()
        sh = self.winfo_screenheight()
        w = int(sw * 0.85)
        h = int(sh * 0.7)
        self.geometry("{}x{}".format(w, h))
        self.minsize(900, 520)
        self.transient(parent)

        self._build()

        if self.order_id:
            self.after(100, self._load_order)
            self.after(200, lambda: setattr(self, '_dirty', False))

        # Keyboard shortcuts
        self.bind("<Control-s>", lambda e: self._save())
        self.bind("<Control-S>", lambda e: self._save())
        self.protocol("WM_DELETE_WINDOW", self._on_close)

    def _safe_ui(self, fn):
        """线程安全的 UI 更新：双重拦截 TclError 与窗口销毁异常"""
        try:
            if self.winfo_exists():
                def _run():
                    try:
                        fn()
                    except tk.TclError:
                        pass
                self.after(0, _run)
        except Exception:
            pass
        self.after(0, _run)

    def _build(self):
        # ---- Big Title ----
        title_bar = tk.Frame(self, bg="#2c3e50", height=36)
        title_bar.pack(fill=tk.X)
        title_bar.pack_propagate(False)
        status_text = u"（只读）" if self.read_only else u"（编辑模式）"
        title_label = tk.Label(
            title_bar, text=u"  订单明细  {}".format(status_text),
            font=("Microsoft YaHei", 14, "bold"),
            bg="#2c3e50", fg="white", anchor=tk.W)
        title_label.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        # ---- Header Row ----
        hdr = tk.Frame(self)
        hdr.pack(fill=tk.X)

        r1 = tk.Frame(hdr)
        r1.pack(fill=tk.X, pady=2)

        tk.Label(r1, text=u"客户:", font=("Microsoft YaHei", 9)).pack(
            side=tk.LEFT)
        self.customer_var = tk.StringVar()
        self.customer_var.trace("w", lambda *a: self._on_customer_change())
        self.customer_cb = ttk.Combobox(
            r1, textvariable=self.customer_var, width=14)
        self.customer_cb["values"] = [c["name"] for c in self._customers]
        self.customer_cb.pack(side=tk.LEFT, padx=(2, 14))
        self.customer_cb.bind("<<ComboboxSelected>>", self._on_customer_change)

        tk.Label(r1, text=u"日期:", font=("Microsoft YaHei", 9)).pack(
            side=tk.LEFT)
        self.date_var = tk.StringVar(value=self._today())
        df = tk.Frame(r1)
        df.pack(side=tk.LEFT, padx=(2, 14))
        self.date_e = ttk.Entry(df, textvariable=self.date_var, width=12)
        self.date_e.pack(side=tk.LEFT)
        ttk.Button(df, text=u"📅", width=2,
                   command=lambda: DatePicker(self, self.date_e)).pack(
                       side=tk.LEFT)

        tk.Label(r1, text=u"订单名称:", font=("Microsoft YaHei", 9)).pack(
            side=tk.LEFT)
        self.name_var = tk.StringVar()
        ttk.Entry(r1, textvariable=self.name_var, width=20).pack(
            side=tk.LEFT, padx=(2, 14))
        self.name_var.trace("w", lambda *a: (self._auto_name(), setattr(self, '_dirty', True)))

        tk.Label(r1, text=u"送货单号:", font=("Microsoft YaHei", 9)).pack(
            side=tk.LEFT)
        self.dn_var = tk.StringVar()
        ttk.Entry(r1, textvariable=self.dn_var, width=16).pack(
            side=tk.LEFT, padx=(2, 0))

        # Field hint
        r2 = tk.Frame(hdr)
        r2.pack(fill=tk.X)
        self.hint_lbl = tk.Label(r2, text="", fg="gray",
                                 font=("Microsoft YaHei", 8))
        self.hint_lbl.pack(side=tk.LEFT)

        # ---- Excel Grid ----
        widths = list(self.COL_WIDTHS)
        saved = self._load_grid_widths()
        if saved and len(saved) == len(widths):
            widths = saved
        self.grid = ExcelGrid(
            self,
            columns=self.COLUMNS,
            col_widths=widths,
            checkbox_col=self.CHECKBOX_COL,
            extra_checkbox_cols={1},  # 推送列
            read_only_cols={10},    # 状态列冻结
            search_cols={})         # 清空弹窗库，恢复直接打字输入

        if self.read_only:
            pass

        self.grid.on_cell_changed = self._on_grid_cell_changed
        self.grid.on_resize_done = self._save_grid_widths

        self.protocol("WM_DELETE_WINDOW", self._on_close)
        self.grid.data_frame.bind("<Button-3>", self._grid_context_menu)
        if self.read_only:
            self.grid.read_only = True
            self.grid.on_edit_blocked = self._on_readonly_edit

        # ---- Filter Bar (放在表头和数据之间) ----
        self._filter_bar = tk.Frame(self, height=42, bg="#e8e8e8")
        self._filter_bar.pack(fill=tk.X, padx=10, pady=(0, 2))
        self._filter_bar.pack_propagate(False)

        self.grid.pack(fill=tk.BOTH, expand=True, padx=10, pady=2)
        self._build_filter_bar()
        # 接管底层表格的 add_row，确保新行同步到 _full_rows
        self.grid.add_row = self._add_row

        # ---- Footer ----
        ft = tk.Frame(self)
        ft.pack(fill=tk.X)

        ttk.Button(ft, text=u"添加行",
                   command=lambda: self._add_row()).pack(
                       side=tk.LEFT, padx=2)
        ttk.Button(ft, text=u"删除选中行",
                   command=self._del_rows).pack(side=tk.LEFT, padx=2)

        tk.Frame(ft, width=12).pack(side=tk.LEFT)

        ttk.Button(ft, text=u"打印勾选行",
                   command=self._print_checked).pack(side=tk.LEFT, padx=2)

        if not self.read_only:
            self.push_btn = ttk.Button(
                ft, text=u"推送到送货单 →",
                command=self._push_to_dn,
                state=tk.NORMAL if self.order_id else tk.DISABLED)
            self.push_btn.pack(side=tk.LEFT, padx=(16, 2))

            # 完单 checkbox
            self.done_var = tk.IntVar(value=0)
            ttk.Checkbutton(ft, text=u"完单", variable=self.done_var).pack(
                side=tk.LEFT, padx=(8, 2))

        tk.Frame(ft, width=12).pack(side=tk.RIGHT)
        ttk.Button(ft, text=u"取消",
                   command=self._on_close).pack(side=tk.RIGHT, padx=2)
        if not self.read_only:
            ttk.Button(ft, text=u"保存草稿",
                       command=self._save).pack(side=tk.RIGHT, padx=4)

    def _grid_context_menu(self, event):
        if self.read_only:
            return
        menu = tk.Menu(self, tearoff=0)
        menu.add_command(label=u"打印全选", command=self._select_all_print)
        menu.add_command(label=u"推送全选", command=self._select_all_push)
        menu.add_separator()
        menu.add_command(label=u"添加行", command=lambda: self._add_row())
        menu.tk_popup(event.x_root, event.y_root)

    def _select_all_print(self):
        for ri in range(self.grid.row_count()):
            if self.grid.get_cell(ri, 6):  # has product name
                self.grid.set_cell(ri, 0, u"☑")
                if self.grid.on_cell_changed:
                    self.grid.on_cell_changed(ri, 0, u"☑")
        self.grid.draw()

    def _select_all_push(self):
        for ri in range(self.grid.row_count()):
            if self.grid.get_cell(ri, 6):  # has product name
                self.grid.set_cell(ri, 1, u"☑")
                if self.grid.on_cell_changed:
                    self.grid.on_cell_changed(ri, 1, u"☑")
        self.grid.draw()

    # ================================================================
    # Customer
    # ================================================================
    def _get_customer(self):
        name = self.customer_var.get().strip()
        if not name:
            return None
        # Exact match first
        for c in self._customers:
            if c["name"] == name:
                return c
        # Fuzzy: customer name contains typed text
        for c in self._customers:
            if name in c["name"]:
                return c
        # Fuzzy: typed text contains customer name
        for c in self._customers:
            if c["name"] in name:
                return c
        return None

    def _on_customer_change(self, e=None):
        self._auto_name()
        self._apply_column_visibility()

    def _auto_name(self):
        cust = self._get_customer()
        dv = self.date_var.get().strip()
        if not cust or not dv:
            return
        try:
            parts = dv.split("-")
            m, d = int(parts[1]), int(parts[2].split(" ")[0])
            new_name = u"{}{}.{}".format(cust["name"], m, d)
        except (ValueError, IndexError):
            return
        current = self.name_var.get().strip()
        if not current:
            self.name_var.set(new_name)
        elif self._last_cust_name and current.startswith(self._last_cust_name):
            # Replace old customer name prefix, keep date suffix
            suffix = current[len(self._last_cust_name):]
            self.name_var.set(cust["name"] + suffix)
        self._last_cust_name = cust["name"]

    def _today(self):
        import datetime
        return datetime.date.today().strftime("%Y-%m-%d")

    # ================================================================
    # Row ops
    # ================================================================
    def _add_row(self, values=None):
        """新增行：支持底层表格敲回车触发，自动继承过滤栏内容防止行消失"""
        if values:
            row = list(values)
        else:
            row = [u"☐" if i == 0 or i == 1 else ""
                   for i in range(len(self.COLUMNS))]
            for ci, (var, _) in self._filter_entries.items():
                val = var.get().strip()
                if val:
                    row[ci] = val
        self._full_rows.append(row)
        self._apply_filters()
        # 滚动到底部
        self.grid.canvas.update_idletasks()
        self.grid.canvas.yview_moveto(1.0)

    def _del_rows(self):
        sel = self.grid.get_selected_rows()
        if sel:
            full_sel = sorted([self._map_row(r) for r in sel], reverse=True)
            for fi in full_sel:
                if fi < len(self._full_rows):
                    self._full_rows.pop(fi)
            self._apply_filters()
        else:
            messagebox.showinfo(u"提示", u"请先勾选要删除的行。")

    def _print_checked(self):
        cust = self._get_customer()
        if not cust:
            messagebox.showwarning(u"提示", u"请先选择客户。")
            return

        self.grid.commit_edit()
        data = self._get_full_data()
        checked = [(ri, row) for ri, row in enumerate(data)
                   if row[0] == u"☑" and row[6]]
        if not checked:
            messagebox.showinfo(u"提示", u"请先勾选要打印的行（点击☐变☑）。")
            return

        # Build print items.  Formula quantities are expanded into
        # sub-items so that each sub-quantity prints the right number
        # of labels without missing any trailing terms.
        items = []
        total_label_count = 0
        row_label_counts = {}

        for ri, row in checked:
            qty_raw = (row[8] if len(row) > 8 else "").strip()
            pc_raw = (row[9] if len(row) > 9 else "").strip()

            # 获取用户在界面上手动指定的张数，默认为 1
            try:
                manual_pc = int(pc_raw) if pc_raw else 1
            except ValueError:
                manual_pc = 1

            # 核心判断：复杂连加公式（含有+或*）才走 expand_formula 拆分；
            # 纯数字严格尊重手动修改的张数，不走公式拆分。
            if u"+" in qty_raw or u"*" in qty_raw:
                from logic.formula import expand_formula
                sub_items = expand_formula(qty_raw)
            else:
                sub_items = []

            row_total = 0

            if sub_items:
                # 带有 + 或 * 的复杂公式流：各拆分项自带张数
                for si in sub_items:
                    items.append({
                        "customer_code": row[2] if len(row) > 2 else "",
                        "order_number": row[3],
                        "item_code": row[4],
                        "model_number": row[5],
                        "product_name": row[6],
                        "color_name": row[7],
                        "quantity": str(si["quantity"]),
                        "print_count": si["count"],
                    })
                    row_total += si["count"]
            else:
                # 纯数字输入流（例如：50米）：数量直接用输入的文本，张数严格采用手动输入修改后的 manual_pc
                items.append({
                    "customer_code": row[2] if len(row) > 2 else "",
                    "order_number": row[3],
                    "item_code": row[4],
                    "model_number": row[5],
                    "product_name": row[6],
                    "color_name": row[7],
                    "quantity": qty_raw,
                    "print_count": manual_pc,
                })
                row_total = manual_pc

            row_label_counts[ri] = row_total
            total_label_count += row_total

        if not items:
            messagebox.showinfo(u"提示", u"没有可打印的标签。")
            return

        if not messagebox.askyesno(
                u"标签打印",
                u"将打印 {} 张标签。\n\n"
                u"请确认 BarTender 已打开标签模板。".format(total_label_count)):
            return

        order_info = {
            "display_name": self.name_var.get().strip(),
            "order_date": self.date_var.get().strip(),
        }

        try:
            from logic.label_print import print_labels
            printed, err = print_labels(items, cust, order_info)
            if err:
                messagebox.showerror(
                    u"打印中断",
                    u"已打印 {} 张，遇到错误：\n{}".format(printed, err))
            else:
                messagebox.showinfo(
                    u"打印完成",
                    u"成功打印 {} 张标签。".format(printed))
        except Exception as e:
            messagebox.showerror(u"打印失败", str(e))
            return

        # Update print count display in grid
        items_db = OrderItem.get_by_order(self.order_id)
        for ri, row in checked:
            self.grid.set_cell(ri, 0, u"☑")
            if ri in row_label_counts:
                self.grid.set_cell(ri, 9, str(row_label_counts[ri]))
            if ri < len(items_db):
                OrderItem.update(items_db[ri]["id"], is_printed=1)
        self.grid.draw()
        self._dirty = True

    def _push_to_dn(self):
        if not self.order_id:
            return
        cust = self._get_customer()
        if not cust:
            messagebox.showwarning(u"提示", u"请先选择客户。")
            return

        data = self._get_full_data()
        checked = [row for row in data if row[1] == u"☑" and row[6]]
        if not checked:
            messagebox.showinfo(u"提示", u"请先勾选要推送的行（点击推送列切换☑）。")
            return

        # 12-row limit check
        if len(checked) > 12:
            if not messagebox.askyesno(
                    u"超出模板限制",
                    u"物料行数（{}行）超过打印模板12行限制！\n\n"
                    u"是否继续推送？".format(len(checked))):
                return

        if not messagebox.askyesno(
                u"推送到送货单",
                u"将勾选的 {} 行推送到新送货单？".format(len(checked))):
            return
        if not self._save(silent=True):
            return

        import threading, datetime

        checked_rows = [i for i, row in enumerate(data)
                       if row[1] == u"☑" and row[6]]
        order_id = self.order_id
        customer_id = cust["id"]
        on_save_cb = self.on_save_callback
        master = self.master
        grid = self.grid

        _safe_btn_config(self.push_btn, state=tk.DISABLED)
        self.config(cursor="watch")

        def do_push():
            from db.database import get_connection
            conn = get_connection()
            try:
                conn.execute("BEGIN IMMEDIATE")
                now = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")

                dn_str = generate_delivery_number(customer_id, None)
                if not dn_str:
                    self._safe_ui(lambda: messagebox.showwarning(
                        u"错误", u"无法生成送货单号。"))
                    return

                cust_row = conn.execute(
                    "SELECT customer_id FROM orders WHERE id=?",
                    (order_id,)).fetchone()
                cust_id = cust_row["customer_id"] if cust_row else None
                today = datetime.date.today().strftime("%Y-%m-%d")
                cur = conn.execute(
                    "INSERT INTO delivery_note (order_id, customer_id,"
                    " delivery_number, delivery_date, status,"
                    " created_at, updated_at)"
                    " VALUES (?,?,?,?,'draft',?,?)",
                    (order_id, cust_id, dn_str, today, now, now))
                dn_id = cur.lastrowid
                conn.execute(
                    "INSERT OR IGNORE INTO delivery_note_order"
                    " (delivery_note_id, order_id) VALUES (?,?)",
                    (dn_id, order_id))

                conn.execute(
                    "UPDATE orders SET status='qty_filled',"
                    " updated_at=? WHERE id=?", (now, order_id))

                # Collect items into dicts for dedup
                items = []
                for ri in checked_rows:
                    row = data[ri]
                    mn = row[5] or ""
                    items.append({
                        "model_number": mn,
                        "_orig_model": mn,
                        "customer_code": row[2] if len(row) > 2 else "",
                        "item_code": row[3] if len(row) > 3 else "",
                        "mfg_number": row[4] if len(row) > 4 else "",
                        "product_name": row[6] or "",
                        "color_name": row[7] or "",
                        "quantity_formula": row[8] if len(row) > 8 else "",
                    })

                # Deduplicate consecutive rows with same 款号+制单号+订单号
                deduplicate_dn_items(items)

                from logic.formula import parse_formula
                for idx, item in enumerate(items):
                    mn = item["_orig_model"]
                    pn = item["product_name"]
                    cn = item["color_name"]
                    conn.execute(
                        "INSERT OR IGNORE INTO product"
                        " (model_number, item_code, product_name)"
                        " VALUES (?,?,?)", (mn, "", pn))
                    pr = conn.execute(
                        "SELECT id, model_number FROM product"
                        " WHERE product_name=? AND model_number=?",
                        (pn, mn)).fetchone()
                    conn.execute(
                        "INSERT OR IGNORE INTO color (name) VALUES (?)", (cn,))

                    # 推送到送货单时，强行将 1+1 融合成汇总数字 2
                    q_raw = item["quantity_formula"]
                    res = parse_formula(q_raw)
                    if not res.get("error"):
                        q = float(res["quantity"])
                        q_formula_pushed = str(int(q)) if q.is_integer() else str(q)
                    else:
                        try:
                            q = float(q_raw) if q_raw else 0.0
                        except ValueError:
                            q = 0.0
                        q_formula_pushed = q_raw

                    conn.execute(
                        "INSERT INTO delivery_note_item (delivery_note_id,"
                        " model_number, item_code, customer_code,"
                        " product_name, color_name,"
                        " quantity_formula, quantity, notes, mfg_number,"
                        " sort_order, created_at, updated_at)"
                        " VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)",
                        (dn_id,
                         "" if not item["model_number"] else (pr["model_number"] if pr else mn),
                         item["item_code"],
                         item.get("customer_code", ""),
                         pn, cn,
                         q_formula_pushed,
                         q,
                         (res.get("notes", "") if q_raw else ""),
                         item["mfg_number"],
                         idx, now, now))

                # Get order item IDs for push_count updates
                item_rows = conn.execute(
                    "SELECT id FROM order_item WHERE order_id=?"
                    " ORDER BY sort_order, id",
                    (order_id,)).fetchall()

                for ri in checked_rows:
                    cur_val = data[ri][10] if len(data[ri]) > 10 else ""
                    try:
                        cnt = int(cur_val) if cur_val else 0
                    except ValueError:
                        cnt = 0
                    cnt += 1
                    if ri < len(item_rows):
                        conn.execute(
                            "UPDATE order_item SET push_count=?,"
                            " updated_at=? WHERE id=?",
                            (cnt, now, item_rows[ri]["id"]))

                conn.commit()

                def update_ui():
                    self._clear_filters()
                    for ri in checked_rows:
                        grid.set_cell(ri, 1, u"☐")
                        cur_val = grid.get_cell(ri, 10)
                        try:
                            cnt = int(cur_val) if cur_val else 0
                        except ValueError:
                            cnt = 0
                        grid.set_cell(ri, 10, str(cnt + 1))
                    grid.draw()
                    self._dirty = False
                    self.destroy()
                    from ui.dn_window import DeliveryNoteWindow
                    DeliveryNoteWindow(
                        master, dn_id=dn_id,
                        on_save_callback=on_save_cb)

                self._safe_ui(update_ui)
            except Exception as e:
                try:
                    conn.rollback()
                except:
                    pass
                self._safe_ui(lambda: messagebox.showerror(
                    u"推送失败", str(e)))
            finally:
                conn.close()
                self._safe_ui(lambda: self._safe_ui(
                    lambda: (self.config(cursor=""),
                             _safe_btn_config(self.push_btn, state=tk.NORMAL))))

        t = threading.Thread(target=do_push)
        t.daemon = True
        t.start()

    # ================================================================
    # Search dialogs
    # ================================================================
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
        lb = tk.Listbox(lb_frame, height=14,
                        font=("Microsoft YaHei", 10))
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
                if p["item_code"]:
                    d = u"({}) {}".format(p["item_code"], d)
                lb.insert(tk.END, d)

        refresh(cur)
        entry.bind("<KeyRelease>", lambda e: refresh(entry.get().strip()))

        def confirm(e=None):
            sel = lb.curselection()
            if sel and sel[0] < len(results):
                p = results[sel[0]]
                self.grid.set_cell(row, 5, p.get("model_number", ""))
                self.grid.set_cell(row, 6, p.get("product_name", ""))
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

        lb = tk.Listbox(dlg, height=12,
                        font=("Microsoft YaHei", 10))
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

    # ================================================================
    # Save / Load
    # ================================================================
    def _save(self, silent=False):
        cust = self._get_customer()
        if not cust:
            if not silent:
                name = self.customer_var.get().strip()
                if name:
                    messagebox.showwarning(
                        u"客户不存在",
                        u"客户「{}」不在档案中，请先在\n"
                        u"基础数据 → 客户管理 中新建客户。".format(name))
                else:
                    messagebox.showwarning(u"提示", u"请选择客户。")
            return False
        dv = self.date_var.get().strip()
        if not dv:
            if not silent:
                messagebox.showwarning(u"提示", u"请输入日期。")
            return False
        nm = self.name_var.get().strip()
        if not nm:
            self._auto_name()
            nm = self.name_var.get().strip()
        if not nm:
            if not silent:
                messagebox.showwarning(u"提示", u"请输入订单名称。")
            return False

        self.grid.commit_edit()
        data = self._get_full_data()
        items = [row for row in data if row[6]]
        if not items:
            if not silent:
                messagebox.showwarning(u"提示", u"请至少添加一行产品。")
            return False

        try:
            dn = self.dn_var.get().strip()
            if self.order_id:
                Order.update(self.order_id, display_name=nm,
                             order_date=dv, delivery_number=dn)
                OrderItem.delete_by_order(self.order_id)
                oid = self.order_id
            else:
                oid = Order.create(customer_id=cust["id"],
                                   display_name=nm, order_date=dv)
                if dn:
                    Order.update(oid, delivery_number=dn)
                self.order_id = oid

            for i, row in enumerate(data):
                if not row[6]:
                    continue
                p = Product.find_or_create(row[5], "", row[6])
                c = Color.find_or_create(row[7])

                # 修复数量与公式存储
                qty_raw = str(row[8]).strip() if row[8] else ""
                from logic.formula import parse_formula
                res = parse_formula(qty_raw)
                if not res.get("error"):
                    qty = float(res["quantity"])
                else:
                    try:
                        qty = float(qty_raw) if qty_raw else 0.0
                    except ValueError:
                        qty = 0.0

                try:
                    pc = int(row[9]) if row[9] else 1
                except ValueError:
                    pc = 1
                try:
                    p_cnt = int(row[10]) if len(row) > 10 and row[10] else 0
                except ValueError:
                    p_cnt = 0

                OrderItem.create(
                    order_id=oid, product_id=p["id"],
                    color_id=c["id"],
                    model_number=p["model_number"],
                    item_code=p["item_code"],
                    product_name=p["product_name"],
                    color_name=c["name"],
                    # 将表格里输入的原始公式字符串（如 1+1）完好无损地存入数据库
                    quantity_formula=qty_raw,
                    quantity=qty,
                    print_count=pc,
                    is_printed=1 if row[0] == u"☑" else 0,
                    push_count=p_cnt,
                    sort_order=i,
                    customer_code=row[2] if len(row) > 2 else "",
                    order_number=row[3] if len(row) > 3 else "",
                    mfg_number=row[4] if len(row) > 4 else "")

            if not silent:
                messagebox.showinfo(u"保存成功", u"订单已保存。")

            # Handle 完单 checkbox
            if hasattr(self, 'done_var') and self.done_var.get():
                if messagebox.askyesno(
                        u"确认完单",
                        u"确定将此订单标记为完单吗？\n完单后将锁定不可编辑。"):
                    Order.update(self.order_id, status="completed")
                    if self.on_save_callback:
                        self.on_save_callback()

            self._dirty = False
            _safe_btn_config(self.push_btn, state=tk.NORMAL)
            if self.on_save_callback:
                self.on_save_callback()
            return True
        except Exception as e:
            if not silent:
                messagebox.showerror(u"保存失败", str(e))
            return False

    # ---- Filter bar ----
    def _build_filter_bar(self):
        """创建与列宽完美对齐的过滤输入框"""
        for w in self._filter_bar.winfo_children():
            w.destroy()
        self._filter_entries.clear()

        filter_cols = [2, 3, 4, 5, 6, 7]  # 客户号..颜色
        for ci in filter_cols:
            var = tk.StringVar()
            var.trace("w", lambda *a, c=ci: self._on_filter_change(c))
            entry = tk.Entry(self._filter_bar, textvariable=var, width=8,
                             font=("Microsoft YaHei", 10), bg="#ffffcc",
                             relief="solid", borderwidth=1)
            self._filter_entries[ci] = (var, entry)

        self._clear_btn = tk.Button(self._filter_bar, text=u"一键清空",
                                    font=("Microsoft YaHei", 8),
                                    command=self._clear_filters)
        self._update_filter_layout()

    def _update_filter_layout(self):
        """根据表格的最新列宽，动态调整过滤输入框的绝对坐标"""
        x_offset = 0
        for i in range(len(self.COLUMNS)):
            w = self.grid.col_widths[i] if i < len(self.grid.col_widths) else 80
            if i in self._filter_entries and w > 0:
                var, entry = self._filter_entries[i]
                entry.place(x=x_offset + 2, y=4, width=w - 4, height=35)
                entry.lift()
            elif i in self._filter_entries:
                self._filter_entries[i][1].place_forget()
            x_offset += w
        self._clear_btn.place(x=x_offset + 10, y=4, height=35)

    def _on_filter_change(self, col):
        """任意输入框内容改变时，实时触发过滤"""
        self._apply_filters()

    def _apply_filters(self):
        """核心计算：执行多列 AND 条件的模糊匹配"""
        if not getattr(self, '_full_rows', None):
            self._filter_index_map = []
            self.grid.set_data([])
            return

        # 1. 收集所有当前填写的非空过滤条件
        filters = {}
        for ci, (var, _) in self._filter_entries.items():
            kw = var.get().strip().lower()
            if kw:
                filters[ci] = kw

        # 2. 计算筛选后的结果集
        if not filters:
            mapped = list(range(len(self._full_rows)))
            filtered = [list(r) for r in self._full_rows]
        else:
            mapped = []
            for fi, row in enumerate(self._full_rows):
                match = True
                for ci, kw in filters.items():
                    val = str(row[ci] if ci < len(row) else "").lower()
                    if kw not in val:
                        match = False
                        break
                if match:
                    mapped.append(fi)
            filtered = [list(self._full_rows[i]) for i in mapped]

        self._filter_index_map = mapped

        # 不挂载任何显示覆盖，列表中原本是什么文本就永远显示什么文本
        self.grid.clear_overrides()
        self.grid.set_data(filtered)

    def _clear_filters(self):
        """一键清空所有过滤框"""
        for var, _ in self._filter_entries.values():
            var.set("")
        self._apply_filters()

    def _map_row(self, filtered_row):
        """Translate filtered row index → full row index."""
        if self._filter_index_map and filtered_row < len(self._filter_index_map):
            return self._filter_index_map[filtered_row]
        return filtered_row  # fallback: identity map

    def _get_full_data(self):
        """Return the master (unfiltered) row list."""
        return [list(r) for r in self._full_rows]

    def _apply_column_visibility(self):
        """Hide code columns whose dn_headers label is empty for this customer."""
        cust = self._get_customer()
        if not cust:
            return
        vis = get_column_visibility(cust.get("dn_headers", ""))
        code_cols = [2, 3, 4, 5]  # 客户号, 订单号, 制单号, 款号
        defaults = [80, 100, 100, 80]
        for i, visible in enumerate(vis):
            if visible:
                self.grid.col_widths[code_cols[i]] = max(
                    self.grid.col_widths[code_cols[i]], defaults[i])
            else:
                self.grid.col_widths[code_cols[i]] = 0
        self.grid.draw()
        self._update_filter_layout()

    def _on_grid_cell_changed(self, row, col, new_val):
        self._dirty = True
        full_row = self._map_row(row)
        if full_row < len(self._full_rows) and col < len(self._full_rows[full_row]):
            self._full_rows[full_row][col] = new_val
        if col == 8:
            self._auto_calc_print_count(row, new_val)
        # Re-apply filter in case the edit changed which rows match
        if any(var.get().strip() for var, _ in self._filter_entries.values()):
            self._apply_filters()

    def _auto_calc_print_count(self, row, qty_raw):
        """当输入数量如 10*2+3 时，自动算出张数并填入第9列"""
        if not qty_raw:
            return
        from logic.formula import parse_formula
        res = parse_formula(qty_raw)
        if not res.get("error") and res.get("piece_count", 0) > 0:
            full_row = self._map_row(row)
            if full_row < len(self._full_rows):
                self._full_rows[full_row][9] = str(res["piece_count"])
            self.grid.set_cell(row, 9, str(res["piece_count"]))
            self.grid.draw()

    def _load_order(self):
        o = Order.get_by_id(self.order_id)
        if not o:
            return
        cust = Customer.get_by_id(o["customer_id"])
        if cust:
            self.customer_var.set(cust["name"])
            self._on_customer_change()
        self.date_var.set(o["order_date"])
        self.name_var.set(o["display_name"])
        self.dn_var.set(o["delivery_number"] or "")
        if o["status"] in ("completed", "dn_generated"):
            if hasattr(self, 'done_var'):
                self.done_var.set(1)

        items = OrderItem.get_by_order(self.order_id)
        rows = []
        for item in items:
            mk = u"☑" if item["is_printed"] else u"☐"
            pc = item["print_count"] if item["print_count"] else 1

            # 优先提取当初存进去的原始公式文本（如 1+1）
            qty = item.get("quantity_formula") or ""
            if not qty:
                q_num = item["quantity"] if item["quantity"] else ""
                qty = str(int(q_num)) if isinstance(q_num, float) and q_num.is_integer() else str(q_num) if q_num else ""

            push_cnt = item.get("push_count") or 0
            rows.append([
                mk,
                u"☐",
                item.get("customer_code", ""),
                item.get("order_number", ""),
                item.get("mfg_number", ""),
                item["model_number"],
                item["product_name"],
                item["color_name"],
                qty,
                str(pc),
                str(push_cnt) if push_cnt else "",
            ])
        self._full_rows = rows
        self._apply_column_visibility()
        self._apply_filters()

    def _on_readonly_edit(self, row, col):
        if self._ask_yesno(
                u"已完单",
                u"此订单已标记为完单，不能编辑。"
                u"\n\n要解除完单状态吗？"):
            Order.update(self.order_id, status="qty_filled")
            self.grid.read_only = False
            self.read_only = False
            # Update title
            self.title(self.title().replace(u"（只读）", u"（编辑模式）"))

    def _ask_yesno(self, title, msg):
        """Custom Yes/No dialog (works on Win7 where askyesno may not)."""
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

        def _yes():
            result[0] = True
            dlg.destroy()

        def _no():
            result[0] = False
            dlg.destroy()

        ttk.Button(btn_frame, text=u"是", command=_yes, width=10).pack(
            side=tk.LEFT, padx=4)
        ttk.Button(btn_frame, text=u"否", command=_no, width=10).pack(
            side=tk.LEFT, padx=4)

        dlg.protocol("WM_DELETE_WINDOW", _no)
        dlg.wait_window()
        return result[0]

    def _save_grid_widths(self):
        try:
            Settings.set("order_grid_widths",
                         json.dumps(self.grid.col_widths))
        except Exception:
            pass
        self._update_filter_layout()

    def _load_grid_widths(self):
        try:
            raw = Settings.get_all().get("order_grid_widths", "")
            if raw:
                return json.loads(raw)
        except Exception:
            pass
        return None

    def _on_close(self):
        self._save_grid_widths()
        if self._dirty and not self.read_only:
            answer = messagebox.askyesnocancel(
                u"未保存",
                u"有未保存的修改，要保存吗？\n\n"
                u"  是(Y) = 保存并退出\n"
                u"  否(N) = 不保存退出\n"
                u"  取消   = 继续编辑")
            if answer is None:
                return
            if answer:
                if not self._save(silent=True):
                    return
        self.destroy()
