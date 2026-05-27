# -*- coding: utf-8 -*-
"""Main window with order list and filters."""
try:
    import Tkinter as tk
    import ttk
    import tkMessageBox as messagebox
except ImportError:
    import tkinter as tk
    from tkinter import ttk
    from tkinter import messagebox

import os
import json
import datetime
from db.models import Customer, Order, OrderItem, DeliveryNote, Settings
from ui.customer_window import CustomerWindow
from ui.settings_window import SettingsWindow
from ui.order_window import OrderWindow
from ui.dn_window import DeliveryNoteWindow
from ui.merge_dn_window import MergeDeliveryNoteWindow
from ui.widgets import DateEntry
from logic.order_number import generate_delivery_number


class MainWindow(object):

    PAGE_SIZE = 50

    def __init__(self, root):
        self.root = root
        self.root.deiconify()
        self.root.title(u"海绵厂订单管理系统 v1.0")
        self.root.geometry("1100x650")
        self.root.minsize(900, 500)

        self.current_page = 1
        self.filters = {}
        self._customers_cache = []
        self.view_mode = "orders"  # "orders" or "delivery_notes"
        self._sort_col = None
        self._sort_desc = False

        # Merge mode state
        self._merge_mode = False
        self._merge_checked = set()       # set of checked order IDs
        self._merge_customer_id = None    # locked customer for merge

        self._build_menu()
        self._build_toolbar()
        self._build_merge_bar()
        self._build_tree()
        self._build_statusbar()
        self._load_data()

        self.root.protocol("WM_DELETE_WINDOW", self._on_close)

    # ---- Menu ----
    def _build_menu(self):
        menubar = tk.Menu(self.root)
        self.root.config(menu=menubar)

        order_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label=u"订单管理", menu=order_menu)
        order_menu.add_command(label=u"新建订单", command=self._new_order,
                               accelerator="Ctrl+N")
        order_menu.add_separator()
        order_menu.add_command(label=u"刷新列表", command=self._load_data,
                               accelerator="F5")
        order_menu.add_separator()
        order_menu.add_command(label=u"合并送货", command=self._merge_delivery)
        order_menu.add_separator()
        order_menu.add_command(label=u"退出", command=self.root.quit)

        data_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label=u"基础数据", menu=data_menu)
        data_menu.add_command(label=u"客户管理", command=self._manage_customers)
        data_menu.add_separator()
        data_menu.add_command(label=u"系统设置", command=self._manage_settings)

        help_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label=u"帮助", menu=help_menu)
        help_menu.add_command(label=u"查看错误日志", command=self._view_error_log)

        self.root.bind_all("<Control-n>", lambda e: self._new_order())
        self.root.bind_all("<F5>", lambda e: self._load_data())

    # ---- Toolbar / Filter Bar (normal mode) ----
    def _build_toolbar(self):
        self._normal_bar = ttk.Frame(self.root, padding=(8, 6, 8, 4))
        self._normal_bar.pack(fill=tk.X)

        ttk.Label(self._normal_bar, text=u"客户:").pack(
            side=tk.LEFT, padx=(0, 2))
        self.filter_customer = ttk.Entry(self._normal_bar, width=14)
        self.filter_customer.pack(side=tk.LEFT, padx=(0, 10))
        self.filter_customer.bind("<Return>", lambda e: self._load_data())

        ttk.Label(self._normal_bar, text=u"状态:").pack(
            side=tk.LEFT, padx=(0, 2))
        self.filter_status = ttk.Combobox(
            self._normal_bar, state="readonly", width=8,
            values=[u"全部", u"未完单", u"完单"])
        self.filter_status.current(0)
        self.filter_status.pack(side=tk.LEFT, padx=(0, 10))
        self.filter_status.bind("<<ComboboxSelected>>",
                                lambda e: self._load_data())

        ttk.Label(self._normal_bar, text=u"日期从:").pack(
            side=tk.LEFT, padx=(0, 2))
        self.filter_date_from = DateEntry(self._normal_bar, width=16)
        self.filter_date_from.pack(side=tk.LEFT, padx=(0, 4))

        ttk.Label(self._normal_bar, text=u"到:").pack(
            side=tk.LEFT, padx=(0, 2))
        self.filter_date_to = DateEntry(self._normal_bar, width=16)
        self.filter_date_to.pack(side=tk.LEFT, padx=(0, 8))

        ttk.Label(self._normal_bar, text=u"名称:").pack(
            side=tk.LEFT, padx=(0, 2))
        self.filter_keyword = ttk.Entry(self._normal_bar, width=16)
        self.filter_keyword.pack(side=tk.LEFT, padx=(0, 4))
        self.filter_keyword.bind("<Return>", lambda e: self._load_data())

        ttk.Button(self._normal_bar, text=u"搜索",
                   command=self._search).pack(side=tk.LEFT, padx=2)
        ttk.Button(self._normal_bar, text=u"重置",
                   command=self._reset_filters).pack(side=tk.LEFT, padx=2)

        ttk.Button(self._normal_bar, text=u"合并送货",
                   command=self._merge_delivery).pack(side=tk.RIGHT, padx=4)
        ttk.Button(self._normal_bar, text=u"新建订单",
                   command=self._new_order).pack(side=tk.RIGHT, padx=4)

        self.view_btn = ttk.Button(
            self._normal_bar, text=u"查看送货单列表",
            command=self._toggle_view)
        self.view_btn.pack(side=tk.RIGHT, padx=8)

        # Big view title with colored background
        self._normal_title_frame = tk.Frame(self.root, bg="#1565c0", height=36)
        self._normal_title_frame.pack(fill=tk.X, padx=12, pady=(6, 0))
        self._normal_title_frame.pack_propagate(False)
        self._normal_title = tk.Label(
            self._normal_title_frame, text=u"📋 订单列表",
            font=("Microsoft YaHei", 16, "bold"),
            bg="#1565c0", fg="white", anchor=tk.CENTER)
        self._normal_title.pack(fill=tk.BOTH, expand=True)

    # ---- Merge mode toolbar (hidden normally) ----
    def _build_merge_bar(self):
        # Merge filter bar
        self._merge_bar = ttk.Frame(self.root, padding=(8, 6, 8, 4))

        ttk.Label(self._merge_bar, text=u"客户:").pack(
            side=tk.LEFT, padx=(0, 2))
        self._merge_cust_var = tk.StringVar()
        self._merge_cust_entry = ttk.Entry(
            self._merge_bar, textvariable=self._merge_cust_var, width=14)
        self._merge_cust_entry.pack(side=tk.LEFT, padx=(0, 10))
        self._merge_cust_entry.bind("<Return>",
                                     lambda e: self._load_data())

        ttk.Label(self._merge_bar, text=u"状态:").pack(
            side=tk.LEFT, padx=(0, 2))
        self._merge_stat_cb = ttk.Combobox(
            self._merge_bar, state="readonly", width=8,
            values=[u"全部", u"未完单", u"完单"])
        self._merge_stat_cb.current(1)  # default: 未完单
        self._merge_stat_cb.pack(side=tk.LEFT, padx=(0, 10))
        self._merge_stat_cb.bind("<<ComboboxSelected>>",
                                 lambda e: self._load_data())

        ttk.Label(self._merge_bar, text=u"日期从:").pack(
            side=tk.LEFT, padx=(0, 2))
        self._merge_date_from = DateEntry(self._merge_bar, width=14)
        self._merge_date_from.pack(side=tk.LEFT, padx=(0, 4))

        ttk.Label(self._merge_bar, text=u"到:").pack(
            side=tk.LEFT, padx=(0, 2))
        self._merge_date_to = DateEntry(self._merge_bar, width=14)
        self._merge_date_to.pack(side=tk.LEFT, padx=(0, 8))

        ttk.Button(self._merge_bar, text=u"查询",
                   command=self._search).pack(side=tk.LEFT, padx=2)

        # Right-side action buttons
        self._merge_confirm_btn = ttk.Button(
            self._merge_bar, text=u"确定",
            command=self._merge_confirm)
        self._merge_confirm_btn.pack(side=tk.RIGHT, padx=4)
        ttk.Button(self._merge_bar, text=u"取消",
                   command=self._exit_merge_mode).pack(side=tk.RIGHT, padx=2)

        # Hint label
        self._merge_hint = tk.Label(
            self._merge_bar, text=u"", fg="#006600",
            font=("Microsoft YaHei", 9))
        self._merge_hint.pack(side=tk.RIGHT, padx=12)

        # Merge title (created but not packed yet)
        self._merge_title_frame = tk.Frame(self.root, bg="#2e7d32", height=36)
        self._merge_title_frame.pack_propagate(False)
        self._merge_title = tk.Label(
            self._merge_title_frame, text=u"☑ 合并送货模式 — 勾选要合并的订单，点确定",
            font=("Microsoft YaHei", 14, "bold"),
            bg="#2e7d32", fg="white", anchor=tk.CENTER)
        self._merge_title.pack(fill=tk.BOTH, expand=True)

    # Column -> DB ORDER BY mapping
    _ORDER_SORT_MAP = {
        "display_name": ("o.display_name", u"订单名称"),
        "customer_name": ("c.name", u"客户"),
        "order_date": ("o.order_date", u"订单日期"),
        "delivery_number": ("o.delivery_number", u"送货单号"),
        "status": ("o.status", u"状态"),
    }
    _DN_SORT_MAP = {
        "delivery_number": ("dn.delivery_number", u"送货单号"),
        "customer_name": ("c.name", u"客户"),
        "delivery_date": ("dn.delivery_date", u"送货日期"),
        "order_display_name": ("o.display_name", u"所属订单"),
        "status": ("dn.status", u"状态"),
    }

    _MERGE_COLS = ("_chk", "display_name", "customer_name", "order_date",
                   "delivery_number", "status", "item_count")

    # ---- Treeview ----
    def _build_tree(self):
        self._tree_container = ttk.Frame(self.root)
        self._tree_container.pack(fill=tk.BOTH, expand=True, padx=8, pady=2)

        style = ttk.Style()
        style.configure("Treeview", rowheight=28,
                        font=("Microsoft YaHei", 10))
        style.configure("Treeview.Heading",
                        font=("Microsoft YaHei", 10, "bold"))

        self._order_cols = ("status", "display_name", "customer_name",
                            "order_date", "delivery_number", "item_count")
        self._dn_cols = ("delivery_number", "customer_name", "delivery_date",
                         "order_display_name", "status")

        cols = self._order_cols
        self.tree = ttk.Treeview(
            self._tree_container, columns=cols, show="headings",
            selectmode="browse")

        self._setup_order_columns()

        scrollbar = ttk.Scrollbar(self._tree_container, orient=tk.VERTICAL,
                                  command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        self.tree.bind("<Button-1>", self._on_single_click)
        self.tree.bind("<Button-3>", self._context_menu)

        # Tag for merge-checked rows
        self.tree.tag_configure("merge_checked", background="#c8e6c9")
        self.tree.tag_configure("merge_diff_cust", foreground="#cccccc")

    # ---- Statusbar / Pagination ----
    def _build_statusbar(self):
        bottom = ttk.Frame(self.root, padding=(8, 2, 8, 4))
        bottom.pack(fill=tk.X)

        self.status_label = ttk.Label(bottom, text=u"就绪")
        self.status_label.pack(side=tk.LEFT)

        ttk.Button(bottom, text=u"上一页",
                   command=self._prev_page).pack(side=tk.RIGHT, padx=2)
        ttk.Button(bottom, text=u"下一页",
                   command=self._next_page).pack(side=tk.RIGHT, padx=2)
        self.page_label = ttk.Label(bottom, text="")
        self.page_label.pack(side=tk.RIGHT, padx=6)

    # ---- Data Loading ----
    def _build_filters(self):
        filters = {}

        if self._merge_mode:
            # Use merge bar controls
            cust_name = self._merge_cust_var.get().strip()
            stat = self._merge_stat_cb.get()
            df = self._merge_date_from.get()
            dt = self._merge_date_to.get()
        else:
            cust_name = self.filter_customer.get().strip()
            stat = self.filter_status.get()
            df = self.filter_date_from.get()
            dt = self.filter_date_to.get()

        # Customer fuzzy search
        if cust_name and cust_name != u"全部":
            for c in self._customers_cache:
                if c["name"] == cust_name:
                    filters["customer_id"] = c["id"]
                    break
            else:
                matching_ids = [
                    c["id"] for c in self._customers_cache
                    if cust_name.lower() in c["name"].lower()
                ]
                if matching_ids:
                    filters["customer_ids"] = matching_ids

        # Status filter
        if self.view_mode == "orders":
            if stat == u"未完单":
                filters["status_not_in"] = ("completed", "dn_generated")
            elif stat == u"完单":
                filters["status_in"] = ("completed", "dn_generated")
        else:
            if stat == u"未完单":
                filters["status_not_in"] = ("completed",)
            elif stat == u"完单":
                filters["status_in"] = ("completed",)

        # Validate date format (YYYY-MM-DD) before applying
        if df and len(df) >= 10 and df[4] == "-" and df[7] == "-":
            filters["date_from"] = df[:10]
        else:
            # 默认只拉取近 30 天数据，防止历史数据过多卡死界面
            import datetime
            filters["date_from"] = (
                datetime.date.today() - datetime.timedelta(days=30)
            ).strftime("%Y-%m-%d")
        if dt and len(dt) >= 10 and dt[4] == "-" and dt[7] == "-":
            filters["date_to"] = dt[:10]

        if not self._merge_mode:
            kw = self.filter_keyword.get().strip()
            if kw:
                filters["keyword"] = kw
        return filters

    def _load_data(self):
        self._customers_cache = Customer.get_all()

        for item in self.tree.get_children():
            self.tree.delete(item)

        self.filters = self._build_filters()
        order_by = self._get_order_by()

        if self.view_mode == "orders":
            result = Order.get_all(
                page=self.current_page, page_size=self.PAGE_SIZE,
                filters=self.filters, order_by=order_by)

            # In merge mode, determine locked customer from checked orders
            if self._merge_mode:
                self._merge_customer_id = self._resolve_merge_customer()

            for row in result["items"]:
                oid = row["id"]
                if row["status"] in ("completed", "dn_generated"):
                    sc, tag = u"完单", "completed"
                else:
                    sc, tag = "", ""

                if self._merge_mode:
                    chk = u"☑" if oid in self._merge_checked else u"☐"
                    vals = (chk, row["display_name"], row["customer_name"],
                            row["order_date"], row["delivery_number"] or "",
                            sc, Order.get_item_count(row["id"]))

                    # Determine row tags
                    tags = []
                    if oid in self._merge_checked:
                        tags.append("merge_checked")
                    elif (self._merge_customer_id is not None and
                          row["customer_id"] != self._merge_customer_id):
                        tags.append("merge_diff_cust")
                    elif tag:
                        tags.append(tag)
                    self.tree.insert("", tk.END, iid=str(oid),
                                     values=vals, tags=tuple(tags))
                else:
                    self.tree.insert("", tk.END, iid=str(row["id"]), values=(
                        sc,
                        row["display_name"], row["customer_name"],
                        row["order_date"], row["delivery_number"] or "",
                        Order.get_item_count(row["id"])), tags=(tag,))
        else:
            result = DeliveryNote.get_all(
                page=self.current_page, page_size=self.PAGE_SIZE,
                filters=self.filters, order_by=order_by)
            for row in result["items"]:
                sc = u"完单" if row["status"] == "completed" else ""
                tag = "dn_completed" if row["status"] == "completed" else ""
                self.tree.insert("", tk.END,
                                 iid="dn_{}".format(row["id"]), values=(
                    row["delivery_number"], row["customer_name"],
                    row["delivery_date"], row["order_display_name"], sc),
                    tags=(tag,))

        self.page_label.config(
            text=u"第{}页/共{}页 ({}条)".format(
                result["page"], result["total_pages"], result["total"]))
        self.status_label.config(
            text=u"共 {} 条记录".format(result["total"]))

        self._update_headings()
        self._update_merge_hint()

    # ---- Merge mode helpers ----
    def _resolve_merge_customer(self):
        """Return the customer_id if all checked orders share one, else None."""
        if not self._merge_checked:
            return None
        cust_ids = set()
        for oid in self._merge_checked:
            order = Order.get_by_id(oid)
            if order:
                cust_ids.add(order["customer_id"])
        return cust_ids.pop() if len(cust_ids) == 1 else None

    def _update_merge_hint(self):
        if not self._merge_mode:
            return
        count = len(self._merge_checked)
        if count == 0:
            self._merge_hint.config(
                text=u"请勾选要合并的订单", fg="#888888")
        else:
            names = []
            for oid in sorted(self._merge_checked):
                o = Order.get_by_id(oid)
                if o:
                    names.append(o["display_name"])
            preview = u", ".join(names[:5])
            if len(names) > 5:
                preview += u"..."
            self._merge_hint.config(
                text=u"已选{}单: {}".format(count, preview), fg="#006600")

    def _enter_merge_mode(self):
        """Switch UI to merge selection mode."""
        if self.view_mode != "orders":
            # Force orders view
            self._save_col_widths()
            self.view_mode = "orders"
            self.view_btn.config(text=u"查看送货单列表")
            self._normal_title_frame.config(bg="#1565c0")
            self._normal_title.config(text=u"📋 订单列表", bg="#1565c0")

        self._merge_mode = True
        self._merge_checked = set()
        self._merge_customer_id = None
        self.current_page = 1
        self._sort_col = None
        self._sort_desc = False

        # Hide normal UI
        self._normal_bar.pack_forget()
        self._normal_title_frame.pack_forget()

        # Show merge UI before the tree container
        self._merge_bar.pack(fill=tk.X, before=self._tree_container)
        self._merge_title_frame.pack(fill=tk.X, padx=12, pady=(6, 0),
                                     before=self._tree_container)

        # Reset filters
        self._merge_cust_var.set("")
        self._merge_stat_cb.current(1)  # default: 未完单
        self._merge_date_from.delete(0, tk.END)
        self._merge_date_to.delete(0, tk.END)

        # Setup merge columns (add checkbox column)
        self._setup_merge_columns()

        # Disable view toggle during merge
        self.view_btn.config(state=tk.DISABLED)

        self._load_data()

    def _exit_merge_mode(self):
        """Restore normal UI."""
        self._merge_mode = False
        self._merge_checked = set()
        self._merge_customer_id = None

        # Hide merge UI
        self._merge_bar.pack_forget()
        self._merge_title_frame.pack_forget()

        # Show normal UI before tree container
        self._normal_bar.pack(fill=tk.X, before=self._tree_container)
        self._normal_title_frame.pack(fill=tk.X, padx=12, pady=(6, 0),
                                      before=self._tree_container)

        # Restore order columns
        self._setup_order_columns()

        # Re-enable view toggle
        self.view_btn.config(state=tk.NORMAL)

        self._load_data()

    def _setup_merge_columns(self):
        """Configure treeview columns for merge mode (with checkbox col)."""
        self.tree["columns"] = self._MERGE_COLS
        self.tree.heading("_chk", text=u"选")
        self.tree.heading("display_name", text=u"订单名称")
        self.tree.heading("customer_name", text=u"客户")
        self.tree.heading("order_date", text=u"订单日期")
        self.tree.heading("delivery_number", text=u"送货单号")
        self.tree.heading("status", text=u"状态")
        self.tree.heading("item_count", text=u"明细数")
        self.tree.column("_chk", width=36, minwidth=36, anchor=tk.CENTER)
        self.tree.column("display_name", width=160, minwidth=100)
        self.tree.column("customer_name", width=100, minwidth=60)
        self.tree.column("order_date", width=100, minwidth=80)
        self.tree.column("delivery_number", width=130, minwidth=100)
        self.tree.column("status", width=70, minwidth=60, anchor=tk.CENTER)
        self.tree.column("item_count", width=60, minwidth=50, anchor=tk.CENTER)

    def _merge_toggle_order(self, oid):
        """Toggle an order's check state. Returns True if toggled on."""
        order = Order.get_by_id(oid)
        if not order:
            return False

        if oid in self._merge_checked:
            self._merge_checked.discard(oid)
            self._merge_customer_id = self._resolve_merge_customer()
            self._load_data()
            return False

        # Check same customer constraint
        if self._merge_checked:
            self._merge_customer_id = self._resolve_merge_customer()
            if (self._merge_customer_id is not None and
                    order["customer_id"] != self._merge_customer_id):
                messagebox.showwarning(
                    u"客户不同",
                    u"只能合并同一客户的订单！\n\n"
                    u"已选客户: {}".format(
                        Customer.get_by_id(self._merge_customer_id)["name"]
                        if self._merge_customer_id else u"未知"))
                return False

        self._merge_checked.add(oid)
        self._merge_customer_id = self._resolve_merge_customer()
        self._load_data()
        return True

    def _merge_confirm(self):
        """Validate selection and open merge window."""
        if not self._merge_checked:
            messagebox.showwarning(u"提示", u"请至少勾选一个订单。")
            return

        if len(self._merge_checked) < 1:
            return

        # Build result list
        selected_orders = []
        for oid in sorted(self._merge_checked):
            o = Order.get_by_id(oid)
            if o:
                selected_orders.append(o)

        if not selected_orders:
            return

        # Verify all same customer
        cust_ids = set(o["customer_id"] for o in selected_orders)
        if len(cust_ids) > 1:
            messagebox.showwarning(
                u"客户不同",
                u"所选订单包含不同客户，请只选择同一客户的订单。")
            return

        # Verify all have items
        for o in selected_orders:
            items = OrderItem.get_by_order(o["id"])
            if not items:
                messagebox.showwarning(
                    u"无明细",
                    u"订单「{}」没有明细数据，不能合并。\n"
                    u"请先打开该订单添加物料明细。".format(o["display_name"]))
                return

        # Open merge window
        MergeDeliveryNoteWindow(self.root, selected_orders)

        # Exit merge mode
        self._exit_merge_mode()

    # ---- Sort ----
    def _sort_by(self, col_key):
        if self._merge_mode and col_key == "_chk":
            return  # don't sort checkbox column
        if self._sort_col == col_key:
            self._sort_desc = not self._sort_desc
        else:
            self._sort_col = col_key
            self._sort_desc = False
        self.current_page = 1
        self._load_data()

    def _get_order_by(self):
        if not self._sort_col:
            return None
        smap = self._ORDER_SORT_MAP if self.view_mode == "orders" else self._DN_SORT_MAP
        entry = smap.get(self._sort_col)
        if not entry:
            return None
        direction = "DESC" if self._sort_desc else "ASC"
        return "{} {}".format(entry[0], direction)

    def _update_headings(self):
        arrow = u" ▼" if self._sort_desc else u" ▲"
        if self._merge_mode:
            smap = dict(self._ORDER_SORT_MAP)
            smap["_chk"] = ("o.id", u"选")
        else:
            smap = self._ORDER_SORT_MAP if self.view_mode == "orders" else self._DN_SORT_MAP
        for col_key, (db_col, label) in smap.items():
            try:
                if col_key == self._sort_col:
                    self.tree.heading(col_key, text=label + arrow)
                else:
                    self.tree.heading(col_key, text=label)
            except Exception:
                pass

    # ---- Column widths persistence ----
    def _save_col_widths(self):
        widths = {}
        cols = self.tree["columns"]
        for col in cols:
            try:
                w = self.tree.column(col, "width")
                if w:
                    widths[col] = w
            except Exception:
                pass
        key = "order_col_widths" if self.view_mode == "orders" else "dn_col_widths"
        try:
            Settings.set(key, json.dumps(widths))
        except Exception:
            pass

    def _restore_col_widths(self):
        key = "order_col_widths" if self.view_mode == "orders" else "dn_col_widths"
        try:
            raw = Settings.get_all().get(key, "")
            if raw:
                widths = json.loads(raw)
                for col, w in widths.items():
                    try:
                        self.tree.column(col, width=w)
                    except Exception:
                        pass
        except Exception:
            pass

    def _on_close(self):
        self._save_col_widths()
        self.root.destroy()

    def _search(self):
        """Search with page reset to 1."""
        self.current_page = 1
        self._load_data()

    def _reset_filters(self):
        self.filter_customer.delete(0, tk.END)
        self.filter_status.set(u"全部")
        self.filter_date_from.delete(0, tk.END)
        self.filter_date_to.delete(0, tk.END)
        self.filter_keyword.delete(0, tk.END)
        self.current_page = 1
        self._load_data()

    def _prev_page(self):
        if self.current_page > 1:
            self.current_page -= 1
            self._load_data()

    def _next_page(self):
        self.current_page += 1
        self._load_data()

    # ---- Actions ----
    def _new_order(self):
        OrderWindow(self.root, on_save_callback=self._load_data)

    def _setup_order_columns(self):
        self.tree["columns"] = self._order_cols
        self.tree.heading("display_name", text=u"订单名称")
        self.tree.heading("customer_name", text=u"客户")
        self.tree.heading("order_date", text=u"订单日期")
        self.tree.heading("delivery_number", text=u"送货单号")
        self.tree.heading("status", text=u"状态")
        self.tree.heading("item_count", text=u"明细数")
        self.tree.column("display_name", width=160, minwidth=100)
        self.tree.column("customer_name", width=100, minwidth=60)
        self.tree.column("order_date", width=100, minwidth=80)
        self.tree.column("delivery_number", width=140, minwidth=100)
        self.tree.column("status", width=80, minwidth=60, anchor=tk.CENTER)
        self.tree.column("item_count", width=60, minwidth=50, anchor=tk.CENTER)
        self._restore_col_widths()

    def _setup_dn_columns(self):
        self.tree["columns"] = self._dn_cols
        self.tree.heading("delivery_number", text=u"送货单号")
        self.tree.heading("customer_name", text=u"客户")
        self.tree.heading("delivery_date", text=u"送货日期")
        self.tree.heading("order_display_name", text=u"所属订单")
        self.tree.heading("status", text=u"状态")
        self.tree.column("delivery_number", width=150, minwidth=100)
        self.tree.column("customer_name", width=100, minwidth=60)
        self.tree.column("delivery_date", width=100, minwidth=80)
        self.tree.column("order_display_name", width=160, minwidth=100)
        self.tree.column("status", width=80, minwidth=60, anchor=tk.CENTER)
        self._restore_col_widths()

    def _toggle_view(self):
        self._save_col_widths()
        if self.view_mode == "orders":
            self.view_mode = "delivery_notes"
            self.view_btn.config(text=u"查看订单列表")
            self._normal_title_frame.config(bg="#e65100")
            self._normal_title.config(text=u"🚚 送货单列表", bg="#e65100")
            self._setup_dn_columns()
        else:
            self.view_mode = "orders"
            self.view_btn.config(text=u"查看送货单列表")
            self._normal_title_frame.config(bg="#1565c0")
            self._normal_title.config(text=u"📋 订单列表", bg="#1565c0")
            self._setup_order_columns()
        self._sort_col = None
        self._sort_desc = False
        self.current_page = 1
        self._load_data()

    # ---- Click handling ----
    def _on_single_click(self, event):
        region = self.tree.identify_region(event.x, event.y)
        if region == "heading":
            col = self.tree.identify_column(event.x)
            col_idx = int(col.replace("#", "")) - 1
            cols = self.tree["columns"]
            if col_idx < len(cols):
                self._sort_by(cols[col_idx])
            return
        if region != "cell":
            return
        col = self.tree.identify_column(event.x)
        col_idx = int(col.replace("#", "")) - 1
        row = self.tree.identify_row(event.y)
        if not row:
            return

        # Merge mode: checkbox column toggles, other columns open order
        if self._merge_mode:
            if row.startswith("dn_"):
                return
            oid = int(row)
            if col_idx == 0:  # checkbox column
                self._merge_toggle_order(oid)
                return
            else:
                self._open_order(oid)
                return

        # Normal mode
        if self.view_mode == "orders":
            oid = int(row)
            if col_idx == 3:  # delivery_number column
                self._open_dn(oid)
            else:
                self._open_order(oid)
        else:
            if row.startswith("dn_"):
                dn_id = int(row[3:])
                self._open_dn_by_id(dn_id)

    def _open_dn_by_id(self, dn_id):
        dn = DeliveryNote.get_by_id(dn_id)
        if not dn:
            return
        DeliveryNoteWindow(
            self.root, dn_id=dn_id,
            on_save_callback=self._load_data,
            read_only=(dn["status"] == "completed"))

    def _open_order(self, order_id=None):
        if order_id is None:
            sel = self.tree.selection()
            if not sel:
                return
            order_id = int(sel[0])

        order = Order.get_by_id(order_id)
        if not order:
            return

        is_done = order["status"] in ("completed", "dn_generated")
        OrderWindow(self.root, order_id=order_id,
                    on_save_callback=self._load_data,
                    read_only=is_done)

    # ---- Context menu ----
    def _context_menu(self, event):
        item = self.tree.identify_row(event.y)
        if not item:
            return
        self.tree.selection_set(item)

        if self._merge_mode:
            return  # no context menu in merge mode

        if self.view_mode == "orders":
            oid = int(item)
            order = Order.get_by_id(oid)
            menu = tk.Menu(self.root, tearoff=0)
            menu.add_command(label=u"📋 订单明细（打标签）",
                             command=lambda: self._open_order(oid))
            menu.add_command(label=u"📄 送货单明细",
                             command=lambda: self._open_dn(oid))
            menu.add_separator()
            if order:
                if order["status"] in ("completed", "dn_generated"):
                    menu.add_command(label=u"↩ 取消完单",
                                     command=lambda: self._toggle_done(oid, False))
                else:
                    menu.add_command(label=u"✓ 标记完单",
                                     command=lambda: self._toggle_done(oid, True))
                menu.add_separator()
                if order["status"] not in ("completed", "dn_generated"):
                    items = OrderItem.get_by_order(oid)
                    if items:
                        menu.add_command(label=u"→ 推送到送货单",
                                         command=lambda: self._push_to_dn(oid))
            menu.add_separator()
            menu.add_command(label=u"删除订单",
                             command=lambda: self._delete_order(oid))
            menu.tk_popup(event.x_root, event.y_root)
        else:
            if item.startswith("dn_"):
                dn_id = int(item[3:])
                dn = DeliveryNote.get_by_id(dn_id)
                menu = tk.Menu(self.root, tearoff=0)
                menu.add_command(label=u"📄 打开送货单",
                                 command=lambda: self._open_dn_by_id(dn_id))
                menu.add_separator()
                if dn and dn["status"] != "completed":
                    menu.add_command(label=u"✓ 标记完单",
                                     command=lambda: self._toggle_dn_done(dn_id))
                menu.add_separator()
                menu.add_command(label=u"删除送货单",
                                 command=lambda: self._delete_dn(dn_id))
                menu.tk_popup(event.x_root, event.y_root)

    def _toggle_dn_done(self, dn_id):
        DeliveryNote.update(dn_id, status="completed")
        self._load_data()

    def _delete_dn(self, dn_id):
        dn = DeliveryNote.get_by_id(dn_id)
        if dn and dn["status"] == "completed":
            if not messagebox.askyesno(
                    u"无法删除",
                    u"送货单 '{}' 已完成，不能删除。\n\n"
                    u"要先取消完单状态吗？".format(dn.get("delivery_number", ""))):
                return
            DeliveryNote.update(dn_id, status="draft")
        if messagebox.askyesno(u"确认删除", u"确定要删除此送货单吗？"):
            DeliveryNote.delete(dn_id)
            self._load_data()

    def _open_dn(self, order_id):
        """Open delivery note for an order."""
        order = Order.get_by_id(order_id)
        is_done = order["status"] in ("completed", "dn_generated") if order else False
        result = DeliveryNote.get_by_order(order_id, page=1, page_size=1)
        if result["items"]:
            dn = result["items"][0]
            read_only = is_done or dn["status"] == "completed"
            DeliveryNoteWindow(
                self.root, dn_id=dn["id"],
                on_save_callback=self._load_data,
                read_only=read_only)
        else:
            items = OrderItem.get_by_order(order_id)
            if not items:
                messagebox.showwarning(
                    u"无明细",
                    u"此订单没有任何明细数据，不能生成送货单。\n"
                    u"请先打开订单添加物料明细。")
                return
            if messagebox.askyesno(u"提示",
                                   u"此订单还没有送货单，要新建吗？"):
                self._push_to_dn(order_id)

    def _toggle_done(self, order_id, mark_done):
        if mark_done:
            Order.update(order_id, status="completed")
        else:
            order = Order.get_by_id(order_id)
            if order.get("delivery_number"):
                Order.update(order_id, status="qty_filled")
            else:
                Order.update(order_id, status="draft")
        self._load_data()

    def _push_to_dn(self, order_id):
        """Push order to a new delivery note."""
        if not messagebox.askyesno(
                u"确认",
                u"将订单推送到新的送货单？\n物料明细将自动带入。"):
            return

        order = Order.get_by_id(order_id)
        if not order:
            return

        items = OrderItem.get_by_order(order_id)
        if not items:
            messagebox.showwarning(
                u"无明细",
                u"此订单没有任何明细数据，不能生成送货单。\n"
                u"请先打开订单添加物料明细。")
            return

        dn_number = generate_delivery_number(
            order["customer_id"],
            order.get("delivery_date") or None)
        if not dn_number:
            messagebox.showwarning(u"错误", u"无法生成送货单号。")
            return

        Order.update(order_id, status="qty_filled")
        dn_id = DeliveryNote.create(
            order_id, dn_number,
            delivery_date=datetime.date.today().strftime("%Y-%m-%d"))

        for i, item in enumerate(items):
            DNItem.create(
                delivery_note_id=dn_id,
                order_item_id=item["id"],
                model_number=item["model_number"],
                item_code=item.get("order_number", ""),
                customer_code=item.get("customer_code", ""),
                product_name=item["product_name"],
                color_name=item["color_name"],
                quantity=item["quantity"],
                notes=item["notes"] if item["notes"] else "",
                mfg_number=item.get("mfg_number", ""),
                sort_order=i)

        self._load_data()
        DeliveryNoteWindow(
            self.root, dn_id=dn_id,
            on_save_callback=self._load_data)

    def _generate_dn(self, order_id):
        messagebox.showinfo(
            u"提示",
            u"Excel导出和打印功能将在后续版本实现。")
        Order.update(order_id, status="dn_generated")
        self._load_data()

    def _export_excel(self, order_id):
        self._generate_dn(order_id)

    def _delete_order(self, order_id):
        order = Order.get_by_id(order_id)
        if not order:
            return
        if order["status"] in ("completed", "dn_generated"):
            if not messagebox.askyesno(
                    u"无法删除",
                    u"订单 '{}' 已完成，不能删除。\n\n"
                    u"要先取消完单状态吗？".format(order["display_name"])):
                return
            Order.update(order_id, status="draft")
        if not messagebox.askyesno(
                u"确认删除",
                u"确定要删除订单 '{}' 吗？".format(order["display_name"])):
            return
        Order.delete(order_id)
        self._load_data()

    def _manage_customers(self):
        CustomerWindow(self.root, on_save_callback=self._load_data)

    def _manage_settings(self):
        SettingsWindow(self.root, on_save_callback=self._load_data)

    def _view_error_log(self):
        from logic.logger import get_log_path, get_recent_errors
        log_path = get_log_path()
        if not os.path.exists(log_path):
            messagebox.showinfo(u"错误日志",
                                u"暂无错误日志。\n\n日志位置: {}".format(log_path))
            return
        try:
            os.startfile(log_path)
        except Exception:
            text = get_recent_errors(50)
            if not text.strip():
                text = u"(日志为空)"
            top = tk.Toplevel(self.root)
            top.title(u"错误日志")
            top.geometry("800x500")
            t = tk.Text(top, wrap=tk.WORD, font=("Consolas", 9))
            t.pack(fill=tk.BOTH, expand=True, padx=4, pady=4)
            t.insert("1.0", text)
            t.config(state=tk.DISABLED)

    def _merge_delivery(self):
        """Enter merge delivery mode on the order list."""
        if self.view_mode != "orders":
            messagebox.showinfo(
                u"提示",
                u"合并送货功能请在「订单列表」中使用。\n\n"
                u"请先切换到订单列表视图。")
            return
        self._enter_merge_mode()

    def _manage_products(self):
        messagebox.showinfo(u"提示", u"产品管理将在后续版本实现。")
