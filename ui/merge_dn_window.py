# -*- coding: utf-8 -*-
"""Merge delivery note window – combine items from multiple orders into one DN."""
try:
    import Tkinter as tk
    import ttk
    import tkMessageBox as messagebox
    import tkFont as tkfont
except ImportError:
    import tkinter as tk
    from tkinter import ttk
    from tkinter import messagebox
    import tkinter.font as tkfont

import threading
import datetime

from db.models import (Customer, Order, OrderItem, Product, Color,
                       DeliveryNote, DNItem, DeliveryNoteOrder,
                       deduplicate_dn_items)
from logic.order_number import generate_delivery_number


ROW_H = 50       # match ExcelGrid row height
SEP_H = 30       # separator bar height
HEADER_H = 36    # header height


class MergeDeliveryNoteWindow(tk.Toplevel):
    """Show items from multiple orders with gray separators, checkboxes."""

    COLUMNS = [u"选", u"客户号", u"订单号", u"制单号", u"款号",
               u"品名", u"颜色",
               u"数量", u"备注"]
    MIN_WIDTHS = [42, 70, 80, 80, 70, 180, 70, 70, 100]
    CHECK_COL = 0

    def __init__(self, parent, orders):
        tk.Toplevel.__init__(self, parent)
        self._orders = orders  # list of order dicts, all same customer
        self._customer = Customer.get_by_id(orders[0]["customer_id"])

        self.title(u"合并送货 – {}".format(self._customer["name"]))
        sw = self.winfo_screenwidth()
        sh = self.winfo_screenheight()
        w = int(sw * 0.92)
        h = int(sh * 0.82)
        self.geometry("{}x{}".format(w, h))
        self.minsize(900, 520)
        self.transient(parent)

        # item data: [(order, order_item_dict, checked_bool), ...]
        self._items = []
        for o in orders:
            oi_list = OrderItem.get_by_order(o["id"])
            for oi in oi_list:
                self._items.append([o, oi, True])  # default all checked

        # Calculate column widths based on content
        self.col_widths = self._calc_widths()

        self._build()
        self._draw_all()

        self.protocol("WM_DELETE_WINDOW", self.destroy)

    def _calc_widths(self):
        """Auto-calculate column widths based on content."""
        try:
            cell_font = tkfont.Font(family="Microsoft YaHei", size=10)
        except Exception:
            return list(self.MIN_WIDTHS)
        hdr_font = cell_font.copy()
        hdr_font.config(weight="bold")

        widths = list(self.MIN_WIDTHS)
        for ci, header in enumerate(self.COLUMNS):
            # Header width
            max_w = hdr_font.measure(header) + 20
            # Data width
            for o, oi, checked in self._items:
                text = self._col_text(ci, o, oi)
                w = cell_font.measure(str(text)) + 20
                if w > max_w:
                    max_w = w
            # Clamp between min and max
            max_col_w = 400 if ci == 4 else 200  # 品名 can be wider
            widths[ci] = max(self.MIN_WIDTHS[ci], min(max_w, max_col_w))
        return widths

    def _col_text(self, ci, order, oi):
        """Return display text for a given column and item."""
        if ci == 0:
            return u"☑"
        elif ci == 1:
            return oi.get("customer_code", "")  # 客户号
        elif ci == 2:
            return oi.get("order_number", "")  # 订单号
        elif ci == 3:
            return oi.get("mfg_number", "")  # 制单号
        elif ci == 4:
            return oi.get("model_number", "")
        elif ci == 5:
            return oi.get("product_name", "")
        elif ci == 6:
            return oi.get("color_name", "")
        elif ci == 7:
            return self._fmt_qty(oi)
        elif ci == 8:
            return oi.get("notes", "")
        return ""

    def _build(self):
        # -- Title bar --
        title_bar = tk.Frame(self, bg="#1a5276", height=38)
        title_bar.pack(fill=tk.X)
        title_bar.pack_propagate(False)
        order_names = u", ".join(o["display_name"] for o in self._orders)
        tk.Label(
            title_bar,
            text=u"  合并送货 – {}  |  来源订单: {}".format(
                self._customer["name"], order_names),
            font=("Microsoft YaHei", 12, "bold"),
            bg="#1a5276", fg="white", anchor=tk.W).pack(
                side=tk.LEFT, fill=tk.BOTH, expand=True)

        # -- Scrollable content area --
        container = ttk.Frame(self)
        container.pack(fill=tk.BOTH, expand=True, padx=8, pady=(4, 0))

        self._canvas = tk.Canvas(container, bg="#f0f0f0",
                                 highlightthickness=0, bd=0)
        sb = ttk.Scrollbar(container, orient=tk.VERTICAL,
                          command=self._canvas.yview)
        self._canvas.configure(yscrollcommand=sb.set)
        sb.pack(side=tk.RIGHT, fill=tk.Y)
        self._canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        self._content_frame = tk.Frame(self._canvas, bg="#f0f0f0")
        self._cw = self._canvas.create_window(
            (0, 0), window=self._content_frame, anchor=tk.NW)

        self._canvas.bind("<Configure>", lambda e: self._canvas.itemconfig(
            self._cw, width=e.width))
        self._content_frame.bind(
            "<Configure>",
            lambda e: self._canvas.configure(
                scrollregion=self._canvas.bbox("all")))

        def _on_mousewheel(e):
            self._canvas.yview_scroll(int(-1 * (e.delta / 120)), "units")
        self._canvas.bind("<MouseWheel>", _on_mousewheel)

        # -- Footer --
        ft = ttk.Frame(self, padding=(8, 4, 8, 8))
        ft.pack(fill=tk.X)

        ttk.Button(ft, text=u"全选", command=self._select_all).pack(
            side=tk.LEFT, padx=2)
        ttk.Button(ft, text=u"取消全选", command=self._deselect_all).pack(
            side=tk.LEFT, padx=2)

        self._gen_btn = ttk.Button(ft, text=u"生成送货单",
                                   command=self._generate_dn)
        self._gen_btn.pack(side=tk.RIGHT, padx=4)
        ttk.Button(ft, text=u"取消",
                   command=self.destroy).pack(side=tk.RIGHT, padx=2)

    def _draw_all(self):
        for w in self._content_frame.winfo_children():
            w.destroy()

        total_w = sum(self.col_widths)

        # Fonts
        hdr_font = tkfont.Font(family="Microsoft YaHei", size=10,
                               weight="bold")
        cell_font = tkfont.Font(family="Microsoft YaHei", size=10)
        sep_font = tkfont.Font(family="Microsoft YaHei", size=10,
                               weight="bold")

        # Column header
        x = 0
        for ci, header in enumerate(self.COLUMNS):
            w = self.col_widths[ci]
            lbl = tk.Label(self._content_frame, text=header,
                           font=hdr_font, bg="#d0d0d0", fg="#333",
                           relief="groove", borderwidth=1,
                           anchor=tk.CENTER,
                           justify=tk.CENTER)
            lbl.place(x=x, y=0, width=w, height=HEADER_H)
            # Center text vertically by using a compound approach
            x += w

        # Draw rows grouped by order
        y = HEADER_H

        for order in self._orders:
            # ---- Separator bar ----
            sep_text = u"▬▬ {}  ({})".format(
                order["display_name"], order.get("order_date", ""))
            sep_frame = tk.Frame(self._content_frame, bg="#9e9e9e",
                                 height=SEP_H)
            sep_frame.place(x=0, y=y, width=total_w, height=SEP_H)
            sep_frame.pack_propagate(False)
            sep_lbl = tk.Label(sep_frame, text=sep_text,
                               font=sep_font, bg="#9e9e9e", fg="white",
                               anchor=tk.W)
            sep_lbl.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=12)
            y += SEP_H

            # ---- Data rows for this order ----
            order_items = [(idx, item) for idx, item in enumerate(self._items)
                          if item[0]["id"] == order["id"]]

            for item_idx, (gi, (o, oi, checked)) in enumerate(order_items):
                bg = "#ffffff" if item_idx % 2 == 0 else "#f6f6f6"
                chk_char = u"☑" if checked else u"☐"
                fg_chk = "#006600" if checked else "#aaaaaa"

                qty_str = self._fmt_qty(oi)

                vals = [
                    (chk_char, fg_chk, tk.CENTER),              # 选
                    (oi.get("customer_code", ""),                # 客户号
                     "#333333", tk.CENTER),
                    (oi.get("order_number", ""),                 # 订单号
                     "#333333", tk.CENTER),
                    (oi.get("mfg_number", ""),                   # 制单号
                     "#333333", tk.CENTER),
                    (oi.get("model_number", ""),                 # 款号
                     "#333333", tk.CENTER),
                    (oi.get("product_name", ""),                 # 品名
                     "#333333", tk.W),
                    (oi.get("color_name", ""),                   # 颜色
                     "#333333", tk.CENTER),
                    (qty_str, "#333333", tk.CENTER),             # 数量
                    (oi.get("notes", ""), "#333333", tk.W),      # 备注
                ]

                x = 0
                for ci, (text, fg, anchor) in enumerate(vals):
                    w = self.col_widths[ci]
                    lbl = tk.Label(self._content_frame, text=str(text),
                                   font=cell_font, bg=bg, fg=fg,
                                   anchor=anchor,
                                   justify=tk.CENTER,
                                   relief="solid",
                                   borderwidth=1)
                    lbl.place(x=x, y=y, width=w, height=ROW_H)
                    if ci == self.CHECK_COL:
                        lbl.bind("<Button-1>",
                                 lambda e, idx=gi: self._toggle_item(idx))
                    x += w
                y += ROW_H

        # Update canvas scroll region
        total_h = y + 20
        self._content_frame.config(width=total_w, height=total_h)
        self._canvas.configure(scrollregion=(0, 0, total_w, total_h))

    @staticmethod
    def _fmt_qty(oi):
        formula = oi.get("quantity_formula") or ""
        if formula:
            return formula
        q = oi.get("quantity", 0)
        if q:
            return str(int(q) if q == int(q) else q)
        return ""

    def _toggle_item(self, global_idx):
        if global_idx < len(self._items):
            self._items[global_idx][2] = not self._items[global_idx][2]
            self._draw_all()

    def _select_all(self):
        for item in self._items:
            item[2] = True
        self._draw_all()

    def _deselect_all(self):
        for item in self._items:
            item[2] = False
        self._draw_all()

    def _get_checked_items(self):
        return [(o, oi) for o, oi, checked in self._items if checked]

    def _generate_dn(self):
        checked = self._get_checked_items()
        if not checked:
            messagebox.showwarning(u"提示", u"请至少勾选一行物料。")
            return

        count = len(checked)
        if count > 12:
            if not messagebox.askyesno(
                    u"超出模板限制",
                    u"物料行数（{}行）超过打印模板12行限制！\n\n"
                    u"是否继续生成？".format(count)):
                return

        order_names = u"、".join(
            sorted(set(o["display_name"] for o, _ in checked)))
        if not messagebox.askyesno(
                u"确认生成",
                u"将以下 {} 行物料合并生成一张送货单：\n\n"
                u"来源订单: {}\n客户: {}\n行数: {}".format(
                    count, order_names, self._customer["name"], count)):
            return

        master = self.master
        customer = self._customer
        orders = self._orders
        checked_data = checked

        self._gen_btn.config(state=tk.DISABLED)
        self.config(cursor="watch")

        def do_generate():
            from db.database import get_connection
            conn = get_connection()
            try:
                conn.execute("BEGIN IMMEDIATE")
                now = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")

                dn_str = generate_delivery_number(customer["id"], None)
                if not dn_str:
                    self.after(0, lambda: messagebox.showwarning(
                        u"错误", u"无法生成送货单号。"))
                    return

                primary_order_id = orders[0]["id"]

                today = datetime.date.today().strftime("%Y-%m-%d")
                conn.execute(
                    "INSERT INTO delivery_note (order_id, customer_id,"
                    " delivery_number, delivery_date, status,"
                    " created_at, updated_at)"
                    " VALUES (?,?,?,?,'draft',?,?)",
                    (primary_order_id, customer["id"], dn_str, today,
                     now, now))
                dn_id = conn.execute(
                    "SELECT last_insert_rowid()").fetchone()[0]

                # Insert all source orders into junction table
                source_order_ids = sorted(
                    set(o["id"] for o, _ in checked_data))
                for oid in source_order_ids:
                    conn.execute(
                        "INSERT OR IGNORE INTO delivery_note_order"
                        " (delivery_note_id, order_id) VALUES (?,?)",
                        (dn_id, oid))

                # Collect items into dicts for dedup
                items = []
                for order, oi in checked_data:
                    mn = oi.get("model_number") or ""
                    items.append({
                        "model_number": mn,
                        "_orig_model": mn,
                        "item_code": oi.get("order_number", ""),
                        "mfg_number": oi.get("mfg_number", ""),
                        "customer_code": oi.get("customer_code", ""),
                        "product_name": oi.get("product_name") or "",
                        "color_name": oi.get("color_name") or "",
                        "quantity_formula": oi.get("quantity_formula") or "",
                        "quantity": oi.get("quantity", 0),
                        "notes": oi.get("notes") or "",
                        "_oi_id": oi["id"],
                        "_push_count": oi.get("push_count") or 0,
                    })

                # Deduplicate consecutive rows with same 款号+制单号+订单号
                deduplicate_dn_items(items)

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

                    conn.execute(
                        "INSERT INTO delivery_note_item"
                        " (delivery_note_id, order_item_id,"
                        " model_number, item_code, customer_code,"
                        " product_name,"
                        " color_name, quantity_formula, quantity,"
                        " notes, mfg_number, sort_order,"
                        " created_at, updated_at)"
                        " VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?)",
                        (dn_id, item["_oi_id"],
                         "" if not item["model_number"] else (pr["model_number"] if pr else mn),
                         item["item_code"],
                         item.get("customer_code", ""),
                         pn, cn, item["quantity_formula"],
                         item["quantity"], item["notes"],
                         item["mfg_number"],
                         idx, now, now))

                    # Increment push_count on source order_item
                    push_cnt = item["_push_count"] + 1
                    conn.execute(
                        "UPDATE order_item SET push_count=?,"
                        " updated_at=? WHERE id=?",
                        (push_cnt, now, item["_oi_id"]))

                # Update order statuses
                for oid in source_order_ids:
                    conn.execute(
                        "UPDATE orders SET status='qty_filled',"
                        " updated_at=? WHERE id=? AND status='draft'",
                        (now, oid))

                conn.commit()

                def open_dn():
                    self.destroy()
                    from ui.dn_window import DeliveryNoteWindow
                    DeliveryNoteWindow(
                        master, dn_id=dn_id,
                        on_save_callback=None)

                self.after(0, open_dn)

            except Exception as e:
                try:
                    conn.rollback()
                except Exception:
                    pass
                import traceback
                self.after(0, lambda: messagebox.showerror(
                    u"生成失败", str(e) + "\n" + traceback.format_exc()))
            finally:
                conn.close()
                self.after(0, lambda: self._safe_ui(
                    lambda: (self.config(cursor=""),
                             self._gen_btn.config(state=tk.NORMAL))))

        t = threading.Thread(target=do_generate)
        t.daemon = True
        t.start()

    def _safe_ui(self, fn):
        def _run():
            try:
                fn()
            except tk.TclError:
                pass
        self.after(0, _run)
