# -*- coding: utf-8 -*-
"""Customer management window — list + edit form."""
try:
    import Tkinter as tk
    import ttk
    import tkMessageBox as messagebox
except ImportError:
    import tkinter as tk
    from tkinter import ttk
    from tkinter import messagebox

from db.models import Customer


class CustomerWindow(tk.Toplevel):

    def __init__(self, parent, on_save_callback=None):
        tk.Toplevel.__init__(self, parent)
        self.on_save_callback = on_save_callback
        self._customers = []
        self._editing_id = None

        self.title(u"客户管理")
        self.geometry("960x640")
        self.minsize(860, 560)
        self.transient(parent)

        self._build()
        self._load_list()

        self.protocol("WM_DELETE_WINDOW", self.destroy)

    def _build(self):
        # Title bar
        title_bar = tk.Frame(self, bg="#2c3e50", height=32)
        title_bar.pack(fill=tk.X)
        title_bar.pack_propagate(False)
        tk.Label(title_bar, text=u"  客户管理",
                 font=("Microsoft YaHei", 12, "bold"),
                 bg="#2c3e50", fg="white", anchor=tk.W).pack(
                     side=tk.LEFT, fill=tk.BOTH, expand=True)

        # Body: left tree + right form
        body = tk.Frame(self)
        body.pack(fill=tk.BOTH, expand=True, padx=10, pady=8)

        # -- Left: Treeview --
        left = tk.Frame(body)
        left.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        # Search box
        search_frame = tk.Frame(left)
        search_frame.pack(fill=tk.X, pady=(0, 4))
        tk.Label(search_frame, text=u"查询:",
                 font=("Microsoft YaHei", 9)).pack(side=tk.LEFT, padx=(0, 4))
        self._search_var = tk.StringVar()
        self._search_var.trace("w", lambda *a: self._filter_list())
        search_entry = ttk.Entry(search_frame, textvariable=self._search_var,
                                 width=20)
        search_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)

        tree_frame = tk.Frame(left)
        tree_frame.pack(fill=tk.BOTH, expand=True)
        self.tree = ttk.Treeview(
            tree_frame, columns=("name", "prefix"), show="headings",
            selectmode="browse", height=14)
        self.tree.heading("name", text=u"客户名称")
        self.tree.heading("prefix", text=u"前缀")
        self.tree.column("name", width=180)
        self.tree.column("prefix", width=60, anchor=tk.CENTER)
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        sb = ttk.Scrollbar(tree_frame, orient=tk.VERTICAL,
                           command=self.tree.yview)
        self.tree.configure(yscrollcommand=sb.set)
        sb.pack(side=tk.RIGHT, fill=tk.Y)

        self.tree.bind("<<TreeviewSelect>>", self._on_select)

        # -- Right: Form --
        # 【修改点3】：把右侧表单区域的宽度从 280 拉宽到 360，彻底解决列3被切掉的问题
        right = tk.Frame(body, width=480)
        right.pack(side=tk.RIGHT, fill=tk.Y, padx=(12, 0))
        right.pack_propagate(False)

        form = tk.Frame(right)
        form.pack(fill=tk.X, pady=(4, 0))

        # Customer name
        tk.Label(form, text=u"客户名称:", font=("Microsoft YaHei", 9)).pack(
            anchor=tk.W, pady=(6, 0))
        self.var_name = tk.StringVar()
        self.e_name = ttk.Entry(form, textvariable=self.var_name, width=30)
        self.e_name.pack(fill=tk.X, pady=(2, 0))

        # Prefix
        tk.Label(form, text=u"前缀 (送货单号):", font=("Microsoft YaHei", 9)).pack(
            anchor=tk.W, pady=(8, 0))
        self.var_prefix = tk.StringVar()
        ttk.Entry(form, textvariable=self.var_prefix, width=10).pack(
            anchor=tk.W, pady=(2, 0))

        # Needs price checkbox
        self.var_needs_price = tk.IntVar()
        ttk.Checkbutton(form, text=u"需要价格（开启后订单明细的单价列高亮提醒）",
                        variable=self.var_needs_price).pack(
            anchor=tk.W, pady=(10, 0))

        # Column name config — presets + manual fields
        tk.Label(form, text=u"送货单列名 (按客户习惯):",
                 font=("Microsoft YaHei", 9)).pack(
            anchor=tk.W, pady=(12, 0))

        presets = [
            u",订单号,,款号",
            u",订单号,制单号,款号",
            u",订单号,系列,款号",
            u"",
        ]
        preset_labels = [
            u"(空) / 订单号 / (空) / 款号",
            u"(空) / 订单号 / 制单号 / 款号",
            u"(空) / 订单号 / 系列 / 款号",
            u"自定义",
        ]
        preset_frame = tk.Frame(form)
        preset_frame.pack(fill=tk.X, pady=(2, 0))
        tk.Label(preset_frame, text=u"预设:",
                 font=("Microsoft YaHei", 8)).pack(side=tk.LEFT, padx=(0, 4))
        self.var_preset = tk.StringVar(value=presets[0])
        self._presets = presets
        self._preset_labels = preset_labels
        cb = ttk.Combobox(preset_frame, textvariable=self.var_preset,
                          values=preset_labels, state="readonly", width=28)
        cb.pack(side=tk.LEFT)
        cb.bind("<<ComboboxSelected>>", self._on_preset_changed)

        cols_frame = tk.Frame(form)
        cols_frame.pack(fill=tk.X, pady=(2, 0))
        self.var_dn_col0 = tk.StringVar(value=u"")
        self.var_dn_col1 = tk.StringVar(value=u"订单号")
        self.var_dn_col2 = tk.StringVar(value=u"制单号")
        self.var_dn_col3 = tk.StringVar(value=u"款号")
        for i, (label, var) in enumerate([
                (u"列1(客户号):", self.var_dn_col0),
                (u"列2(订单号):", self.var_dn_col1),
                (u"列3(制单号):", self.var_dn_col2),
                (u"列4(款号):", self.var_dn_col3)]):
            r, c = divmod(i, 2)  # 2 per row
            col = c * 2
            tk.Label(cols_frame, text=label,
                     font=("Microsoft YaHei", 8)).grid(
                row=r, column=col, padx=(0 if col == 0 else 12, 2))
            ttk.Entry(cols_frame, textvariable=var, width=10).grid(
                row=r, column=col + 1)

        # Buttons
        btn_frame = tk.Frame(right)
        btn_frame.pack(fill=tk.X, pady=(16, 0))
        ttk.Button(btn_frame, text=u"新建", command=self._new).pack(
            side=tk.LEFT, padx=2)
        ttk.Button(btn_frame, text=u"保存", command=self._save).pack(
            side=tk.LEFT, padx=2)
        ttk.Button(btn_frame, text=u"删除", command=self._delete).pack(
            side=tk.LEFT, padx=2)

        # Close button at bottom
        ttk.Button(right, text=u"关闭", command=self.destroy).pack(
            side=tk.BOTTOM, pady=(12, 0))

    # ---- Data ----
    def _load_list(self):
        self._customers = Customer.get_all()
        self._filter_list()

    def _filter_list(self):
        kw = self._search_var.get().strip().lower() if hasattr(
            self, '_search_var') else ""
        for item in self.tree.get_children():
            self.tree.delete(item)
        for c in self._customers:
            if kw and kw not in c["name"].lower() and kw not in c.get(
                    "prefix", "").lower():
                continue
            self.tree.insert("", tk.END, iid=str(c["id"]),
                             values=(c["name"], c["prefix"]))

    def _on_preset_changed(self, event=None):
        """When user picks a preset, fill the three column name fields."""
        val = self.var_preset.get()
        idx = self._preset_labels.index(val) if val in self._preset_labels else 0
        if idx == len(self._presets) - 1:
            return  # "自定义" — don't overwrite
        headers = self._presets[idx].split(",")
        for i, var in enumerate([self.var_dn_col0, self.var_dn_col1,
                                  self.var_dn_col2, self.var_dn_col3]):
            var.set(headers[i].strip() if i < len(headers) else "")

    def _on_select(self, event):
        sel = self.tree.selection()
        if not sel:
            return
        cid = int(sel[0])
        for c in self._customers:
            if c["id"] == cid:
                self._editing_id = cid
                self.var_name.set(c["name"])
                self.var_prefix.set(c["prefix"])
                self.var_needs_price.set(c.get("needs_price", 0))
                # Parse dn_headers and set fields + preset dropdown
                raw = c.get("dn_headers") or u",订单号,,款号"
                headers = raw.split(",")
                if len(headers) < 4:
                    headers = [""] + headers  # old 3-field → prepend 客户号
                for i, var in enumerate([self.var_dn_col0, self.var_dn_col1,
                                          self.var_dn_col2, self.var_dn_col3]):
                    var.set(headers[i].strip() if i < len(headers) else "")
                # Match to a preset, or fall back to "自定义"
                joined = ",".join([self.var_dn_col0.get().strip(),
                                   self.var_dn_col1.get().strip(),
                                   self.var_dn_col2.get().strip(),
                                   self.var_dn_col3.get().strip()])
                try:
                    pi = self._presets.index(joined)
                    self.var_preset.set(self._preset_labels[pi])
                except ValueError:
                    self.var_preset.set(self._preset_labels[-1])
                return

    def _new(self):
        self._editing_id = None
        self.var_name.set("")
        self.var_prefix.set("")
        self.var_needs_price.set(0)
        self.var_dn_col0.set(u"")
        self.var_dn_col1.set(u"订单号")
        self.var_dn_col2.set(u"")
        self.var_dn_col3.set(u"款号")
        self.e_name.focus_set()

    def _save(self):
        name = self.var_name.get().strip()
        if not name:
            messagebox.showwarning(u"提示", u"请输入客户名称。")
            return
        prefix = self.var_prefix.get().strip()
        if not prefix:
            messagebox.showwarning(u"提示", u"请输入前缀（如 HR、YL）。")
            return

        # Duplicate name check
        for c in self._customers:
            if c["name"] == name and c["id"] != self._editing_id:
                messagebox.showwarning(u"提示", u"客户名称 '{}' 已存在，不能重复创建。".format(name))
                return

        kwargs = {
            "dn_headers": ",".join([
                self.var_dn_col0.get().strip(),
                self.var_dn_col1.get().strip(),
                self.var_dn_col2.get().strip(),
                self.var_dn_col3.get().strip(),
            ]),
            "needs_price": int(self.var_needs_price.get()),
        }

        try:
            if self._editing_id:
                Customer.update(self._editing_id, name=name,
                                prefix=prefix, **kwargs)
            else:
                Customer.create(name=name, prefix=prefix, **kwargs)
        except Exception as e:
            messagebox.showerror(u"保存失败", str(e))
            return

        self._load_list()
        if self.on_save_callback:
            self.on_save_callback()

    def _delete(self):
        if not self._editing_id:
            messagebox.showinfo(u"提示", u"请先选择一个客户。")
            return
        name = self.var_name.get()
        if not messagebox.askyesno(
                u"确认删除",
                u"确定要删除客户 '{}' 吗？\n\n此操作不可撤销。".format(name)):
            return
        try:
            Customer.delete(self._editing_id)
        except Exception as e:
            messagebox.showerror(u"删除失败", str(e))
            return
        self._new()
        self._load_list()
        if self.on_save_callback:
            self.on_save_callback()
