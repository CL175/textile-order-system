# -*- coding: utf-8 -*-
"""System settings window — company info, export folder."""
try:
    import Tkinter as tk
    import ttk
    import tkMessageBox as messagebox
    import tkFileDialog as filedialog
except ImportError:
    import tkinter as tk
    from tkinter import ttk
    from tkinter import messagebox
    from tkinter import filedialog

from db.models import Settings


class SettingsWindow(tk.Toplevel):

    FIELDS = [
        ("company_name", u"公司名称:", 40),
        ("company_address", u"公司地址:", 50),
        ("company_email", u"公司邮箱:", 40),
        ("maker_name", u"制表人:", 30),
        ("footer_text", u"页脚文本:", 50),
    ]

    FOLDER_KEY = "default_excel_folder"
    TEMPLATE_KEY = "template_dir"

    def __init__(self, parent, on_save_callback=None):
        tk.Toplevel.__init__(self, parent)
        self.on_save_callback = on_save_callback
        self._vars = {}

        self.title(u"系统设置")
        self.geometry("580x420")
        self.minsize(460, 360)
        self.transient(parent)

        self._build()
        self._load()

        self.protocol("WM_DELETE_WINDOW", self.destroy)

    def _build(self):
        # Title bar
        title_bar = tk.Frame(self, bg="#2c3e50", height=32)
        title_bar.pack(fill=tk.X)
        title_bar.pack_propagate(False)
        tk.Label(title_bar, text=u"  系统设置",
                 font=("Microsoft YaHei", 12, "bold"),
                 bg="#2c3e50", fg="white", anchor=tk.W).pack(
                     side=tk.LEFT, fill=tk.BOTH, expand=True)

        # Form
        form = tk.Frame(self)
        form.pack(fill=tk.BOTH, expand=True, padx=20, pady=12)

        # Standard text fields
        for key, label, width in self.FIELDS:
            row = tk.Frame(form)
            row.pack(fill=tk.X, pady=(8, 0))
            tk.Label(row, text=label, font=("Microsoft YaHei", 9),
                     width=10, anchor=tk.W).pack(side=tk.LEFT)
            var = tk.StringVar()
            self._vars[key] = var
            ttk.Entry(row, textvariable=var, width=width).pack(
                side=tk.LEFT, fill=tk.X, expand=True)

        # Folder fields (with browse button)
        folders = [
            (self.FOLDER_KEY, u"导出文件夹:"),
            (self.TEMPLATE_KEY, u"模板文件夹:"),
        ]
        for key, label in folders:
            row = tk.Frame(form)
            row.pack(fill=tk.X, pady=(12, 0))
            tk.Label(row, text=label, font=("Microsoft YaHei", 9),
                     width=10, anchor=tk.W).pack(side=tk.LEFT)
            var = tk.StringVar()
            self._vars[key] = var
            ttk.Entry(row, textvariable=var).pack(
                side=tk.LEFT, fill=tk.X, expand=True)
            ttk.Button(row, text=u"浏览...",
                       command=lambda k=key: self._browse_folder(k)).pack(
                           side=tk.LEFT, padx=(4, 0))

        # Buttons
        btn_frame = tk.Frame(self)
        btn_frame.pack(fill=tk.X, padx=20, pady=(10, 14))
        ttk.Button(btn_frame, text=u"保存", command=self._save).pack(
            side=tk.RIGHT, padx=4)
        ttk.Button(btn_frame, text=u"取消", command=self.destroy).pack(
            side=tk.RIGHT, padx=4)

    def _load(self):
        data = Settings.get_all()
        for key, var in self._vars.items():
            var.set(data.get(key, ""))

    def _browse_folder(self, key):
        path = filedialog.askdirectory(
            title=u"选择文件夹",
            initialdir=self._vars[key].get() or "D:/xxm")
        if path:
            self._vars[key].set(path)

    def _save(self):
        try:
            for key, var in self._vars.items():
                Settings.set(key, var.get().strip())
        except Exception as e:
            messagebox.showerror(u"保存失败", str(e))
            return
        messagebox.showinfo(u"保存成功", u"系统设置已保存。")
        if self.on_save_callback:
            self.on_save_callback()
