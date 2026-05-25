# -*- coding: utf-8 -*-
"""Reusable tkinter widgets."""
try:
    import Tkinter as tk
    import ttk
except ImportError:
    import tkinter as tk
    from tkinter import ttk
import calendar
import datetime


class DatePicker(object):
    """Simple calendar date picker popup."""

    def __init__(self, parent, entry_widget):
        self.parent = parent
        self.entry = entry_widget
        self.top = tk.Toplevel(parent)
        self.top.title(u"选择日期")
        self.top.geometry("460x380")
        self.top.transient(parent)
        self.top.grab_set()
        self.top.resizable(False, False)

        self._selected = None
        today = datetime.date.today()
        self._year = today.year
        self._month = today.month
        self._day = today.day

        self._build()
        self._draw_calendar()

        # Position near entry
        self.top.update_idletasks()
        x = entry_widget.winfo_rootx()
        y = entry_widget.winfo_rooty() + entry_widget.winfo_height()
        self.top.geometry("+{}+{}".format(x, y))

    def _build(self):
        # Month/year navigation
        nav = ttk.Frame(self.top)
        nav.pack(fill=tk.X, padx=4, pady=4)
        ttk.Button(nav, text=u"◀", width=2,
                   command=self._prev_month).pack(side=tk.LEFT)
        self.month_label = ttk.Label(nav, text="", width=14, anchor=tk.CENTER)
        self.month_label.pack(side=tk.LEFT, padx=4)
        ttk.Button(nav, text=u"▶", width=2,
                   command=self._next_month).pack(side=tk.LEFT)

        # Weekday headers
        hdr = ttk.Frame(self.top)
        hdr.pack(fill=tk.X, padx=4)
        for d in [u"一", u"二", u"三", u"四", u"五", u"六", u"日"]:
            ttk.Label(hdr, text=d, width=5, anchor=tk.CENTER).pack(
                side=tk.LEFT, padx=1)

        # Day grid
        self.day_frame = ttk.Frame(self.top)
        self.day_frame.pack(fill=tk.BOTH, expand=True, padx=4, pady=2)

        # Today button
        bottom = ttk.Frame(self.top)
        bottom.pack(fill=tk.X, padx=4, pady=4)
        ttk.Button(bottom, text=u"今天",
                   command=self._select_today).pack(side=tk.LEFT)
        ttk.Button(bottom, text=u"清除",
                   command=self._clear).pack(side=tk.LEFT, padx=4)

    def _draw_calendar(self):
        for w in self.day_frame.winfo_children():
            w.destroy()

        self.month_label.config(
            text=u"{}年 {}月".format(self._year, self._month))

        cal = calendar.monthcalendar(self._year, self._month)
        today = datetime.date.today()

        for week_idx, week in enumerate(cal):
            for day_idx, day in enumerate(week):
                if day == 0:
                    ttk.Label(self.day_frame, text="", width=5).grid(
                        row=week_idx, column=day_idx, padx=1, pady=1)
                    continue

                is_today = (self._year == today.year and
                           self._month == today.month and
                           day == today.day)
                is_selected = (self._selected and
                              self._selected[0] == self._year and
                              self._selected[1] == self._month and
                              self._selected[2] == day)

                if is_selected:
                    btn = tk.Button(
                        self.day_frame, text=str(day), width=5,
                        bg="#3399ff", fg="white",
                        command=lambda d=day: self._pick_day(d))
                elif is_today:
                    btn = tk.Button(
                        self.day_frame, text=str(day), width=5,
                        bg="#e0e0e0",
                        command=lambda d=day: self._pick_day(d))
                else:
                    btn = tk.Button(
                        self.day_frame, text=str(day), width=5,
                        command=lambda d=day: self._pick_day(d))
                btn.grid(row=week_idx, column=day_idx, padx=1, pady=1)

    def _prev_month(self):
        if self._month == 1:
            self._year -= 1
            self._month = 12
        else:
            self._month -= 1
        self._selected = None
        self._draw_calendar()

    def _next_month(self):
        if self._month == 12:
            self._year += 1
            self._month = 1
        else:
            self._month += 1
        self._selected = None
        self._draw_calendar()

    def _pick_day(self, day):
        self._selected = (self._year, self._month, day)
        date_str = "{:04d}-{:02d}-{:02d}".format(
            self._year, self._month, day)
        self.entry.delete(0, tk.END)
        self.entry.insert(0, date_str)
        self.top.destroy()

    def _select_today(self):
        today = datetime.date.today()
        self._selected = (today.year, today.month, today.day)
        date_str = today.strftime("%Y-%m-%d")
        self.entry.delete(0, tk.END)
        self.entry.insert(0, date_str)
        self.top.destroy()

    def _clear(self):
        self.entry.delete(0, tk.END)
        self.top.destroy()


class DateEntry(ttk.Frame):
    """Entry with a calendar button for date picking."""

    def __init__(self, parent, width=12, **kwargs):
        ttk.Frame.__init__(self, parent)
        self.var = tk.StringVar()
        self.entry = ttk.Entry(self, font=("Microsoft YaHei", 11),
                               width=width, textvariable=self.var)
        self.entry.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.btn = ttk.Button(self, text=u"📅", width=2,
                              command=self._open_picker)
        self.btn.pack(side=tk.RIGHT, fill=tk.Y)

    def _open_picker(self):
        DatePicker(self, self.entry)

    def get(self):
        return self.var.get().strip()

    def set(self, value):
        self.var.set(value)

    def delete(self, first, last=None):
        self.entry.delete(first, last)

    def config(self, **kwargs):
        self.entry.config(**kwargs)

    def bind(self, *args, **kwargs):
        self.entry.bind(*args, **kwargs)
