# -*- coding: utf-8 -*-
"""Excel-like data grid with visible borders, inline editing, scrollbar."""
try:
    import Tkinter as tk
    import tkFont as tkfont
except ImportError:
    import tkinter as tk
    import tkinter.font as tkfont


ROW_H = 50
HEADER_H = 36


class ExcelGrid(tk.Frame):

    def __init__(self, parent, columns, col_widths, checkbox_col=None,
                 search_cols=None, extra_checkbox_cols=None,
                 read_only_cols=None):
        """
        columns: list of header strings
        col_widths: list of pixel widths per column
        checkbox_col: index of checkbox toggle column (optional)
        search_cols: dict {col_idx: callback(row, col)} for double-click search
        """
        tk.Frame.__init__(self, parent, bg="white")

        self.columns = columns
        self.col_widths = col_widths
        self.checkbox_col = checkbox_col
        self.search_cols = search_cols or {}
        self.extra_checkbox_cols = set(extra_checkbox_cols or [])
        self.read_only_cols = set(read_only_cols or [])

        # Internal data: list of lists
        self._data = []

        # Editing state
        self._edit_entry = None
        self._edit_row = None
        self._edit_col = None

        # Fonts
        self._hdr_font = tkfont.Font(family="Microsoft YaHei", size=10,
                                      weight="bold")
        self._cell_font = tkfont.Font(family="Microsoft YaHei", size=10)
        self._bold_font = tkfont.Font(family="Microsoft YaHei", size=10,
                                      weight="bold")

        # Colors
        self._hdr_bg = "#e8e8e8"
        self._cell_white = "#ffffff"
        self._cell_alt = "#f8f8f8"
        self._border_color = "#c0c0c0"

        self.on_cell_changed = None  # callback(row, col, new_value)
        self.summary_text = ""       # set to show summary row at bottom
        self._display_overrides = {}  # {(row,col): display_text} for formula cols
        self.read_only = False
        self.on_edit_blocked = None  # called when edit attempted in read_only mode

        # Sort state
        self._sort_col = None
        self._sort_desc = False

        self._build()
        self.draw()

    def _build(self):
        # Header
        self.header_frame = tk.Frame(self, height=HEADER_H)
        self.header_frame.pack(fill=tk.X, side=tk.TOP)
        self.header_frame.pack_propagate(False)

        # Scrollable data
        container = tk.Frame(self)
        container.pack(fill=tk.BOTH, expand=True, side=tk.TOP)

        self.canvas = tk.Canvas(
            container, bg="white", highlightthickness=0, bd=0)
        self.scrollbar = tk.Scrollbar(
            container, orient=tk.VERTICAL, command=self.canvas.yview)
        self.canvas.configure(yscrollcommand=self.scrollbar.set)

        self.scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        self.data_frame = tk.Frame(self.canvas, bg="white")
        self._canvas_win = self.canvas.create_window(
            (0, 0), window=self.data_frame, anchor=tk.NW)

        self.canvas.bind("<Configure>",
                         lambda e: self.canvas.itemconfig(
                             self._canvas_win, width=e.width))
        self.data_frame.bind(
            "<Configure>",
            lambda e: self.canvas.configure(
                scrollregion=self.canvas.bbox("all")))
        self._mw_binding = self.canvas.bind("<MouseWheel>", self._on_mousewheel)
        self.data_frame.bind("<MouseWheel>", self._on_mousewheel)
        # Bind <Enter> to ensure focus for mouse wheel on Windows
        self.canvas.bind("<Enter>", lambda e: self.canvas.focus_set())

    def _on_mousewheel(self, event):
        """Mouse wheel scroll handler — works on canvas and all children."""
        self.canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

    # ================================================================
    # Drawing
    # ================================================================
    def draw(self):
        self._clear()
        self._draw_header()
        self._draw_rows()
        self._draw_summary()
        tw = sum(self.col_widths)
        n = len(self._data)
        extra = 2 if self.summary_text else 0
        th = (n + extra) * ROW_H + 4
        tw = max(tw, 100)
        th = max(th, 20)
        self.data_frame.config(width=tw, height=th)
        self.canvas.configure(scrollregion=(0, 0, tw, th))
        self.data_frame.update_idletasks()

    def _draw_summary(self):
        if not self.summary_text:
            return
        n = len(self._data)
        # Blank separator row
        y_sep = n * ROW_H
        tk.Label(self.data_frame, text="", bg="#f0f0f0",
                 relief="solid", borderwidth=1).place(
                     x=0, y=y_sep,
                     width=sum(self.col_widths), height=ROW_H)
        # Summary row
        y_sum = (n + 1) * ROW_H
        lbl = tk.Label(self.data_frame, text=self.summary_text,
                       font=self._bold_font, bg="#e8e8e8", fg="#333333",
                       relief="solid", borderwidth=1, anchor=tk.W)
        lbl.place(x=0, y=y_sum,
                  width=sum(self.col_widths), height=ROW_H)

    def _clear(self):
        for w in self.header_frame.winfo_children():
            w.destroy()
        for w in self.data_frame.winfo_children():
            w.destroy()

    def _draw_header(self):
        x = 0
        arrow = u" ▼" if self._sort_desc else u" ▲"
        for ci, header in enumerate(self.columns):
            w = self.col_widths[ci]
            text = header + arrow if ci == self._sort_col else header
            lbl = tk.Label(
                self.header_frame, text=text, font=self._hdr_font,
                bg=self._hdr_bg, fg="#333333",
                relief="groove", borderwidth=1, anchor=tk.CENTER,
                justify=tk.CENTER)
            lbl.place(x=x, y=0, width=w, height=HEADER_H)
            lbl.bind("<Button-1>", lambda e, col=ci: self._header_click(col))
            # Double-click header to select all (for checkbox cols)
            if ci == self.checkbox_col or ci in self.extra_checkbox_cols:
                lbl.bind("<Double-1>", lambda e, col=ci: self._header_dbl_click(col))

            # Resize handle at right edge (20px wide, easier to grab)
            handle = tk.Frame(self.header_frame, bg="#a0a0a0",
                              cursor="sb_h_double_arrow")
            handle.place(x=x + w - 10, y=0, width=20, height=HEADER_H)
            handle.bind("<Button-1>",
                       lambda e, col=ci: self._resize_start(e, col))
            handle.bind("<B1-Motion>", self._resize_move)
            handle.bind("<ButtonRelease-1>", self._resize_end)
            handle.lift()

            x += w

    def _header_dbl_click(self, col):
        """Double-click checkbox header → toggle all."""
        current = self._data[0][col] if self._data else u"☐"
        new_val = u"☑" if current == u"☐" else u"☐"
        for ri in range(len(self._data)):
            if len(self._data[ri]) > col:
                self._data[ri][col] = new_val
        self.draw()

    def _header_click(self, col):
        """Single-click header → sort by that column."""
        if col == self.checkbox_col or col in self.extra_checkbox_cols:
            return  # don't sort checkbox cols
        if self._sort_col == col:
            self._sort_desc = not self._sort_desc
        else:
            self._sort_col = col
            self._sort_desc = False
        self._sort_data()
        self.draw()

    def _sort_data(self):
        if self._sort_col is None or not self._data:
            return
        ci = self._sort_col

        def _key(row):
            if ci >= len(row):
                return ""
            v = row[ci]
            if isinstance(v, (int, float)):
                return (0, v, "")
            s = str(v)
            # Try numeric sort
            try:
                return (0, float(s), "")
            except ValueError:
                return (1, 0, s.lower())

        self._data.sort(key=_key, reverse=self._sort_desc)

    # ---- Column resize ----
    _resize_col = None
    _resize_start_x = 0
    _resize_start_w = 0

    def _resize_start(self, event, col):
        self._resize_col = col
        self._resize_start_x = event.x_root
        self._resize_start_w = self.col_widths[col]

    def _resize_move(self, event):
        if self._resize_col is None:
            return
        dx = event.x_root - self._resize_start_x
        new_w = max(20, self._resize_start_w + dx)
        self.col_widths[self._resize_col] = new_w
        self.draw()

    def _resize_end(self, event):
        self._resize_col = None
        if hasattr(self, 'on_resize_done'):
            self.on_resize_done()

    def _draw_rows(self):
        for ri in range(len(self._data)):
            self._draw_row(ri)

    def _fit_text(self, text, max_width):
        """Truncate text with … so it fits within max_width pixels."""
        if not text:
            return ""
        text = str(text)
        if self._cell_font.measure(text) <= max_width - 8:
            return text
        # Binary search for the right length
        lo, hi = 0, len(text)
        while lo < hi:
            mid = (lo + hi + 1) // 2
            if self._cell_font.measure(text[:mid] + u"…") <= max_width - 8:
                lo = mid
            else:
                hi = mid - 1
        return text[:lo] + u"…" if lo > 0 else u"…"

    def _draw_row(self, ri):
        bg = self._cell_white if ri % 2 == 0 else self._cell_alt
        row_data = self._data[ri]
        y = ri * ROW_H
        x = 0

        for ci in range(len(self.columns)):
            w = self.col_widths[ci]
            # Use display override if exists, otherwise actual data
            override = self._display_overrides.get((ri, ci))
            val = override if override is not None else (
                row_data[ci] if ci < len(row_data) else "")

            if ci == self.checkbox_col or ci in self.extra_checkbox_cols:
                is_on = (val == u"☑")
                fg = "#006600" if is_on else "#aaaaaa"
                c = tk.Label(self.data_frame, text=str(val),
                            font=self._cell_font, bg=bg, fg=fg,
                            anchor=tk.CENTER,
                            relief="solid", borderwidth=1)
                c.place(x=x, y=y, width=w, height=ROW_H)
                c.bind("<Button-1>", lambda e, r=ri, col=ci:
                       self._toggle_chk(r, col))
                c.bind("<MouseWheel>", self._on_mousewheel)
            else:
                display = self._fit_text(val, w)
                c = tk.Label(self.data_frame, text=display,
                            font=self._cell_font, bg=bg, fg="#333333",
                            anchor=tk.CENTER,
                            relief="solid", borderwidth=1)
                c.place(x=x, y=y, width=w, height=ROW_H)
                c.bind("<Button-1>",
                       lambda e, r=ri, col=ci: self._click(r, col))
                c.bind("<MouseWheel>", self._on_mousewheel)
                if ci in self.search_cols:
                    c.bind("<Double-1>",
                           lambda e, r=ri, col=ci: self._dbl_click(r, col))

            x += w

    # ================================================================
    # Data access
    # ================================================================
    def get_data(self):
        return [list(row) for row in self._data]

    def set_data(self, data):
        self._data = [list(row) for row in data]
        self.draw()

    def get_cell(self, row, col):
        if row < len(self._data) and col < len(self._data[row]):
            return self._data[row][col]
        return ""

    def set_cell(self, row, col, value):
        while row >= len(self._data):
            self._data.append([u"☐" if i == self.checkbox_col else ""
                              for i in range(len(self.columns))])
        while col >= len(self._data[row]):
            self._data[row].append("")
        self._data[row][col] = value

    def set_display_override(self, row, col, text):
        """Set a display override (what's shown vs what's stored)."""
        if text:
            self._display_overrides[(row, col)] = text
        else:
            self._display_overrides.pop((row, col), None)

    def clear_overrides(self):
        self._display_overrides.clear()

    def row_count(self):
        return len(self._data)

    # ================================================================
    # Row operations
    # ================================================================
    def add_row(self, values=None):
        if values:
            self._data.append(list(values))
        else:
            row = []
            for i in range(len(self.columns)):
                if i == self.checkbox_col or i in self.extra_checkbox_cols:
                    row.append(u"☐")
                else:
                    row.append("")
            self._data.append(row)
        self.draw()
        self._scroll_bottom()

    def delete_rows(self, indices):
        for ri in sorted(indices, reverse=True):
            if ri < len(self._data):
                self._data.pop(ri)
        self.draw()

    def get_selected_rows(self):
        """Return indices of rows with checkbox checked."""
        if self.checkbox_col is None:
            return []
        return [ri for ri, row in enumerate(self._data)
                if ri < len(row) and row[self.checkbox_col] == u"☑"]

    def get_all_rows(self):
        return self.get_data()

    # ================================================================
    # Interaction
    # ================================================================
    def commit_edit(self):
        """Flush any pending edit to underlying data.  Call before get_data()."""
        self._commit_edit()

    def _click(self, row, col):
        if self.read_only:
            if self.on_edit_blocked:
                self.on_edit_blocked(row, col)
            return
        if col in self.read_only_cols:
            return  # frozen column
        self._start_edit(row, col)

    def _dbl_click(self, row, col):
        if self.read_only:
            if self.on_edit_blocked:
                self.on_edit_blocked(row, col)
            return
        if col in self.search_cols:
            self.search_cols[col](row, col)

    def _toggle_chk(self, row, col):
        if self.read_only:
            if self.on_edit_blocked:
                self.on_edit_blocked(row, col)
            return
        if row < len(self._data) and col < len(self._data[row]):
            cur = self._data[row][col]
            self._data[row][col] = u"☐" if cur == u"☑" else u"☑"
            self.draw()

    def _start_edit(self, row, col):
        self._commit_edit()
        self._edit_row = row
        self._edit_col = col

        w = self.col_widths[col]
        y = row * ROW_H
        x = sum(self.col_widths[i] for i in range(col))

        # Show actual stored value (not display override) for editing
        if row < len(self._data) and col < len(self._data[row]):
            current = self._data[row][col]
        else:
            current = ""

        e = tk.Entry(self.data_frame, font=self._cell_font,
                     bg="#ffffcc", relief="solid", borderwidth=2)
        e.place(x=x + 1, y=y + 1, width=w - 2, height=ROW_H - 2)
        e.insert(0, str(current) if current else "")
        e.select_range(0, tk.END)
        e.focus_set()

        e.bind("<Return>", lambda ev: self._commit_edit())
        e.bind("<Escape>", lambda ev: self._cancel_edit())
        e.bind("<FocusOut>", lambda ev: self._commit_edit())
        e.bind("<Tab>",
               lambda ev, r=row, c=col: self._tab_next(r, c))

        self._edit_entry = e

    def _commit_edit(self):
        if not self._edit_entry:
            return
        try:
            new_val = self._edit_entry.get()
        except tk.TclError:
            new_val = ""
        self._edit_entry.destroy()
        self._edit_entry = None

        if self._edit_row is not None:
            self.set_cell(self._edit_row, self._edit_col, new_val)
            self.draw()
            if self.on_cell_changed:
                self.on_cell_changed(self._edit_row, self._edit_col, new_val)
        self._edit_row = None
        self._edit_col = None

    def _cancel_edit(self):
        if self._edit_entry:
            self._edit_entry.destroy()
            self._edit_entry = None
        self._edit_row = None
        self._edit_col = None

    def _tab_next(self, row, col):
        self._commit_edit()
        nc = col + 1
        nr = row
        if nc >= len(self.columns):
            nc = 1 if self.checkbox_col == 0 else 0
            nr += 1
        if nr < len(self._data):
            self.after(50, lambda: self._start_edit(nr, nc))

    def _scroll_bottom(self):
        self.canvas.update_idletasks()
        self.canvas.yview_moveto(1.0)
