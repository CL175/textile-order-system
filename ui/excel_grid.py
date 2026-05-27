# -*- coding: utf-8 -*-
"""Excel-like data grid with Dual-Mode (Selection/Edit), inline editing, and IME protection."""
try:
    import Tkinter as tk
    import tkFont as tkfont
except ImportError:
    import tkinter as tk
    import tkinter.font as tkfont
import time

ROW_H = 50
HEADER_H = 36


class ExcelGrid(tk.Frame):

    def __init__(self, parent, columns, col_widths, checkbox_col=None,
                 search_cols=None, extra_checkbox_cols=None,
                 read_only_cols=None, highlight_cols=None):
        tk.Frame.__init__(self, parent, bg="white")

        self.columns = columns
        self.col_widths = col_widths
        self.checkbox_col = checkbox_col
        self.search_cols = search_cols or {}
        self.extra_checkbox_cols = set(extra_checkbox_cols or [])
        self.read_only_cols = set(read_only_cols or [])
        self.highlight_cols = set(highlight_cols or [])

        # Data
        self._data = []

        # Dual-Mode State
        self._selected_cell = None  # tuple: (row, col)
        self._edit_entry = None
        self._edit_row = None
        self._edit_col = None

        # IME Protection
        self._last_edit_time = 0

        # Fonts
        self._hdr_font = tkfont.Font(family="Microsoft YaHei", size=10, weight="bold")
        self._cell_font = tkfont.Font(family="Microsoft YaHei", size=10)
        self._bold_font = tkfont.Font(family="Microsoft YaHei", size=10, weight="bold")

        # Colors
        self._hdr_bg = "#e8e8e8"
        self._cell_white = "#ffffff"
        self._cell_alt = "#f8f8f8"
        self._cell_highlight = "#fff9c4"  # 浅黄色高亮，提醒填写
        self._border_color = "#c0c0c0"
        self._sel_color = "#00B050"  # Office 鲜绿色

        self.on_cell_changed = None
        self.summary_text = ""
        self._display_overrides = {}
        self.read_only = False
        self.on_edit_blocked = None

        self._sort_col = None
        self._sort_desc = False

        self._build()
        self.draw()

    def _build(self):
        # Header
        self.header_frame = tk.Frame(self, height=HEADER_H)
        self.header_frame.pack(fill=tk.X, side=tk.TOP)
        self.header_frame.pack_propagate(False)

        # Container
        container = tk.Frame(self)
        container.pack(fill=tk.BOTH, expand=True, side=tk.TOP)

        self.canvas = tk.Canvas(container, bg="white", highlightthickness=0, bd=0)
        self.canvas.config(takefocus=True)  # CRITICAL: Allow canvas to capture key events

        self.scrollbar = tk.Scrollbar(container, orient=tk.VERTICAL, command=self.canvas.yview)
        self.canvas.configure(yscrollcommand=self.scrollbar.set)

        self.scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        self.data_frame = tk.Frame(self.canvas, bg="white")
        self._canvas_win = self.canvas.create_window((0, 0), window=self.data_frame, anchor=tk.NW)

        # Canvas Resizing & Scrolling
        self.canvas.bind("<Configure>", lambda e: self.canvas.itemconfig(self._canvas_win, width=e.width))
        self.data_frame.bind("<Configure>", lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all")))

        # Mousewheel: bind to ALL layers so scrolling works wherever the mouse is
        self.canvas.bind("<MouseWheel>", self._on_mousewheel)
        self.data_frame.bind("<MouseWheel>", self._on_mousewheel)
        self.bind("<MouseWheel>", self._on_mousewheel)

        # Click empty space to gain focus
        self.canvas.bind("<Button-1>", lambda e: self.canvas.focus_set())
        self.canvas.bind("<Enter>", lambda e: self.canvas.focus_set())

        # ==========================================
        # Navigation Keyboard Bindings (Selection Mode)
        # ==========================================
        self.canvas.bind("<Up>", self._nav_up)
        self.canvas.bind("<Down>", self._nav_down)
        self.canvas.bind("<Left>", self._nav_left)
        self.canvas.bind("<Right>", self._nav_right)
        self.canvas.bind("<Return>", self._nav_enter)
        self.canvas.bind("<Tab>", self._nav_tab)
        self.canvas.bind("<Control-c>", self._copy_cell)
        self.canvas.bind("<Control-C>", self._copy_cell)
        self.canvas.bind("<Control-x>", self._cut_cell)
        self.canvas.bind("<Control-X>", self._cut_cell)
        self.canvas.bind("<Control-v>", self._paste_cell)
        self.canvas.bind("<Control-V>", self._paste_cell)
        self.canvas.bind("<Key>", self._on_canvas_key)

    def _on_mousewheel(self, event):
        self.canvas.yview_scroll(int(-3 * (event.delta / 120)), "units")

    # ================================================================
    # Drawing & Selection Box
    # ================================================================
    def draw(self):
        if self._edit_entry:
            self._commit_edit()

        self._clear()
        self._draw_header()
        self._draw_rows()
        self._draw_summary()
        self._update_selection_box()

        tw = sum(self.col_widths)
        n = len(self._data)
        th = (n + (2 if self.summary_text else 0)) * ROW_H + 4
        self.data_frame.config(width=max(tw, 100), height=max(th, 20))
        self.canvas.configure(scrollregion=(0, 0, max(tw, 100), max(th, 20)))
        self.data_frame.update_idletasks()

    def _clear(self):
        for w in self.header_frame.winfo_children(): w.destroy()
        for w in self.data_frame.winfo_children(): w.destroy()

    def _update_selection_box(self):
        """Draw the Excel-style green selection border."""
        if hasattr(self, '_sel_borders'):
            for b in self._sel_borders:
                try: b.destroy()
                except Exception: pass
        self._sel_borders = []

        if not self._selected_cell: return
        r, c = self._selected_cell
        if r >= len(self._data) or c >= len(self.columns): return

        # Create 4 independent lines so we don't cover cell text
        for _ in range(4):
            self._sel_borders.append(tk.Frame(self.data_frame, bg=self._sel_color))

        x = sum(self.col_widths[:c])
        y = r * ROW_H
        w = self.col_widths[c]

        bw = 3  # 边框粗细
        self._sel_borders[0].place(x=x, y=y, width=w, height=bw)                 # Top
        self._sel_borders[1].place(x=x, y=y+ROW_H-bw, width=w, height=bw)        # Bottom
        self._sel_borders[2].place(x=x, y=y, width=bw, height=ROW_H)             # Left
        self._sel_borders[3].place(x=x+w-bw, y=y, width=bw, height=ROW_H)        # Right

        for b in self._sel_borders:
            b.bind("<MouseWheel>", self._on_mousewheel)
            b.lift()

    def _ensure_visible(self, r, c):
        """Auto-scroll if the selection moves out of bounds."""
        self.canvas.update_idletasks()
        y_top = r * ROW_H
        y_bottom = y_top + ROW_H
        canvas_h = self.canvas.winfo_height()
        if canvas_h <= 1: return

        yview = self.canvas.yview()
        total_h = self.data_frame.winfo_height()
        if total_h <= 0: return

        view_top = yview[0] * total_h
        view_bottom = yview[1] * total_h

        if y_top < view_top:
            self.canvas.yview_moveto(float(y_top) / total_h)
        elif y_bottom > view_bottom:
            self.canvas.yview_moveto(float(y_bottom - canvas_h + 10) / total_h)

    def _draw_summary(self):
        if not self.summary_text: return
        n = len(self._data)
        tk.Label(self.data_frame, text="", bg="#f0f0f0", relief="solid", borderwidth=1).place(
            x=0, y=n*ROW_H, width=sum(self.col_widths), height=ROW_H)
        tk.Label(self.data_frame, text=self.summary_text, font=self._bold_font,
                 bg="#e8e8e8", fg="#333", relief="solid", borderwidth=1, anchor=tk.W).place(
                     x=0, y=(n+1)*ROW_H, width=sum(self.col_widths), height=ROW_H)

    # ---- Header & Resize ----
    def _draw_header(self):
        x = 0
        arrow = u" ▼" if self._sort_desc else u" ▲"
        for ci, header in enumerate(self.columns):
            w = self.col_widths[ci]
            text = header + arrow if ci == self._sort_col else header
            lbl = tk.Label(self.header_frame, text=text, font=self._hdr_font,
                           bg=self._hdr_bg, fg="#333333", relief="groove", borderwidth=1)
            lbl.place(x=x, y=0, width=w, height=HEADER_H)
            lbl.bind("<Button-1>", lambda e, col=ci: self._header_click(col))
            if ci == self.checkbox_col or ci in self.extra_checkbox_cols:
                lbl.bind("<Double-1>", lambda e, col=ci: self._header_dbl_click(col))

            handle = tk.Frame(self.header_frame, bg="#a0a0a0", cursor="sb_h_double_arrow")
            handle.place(x=x + w - 10, y=0, width=20, height=HEADER_H)
            handle.bind("<Button-1>", lambda e, col=ci: self._resize_start(e, col))
            handle.lift()
            x += w

    def _header_dbl_click(self, col):
        current = self._data[0][col] if self._data else u"☐"
        new_val = u"☑" if current == u"☐" else u"☐"
        for ri in range(len(self._data)):
            if len(self._data[ri]) > col:
                self._data[ri][col] = new_val
                if self.on_cell_changed:
                    self.on_cell_changed(ri, col, new_val)
        self.draw()

    def _header_click(self, col):
        if col == self.checkbox_col or col in self.extra_checkbox_cols: return
        if self._sort_col == col: self._sort_desc = not self._sort_desc
        else: self._sort_col = col; self._sort_desc = False
        self._sort_data()
        self.draw()

    def _sort_data(self):
        if self._sort_col is None or not self._data: return
        ci = self._sort_col
        def _key(row):
            v = row[ci] if ci < len(row) else ""
            if isinstance(v, (int, float)): return (0, v, "")
            try: return (0, float(str(v)), "")
            except ValueError: return (1, 0, str(v).lower())
        self._data.sort(key=_key, reverse=self._sort_desc)

    _resize_col = None
    _resize_start_x = 0
    _resize_start_w = 0

    def _resize_start(self, event, col):
        self._resize_col = col
        self._resize_start_x = event.x_root
        self._resize_start_w = self.col_widths[col]
        self._rb_motion = self.bind_all("<B1-Motion>", self._resize_move)
        self._rb_release = self.bind_all("<ButtonRelease-1>", self._resize_end)

    def _resize_move(self, event):
        if self._resize_col is None: return
        dx = event.x_root - self._resize_start_x
        self.col_widths[self._resize_col] = max(20, self._resize_start_w + dx)
        self.draw()

    def _resize_end(self, event):
        self._resize_col = None
        self.unbind_all("<B1-Motion>")
        self.unbind_all("<ButtonRelease-1>")
        if hasattr(self, 'on_resize_done'): self.on_resize_done()

    # ---- Cell Rendering ----
    def _draw_rows(self):
        for ri in range(len(self._data)):
            self._draw_row(ri)

    def _fit_text(self, text, max_width):
        if not text: return ""
        text = str(text)
        if self._cell_font.measure(text) <= max_width - 8: return text
        lo, hi = 0, len(text)
        while lo < hi:
            mid = (lo + hi + 1) // 2
            if self._cell_font.measure(text[:mid] + u"…") <= max_width - 8: lo = mid
            else: hi = mid - 1
        return text[:lo] + u"…" if lo > 0 else u"…"

    def _draw_row(self, ri):
        row_data = self._data[ri]
        y = ri * ROW_H
        x = 0

        for ci in range(len(self.columns)):
            w = self.col_widths[ci]
            override = self._display_overrides.get((ri, ci))
            val = override if override is not None else (row_data[ci] if ci < len(row_data) else "")
            bg = self._cell_highlight if ci in self.highlight_cols else (
                self._cell_white if ri % 2 == 0 else self._cell_alt)

            if ci == self.checkbox_col or ci in self.extra_checkbox_cols:
                fg = "#006600" if val == u"☑" else "#aaaaaa"
                c = tk.Label(self.data_frame, text=str(val), font=self._cell_font, bg=bg, fg=fg,
                             relief="solid", borderwidth=1)
                c.bind("<Button-1>", lambda e, r=ri, col=ci: self._on_chk_click(r, col))
            else:
                c = tk.Label(self.data_frame, text=self._fit_text(val, w), font=self._cell_font,
                             bg=bg, fg="#333", relief="solid", borderwidth=1)
                c.bind("<Button-1>", lambda e, r=ri, col=ci: self._on_cell_click(r, col))
                c.bind("<Double-1>", lambda e, r=ri, col=ci: self._on_cell_dbl_click(r, col))

            c.place(x=x, y=y, width=w, height=ROW_H)
            c.bind("<MouseWheel>", self._on_mousewheel)
            x += w

    # ================================================================
    # Interaction & Navigation (The Excel Logic)
    # ================================================================
    def _set_selection(self, r, c):
        if self._edit_entry: self._commit_edit()
        self._selected_cell = (r, c)
        self._update_selection_box()
        self._ensure_visible(r, c)
        self.canvas.focus_set()

    def _on_cell_click(self, r, c):
        self._set_selection(r, c)

    def _on_chk_click(self, r, c):
        self._set_selection(r, c)
        if self.read_only:
            if self.on_edit_blocked: self.on_edit_blocked(r, c)
            return
        cur = self._data[r][c]
        new_val = u"☐" if cur == u"☑" else u"☑"
        self.set_cell(r, c, new_val)
        self.draw()
        if self.on_cell_changed: self.on_cell_changed(r, c, new_val)

    def _on_cell_dbl_click(self, r, c):
        if self.read_only or c in self.read_only_cols:
            if self.on_edit_blocked: self.on_edit_blocked(r, c)
            return
        if c in self.search_cols:
            self._set_selection(r, c)
            self.search_cols[c](r, c)
        else:
            self._start_edit(r, c, clear=False)

    def _nav_up(self, e):
        if self._selected_cell:
            r, c = self._selected_cell
            if r > 0: self._set_selection(r-1, c)
        return "break"

    def _nav_down(self, e):
        if self._selected_cell:
            r, c = self._selected_cell
            if r < len(self._data) - 1: self._set_selection(r+1, c)
            else: self.add_row(); self._set_selection(r+1, c)
        return "break"

    def _nav_left(self, e):
        if self._selected_cell:
            r, c = self._selected_cell
            if c > 0: self._set_selection(r, c-1)
        return "break"

    def _nav_right(self, e):
        if self._selected_cell:
            r, c = self._selected_cell
            if c < len(self.columns) - 1: self._set_selection(r, c+1)
        return "break"

    def _nav_enter(self, e):
        return self._nav_down(e)

    def _nav_tab(self, e):
        if self._selected_cell:
            r, c = self._selected_cell
            if c < len(self.columns) - 1:
                self._set_selection(r, c+1)
            elif r < len(self._data) - 1:
                self._set_selection(r+1, 0)
        return "break"

    def _on_canvas_key(self, e):
        """Direct typing to edit cell (Excel feature)."""
        if not self._selected_cell or self._edit_entry: return

        # 强行捕获空格键，直接激活编辑模式（Win7 上空格可能不被 isprintable 识别）
        if e.keysym == "space":
            r, c = self._selected_cell
            if c in self.read_only_cols or self.read_only: return
            if c == self.checkbox_col or c in self.extra_checkbox_cols: return
            if c in self.search_cols: return
            self._start_edit(r, c, clear=True, initial_char=" ")
            return "break"

        # 如果按住 Ctrl 或 Alt，不触发打字（放行给复制粘贴快捷键）
        if e.state & 0x0004 or e.state & 0x0008:
            return

        # Ignore control keys
        if e.keysym in ("Up", "Down", "Left", "Right", "Return", "Tab", "Escape",
                        "Shift_L", "Shift_R", "Control_L", "Alt_L", "BackSpace", "Delete"):
            if e.keysym in ("BackSpace", "Delete"):
                r, c = self._selected_cell
                if not self.read_only and c not in self.read_only_cols:
                    self.set_cell(r, c, "")
                    self.draw()
                    if self.on_cell_changed: self.on_cell_changed(r, c, "")
            return

        # Printable character check
        if e.char and e.char.isprintable():
            r, c = self._selected_cell
            if c in self.read_only_cols or self.read_only: return
            if c == self.checkbox_col or c in self.extra_checkbox_cols: return
            if c in self.search_cols: return # Search cols must be double clicked

            self._start_edit(r, c, clear=True, initial_char=e.char)

    # ================================================================
    # Edit Mode & IME Defense
    # ================================================================
    def _copy_cell(self, event=None):
        """Ctrl+C 复制当前格内容到剪贴板"""
        if not self._selected_cell or self._edit_entry:
            return "break"
        r, c = self._selected_cell
        val = self.get_cell(r, c)
        self.clipboard_clear()
        self.clipboard_append(str(val) if val is not None else "")
        return "break"

    def _cut_cell(self, event=None):
        """Ctrl+X 剪切"""
        if not self._selected_cell or self._edit_entry:
            return "break"
        self._copy_cell()
        r, c = self._selected_cell
        if not self.read_only and c not in self.read_only_cols:
            self.set_cell(r, c, "")
            self.draw()
            if self.on_cell_changed:
                self.on_cell_changed(r, c, "")
        return "break"

    def _paste_cell(self, event=None):
        """Ctrl+V 粘贴剪贴板内容到当前格"""
        if not self._selected_cell or self._edit_entry:
            return "break"
        r, c = self._selected_cell
        if self.read_only or c in self.read_only_cols:
            if self.on_edit_blocked: self.on_edit_blocked(r, c)
            return "break"
        if c == self.checkbox_col or c in self.extra_checkbox_cols:
            return "break"
        try:
            text = self.clipboard_get()
        except tk.TclError:
            text = ""
        if text:
            text = text.replace("\n", " ").replace("\r", "").strip()
            self.set_cell(r, c, text)
            self.draw()
            if self.on_cell_changed:
                self.on_cell_changed(r, c, text)
        return "break"

    def _start_edit(self, row, col, clear=False, initial_char=""):
        self._commit_edit()
        self._selected_cell = (row, col)
        self._edit_row = row
        self._edit_col = col

        w = self.col_widths[col]
        y = row * ROW_H
        x = sum(self.col_widths[i] for i in range(col))

        current = self._data[row][col] if row < len(self._data) and col < len(self._data[row]) else ""

        # Using Excel green highlight for typing indicator
        e = tk.Entry(self.data_frame, font=self._cell_font, bg="#ffffff",
                     relief="solid", borderwidth=1, highlightthickness=2,
                     highlightcolor=self._sel_color)
        e.place(x=x, y=y, width=w, height=ROW_H)

        e.insert(0, str(current) if current else "")
        if clear:
            e.delete(0, tk.END)
            e.insert(0, initial_char)
            e.icursor(tk.END)
        else:
            e.select_range(0, tk.END)
            e.icursor(tk.END)

        e.focus_set()

        self._last_edit_time = time.time()

        def _on_key(ev):
            # Track typing time to defend against IME Return event
            self._last_edit_time = time.time()

        e.bind("<Key>", _on_key)
        e.bind("<Return>", self._on_entry_return)
        e.bind("<Escape>", lambda ev: self._cancel_edit())
        e.bind("<Up>", self._on_entry_up)
        e.bind("<Down>", self._on_entry_down)
        e.bind("<Tab>", self._on_entry_tab)
        self._edit_entry = e

    def _on_entry_return(self, ev):
        """Smart Enter: Ignores IME composition commit, otherwise moves down."""
        # IME Protection: If text was changed in the last 50ms, it's likely an IME commit.
        if time.time() - self._last_edit_time < 0.05:
            return "break"

        self._commit_edit()
        r, c = self._selected_cell
        if r < len(self._data) - 1:
            self._set_selection(r+1, c)
        else:
            self.add_row()
            self._set_selection(r+1, c)
        return "break"

    def _on_entry_up(self, ev):
        self._commit_edit()
        r, c = self._selected_cell
        if r > 0: self._set_selection(r-1, c)
        return "break"

    def _on_entry_down(self, ev):
        return self._on_entry_return(ev)

    def _on_entry_tab(self, ev):
        self._commit_edit()
        r, c = self._selected_cell
        if c < len(self.columns) - 1: self._set_selection(r, c+1)
        return "break"

    def commit_edit(self):
        self._commit_edit()

    def _commit_edit(self):
        if not self._edit_entry: return
        try: new_val = self._edit_entry.get()
        except tk.TclError: new_val = ""

        self._edit_entry.destroy()
        self._edit_entry = None
        r, c = self._edit_row, self._edit_col
        self._edit_row = None
        self._edit_col = None

        if r is not None and c is not None:
            self.set_cell(r, c, new_val)
            self.draw()
            if self.on_cell_changed:
                self.on_cell_changed(r, c, new_val)

        self.canvas.focus_set()

    def _cancel_edit(self):
        if self._edit_entry:
            self._edit_entry.destroy()
            self._edit_entry = None
        self._edit_row = None
        self._edit_col = None
        self.canvas.focus_set()

    # ================================================================
    # Data Access APIs
    # ================================================================
    def get_data(self): return [list(row) for row in self._data]
    def set_data(self, data): self._data = [list(row) for row in data]; self.draw()
    def get_cell(self, row, col):
        return self._data[row][col] if row < len(self._data) and col < len(self._data[row]) else ""
    def set_cell(self, row, col, value):
        while row >= len(self._data):
            self._data.append([u"☐" if i == self.checkbox_col or i in self.extra_checkbox_cols else "" for i in range(len(self.columns))])
        while col >= len(self._data[row]): self._data[row].append("")
        self._data[row][col] = value
    def set_display_override(self, row, col, text):
        if text: self._display_overrides[(row, col)] = text
        else: self._display_overrides.pop((row, col), None)
    def clear_overrides(self): self._display_overrides.clear()
    def set_highlight_cols(self, cols):
        self.highlight_cols = set(cols or [])
        self.draw()
    def row_count(self): return len(self._data)

    def add_row(self, values=None):
        if values: self._data.append(list(values))
        else: self._data.append([u"☐" if i == self.checkbox_col or i in self.extra_checkbox_cols else "" for i in range(len(self.columns))])
        self.draw()

    def delete_rows(self, indices):
        for ri in sorted(indices, reverse=True):
            if ri < len(self._data): self._data.pop(ri)
        self.draw()

    def get_selected_rows(self):
        if self.checkbox_col is None: return []
        return [ri for ri, row in enumerate(self._data) if ri < len(row) and row[self.checkbox_col] == u"☑"]
