# -*- coding: utf-8 -*-
"""探查 WPS 打印对话框的控件结构 (加强版)."""
import sys, os
sys.stdout.reconfigure(encoding='utf-8') if hasattr(sys.stdout, 'reconfigure') else None

try:
    import pywinauto
except ImportError:
    print("pip install pywinauto")
    sys.exit(1)

from pywinauto import Desktop

print("=== 1. 所有顶层窗口 ===")
desktop = Desktop(backend="win32")
for w in desktop.windows():
    try:
        t = w.window_text()
        if t and len(t.strip()) > 0:
            print("  class=[{}] text=[{}] rect={}".format(
                w.class_name(), t, w.rectangle()))
    except:
        pass

print("\n=== 2. 查找打印对话框 ===")
dlg = None
# Try various titles
for pattern in [".*打印.*", ".*Print.*", ".*WPS.*", ".*wps.*"]:
    try:
        found = desktop.window(title_re=pattern)
        if found.exists():
            print("匹配 [{}]: class={} text={}".format(
                pattern, found.class_name(), found.window_text()))
            if not dlg:
                dlg = found
    except:
        pass

if dlg is None:
    print("\n未找到打印对话框！请确认 WPS 中 Ctrl+P 的打印对话框已打开。")
    sys.exit(1)

print("\n=== 3. win32 控件树 (depth=3) ===")
try:
    dlg.print_control_identifiers(depth=3)
except Exception as e:
    print("出错: {}".format(e))

print("\n=== 4. 尝试 UIA backend ===")
try:
    desktop_uia = Desktop(backend="uia")
    for w in desktop_uia.windows():
        try:
            t = w.window_text()
            if t and "打印" in t:
                print("UIA 找到: class={} text={} rect={}".format(
                    w.class_name(), t, w.rectangle()))
        except:
            pass
    # Try directly
    dlg_uia = desktop_uia.window(title_re=".*打印.*")
    if dlg_uia.exists():
        print("\nUIA 控件树:")
        dlg_uia.print_control_identifiers(depth=3)
    else:
        print("UIA title_re 未匹配到打印对话框。")
except Exception as e:
    print("UIA 出错: {}".format(e))

print("\n=== 5. 尝试通过窗口类名查找 Qt5 内部控件 ===")
# Qt5 windows sometimes expose children via window enumeration
try:
    def enum_children(hwnd, depth=0):
        import win32gui
        result = []
        try:
            text = win32gui.GetWindowText(hwnd)
            cls = win32gui.GetClassName(hwnd)
            rect = win32gui.GetWindowRect(hwnd)
            if text or depth <= 1:
                result.append((depth, cls, text, rect))
        except:
            pass
        try:
            child = win32gui.GetWindow(hwnd, 5)  # GW_CHILD
            while child:
                result.extend(enum_children(child, depth + 1))
                child = win32gui.GetWindow(child, 2)  # GW_HWNDNEXT
        except:
            pass
        return result

    import win32gui
    hwnd = win32gui.FindWindow(None, None)
    dlg_hwnd = None
    while hwnd:
        try:
            if "打印" in win32gui.GetWindowText(hwnd):
                cls = win32gui.GetClassName(hwnd)
                if "Qt5" in cls or "QWindow" in cls:
                    dlg_hwnd = hwnd
                    print("找到对话框 HWND: {} class={}".format(hwnd, cls))
                    break
        except:
            pass
        hwnd = win32gui.GetWindow(hwnd, 2)  # GW_HWNDNEXT

    if dlg_hwnd:
        print("\n对话框子窗口树:")
        for depth, cls, text, rect in enum_children(dlg_hwnd):
            if depth <= 4:
                indent = "  " * depth
                print("{}class=[{}] text=[{}] rect={}".format(indent, cls, text, rect))
except Exception as e:
    print("win32gui 枚举出错: {}".format(e))

print("\n=== 完成 ===")
