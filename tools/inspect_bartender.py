# -*- coding: utf-8 -*-
"""
BarTender 控件探查工具

运行前请确保:
1. BarTender 已打开，标签模板已加载
2. 标签模板中的文本框都可见

使用方法:
    python inspect_bartender.py

输出文件: D:/xxm/bartender_controls.txt
打开这个文件，找到各字段对应的控件名，告诉我即可。
"""
import os
import sys
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from pywinauto.application import Application
from pywinauto.findwindows import find_windows

OUTPUT = "D:/xxm/bartender_controls.txt"


def inspect_window(app, name, f):
    """Print all controls in a window."""
    try:
        window = app.window(title=name)
        window.wait("exists", timeout=2)
    except Exception:
        try:
            window = app[name]
        except Exception:
            return

    f.write(u"\n{'='*60}\n")
    f.write(u"窗口: {}\n".format(name))
    f.write(u"{}\n".format('='*60))

    try:
        # Print full control tree
        f.write(u"\n--- 控件树 (print_control_identifiers) ---\n")
        f.write(window.print_control_identifiers())
    except Exception as e:
        f.write(u"print_control_identifiers 失败: {}\n".format(e))

    # List all descendants manually
    f.write(u"\n--- 子孙控件明细 ---\n")
    try:
        for child in window.descendants():
            info = (
                u"  [{}] class='{}' title='{}' auto_id='{}' rect={}"
                .format(
                    child.element_info.control_id,
                    child.element_info.class_name,
                    child.element_info.name or "",
                    child.element_info.automation_id or "",
                    child.element_info.rectangle,
                ))
            f.write(info + "\n")
    except Exception as e:
        f.write(u"descendants 失败: {}\n".format(e))


def main():
    # Ensure output directory exists
    out_dir = os.path.dirname(OUTPUT)
    if out_dir and not os.path.exists(out_dir):
        os.makedirs(out_dir)

    with open(OUTPUT, "w", encoding="utf-8") as f:
        f.write(u"BarTender 控件探查结果\n")
        f.write(u"=" * 60 + "\n")

        # ---- Method 1: Find BarTender by window title ----
        f.write(u"\n\n### 方法1: 按标题查找 BarTender 窗口\n")
        try:
            bartender = Application(backend="win32").connect(
                title_re=".*BarTender.*", timeout=3)
            f.write(u"连接成功 (win32 backend)\n")

            for w in bartender.windows():
                f.write(u"\n顶层窗口: '{}' class='{}'\n".format(
                    w.element_info.name, w.element_info.class_name))

            # Dump control tree of the first real window
            for w in bartender.windows():
                name = w.element_info.name
                if name and len(name) > 2:
                    inspect_window(bartender, name, f)
        except Exception as e:
            f.write(u"方法1失败: {}\n".format(e))

        # ---- Method 2: Try UIA backend ----
        f.write(u"\n\n### 方法2: UIA backend\n")
        try:
            bartender = Application(backend="uia").connect(
                title_re=".*BarTender.*", timeout=3)
            f.write(u"连接成功 (uia backend)\n")
            for w in bartender.windows():
                name = w.element_info.name
                if name and len(name) > 2:
                    inspect_window(bartender, name, f)
        except Exception as e:
            f.write(u"方法2失败: {}\n".format(e))

        # ---- Method 3: List ALL top-level windows ----
        f.write(u"\n\n### 方法3: 系统所有顶层窗口（找BarTender相关）\n")
        try:
            import ctypes
            from pywinauto.findwindows import find_elements

            all_wins = find_elements()
            for w in all_wins:
                name = w.name or ""
                class_name = w.class_name or ""
                if "bartender" in name.lower() or "bartend" in class_name.lower():
                    f.write(u"  找到: '{}' class='{}'\n".format(name, class_name))
            if not any("bartender" in (w.name or "").lower() for w in all_wins):
                f.write(u"  未找到含有 'BarTender' 的窗口。\n")
                f.write(u"  正在列出所有可见顶层窗口:\n")
                for w in all_wins:
                    name = w.name or ""
                    if name and len(name) > 1:
                        f.write(u"    '{}' (class={})\n".format(
                            name, w.class_name or ""))
        except Exception as e:
            f.write(u"方法3失败: {}\n".format(e))

    print(u"探查完成！结果已写入: {}".format(OUTPUT))
    print(u"请打开该文件，找各字段对应的控件名。")


if __name__ == "__main__":
    main()
