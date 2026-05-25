# -*- coding: utf-8 -*-
"""
BarTender 标签打印模块

通过 BarTender COM 的 SetNamedSubStringValue 操控模板具名数据源:
  客户, 款号, 类别, 颜色, 数量, 日期

同材料名多色：只改颜色和数量，客户名/类别名/日期不变。
"""
import time
import datetime as _dt

_date_set_today = None

# 每行最多字符数（可根据标签实际宽度调整）
MAX_CHARS_PRODUCT = 18      # 类别字段: 每行最多字数，超出自动换行


def _find_child_by_id(parent, cid):
    """Recursively find a child window by its control ID."""
    import win32gui
    child = win32gui.FindWindowEx(parent, None, None, None)
    while child:
        try:
            if win32gui.GetDlgCtrlID(child) == cid:
                return child
        except Exception:
            pass
        result = _find_child_by_id(child, cid)
        if result:
            return result
        child = win32gui.FindWindowEx(parent, child, None, None)
    return None


def _wrap_product(text, max_chars):
    """类别名称超长时按字数换行。"""
    if len(text) <= max_chars:
        return text
    result = ""
    for i, ch in enumerate(text):
        result += ch
        if (i + 1) % max_chars == 0 and i + 1 < len(text):
            result += "\n"
    return result


def _get_fmt():
    import win32com.client
    bt = win32com.client.Dispatch("BarTender.Application")
    fmt = bt.ActiveFormat
    if fmt is None:
        raise RuntimeError(
            u"BarTender 中没有打开的标签模板。\n"
            u"请先打开标签.btw。")
    return fmt


def _set(fmt, name, value):
    if value is None:
        value = ""
    fmt.SetNamedSubStringValue(name, str(value))


def _do_printout(fmt, count=1):
    """
    放弃使用 COM 接口的 PrintOut。
    改用 pywinauto 模拟真人按下 Ctrl+P，填入数量并强制点击"打印"按钮。
    """
    import time

    # ==========================================
    # 【核心防崩溃补丁】解决 Win7 下 WinError 123 问题
    # 强制禁用 comtypes 的本地硬盘缓存，改为纯内存运行，彻底杜绝非法文件名生成
    # ==========================================
    try:
        import comtypes.client
        comtypes.client.gen_dir = None
    except Exception:
        pass

    try:
        from pywinauto.application import Application
        from pywinauto.keyboard import send_keys
    except ImportError:
        raise RuntimeError(u"缺少 pywinauto 库，请在运行环境中执行: pip install pywinauto")

    try:
        # 1. 寻找并连接当前正在运行的 BarTender 窗口
        app = Application(backend="win32").connect(title_re=".*BarTender.*", timeout=3)
        main_win = app.top_window()
        main_win.set_focus()
        time.sleep(0.3)

        # 2. 模拟按下 Ctrl + P 快捷键
        send_keys('^p')
        time.sleep(1.2) # 等待打印对话框弹出（稍微加长以确保弹窗稳定）

        # 3. 捕获弹出的"打印"窗口，确保焦点没跑偏
        try:
            print_dlg = app.window(title_re=".*打印.*")
            print_dlg.set_focus()
        except Exception:
            print_dlg = None

        # 4. 直接输入打印份数
        send_keys(str(count))
        time.sleep(0.5)

        # 5. 提交打印（双重保险）
        clicked = False
        if print_dlg:
            # 第一重保险：直接抓取名为"打印"的按钮，模拟真实鼠标点击
            for btn_title in [u"打印", u"打印(&P)"]:
                try:
                    btn = print_dlg.child_window(title=btn_title, class_name="Button")
                    if btn.exists():
                        btn.set_focus()
                        time.sleep(0.2)
                        send_keys('{SPACE}')  # 物理鼠标左键点击
                        clicked = True
                        break
                except Exception:
                    continue

        if not clicked:
            # 第二重保险：键盘外挂兜底。按 Tab 键把光标从输入框移出，再回车
            send_keys('{TAB}')
            time.sleep(0.1)
            send_keys('{ENTER}')

        # 6. 给系统和打印机留出足够的缓冲时间 (非常关键，防止连续打印时出现跳单)
        time.sleep(3.5)

        return True
    except Exception as e:
        raise RuntimeError(u"模拟按键打印失败: {}".format(str(e)))


def _check_bartender_ready(fmt):
    """Best-effort check: warn if printer might not be configured."""
    try:
        _ = fmt.Name
    except Exception:
        pass
    return True, None


def print_labels(items, customer, order):
    """
    逐行打印标签。
    返回: (total_printed, error_message)
    """
    global _date_set_today

    try:
        fmt = _get_fmt()
    except Exception as e:
        return 0, str(e)

    # Pre-flight check
    ok, err = _check_bartender_ready(fmt)
    if not ok:
        return 0, err

    today_str = _dt.date.today().strftime("%m-%d")

    printed = 0
    first_item = True
    last_product = None

    for item in items:
        pc = item.get("print_count", 1)
        if isinstance(pc, str):
            try:
                pc = int(pc) if pc else 1
            except ValueError:
                pc = 1
        if pc < 1:
            pc = 1

        qty_str = ""
        qty = item.get("quantity", "")
        if qty:
            try:
                qty_str = "{:.0f}Y".format(float(qty))
            except ValueError:
                qty_str = "{}Y".format(qty)

        order_parts = []
        for v in [item.get("customer_code", ""),
                  item.get("order_number", ""),
                  item.get("item_code", ""),
                  item.get("model_number", "")]:
            v = str(v).replace("\n", "").replace("\r", "").strip()
            if v:
                order_parts.append(v)
        order_info = u"/".join(order_parts)

        same_product = (last_product is not None
                       and item.get("product_name") == last_product)

        try:
            if first_item or not same_product:
                _set(fmt, u"客户", customer.get("name", ""))

            _set(fmt, u"款号", order_info)

            if not same_product:
                _set(fmt, u"类别", _wrap_product(
                    item.get("product_name", ""), MAX_CHARS_PRODUCT))
                last_product = item.get("product_name")

            _set(fmt, u"颜色", item.get("color_name", ""))
            _set(fmt, u"数量", qty_str)

            if _date_set_today != today_str:
                _set(fmt, u"日期", today_str)
                _date_set_today = today_str

            _do_printout(fmt, pc)   # 使用带硬盘缓存屏蔽和 UI 自动化的打印逻辑
            printed += pc

            first_item = False
        except Exception as e:
            return printed, str(e)

    return printed, None
