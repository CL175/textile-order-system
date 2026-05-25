# -*- coding: utf-8 -*-
"""测试 BarTender 具名数据源。"""
import sys, os, time
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))


def test():
    from logic.label_print import _wrap_product

    # 测试折行效果
    product = u"D4530/7MM一面贴超细佳积布一面贴超细鸟眼布底领面过快干胶"
    wrapped = _wrap_product(product, 18)
    print("类别原文 ({}字):".format(len(product)))
    print("  {}".format(product))
    print("折行后:")
    for line in wrapped.split("\n"):
        print("  {}".format(line))

    print()

    # COM test
    import win32com.client

    print("COM connect...")
    bt = win32com.client.Dispatch("BarTender.Application")
    fmt = bt.ActiveFormat
    print("OK")

    order_info = u"ZQ-WM3715/ZQ-XY11028/BU-71092/1"

    tests = [
        (u"客户", u"测试客户"),
        (u"款号", order_info),
        (u"类别", wrapped),
        (u"颜色", u"黑色"),
        (u"数量", u"100Y"),
        (u"日期", u"05-14"),
    ]
    for name, val in tests:
        try:
            fmt.SetNamedSubStringValue(name, val)
            print("   {} = {}  OK".format(name, val.replace("\n", "\\n")))
        except Exception as e:
            print("   {} FAIL: {}".format(name, str(e)[:120]))

    print("\nCheck BarTender template. Close WITHOUT saving.")


if __name__ == "__main__":
    test()
