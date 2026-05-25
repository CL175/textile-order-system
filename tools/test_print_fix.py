# -*- coding: utf-8 -*-
"""测试修复后的打印逻辑 —— 不依赖打印机，仅验证 COM 调用是否正常。"""
import sys, os, time
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import win32com.client

print("=" * 60)
print("测试 1: 连接 BarTender COM...")
try:
    bt = win32com.client.Dispatch("BarTender.Application")
    print("  连接成功 [OK]")
except Exception as e:
    print("  连接失败: {}".format(e))
    sys.exit(1)

print()
print("测试 2: 获取活动格式...")
fmt = bt.ActiveFormat
if fmt is None:
    print("  失败: BarTender 中没有打开的标签模板！")
    print("  请先打开标签模板(.btw)再运行测试。")
    sys.exit(1)
print("  获取成功: {}".format(fmt))
print("  模板文件: {}".format(fmt.FileName if hasattr(fmt, 'FileName') else "未知"))

print()
print("测试 3: 设置副本数 (IdenticalCopiesOfLabel)...")
try:
    fmt.IdenticalCopiesOfLabel = 3
    actual = fmt.IdenticalCopiesOfLabel
    print("  设置成功 [OK]  设置为 3, 读回 = {}".format(actual))
except Exception as e:
    print("  不支持 IdenticalCopiesOfLabel: {}".format(e))
    print("  (BarTender UltraLite 可能不支持此属性)")

print()
print("测试 4: NumberOfSerializedLabels (另一种副本控制)...")
try:
    current = fmt.NumberOfSerializedLabels
    fmt.NumberOfSerializedLabels = 3
    actual = fmt.NumberOfSerializedLabels
    print("  设置成功 [OK]  设置为 3, 读回 = {}".format(actual))
except Exception as e:
    print("  不支持 NumberOfSerializedLabels: {}".format(e))

print()
print("测试 5: 设置具名数据源...")
tests = [
    (u"客户", u"测试客户"),
    (u"款号", u"TEST-001/M-001"),
    (u"类别", u"测试面料"),
    (u"颜色", u"红色"),
    (u"数量", u"100Y"),
    (u"日期", u"05-20"),
]
all_ok = True
for name, val in tests:
    try:
        fmt.SetNamedSubStringValue(name, val)
        print("  {} = {}  [OK]".format(name, val))
    except Exception as e:
        print("  {} = {}  [FAIL] ({})".format(name, val, str(e)[:80]))
        all_ok = False

print()
print("测试 6: PrintOut(False, True) -- 不弹对话框打印...")
print("  (你没有打印机，这个调用可能会报错，这是正常的)")
try:
    fmt.PrintOut(False, True)
    print("  调用成功 [OK] (可能打印到了虚拟打印机)")
except Exception as e:
    print("  调用失败 (预期): {}".format(str(e)[:120]))

print()
print("=" * 60)
print("测试完成！")
print()
print("请把上面的输出复制给我，我来判断 BarTender UltraLite 支持哪些功能。")
print("特别关注测试 3 和测试 4 的结果。")
