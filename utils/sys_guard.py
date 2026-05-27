# -*- coding: utf-8 -*-
"""System guard: automatic backup, database vacuum, monthly export routing.

All functions are wrapped in try/except so that any error (missing drive,
permission denied, disk full) is silently skipped — never crash the main app.
"""
import os
import shutil
import datetime
import glob
import sqlite3


# ============================================================
# 风险一：数据库防损坏 + 自动瘦身
# ============================================================

def auto_backup_db(db_path="textile.db", backup_dir="backup", keep_days=7):
    """每日首次启动时自动备份数据库，并清理超过 keep_days 天的旧备份。"""
    if not os.path.exists(db_path):
        return

    if not os.path.exists(backup_dir):
        try:
            os.makedirs(backup_dir)
        except Exception:
            return

    today_str = datetime.date.today().strftime("%Y%m%d")
    backup_path = os.path.join(backup_dir,
                               "textile_backup_{}.db".format(today_str))

    # 今天还没备份过才复制
    if not os.path.exists(backup_path):
        try:
            shutil.copy2(db_path, backup_path)
        except Exception:
            pass

    # 清理超期旧备份
    try:
        pattern = os.path.join(backup_dir, "textile_backup_*.db")
        now = datetime.datetime.now()
        for f in glob.glob(pattern):
            mtime = datetime.datetime.fromtimestamp(os.path.getmtime(f))
            if (now - mtime).days > keep_days:
                os.remove(f)
    except Exception:
        pass


def vacuum_database(db_path="textile.db"):
    """整理并压缩数据库体积。可在设置界面加按钮手动调用。"""
    try:
        conn = sqlite3.connect(db_path)
        conn.execute("VACUUM")
        conn.commit()
        conn.close()
        return True, u"数据库整理完成，体积已优化！"
    except Exception as e:
        return False, u"数据库整理失败: {}".format(str(e))


# ============================================================
# 风险二：导出文件按月分流，防止单个文件夹堆积上万个 Excel
# ============================================================

def get_monthly_export_dir(base_dir=u"D:\\送货单导出"):
    """按月份自动生成文件夹路径（如 D:\送货单导出\2026-05）。"""
    today = datetime.date.today()
    month_folder = today.strftime("%Y-%m")
    target_dir = os.path.join(base_dir, month_folder)

    if not os.path.exists(target_dir):
        try:
            os.makedirs(target_dir)
        except Exception:
            target_dir = os.path.join(os.getcwd(), "export", month_folder)
            if not os.path.exists(target_dir):
                os.makedirs(target_dir)

    return target_dir
