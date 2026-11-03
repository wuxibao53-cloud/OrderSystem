# -*- coding: utf-8 -*-
"""每日播报 v2：生成日报 + 数据自动备份 + 库存数据时效检查"""
import os
import sys
import json
import shutil
import ctypes
import datetime

APP = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, APP)

from stock import generate_daily  # noqa: E402


def backup_core():
    """每天首次播报时，把三份核心数据备份到 data/backup/日期/"""
    date8 = datetime.datetime.now().strftime("%Y%m%d")
    bdir = os.path.join(APP, "data", "backup", date8)
    os.makedirs(bdir, exist_ok=True)
    for fn in ("recipes.json", "stock.json", "mapping.json"):
        src = os.path.join(APP, fn)
        if os.path.exists(src):
            shutil.copy2(src, os.path.join(bdir, fn))
    return bdir


def stock_freshness():
    """检查库存数据是否是今天导入的"""
    p = os.path.join(APP, "stock.json")
    if not os.path.exists(p):
        return "missing"
    data = json.load(open(p, encoding="utf-8"))
    sd = (data.get("_meta") or {}).get("stock_date", "")
    today = datetime.datetime.now().strftime("%Y-%m-%d")
    if not sd:
        return "unknown"
    if sd.strip() != today:
        return "stale:" + sd
    return "fresh"


def main():
    # 1) 数据备份
    bdir = backup_core()

    # 2) 库存时效
    fresh = stock_freshness()

    # 3) 日报
    try:
        path, alert, warn, n = generate_daily()
    except Exception as e:
        ctypes.windll.user32.MessageBoxW(None, "日报生成失败：%s" % e, "跟单日报", 0x10)
        return

    lines = []
    if fresh.startswith("stale"):
        lines.append("提醒：库存数据还是 %s 的，今天还没导入新的库存余额表。" % fresh.split(":")[1])
        lines.append("（从柠檬云导出后：库存管理 → 5 导入）")
        lines.append("")
    try:
        from orders import urgent_orders
        late, soon = urgent_orders()
        if late:
            lines.append("⚠ 逾期订单 %d 单（必须当日处理）：" % len(late))
            for o, dd in late[:5]:
                lines.append("  · %s %s　逾期 %d 天　%s" %
                             (o.get("serial"), o.get("product"), -dd, o.get("status")))
            lines.append("")
        if soon:
            lines.append("3 天内到期 %d 单：" % len(soon))
            for o, dd in soon[:5]:
                lines.append("  · %s %s　剩 %d 天" % (o.get("serial"), o.get("product"), dd))
            lines.append("")
    except Exception:
        pass
    lines.append("今日指令单：%d 单" % n)
    lines.append("库存警报：%d 项　偏低：%d 项" % (len(alert), len(warn)))
    if alert:
        lines.append("")
        for r in alert[:6]:
            lines.append("  · %s 现有 %s%s（警戒 %s）" %
                         (r["short"], r["qty"], r["unit"], r["line"]))
        if len(alert) > 6:
            lines.append("  …共 %d 项" % len(alert))
    lines.append("")
    lines.append("日报：%s" % path)
    lines.append("数据已备份：%s" % bdir)
    ctypes.windll.user32.MessageBoxW(None, "\n".join(lines), "跟单日报", 0x30)


if __name__ == "__main__":
    main()
