# -*- coding: utf-8 -*-
"""命令行自动更新库存（供 Windows 任务计划程序调用，不依赖界面是否打开）

用法：
    python auto_update.py            # 自动模式：有新表才导入
    python auto_update.py force      # 强制：忽略指纹重新导入（应急用）
"""
import os
import sys
import json
import datetime

APP = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, APP)

LOG = os.path.join(APP, "data", "auto_import.log")


def log(text):
    line = "%s  %s" % (datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"), text)
    try:
        os.makedirs(os.path.dirname(LOG), exist_ok=True)
        with open(LOG, "a", encoding="utf-8") as f:
            f.write(line + "\n")
    except Exception:
        pass
    print(line)


def main():
    force = len(sys.argv) > 1 and sys.argv[1] == "force"
    fail = 0
    try:
        import app_gui
    except Exception as e:
        log("导入失败：无法加载业务模块 %s" % e)
        return 2

    # 1) 库存表：内部已带 3 次重试；失败返回 fail_count
    ok = False
    for attempt in range(3):
        r = app_gui.auto_import_stock(force)
        if r.get("ok"):
            if r.get("updated"):
                log("库存表已更新：%s 项，库存日期 %s，来源 %s" % (r.get("n"), r.get("sd"), r.get("source")))
            else:
                log("库存表无新表（%s），上次导入 %s" % (r.get("reason", ""), r.get("last_time", "无")))
            fail = int(r.get("fail_count", 0))
            ok = True
            break
        fail = int(r.get("fail_count", 0))
        log("库存表导入失败（第 %d 次）：%s" % (attempt + 1, r.get("msg", "")))
        import time
        time.sleep(5)

    # 2) 批次表
    try:
        b = app_gui.import_batches(None, force)
        if b.get("ok") and b.get("updated"):
            log("批次表已更新：%s 个物料 / %s 个批号，来源 %s" % (b.get("items"), b.get("batches"), b.get("source")))
        elif b.get("ok"):
            log("批次表无新表（%s）" % b.get("reason", ""))
        else:
            log("批次表：%s" % b.get("msg", ""))
    except Exception as e:
        log("批次表导入异常：%s" % e)

    if not ok:
        log("★ 连续失败 %d 次，请手动更新：打开系统 → 库存页 →「导入最新库存表（手动）」" % fail)
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())
