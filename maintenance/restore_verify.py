# -*- coding: utf-8 -*-
"""季度恢复演练：取最近一份备份，恢复到隔离临时目录并校验结构，不触碰真实数据。
用法：python maintenance/restore_verify.py  （退出码 0=通过）"""
import datetime
import json
import os
import sys
import tempfile

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, BASE)
os.chdir(BASE)

from backup_manager import BACKUP_DIR, list_backups, _validated_backup  # noqa: E402

LOG = os.path.join(BASE, "data", "restore_verify.log")


def main():
    stamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    try:
        backups = [b for b in list_backups() if b.get("valid")]
        if not backups:
            raise ValueError("没有可用的有效备份（data\\backups 为空或全部损坏）")
        latest = backups[0]
        path = latest["path"]
        with tempfile.TemporaryDirectory(prefix="restore-verify-") as tmp:
            restored = _validated_backup(path)
            # 写入隔离目录并回读，证明文件本身可完整恢复
            for name, payload in restored.items():
                target = os.path.join(tmp, name)
                with open(target, "w", encoding="utf-8") as fp:
                    json.dump(payload, fp, ensure_ascii=False)
                with open(target, encoding="utf-8") as fp:
                    json.load(fp)
            orders = restored["orders.json"]
            order_list = orders.get("orders", []) if isinstance(orders, dict) else orders
            stock_items = restored["stock.json"].get("items", [])
            line = ("%s OK 恢复演练通过：%s（%s 生成）→ 隔离目录校验 "
                    "orders=%d 台账 / stock=%d 项 / recipes / mapping 全部可解析"
                    % (stamp, os.path.basename(path), latest.get("time", ""),
                       len(order_list), len(stock_items)))
            code = 0
    except Exception as exc:
        line = "%s FAIL %s: %s" % (stamp, type(exc).__name__, exc)
        code = 1
    os.makedirs(os.path.dirname(LOG), exist_ok=True)
    with open(LOG, "a", encoding="utf-8") as fp:
        fp.write(line + "\n")
    print(line)
    return code


if __name__ == "__main__":
    sys.exit(main())
