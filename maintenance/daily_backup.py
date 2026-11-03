# -*- coding: utf-8 -*-
"""Create a validated business-data snapshot for the daily scheduled task."""
import csv
import datetime
import json
import os
import sys

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, BASE)
os.chdir(BASE)

from backup_manager import create_backup  # noqa: E402

LOG = os.path.join(BASE, "data", "backup_task.log")
os.makedirs(os.path.dirname(LOG), exist_ok=True)


def write_stock_snapshot(base=None):
    """每天存一份各物料数量快照（JSON 按天 + CSV 累计），用于观察消耗与异常损耗。"""
    base = base or BASE
    stock_path = os.path.join(base, "stock.json")
    if not os.path.exists(stock_path):
        return None
    snap_dir = os.path.join(base, "data", "snapshots")
    os.makedirs(snap_dir, exist_ok=True)
    with open(stock_path, encoding="utf-8") as fp:
        data = json.load(fp)
    items = data.get("items", [])
    today = datetime.date.today().strftime("%Y%m%d")
    payload = {
        "date": today,
        "stock_date": str(data.get("stock_date") or data.get("date") or ""),
        "items": [{"short": (it.get("short") or it.get("name") or ""),
                   "qty": it.get("qty"), "unit": it.get("unit", ""),
                   "cat": it.get("cat", "")}
                  for it in items if not str(it.get("cat") or "").startswith("产成品")],
    }
    json_path = os.path.join(snap_dir, "stock_%s.json" % today)
    with open(json_path, "w", encoding="utf-8") as fp:
        json.dump(payload, fp, ensure_ascii=False, indent=1)
    csv_path = os.path.join(snap_dir, "stock_history.csv")
    existing_days = set()
    if os.path.exists(csv_path):
        with open(csv_path, encoding="utf-8", newline="") as fp:
            for row in csv.reader(fp):
                if len(row) >= 4 and row[0] == today:
                    existing_days.add((row[0], row[1]))
    new_file = not os.path.exists(csv_path)
    appended = 0
    with open(csv_path, "a", encoding="utf-8", newline="") as fp:
        writer = csv.writer(fp)
        if new_file:
            writer.writerow(["date", "short", "qty", "unit"])
        for item in payload["items"]:
            if (today, item["short"]) in existing_days:
                continue
            writer.writerow([today, item["short"], item["qty"], item["unit"]])
            appended += 1
    # JSON 快照最多保留 180 天
    cutoff = (datetime.date.today() - datetime.timedelta(days=180)).strftime("%Y%m%d")
    for name in os.listdir(snap_dir):
        if name.startswith("stock_") and name.endswith(".json"):
            day = name[6:-5]
            if day.isdigit() and day < cutoff:
                os.remove(os.path.join(snap_dir, name))
    return json_path


def main():
    stamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    try:
        result = create_backup()
        line = "%s OK %s" % (stamp, result.get("path"))
    except Exception as exc:
        line = "%s FAIL %s: %s" % (stamp, type(exc).__name__, exc)
    try:
        snapshot = write_stock_snapshot()
        if snapshot:
            line += " 快照=%s" % snapshot
    except Exception as exc:
        line += " 快照失败:%s: %s" % (type(exc).__name__, exc)
    with open(LOG, "a", encoding="utf-8") as fp:
        fp.write(line + "\n")
    print(json.dumps({"ok": " OK " in line, "line": line}, ensure_ascii=False))


if __name__ == "__main__":
    main()
