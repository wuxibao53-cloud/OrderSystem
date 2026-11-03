# -*- coding: utf-8 -*-
"""柠檬云库存同步：读取 D:\\Downloads 的两个导出表（商品库存余额表 + 批次跟踪表），
与系统库存核对 → 安全自动修正（编码精确匹配才改，改前备份、留修正历史）→
生成核对报告与采购建议（按历史订单消耗速率 + 待生产占用推算）。
设计红线：模糊匹配/单位冲突一律只报告不改；任何修正都可在备份里回滚。"""
import csv
import datetime
import json
import os
import re
import shutil

import json_store

APP = os.path.dirname(os.path.abspath(__file__))
LEMON_DIR = os.path.join(APP, "data", "lemon")
CONFIG_PATH = os.path.join(LEMON_DIR, "lemon_config.json")
HISTORY_CSV = os.path.join(LEMON_DIR, "修正历史.csv")

CATEGORY_PREFIXES = ("直接原料-", "添加剂-", "辅料-", "包材-", "原材料-",
                     "产成品-", "半成品-", "包装材料-", "低值易耗-")
UNIT_ALIASES = {"kg": "kg", "千克": "kg", "公斤": "kg", "KG": "kg",
                "g": "g", "克": "g", "瓶": "瓶", "个": "个", "套": "套",
                "张": "张", "只": "只", "袋": "袋", "盒": "盒", "片": "片"}


def load_config():
    cfg = {"balance_file": "商品库存余额表.xls",
           "batch_file": "批次跟踪表.xls",
           "downloads_dir": r"D:\Downloads",
           "auto_fix": True,
           "min_diff": 0.01}
    if os.path.exists(CONFIG_PATH):
        try:
            with open(CONFIG_PATH, encoding="utf-8-sig") as fp:
                user = json.load(fp)
            if isinstance(user, dict):
                cfg.update({k: v for k, v in user.items() if k in cfg})
        except (OSError, ValueError):
            pass
    return cfg


def _source_files(cfg):
    balance = os.path.join(cfg["downloads_dir"], cfg["balance_file"])
    batch = os.path.join(cfg["downloads_dir"], cfg["batch_file"])
    return balance, batch


def normalize_name(name):
    text = str(name or "").strip()
    for prefix in CATEGORY_PREFIXES:
        if text.startswith(prefix):
            return text[len(prefix):].strip()
    return text


def norm_unit(unit):
    return UNIT_ALIASES.get(str(unit or "").strip(), str(unit or "").strip())


def _num(value):
    try:
        number = float(value)
    except (TypeError, ValueError):
        return None
    return number if number == number else None  # 过滤 NaN


def read_balance(path):
    """商品库存余额表 → [{code, name, cat, unit, qty, cost}]（全部仓库.单位数量）。"""
    import xlrd
    book = xlrd.open_workbook(path)
    sheet = book.sheet_by_index(0)
    rows = []
    for r in range(5, sheet.nrows):
        code = str(sheet.cell_value(r, 1)).strip()
        if not code:
            continue
        rows.append({
            "code": code,
            "name": str(sheet.cell_value(r, 2)).strip(),
            "cat": str(sheet.cell_value(r, 3)).strip(),
            "unit": str(sheet.cell_value(r, 5)).strip(),
            "qty": _num(sheet.cell_value(r, 8)) or 0.0,
            "cost": _num(sheet.cell_value(r, 11)),
        })
    return rows


def read_batches(path):
    """批次跟踪表 → 流水行 [{name, batch, mfg, exp, doc_date, biz, warehouse, qty}]。"""
    import xlrd
    book = xlrd.open_workbook(path)
    sheet = book.sheet_by_index(0)
    rows = []
    for r in range(4, sheet.nrows):
        name = str(sheet.cell_value(r, 2)).strip()
        batch = str(sheet.cell_value(r, 6)).strip()
        if not name or not batch:
            continue
        rows.append({
            "code": str(sheet.cell_value(r, 1)).strip(),
            "name": name,
            "batch": batch,
            "mfg": str(sheet.cell_value(r, 7)).strip(),
            "exp": str(sheet.cell_value(r, 10)).strip(),
            "doc_date": str(sheet.cell_value(r, 11)).strip(),
            "biz": str(sheet.cell_value(r, 13)).strip(),
            "warehouse": str(sheet.cell_value(r, 14)).strip(),
            "qty": _num(sheet.cell_value(r, 16)),
        })
    return rows


def _date_short(value):
    text = str(value or "").strip()
    digits = re.sub(r"\D", "", text)
    if len(digits) >= 8:
        return "%s-%s-%s" % (digits[:4], digits[4:6], digits[6:8])
    return text


def reconcile(stock_items, balance_rows, batch_rows, min_diff=0.01):
    """纯核对（不改数据）。返回 dict：diffs / new_in_lemon / missing_in_lemon /
    unit_conflicts / batch_facts。diffs 元素：{code, name, sys_qty, lemon_qty, diff, exact}"""
    by_code = {}
    by_name = {}
    for item in stock_items:
        if item.get("code"):
            by_code[str(item["code"]).strip()] = item
        key = normalize_name(item.get("short") or item.get("name"))
        if key:
            by_name.setdefault(key, item)
    seen_codes = set()
    diffs, new_in_lemon, unit_conflicts, batch_facts = [], [], [], {}
    for info in balance_rows:
        seen_codes.add(info["code"])
        item = by_code.get(info["code"])
        exact = True
        if item is None:
            item = by_name.get(normalize_name(info["name"]))
            exact = False
        if item is None:
            new_in_lemon.append(info)
            continue
        if norm_unit(info["unit"]) != norm_unit(item.get("unit")):
            unit_conflicts.append({"code": info["code"], "name": info["name"],
                                   "lemon_unit": info["unit"],
                                   "sys_unit": item.get("unit")})
            continue
        sys_qty = float(item.get("qty") or 0)
        diff = round(info["qty"] - sys_qty, 3)
        if abs(diff) > max(min_diff, abs(info["qty"]) * 0.001):
            diffs.append({"code": info["code"], "name": info["name"],
                          "short": item.get("short") or item.get("name"),
                          "sys_qty": sys_qty, "lemon_qty": info["qty"],
                          "diff": diff, "exact": exact})
    missing_in_lemon = [item for code, item in by_code.items()
                        if code not in seen_codes]
    for row in batch_rows:
        # 同一行既按编码又按名称索引，便于修正时两种键都能查到
        batch_facts.setdefault(row.get("code") or row["name"], []).append(row)
        batch_facts.setdefault(row["name"], []).append(row)
    return {"diffs": diffs, "new_in_lemon": new_in_lemon,
            "missing_in_lemon": missing_in_lemon,
            "unit_conflicts": unit_conflicts, "batch_facts": batch_facts}


def _apply_fixes(stock_data, report, cfg, balance_rows):
    """安全自动修正：仅编码精确匹配 + 单位一致。返回 (fixes, warnings)。"""
    fixes, warnings = [], []
    items = stock_data.get("items", [])
    by_code = {str(i.get("code") or "").strip(): i for i in items}
    name_to_code = {info["name"]: info["code"] for info in balance_rows}
    batch_by_code = {}
    for group_key, rows in report["batch_facts"].items():
        code_key = name_to_code.get(group_key, "")
        for row in rows:
            for key in {row.get("code") or "", row.get("name") or "",
                        group_key, code_key}:
                if key:
                    batch_by_code.setdefault(key, []).append(row)
    if not cfg.get("auto_fix"):
        return fixes, warnings
    for diff in report["diffs"]:
        if not diff["exact"]:
            warnings.append("「%s」账面差 %+.3f，但系统编码未精确匹配，未自动修改，请核对"
                            % (diff["short"], diff["diff"]))
            continue
        item = by_code.get(diff["code"])
        if item is None:
            continue
        old_qty = float(item.get("qty") or 0)
        item["qty"] = diff["lemon_qty"]
        rows = batch_by_code.get(diff["code"]) or \
            batch_by_code.get(name_to_code.get(diff["name"], ""), [])
        codes = sorted({row["batch"] for row in rows if row.get("batch")})
        if not item.get("batches") and len(codes) == 1:
            row0 = rows[0]
            item["batches"] = [{
                "batch": codes[0], "qty": diff["lemon_qty"],
                "mfg": _date_short(row0.get("mfg")),
                "exp": _date_short(row0.get("exp")),
                "in_date": _date_short(row0.get("mfg")),
            }]
            fixes.append("「%s」账面 %g→%g，并补批次 %s"
                         % (diff["short"], old_qty, diff["lemon_qty"], codes[0]))
            continue
        fixes.append("「%s」账面 %g→%g" % (diff["short"], old_qty,
                                          diff["lemon_qty"]))
        batch_sum = round(sum(float(b.get("qty") or 0)
                              for b in item.get("batches") or []), 3)
        if item.get("batches") and abs(
                batch_sum - diff["lemon_qty"]) > max(0.01, abs(diff["lemon_qty"]) * 0.001):
            warnings.append("「%s」批次明细合计 %.3f 与新账面 %.3f 不一致，请核对批次数量"
                            % (diff["short"], batch_sum, diff["lemon_qty"]))
        elif not item.get("batches") and len(codes) > 1:
            warnings.append("「%s」柠檬云有 %d 个批次（%s），未自动拆分批次数量，请手工核对"
                            % (diff["short"], len(codes), "、".join(codes)))
    return fixes, warnings


def _write_history(history_csv, fixes):
    if not fixes:
        return
    is_new = not os.path.exists(history_csv)
    with open(history_csv, "a", encoding="utf-8-sig", newline="") as fp:
        writer = csv.writer(fp)
        if is_new:
            writer.writerow(["时间", "修正内容"])
        stamp = datetime.datetime.now().isoformat(timespec="seconds")
        for fix in fixes:
            writer.writerow([stamp, fix])


def forecast(stock_items, orders, horizon_days=14, recent_days=60):
    """采购预测：历史订单消耗速率 + 待生产/生产中占用 → 建议采购量。"""
    now = datetime.datetime.now()
    consumed = {}
    dates = []
    committed = {}
    for order in orders:
        status = str(order.get("status") or "")
        run_materials = []
        for run in order.get("production_runs") or []:
            run_materials.extend(run.get("materials") or [])
        if run_materials and order.get("order_date"):
            try:
                order_dt = datetime.datetime.strptime(
                    str(order["order_date"])[:8], "%Y%m%d")
            except ValueError:
                order_dt = None
            if order_dt and (now - order_dt).days <= recent_days:
                dates.append(order_dt)
                for m in run_materials:
                    name = normalize_name(m.get("name") or "")
                    if name:
                        consumed[name] = consumed.get(name, 0.0) + \
                            float(m.get("kg") or 0)
        if status in ("待生产", "生产中") and order.get("materials"):
            for m in order["materials"]:
                name = normalize_name(m.get("name") or "")
                if name:
                    committed[name] = committed.get(name, 0.0) + \
                        float(m.get("kg") or 0)
    if dates:
        span = max((max(dates) - min(dates)).days, 14)
    else:
        span = recent_days
    rows = []
    for item in stock_items:
        if str(item.get("cat") or "").startswith("产成品"):
            continue
        key = normalize_name(item.get("short") or item.get("name"))
        qty = float(item.get("qty") or 0)
        need = committed.get(key, 0.0)
        daily = round(consumed.get(key, 0.0) / span, 4)
        cover = round(qty / daily, 1) if daily > 0 else None
        suggest = max(0.0, round(need + daily * horizon_days - qty, 2))
        if suggest <= 0 and need <= 0:
            continue
        line = float(item.get("safe") or item.get("line") or 0)
        urgent = need > qty or (cover is not None and cover < 7)
        rows.append({"name": item.get("short") or item.get("name"),
                     "unit": item.get("unit"), "qty": qty, "safe": line,
                     "committed": round(need, 3), "daily": daily,
                     "cover_days": cover, "suggest": suggest,
                     "urgent": urgent})
    rows.sort(key=lambda r: (not r["urgent"], -r["suggest"]))
    return {"horizon_days": horizon_days, "recent_days": recent_days,
            "span_days": span, "rows": rows}


def run(app_dir=None, mode="auto"):
    """主入口：读表 → 核对 → 修正 → 报告。mode: auto / report（只核对不改）。"""
    app_dir = app_dir or APP
    lemon_dir = os.path.join(app_dir, "data", "lemon")
    os.makedirs(lemon_dir, exist_ok=True)
    history_csv = os.path.join(lemon_dir, "修正历史.csv")
    cfg = load_config()
    balance_path, batch_path = _source_files(cfg)
    missing = [p for p in (balance_path, batch_path) if not os.path.exists(p)]
    if missing:
        return {"ok": False, "msg": "找不到柠檬云导出文件：%s（请先从柠檬云导出到 %s）"
                % ("、".join(os.path.basename(p) for p in missing),
                   cfg["downloads_dir"])}
    try:
        balance_rows = read_balance(balance_path)
        batch_rows = read_batches(batch_path)
    except Exception as exc:
        return {"ok": False, "msg": "读取柠檬云导出失败：%s: %s"
                % (type(exc).__name__, exc)}
    stock_path = os.path.join(app_dir, "stock.json")
    with open(stock_path, encoding="utf-8") as fp:
        stock_data = json.load(fp)
    items = stock_data.get("items", [])
    report = reconcile(items, balance_rows, batch_rows,
                       min_diff=float(cfg.get("min_diff") or 0.01))
    fixes, warnings = [], []
    if mode != "report":
        backup = os.path.join(lemon_dir, "stock_backup_%s.json"
                              % datetime.datetime.now().strftime("%Y%m%d_%H%M"))
        shutil.copy2(stock_path, backup)
        backups = sorted(f for f in os.listdir(lemon_dir)
                         if f.startswith("stock_backup_"))
        for old in backups[:-30]:
            os.remove(os.path.join(lemon_dir, old))
        fixes, warnings = _apply_fixes(stock_data, report, cfg, balance_rows)
        if fixes:
            json_store.save_json_atomic(stock_path, stock_data, indent=1)
            _write_history(history_csv, fixes)
    stamp = datetime.datetime.now().strftime("%Y%m%d_%H%M")
    report_path = os.path.join(lemon_dir, "核对_%s.md" % stamp)
    lines = ["# 柠檬云库存核对 %s" % datetime.datetime.now()
             .strftime("%Y-%m-%d %H:%M"),
             "",
             "- 数据来源：%s、%s" % (os.path.basename(balance_path),
                                    os.path.basename(batch_path)),
             "- 模式：%s" % ("自动修正" if mode != "report" else "仅核对不修改"),
             "",
             "## 一、总量差异（%d 项）" % len(report["diffs"])]
    for diff in report["diffs"]:
        action = ""
        if fixes and diff["exact"]:
            action = "（已自动改为柠檬云值）"
        elif not diff["exact"]:
            action = "（匹配不精确，未自动修改）"
        lines.append("- %s「%s」：系统 %.3f → 柠檬云 %.3f（差 %+.3f）%s"
                     % (diff["code"], diff["name"], diff["sys_qty"],
                        diff["lemon_qty"], diff["diff"], action))
    if not report["diffs"]:
        lines.append("- 全部一致 ✓")
    lines.append("")
    lines.append("## 二、批次补号与提醒")
    if warnings:
        lines.extend("- " + w for w in warnings)
    else:
        lines.append("- 无")
    if fixes:
        lines.append("")
        lines.append("## 三、本次自动修正（%d 项）" % len(fixes))
        lines.extend("- " + fix for fix in fixes)
    if report["new_in_lemon"]:
        lines.append("")
        lines.append("## 四、柠檬云有而系统没有（%d 项，建议新增到系统库存）"
                     % len(report["new_in_lemon"]))
        lines.extend("- %s %s（%s %g）" % (i["code"], i["name"], i["unit"],
                                          i["qty"]) for i in report["new_in_lemon"])
    if report["missing_in_lemon"]:
        lines.append("")
        lines.append("## 五、系统有而柠檬云没有（%d 项，可能已删档）"
                     % len(report["missing_in_lemon"]))
        lines.extend("- %s（账面 %g）" % (i.get("short") or i.get("name"),
                                        float(i.get("qty") or 0))
                     for i in report["missing_in_lemon"])
    with open(report_path, "w", encoding="utf-8") as fp:
        fp.write("\n".join(lines) + "\n")
    state = {"at": datetime.datetime.now().isoformat(timespec="seconds"),
             "mode": mode, "report": report_path,
             "balance_file": balance_path, "batch_file": batch_path,
             "checked": len(balance_rows), "diffs": len(report["diffs"]),
             "fixes": fixes, "warnings": warnings,
             "new_in_lemon": len(report["new_in_lemon"]),
             "missing_in_lemon": len(report["missing_in_lemon"]),
             "unit_conflicts": len(report["unit_conflicts"])}
    json_store.save_json_atomic(os.path.join(lemon_dir, "last_sync.json"),
                                state, indent=1)
    return {"ok": True, "report": report_path, "checked": len(balance_rows),
            "diffs": len(report["diffs"]), "fixes": fixes,
            "warnings": warnings, "new_in_lemon": len(report["new_in_lemon"]),
            "missing_in_lemon": len(report["missing_in_lemon"])}


if __name__ == "__main__":
    import sys
    result = run(mode=sys.argv[1] if len(sys.argv) > 1 else "auto")
    print(json.dumps(result, ensure_ascii=False, indent=1))
