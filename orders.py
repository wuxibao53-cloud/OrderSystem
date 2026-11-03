# -*- coding: utf-8 -*-
"""订单台账：对接销售与生产的交期跟踪"""
import os
import sys
import json
import math
import re
import ctypes
import datetime
import threading
from json_store import save_json_atomic

APP = os.path.dirname(os.path.abspath(__file__))   # 便携：系统放哪个目录都能跑
ORDERS = os.path.join(APP, "orders.json")
RECIPE = os.path.join(APP, "recipes.json")
SALES = ["曾帅", "付豪", "房浩绅", "马士龙"]
STATUS = ["待生产", "生产中", "已完成", "已发货"]
DONE = ("已完成", "已发货")
VOID = "已作废"
ORDER_TYPES = ("正式生产", "测试单", "实验单")
PRODUCTION_STEPS = ("混料", "压片", "包装")
STEP_PROGRESS = ("未开始", "进行中", "已完成")
QC_RESULTS = ("未检", "合格", "异常")
_ORDER_LOCK = threading.RLock()

try:
    _k = ctypes.windll.kernel32
    _h = _k.GetStdHandle(-11)
    _m = ctypes.c_uint32()
    _k.GetConsoleMode(_h, ctypes.byref(_m))
    _k.SetConsoleMode(_h, _m.value | 0x0004)
    _k.SetConsoleOutputCP(65001)
    _k.SetConsoleCP(65001)
except Exception:
    pass
try:
    sys.stdout.reconfigure(encoding="utf-8")
except Exception:
    pass

C = {
    "b": "\033[1m", "r": "\033[0m", "dim": "\033[2m",
    "red": "\033[91m", "yel": "\033[93m", "grn": "\033[92m",
    "cyn": "\033[96m", "blu": "\033[94m", "gry": "\033[90m",
}


def clear():
    os.system("cls" if os.name == "nt" else "clear")


def bar(title, width=66):
    print("%s╔%s╗%s" % (C["blu"], "═" * width, C["r"]))
    pad = width - len(title) * 2
    left = max(0, pad // 2)
    print("%s║%s%s%s║%s" % (C["blu"], " " * left, C["b"] + title + C["r"],
                            " " * max(0, pad - left), C["r"]))
    print("%s╚%s╝%s" % (C["blu"], "═" * width, C["r"]))


def load_json(path, default=None):
    if not os.path.exists(path):
        return default if default is not None else {}
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)


def load():
    if not os.path.exists(ORDERS):
        return {"orders": []}
    with open(ORDERS, encoding="utf-8") as stream:
        data = json.load(stream)
    if isinstance(data, list):
        return {"orders": data}
    if not isinstance(data, dict):
        raise ValueError("订单台账必须是对象或订单列表")
    orders = data.setdefault("orders", [])
    if not isinstance(orders, list):
        raise ValueError("订单台账的 orders 必须是列表")
    return data


def save(d):
    save_json_atomic(ORDERS, d)


def due_days(o):
    due = (o.get("due") or "").strip()
    if not due:
        return None
    try:
        d = datetime.datetime.strptime(due.replace("-", "").replace("/", ""), "%Y%m%d").date()
        return (d - datetime.date.today()).days
    except Exception:
        return None


def urgency(o):
    if o.get("status") in DONE or o.get("status") == VOID or o.get("archived"):
        return ("完成", None)
    dd = due_days(o)
    if dd is None:
        return ("无交期", None)
    if dd < 0:
        return ("逾期", dd)
    if dd <= 3:
        return ("临期", dd)
    return ("", dd)


def add_order(o):
    with _ORDER_LOCK:
        d = load()
        for x in d["orders"]:
            if x.get("serial") == o.get("serial") and x.get("product") == o.get("product"):
                return False, "台账中已有 %s" % o.get("serial")
        order_type = o.get("order_type", "正式生产")
        if order_type not in ORDER_TYPES:
            return False, "订单分类无效"
        o["order_type"] = order_type
        o.setdefault("status", "待生产")
        o.setdefault("customer", "")
        o.setdefault("note", "")
        o.setdefault("tracking_created_at", datetime.datetime.now().isoformat(timespec="seconds"))
        d["orders"].append(o)
        save(d)
        return True, "已加入台账：%s %s" % (o.get("serial"), o.get("product"))


def _tracking_age_days(order, now=None):
    created = order.get("tracking_created_at")
    if not created:
        return None
    try:
        created_at = datetime.datetime.fromisoformat(str(created))
    except (TypeError, ValueError):
        return None
    now = now or datetime.datetime.now()
    if created_at.tzinfo is not None and now.tzinfo is None:
        now = now.astimezone(created_at.tzinfo)
    elif created_at.tzinfo is None and now.tzinfo is not None:
        now = now.replace(tzinfo=None)
    return max(0, (now - created_at).days)


def confirm_tracking_order(serial, product):
    with _ORDER_LOCK:
        data = load()
        matches = [order for order in data["orders"]
                   if order.get("serial") == serial and order.get("product") == product]
        if len(matches) != 1:
            raise ValueError("订单不存在或单号匹配不唯一，未确认")
        order = matches[0]
        if order.get("tracking_confirmed_at"):
            return {"ok": True, "changed": False,
                    "confirmed_at": order["tracking_confirmed_at"]}
        now = datetime.datetime.now().isoformat(timespec="seconds")
        order["tracking_confirmed_at"] = now
        order["tracking_confirmation_source"] = "手动确认"
        history = order.setdefault("tracking_history", [])
        if not isinstance(history, list):
            raise ValueError("跟踪确认历史格式无效")
        history.append({"at": now, "source": "手动确认"})
        save(data)
        return {"ok": True, "changed": True, "confirmed_at": now}


def import_production_sheet(order, source_name, content_sha256=None):
    """Create or confirm a tracking order from a parsed canonical production sheet."""
    if not isinstance(order, dict):
        raise ValueError("生产表信息格式无效")
    serial = str(order.get("serial") or "").strip()
    product = str(order.get("product") or "").strip()
    source_name = os.path.basename(str(source_name or "").replace("\\", "/"))
    if not serial or not product or not source_name or len(source_name) > 200:
        raise ValueError("生产单号、产品名称或文件名无效")
    content_sha256 = str(content_sha256 or "").lower()
    if content_sha256 and not re.fullmatch(r"[0-9a-f]{64}", content_sha256):
        raise ValueError("生产表文件校验值无效")
    order_type = order.get("order_type", "")
    if order_type not in ORDER_TYPES:
        raise ValueError("请为导入的生产表选择有效分类")
    materials = order.get("materials")
    if not isinstance(materials, list) or not materials:
        raise ValueError("生产表未识别到有效物料及用量，未导入")
    for material in materials:
        if (not isinstance(material, dict) or not str(material.get("name") or "").strip()
                or not math.isfinite(float(material.get("kg", 0) or 0))
                or float(material.get("kg", 0) or 0) <= 0):
            raise ValueError("生产表含无效物料或用量，未导入")

    with _ORDER_LOCK:
        data = load()
        matches = [existing for existing in data["orders"]
                   if existing.get("serial") == serial]
        if len(matches) > 1:
            raise ValueError("台账中该生产单号不唯一，未导入")
        now = datetime.datetime.now().isoformat(timespec="seconds")
        if matches:
            existing = matches[0]
            if existing.get("product") != product:
                raise ValueError("台账中已有同单号但产品名称不同的记录，未修改")
            source = existing.get("tracking_confirmation_source")
            warnings = []
            existing_total = existing.get("total_kg")
            imported_total = order.get("total_kg")
            if existing_total is not None and imported_total is not None:
                try:
                    quantity_differs = abs(float(existing_total) - float(imported_total)) > 0.001
                except (TypeError, ValueError):
                    quantity_differs = True
                if quantity_differs:
                    warnings.append("生产表计划重量与台账不同；台账未覆盖")
            def material_totals(records):
                totals = {}
                if not isinstance(records, list):
                    return None
                try:
                    for material in records:
                        if not isinstance(material, dict) or not material.get("name"):
                            return None
                        name = str(material["name"])
                        totals[name] = totals.get(name, 0.0) + float(material.get("kg", 0) or 0)
                    return {name: round(quantity, 3) for name, quantity in totals.items()}
                except (TypeError, ValueError):
                    return None
            existing_totals = material_totals(existing.get("materials", []))
            imported_totals = material_totals(materials)
            if existing_totals is not None and existing_totals != imported_totals:
                warnings.append("生产表物料或用量与台账不同；台账预占未覆盖")
            imports = existing.setdefault("production_sheet_imports", [])
            if not isinstance(imports, list):
                raise ValueError("生产表导入历史格式无效，未修改台账")
            repeated = bool(content_sha256) and any(
                isinstance(item, dict) and item.get("sha256") == content_sha256
                for item in imports)
            changed = False
            if not repeated:
                imports.append({"at": now, "file": source_name,
                                "sha256": content_sha256})
                existing["production_sheet_name"] = source_name
                existing["production_sheet_imported_at"] = now
                existing["production_sheet_sha256"] = content_sha256
                changed = True
            if not existing.get("tracking_confirmed_at"):
                existing["tracking_confirmed_at"] = now
                existing["tracking_confirmation_source"] = "生产表导入"
                history = existing.setdefault("tracking_history", [])
                if not isinstance(history, list):
                    raise ValueError("跟踪确认历史格式无效，未修改台账")
                history.append({"at": now, "source": "生产表导入",
                                "file": source_name})
                changed = True
            if changed:
                save(data)
            return {
                "ok": True, "created": False, "confirmed": True,
                "changed": changed, "repeated_import": repeated,
                "order": dict(existing),
                "message": "已匹配现有台账；未覆盖订单数量、物料预占、生产状态或交接记录",
                "already_confirmed": bool(source),
                "warnings": warnings,
            }

        imported = {
            "serial": serial,
            "product": product,
            "wan": order.get("wan"),
            "tablet": order.get("tablet"),
            "total_kg": order.get("total_kg"),
            "customer": str(order.get("customer") or ""),
            "due": str(order.get("due") or ""),
            "sales": str(order.get("sales") or ""),
            "order_no": str(order.get("order_no") or ""),
            "order_date": str(order.get("order_date") or ""),
            "batch": str(order.get("batch") or "01"),
            "materials": [dict(material) for material in materials],
            "packaging": [],
            "order_type": order_type,
            "status": "待生产",
            "manual_tracking": False,
            "production_sheet_import": True,
            "production_sheet_name": source_name,
            "production_sheet_imported_at": now,
            "production_sheet_sha256": content_sha256,
            "production_sheet_imports": [{
                "at": now, "file": source_name, "sha256": content_sha256,
            }],
            "tracking_created_at": now,
            "tracking_confirmed_at": now,
            "tracking_confirmation_source": "生产表导入",
            "tracking_history": [{"at": now, "source": "生产表导入",
                                 "file": source_name}],
        }
        data["orders"].append(imported)
        save(data)
        return {"ok": True, "created": True, "confirmed": True, "order": dict(imported),
                "message": "已从生产表新建跟踪记录"}


def _pending_materials(runs):
    """未完成批次的物料合计（台账预占 = 这些批次的物料）。"""
    totals = {}
    for run in runs:
        if not isinstance(run, dict) or run.get("status") == "已完成":
            continue
        for material in run.get("materials") or []:
            name = str((material or {}).get("name") or "").strip()
            batch = str((material or {}).get("batch") or "").strip()
            if not name:
                continue
            key = (name, batch)
            entry = totals.setdefault(key, {"name": name, "batch": batch, "kg": 0.0})
            try:
                entry["kg"] += float((material or {}).get("kg") or 0)
            except (TypeError, ValueError):
                continue
    return [{"name": key[0], "batch": key[1], "kg": round(value["kg"], 3)}
            for key, value in totals.items()]


def mark_production_run(serial, run_no, completed=True):
    """核销/撤销某一生产批次：核销后该批物料预占释放；全部核销 → 订单已完成。"""
    serial = str(serial or "").strip()
    try:
        run_no = int(run_no)
    except (TypeError, ValueError):
        raise ValueError("生产批次号无效")
    with _ORDER_LOCK:
        data = load()
        matches = [order for order in data["orders"]
                   if order.get("serial") == serial]
        if len(matches) != 1:
            raise ValueError("订单不存在或单号匹配不唯一")
        order = matches[0]
        if order.get("status") == VOID or order.get("archived"):
            raise ValueError("已作废/已归档订单不能核销批次")
        runs = order.get("production_runs")
        if not isinstance(runs, list) or not runs:
            raise ValueError("该订单没有生产批次记录；请先通过销售单导入或生产表导入登记批次")
        target = None
        for run in runs:
            if int(run.get("run_no") or 0) == run_no:
                target = run
                break
        if target is None:
            raise ValueError("没有第 %d 批的记录" % run_no)
        now = datetime.datetime.now().isoformat(timespec="seconds")
        if completed:
            if target.get("status") == "已完成":
                return {"ok": True, "changed": False, "run_no": run_no,
                        "status": order.get("status", ""), "runs": runs}
            target["status"] = "已完成"
            target["completed_at"] = now
            note = "批次核销"
        else:
            if target.get("status") != "已完成":
                return {"ok": True, "changed": False, "run_no": run_no,
                        "status": order.get("status", ""), "runs": runs}
            target["status"] = ""
            target.pop("completed_at", None)
            if not target["status"]:
                target.pop("status", None)
            note = "批次撤销完成"
        pending = [run for run in runs if run.get("status") != "已完成"]
        order["production_runs"] = runs
        order["materials"] = _pending_materials(runs)
        history = order.setdefault("tracking_history", [])
        if not isinstance(history, list):
            raise ValueError("跟踪确认历史格式无效，未修改台账")
        history.append({"at": now, "source": note, "run_no": run_no})
        current = order.get("status", "待生产")
        if completed and not pending and current in ("", "待生产", "生产中"):
            order["status"] = "已完成"
            order.setdefault("status_history", []).append(
                {"at": now, "from": current, "to": "已完成",
                 "reason": "全部生产批次已核销"})
        elif not pending:
            pass
        elif completed and current == "待生产":
            order["status"] = "生产中"
            order.setdefault("status_history", []).append(
                {"at": now, "from": current, "to": "生产中",
                 "reason": "第 %d 批核销，仍有未完成批次" % run_no})
        elif not completed and current == "已完成":
            order["status"] = "生产中"
            order.setdefault("status_history", []).append(
                {"at": now, "from": current, "to": "生产中",
                 "reason": "撤销第 %d 批核销" % run_no})
        save(data)
        return {"ok": True, "changed": True, "run_no": run_no,
                "status": order.get("status", ""), "runs": runs,
                "materials": order["materials"],
                "message": "%s第 %d 批已%s" % (serial, run_no,
                                               "核销完成" if completed else "恢复待生产")}


def attach_sales_import(serial, product, total_kg, materials, source_name,
                        content_sha256=None, manager="", run_date=""):
    """Update a pending tracking order from an imported sales form.

    The batch keeps its serial, type and status. Every import is recorded as
    one production run (生产批次): same batch produced on several days simply
    adds runs — run 1, run 2, ... each with its own plan, recipe and sheet.
    Pending material reservations always reflect the latest run.
    """
    serial = str(serial or "").strip()
    product = str(product or "").strip()
    source_name = os.path.basename(str(source_name or "").replace("\\", "/"))
    if not serial or not product or not source_name or len(source_name) > 200:
        raise ValueError("生产单号、产品名称或文件名无效")
    content_sha256 = str(content_sha256 or "").lower()
    if content_sha256 and not re.fullmatch(r"[0-9a-f]{64}", content_sha256):
        raise ValueError("文件校验值无效")
    if not isinstance(materials, list) or not materials:
        raise ValueError("销售单未识别到有效物料及用量")
    clean_materials = []
    for material in materials:
        name = str((material or {}).get("name") or "").strip()
        kg = float((material or {}).get("kg") or 0)
        if not name or not math.isfinite(kg) or kg <= 0:
            raise ValueError("销售单含无效物料或用量")
        clean_materials.append({"name": name, "kg": round(kg, 3),
                                "batch": str((material or {}).get("batch") or "").strip()})
    with _ORDER_LOCK:
        data = load()
        matches = [existing for existing in data["orders"]
                   if existing.get("serial") == serial]
        if not matches:
            raise ValueError("台账中没有 %s 的记录" % serial)
        existing = matches[0]
        if existing.get("status") not in ("", "待生产"):
            raise ValueError("台账中 %s 状态为「%s」，仅待生产订单可按销售单更新"
                             % (serial, existing.get("status")))
        now = datetime.datetime.now().isoformat(timespec="seconds")
        warnings = []
        old_product = existing.get("product")
        if old_product != product:
            warnings.append("台账产品名称由「%s」更新为「%s」" % (old_product, product))
        runs = existing.get("production_runs")
        if not isinstance(runs, list):
            runs = []
            prior_plan = existing.get("total_kg")
            prior_materials = existing.get("materials")
            if prior_plan or existing.get("production_sheet_imports"):
                runs.append({
                    "run_no": 1,
                    "run_date": str(existing.get("order_date") or ""),
                    "plan_kg": prior_plan,
                    "file": existing.get("production_sheet_name") or "",
                    "sha256": existing.get("production_sheet_sha256") or "",
                    "materials": [dict(item) for item in (prior_materials or [])]
                    if isinstance(prior_materials, list) else [],
                    "recorded_at": existing.get("production_sheet_imported_at") or now,
                    "note": "历史导入补记",
                })
        run_no = len(runs) + 1
        runs.append({
            "run_no": run_no,
            "run_date": str(run_date or ""),
            "plan_kg": round(float(total_kg), 3),
            "file": source_name,
            "sha256": content_sha256,
            "materials": clean_materials,
            "recorded_at": now,
            "note": "销售单导入",
        })
        existing["production_runs"] = runs
        existing["product"] = product
        try:
            existing["total_kg"] = round(sum(
                float(run.get("plan_kg") or 0) for run in runs), 3)
        except (TypeError, ValueError):
            raise ValueError("计划产量无效")
        existing["materials"] = _pending_materials(runs)
        existing["batch"] = "%02d" % run_no
        if manager:
            existing["sales"] = str(manager)
        existing["production_sheet_name"] = source_name
        existing["production_sheet_imported_at"] = now
        existing["production_sheet_sha256"] = content_sha256
        imports = existing.setdefault("production_sheet_imports", [])
        if not isinstance(imports, list):
            imports = existing["production_sheet_imports"] = []
        if not (content_sha256 and any(
                isinstance(item, dict) and item.get("sha256") == content_sha256
                for item in imports)):
            imports.append({"at": now, "file": source_name,
                            "sha256": content_sha256, "source": "销售单导入"})
        history = existing.setdefault("tracking_history", [])
        if not isinstance(history, list):
            raise ValueError("跟踪确认历史格式无效，未修改台账")
        history.append({"at": now, "source": "销售单导入", "file": source_name,
                        "run_no": run_no})
        if not existing.get("tracking_confirmed_at"):
            existing["tracking_confirmed_at"] = now
            existing["tracking_confirmation_source"] = "销售单导入"
        save(data)
        return {"ok": True, "order": dict(existing), "run_no": run_no,
                "message": "已按新销售单更新批次 %s（第 %d 批）：计划产量、物料与批号预占"
                           % (serial, run_no),
                "warnings": warnings}


def update_order_info(serial, product, new_product=None, total_kg=None, note=""):
    """Correct a ledger record's product name and/or planned kg.

    Material reservations scale proportionally with the planned weight; the
    change is logged in edit_history. Generated Excel files are not rewritten.
    """
    new_product = str(new_product).strip() if new_product is not None else None
    note = str(note or "").strip()
    if len(note) > 500:
        raise ValueError("修改说明最多 500 字")
    if new_product is not None and (not new_product or len(new_product) > 100):
        raise ValueError("产品名称不能为空且最多 100 字")
    new_kg = None
    if total_kg not in (None, ""):
        try:
            new_kg = round(float(total_kg), 3)
        except (TypeError, ValueError):
            raise ValueError("计划产量必须是数字")
        if not math.isfinite(new_kg) or new_kg <= 0 or new_kg > 100000:
            raise ValueError("计划产量必须大于 0 且不超过 100000 kg")
    with _ORDER_LOCK:
        data = load()
        matches = [order for order in data["orders"]
                   if order.get("serial") == serial and order.get("product") == product]
        if len(matches) != 1:
            raise ValueError("订单不存在或单号匹配不唯一，未修改")
        order = matches[0]
        if (order.get("status") in DONE or order.get("status") == VOID
                or order.get("archived") or order.get("status") == "已发货"):
            raise ValueError("已完成、已发货、已作废或已归档的订单不能修改名称和数量")
        history = order.setdefault("edit_history", [])
        if not isinstance(history, list):
            raise ValueError("订单修改历史格式无效，未修改")
        changes = {}
        if new_product is not None and new_product != order.get("product"):
            if any(other is not order and other.get("serial") == serial
                   and other.get("product") == new_product for other in data["orders"]):
                raise ValueError("台账中已有同单号同名称的记录，未修改")
            changes["product"] = {"from": order.get("product"), "to": new_product}
        warnings = []
        if new_kg is not None:
            try:
                old_kg = float(order.get("total_kg"))
            except (TypeError, ValueError):
                old_kg = None
            if old_kg is None or old_kg <= 0:
                raise ValueError("该订单原计划重量未知，无法按比例调整物料预占，请作废后重建")
            if abs(old_kg - new_kg) > 0.0005:
                ratio = new_kg / old_kg
                scaled = []
                for material in order.get("materials") or []:
                    if not isinstance(material, dict):
                        raise ValueError("订单物料记录格式无效，未修改")
                    updated = dict(material)
                    if "kg" in updated:
                        updated["kg"] = round(float(updated.get("kg") or 0) * ratio, 3)
                    scaled.append(updated)
                if order.get("packaging"):
                    warnings.append("包材数量未自动调整，请人工核对")
                order["materials"] = scaled
                order["total_kg"] = new_kg
                try:
                    tablet = float(order.get("tablet"))
                    if tablet > 0:
                        order["wan"] = round(new_kg * 1000 / tablet / 10000, 4)
                except (TypeError, ValueError):
                    pass
                changes["total_kg"] = {"from": old_kg, "to": new_kg}
        if not changes:
            return {"ok": True, "changed": False, "order": dict(order), "warnings": []}
        if "product" in changes:
            order["product"] = new_product
        entry = {"at": datetime.datetime.now().isoformat(timespec="seconds"),
                 "changes": changes}
        if note:
            entry["note"] = note
        history.append(entry)
        save(data)
        return {"ok": True, "changed": True, "order": dict(order),
                "warnings": warnings}


def order_exists(serial, product):
    with _ORDER_LOCK:
        return any(
            order.get("serial") == serial and order.get("product") == product
            for order in load().get("orders", [])
        )


def _can_delete_order(order):
    return (
        _effective_order_type(order) == "正式生产"
        and order.get("status", "待生产") == "待生产"
        and not order.get("production_steps")
        and not order.get("production_history")
        and not order.get("status_history")
        and not order.get("archived")
    )


def _effective_order_type(order):
    order_type = order.get("order_type")
    if order_type in ORDER_TYPES:
        return order_type
    legacy_reason = " ".join(
        str(order.get(field) or "")
        for field in ("archive_reason", "note")
    )
    if "实验" in legacy_reason:
        return "实验单"
    if "测试" in legacy_reason:
        return "测试单"
    return "正式生产"


def _can_delete_test_order(order):
    if _effective_order_type(order) not in ("测试单", "实验单"):
        return False
    if order.get("archived") or order.get("status") in DONE or order.get("status") == "已发货":
        return False
    if order.get("production_steps") or order.get("production_history"):
        return False
    history = order.get("status_history") or []
    if order.get("status") == VOID:
        return bool(history) and all(
            isinstance(event, dict) and event.get("to") == VOID
            for event in history
        )
    return order.get("status", "待生产") == "待生产" and not history


def production_dashboard(include_closed=False):
    with _ORDER_LOCK:
        orders = load().get("orders", [])
    result = []
    for order in orders:
        if not include_closed and (
                order.get("status") in DONE or order.get("status") == VOID
                or order.get("archived")):
            continue
        item = dict(order)
        item["order_type"] = _effective_order_type(order)
        item["can_delete"] = _can_delete_order(order)
        item["can_delete_test"] = _can_delete_test_order(order)
        item["production_steps"] = dict(order.get("production_steps") or {})
        age_days = _tracking_age_days(order)
        item["tracking_age_days"] = age_days
        item["tracking_confirmed"] = bool(order.get("tracking_confirmed_at"))
        item["tracking_stale"] = bool(
            age_days is not None and age_days >= 3
            and not order.get("tracking_confirmed_at")
            and order.get("status", "待生产") == "待生产"
            and not order.get("archived")
        )
        result.append(item)
    result.sort(key=lambda item: (
        due_days(item) is None,
        due_days(item) if due_days(item) is not None else 9999,
        item.get("serial", ""),
    ))
    return result


def update_order_status(serial, product, status):
    if status not in STATUS:
        raise ValueError("订单状态无效")
    with _ORDER_LOCK:
        data = load()
        matches = [order for order in data["orders"]
                   if order.get("serial") == serial and order.get("product") == product]
        if len(matches) != 1:
            raise ValueError("订单不存在或单号匹配不唯一，未更新")
        order = matches[0]
        current = order.get("status", "待生产")
        if current == VOID or current == "已发货":
            raise ValueError("已作废或已发货的订单不能变更状态")
        if current == status:
            return {"ok": True, "changed": False, "status": status}
        if current == "已完成" and status != "已发货":
            raise ValueError("已完成订单只能继续转为已发货，不能重新打开")
        now = datetime.datetime.now().isoformat(timespec="seconds")
        history = order.setdefault("status_history", [])
        if not isinstance(history, list):
            raise ValueError("订单状态历史记录格式无效")
        history.append({"at": now, "from": current, "to": status})
        order["status"] = status
        save(data)
        return {"ok": True, "changed": True, "status": status}


def archive_or_void_order(serial, product, reason):
    reason = (reason or "").strip()
    if not reason:
        raise ValueError("请填写作废或归档原因")
    if len(reason) > 500:
        raise ValueError("原因最多 500 字")
    with _ORDER_LOCK:
        data = load()
        matches = [order for order in data["orders"]
                   if order.get("serial") == serial and order.get("product") == product]
        if len(matches) != 1:
            raise ValueError("订单不存在或单号匹配不唯一，未更新")
        order = matches[0]
        if order.get("status") == "已发货":
            raise ValueError("已发货订单不能作废或归档")
        if order.get("archived") or order.get("status") == VOID:
            return {"ok": True, "changed": False, "status": order.get("status", "")}
        now = datetime.datetime.now().isoformat(timespec="seconds")
        history = order.setdefault("status_history", [])
        if not isinstance(history, list):
            raise ValueError("订单状态历史记录格式无效")
        current = order.get("status", "待生产")
        if current in DONE:
            order["archived"] = True
            order["archived_at"] = now
            result_status = current
        else:
            order["status"] = VOID
            order["voided_at"] = now
            result_status = VOID
        order["archive_reason"] = reason
        history.append({"at": now, "from": current, "to": result_status,
                        "reason": reason})
        save(data)
        return {"ok": True, "changed": True, "status": result_status}


def delete_unstarted_order(serial, product):
    with _ORDER_LOCK:
        data = load()
        matches = [order for order in data["orders"]
                   if order.get("serial") == serial and order.get("product") == product]
        if len(matches) != 1:
            raise ValueError("订单不存在或单号匹配不唯一，未删除")
        order = matches[0]
        if not _can_delete_order(order):
            raise ValueError("仅可删除未开工且没有进度、质量或状态记录的待生产单；其他记录请作废/归档")
        data["orders"].remove(order)
        save(data)
        return {"ok": True, "deleted": True}


def delete_test_order(serial, product):
    with _ORDER_LOCK:
        data = load()
        matches = [order for order in data["orders"]
                   if order.get("serial") == serial and order.get("product") == product]
        if len(matches) != 1:
            raise ValueError("测试/实验单不存在或单号匹配不唯一，未删除")
        order = matches[0]
        if not _can_delete_test_order(order):
            raise ValueError(
                "仅可彻底删除未开工、无工序或质量记录的测试/实验单；"
                "已开始或已完成的记录必须保留并归档"
            )
        data["orders"].remove(order)
        save(data)
        return {"ok": True, "deleted": True}


def set_order_file(serial, product, file_path, details=None, run_note="出单生成"):
    with _ORDER_LOCK:
        data = load()
        matches = [order for order in data["orders"]
                   if order.get("serial") == serial and order.get("product") == product]
        if len(matches) != 1:
            raise ValueError("订单不存在或单号匹配不唯一，文件信息未写入台账")
        order = matches[0]
        order["file"] = os.path.abspath(file_path)
        if details:
            order.update(details)
        if not isinstance(order.get("production_runs"), list) and (
                details or {}).get("materials"):
            # 出单链路：生成的第一张交接单即第 1 批
            runs = [{
                "run_no": 1,
                "run_date": str(order.get("order_date") or ""),
                "plan_kg": order.get("total_kg"),
                "file": os.path.basename(str(file_path)),
                "sha256": "",
                "materials": [dict(m) for m in (details.get("materials") or [])
                              if isinstance(m, dict) and m.get("name")],
                "recorded_at": datetime.datetime.now().isoformat(timespec="seconds"),
                "note": run_note,
            }]
            order["production_runs"] = runs
            order["materials"] = _pending_materials(runs)
        save(data)
        return {"ok": True}


def update_production_step(serial, product, step, progress=None, qc=None,
                           operator=None, note=None):
    if step not in PRODUCTION_STEPS:
        raise ValueError("未知生产工序")
    if progress is not None and progress not in STEP_PROGRESS:
        raise ValueError("工序状态无效")
    if qc is not None and qc not in QC_RESULTS:
        raise ValueError("质量交接结果无效")
    operator = (operator or "").strip()
    note = (note or "").strip()
    if len(operator) > 40 or len(note) > 500:
        raise ValueError("操作人最多 40 字，备注最多 500 字")
    with _ORDER_LOCK:
        data = load()
        matches = [order for order in data["orders"]
                   if order.get("serial") == serial and order.get("product") == product]
        if len(matches) != 1:
            raise ValueError("订单不存在或单号匹配不唯一，未更新")
        order = matches[0]
        if order.get("status") in DONE or order.get("status") == VOID or order.get("archived"):
            raise ValueError("已完成、已作废或已归档的订单不能继续编辑工序")
        steps = order.setdefault("production_steps", {})
        if not isinstance(steps, dict):
            raise ValueError("订单工序记录格式无效")
        record = steps.setdefault(step, {
            "progress": "未开始", "qc": "未检", "operator": "", "note": "",
        })
        if not isinstance(record, dict):
            raise ValueError("订单工序记录格式无效")
        changes = {}
        for key, value in (("progress", progress), ("qc", qc),
                           ("operator", operator), ("note", note)):
            if value is not None and record.get(key) != value:
                changes[key] = {"from": record.get(key, ""), "to": value}
                record[key] = value
        if not changes:
            return {"ok": True, "changed": False, "step": step, "record": record}
        now = datetime.datetime.now().isoformat(timespec="seconds")
        record["updated_at"] = now
        history = order.setdefault("production_history", [])
        if not isinstance(history, list):
            raise ValueError("订单工序历史记录格式无效")
        history.append({"step": step, "at": now, "changes": changes})
        save(data)
        return {"ok": True, "changed": True, "step": step, "record": record}


def view():
    d = load()
    orders = d.get("orders", [])
    open_o = [o for o in orders if o.get("status") not in DONE
              and o.get("status") != VOID and not o.get("archived")]
    done_o = [o for o in orders if o.get("status") in DONE
              or o.get("status") == VOID or o.get("archived")]
    open_o.sort(key=lambda x: (due_days(x) is None, due_days(x) if due_days(x) is not None else 9999))
    clear()
    bar("订单台账  %s　在跟 %d / 已完结 %d" %
        (datetime.datetime.now().strftime("%Y-%m-%d"), len(open_o), len(done_o)))
    print()
    print("  %s在跟订单（按交期排序）%s" % (C["b"], C["r"]))
    print("  %s%-14s %-14s %-9s %-11s %-8s %-6s%s" %
          (C["dim"], "单号", "产品", "产量", "交期", "距交期", "状态", C["r"]))
    if not open_o:
        print("  %s（空）%s" % (C["gry"], C["r"]))
    for o in open_o:
        tag, dd = urgency(o)
        if tag == "逾期":
            dd_s = "%s逾期%d天%s" % (C["red"], -dd, C["r"])
        elif tag == "临期":
            dd_s = "%s剩%d天%s" % (C["yel"], dd, C["r"])
        elif tag == "无交期":
            dd_s = "%s未填%s" % (C["gry"], C["r"])
        else:
            dd_s = "剩%d天" % dd
        st = o.get("status", "")
        stc = {"待生产": C["yel"], "生产中": C["cyn"]}.get(st, "")
        print("  %-14s %-14s %7s万 %-11s %s %s%s%s" %
              (o.get("serial", ""), (o.get("product") or "")[:13],
               o.get("wan", ""), o.get("due", "-"), dd_s,
               stc, st, C["r"]))
    print()
    print("  %s已完结%s" % (C["b"], C["r"]))
    if not done_o:
        print("  %s（空）%s" % (C["gry"], C["r"]))
    for o in done_o[-8:]:
        print("  %s%-14s %-14s %7s万 %s%s" %
              (C["gry"], o.get("serial", ""), (o.get("product") or "")[:13],
               o.get("wan", ""), o.get("status", ""), C["r"]))
    print()


def add_flow():
    recipes = load_json(RECIPE, {})
    names = [k for k in recipes if not k.startswith("_")]
    clear()
    bar("添加订单")
    print()
    for i, k in enumerate(names, 1):
        print("  %s%2d%s  %s" % (C["dim"], i, C["r"], k))
    print()
    s = input("  产品序号：").strip()
    try:
        product = names[int(s) - 1]
    except Exception:
        print("  %s无效%s" % (C["red"], C["r"]))
        input("  按回车返回")
        return
    r = recipes[product]
    serial = input("  单号（回车 = 今天 + 01）：").strip()
    if not serial:
        serial = "%s-01" % datetime.datetime.now().strftime("%Y%m%d")
    wan = input("  计划产量（万片）：").strip()
    try:
        wan = float(wan)
    except Exception:
        wan = 0.0
    tablet = input("  片重 g/片（回车用 %s）：" % r.get("tablet_g", 0.5)).strip()
    try:
        tablet = float(tablet)
    except Exception:
        tablet = float(r.get("tablet_g", 0.5))
    customer = input("  客户（可空）：").strip()
    due = input("  交货日期 8 位（如 20261105）：").strip()
    print("  销售经理：")
    for i, m in enumerate(SALES, 1):
        print("    %d %s" % (i, m))
    mi = input("  序号：").strip()
    try:
        sales = SALES[int(mi) - 1]
    except Exception:
        sales = ""
    ok, msg = add_order({"serial": serial, "product": product, "wan": wan,
                         "tablet": tablet, "customer": customer, "due": due,
                         "sales": sales, "status": "待生产",
                         "order_date": datetime.datetime.now().strftime("%Y%m%d")})
    print("  %s%s%s" % (C["grn"] if ok else C["red"], msg, C["r"]))
    input("  按回车返回")


def change_flow():
    d = load()
    orders = [o for o in d.get("orders", [])
              if o.get("status") not in DONE and o.get("status") != VOID
              and not o.get("archived")]
    if not orders:
        print("  没有在跟订单")
        input("  按回车返回")
        return
    view()
    print("  %s输入要改动的单号（或单号前几位）%s" % (C["dim"], C["r"]))
    s = input("  > ").strip()
    hit = [o for o in orders if o.get("serial", "").startswith(s)]
    if len(hit) != 1:
        print("  %s匹配到 %d 个%s" % (C["red"], len(hit), C["r"]))
        input("  按回车返回")
        return
    o = hit[0]
    print("  %s（回车跳过）%s" % (C["dim"], C["r"]))
    print("  状态：1待生产 2生产中 3已完成 4已发货　当前：%s" % o.get("status"))
    st = input("  新状态序号：").strip()
    new_status = STATUS[int(st) - 1] if st in ("1", "2", "3", "4") else None
    due = input("  交货日期 8 位（当前 %s）：" % (o.get("due") or "-")).strip()
    note = input("  备注（当前：%s）：" % (o.get("note") or "-")).strip()
    if new_status and new_status != o.get("status", "待生产"):
        try:
            update_order_status(o.get("serial"), o.get("product"), new_status)
            d = load()
            o = next(order for order in d["orders"]
                     if order.get("serial") == o.get("serial")
                     and order.get("product") == o.get("product"))
        except Exception as exc:
            print("  %s状态更新失败：%s%s" % (C["red"], exc, C["r"]))
            input("  按回车返回")
            return
    if due:
        o["due"] = due
    if note:
        o["note"] = note
    save(d)
    print("  %s已更新%s" % (C["grn"], C["r"]))
    input("  按回车返回")


def del_flow():
    d = load()
    done = [o for o in d.get("orders", [])
            if o.get("status") in DONE and not o.get("archived")]
    if not done:
        print("  没有可归档的完结订单")
        input("  按回车返回")
        return
    view()
    s = input("  输入要归档的完结单号：").strip()
    hit = [o for o in done if o.get("serial") == s]
    if len(hit) != 1:
        print("  %s未找到或单号不唯一；未修改台账%s" % (C["red"], C["r"]))
    else:
        try:
            archive_or_void_order(s, hit[0].get("product", ""), "控制台归档")
            print("  %s已归档 %s（历史记录保留）%s" % (C["grn"], s, C["r"]))
        except Exception as exc:
            print("  %s归档失败：%s%s" % (C["red"], exc, C["r"]))
    input("  按回车返回")


def urgent_orders():
    """供日报调用：返回（逾期列表, 临期列表）"""
    out = ([], [])
    for o in load().get("orders", []):
        tag, dd = urgency(o)
        if tag == "逾期":
            out[0].append((o, dd))
        elif tag == "临期":
            out[1].append((o, dd))
    return out


def main():
    while True:
        clear()
        bar("订单台账")
        print()
        print("  %s1%s  查看台账" % (C["cyn"], C["r"]))
        print("  %s2%s  添加订单" % (C["cyn"], C["r"]))
        print("  %s3%s  改状态 / 交期 / 备注" % (C["cyn"], C["r"]))
        print("  %s4%s  归档已完结订单（保留记录）" % (C["cyn"], C["r"]))
        print("  %s0%s  退出" % (C["gry"], C["r"]))
        print()
        c = input("  请选择：").strip()
        if c == "1":
            view()
            input("  按回车返回")
        elif c == "2":
            add_flow()
        elif c == "3":
            change_flow()
        elif c == "4":
            del_flow()
        elif c == "0":
            break


if __name__ == "__main__":
    main()
