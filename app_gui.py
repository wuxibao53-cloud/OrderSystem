# -*- coding: utf-8 -*-
"""跟单助手图形版（Eel + Edge 窗口）：出单 / 库存 / 新增产品"""
import os
import sys
import json
import math
import re
import io
import hashlib
import zipfile


def _read_json(path):
    with open(path, encoding="utf-8") as stream:
        return json.load(stream)


def _positive_int_or_default(value, default, label):
    if value in (None, ""):
        return int(default)
    try:
        parsed = int(value)
    except (TypeError, ValueError):
        raise ValueError("%s必须是正整数" % label)
    if parsed <= 0:
        raise ValueError("%s必须是正整数" % label)
    return parsed

import time
import base64
import tempfile
import datetime
import threading
from functools import wraps
from json_store import save_json_atomic

APP = os.path.dirname(os.path.abspath(__file__))   # 便携：系统放哪个目录都能跑
sys.path.insert(0, APP)

import eel  # noqa: E402
from generator import generate, load_recipes, TEMPLATE, SHEET  # noqa: E402
from stock import (build_report, compare_stock, latest_inbox, find_latest_stock_file,  # noqa: E402
                   allocate, batch_fifo_key, batch_summary, batches_at_book_balance,
                   import_batch_xls, find_latest_batch_file)

RECIPE_JSON = os.path.join(APP, "recipes.json")
_STOCK_WRITE_LOCK = threading.RLock()


def _serialize_stock_write(fn):
    @wraps(fn)
    def wrapped(*args, **kwargs):
        with _STOCK_WRITE_LOCK:
            return fn(*args, **kwargs)
    return wrapped

# 禁用浏览器缓存：每次都加载最新界面，避免改版后页面不更新
try:
    @eel.bottle_app.hook("after_request")
    def _no_cache():
        eel.bottle.response.set_header("Cache-Control", "no-store, must-revalidate")
        eel.bottle.response.set_header("Pragma", "no-cache")
except Exception:
    pass


def _load_orders():
    p = os.path.join(APP, "orders.json")
    if not os.path.exists(p):
        return {"orders": []}
    with open(p, encoding="utf-8") as stream:
        return json.load(stream)


@eel.expose
def ping():
    return "pong"


@eel.expose
def report_js_error(msg):
    """页面 JS 错误写入日志，供远程诊断"""
    try:
        d = os.path.join(APP, "data")
        os.makedirs(d, exist_ok=True)
        with open(os.path.join(d, "js_errors.log"), "a", encoding="utf-8") as f:
            f.write("%s  %s\n" % (datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"), msg))
    except Exception:
        pass
    return {"ok": True}


def _classify_frontend_errors(lines):
    summary = {"network": 0, "dialog_permission": 0, "runtime": 0}
    for line in lines:
        if "Failed to fetch" in line or "接口无响应/异常" in line:
            summary["network"] += 1
        elif "dialog|confirm" in line:
            summary["dialog_permission"] += 1
        else:
            summary["runtime"] += 1
    return summary


@eel.expose
def get_products():
    try:
        recipes = load_recipes()
        out = []
        for k, v in recipes.items():
            if k.startswith("_") or not isinstance(v, dict):
                continue
            out.append({"name": k, "tablet_g": v.get("tablet_g", 0.5),
                        "category": v.get("category", ""),
                        "materials": v.get("materials", []),
                        "packaging": v.get("packaging", {})})
        return {"ok": True, "products": out}
    except Exception as e:
        return {"ok": False, "msg": "%s: %s" % (type(e).__name__, e)}


@eel.expose
def suggest_batch(product, date8):
    """同一产品同一天已出几单 → 建议批次号"""
    try:
        n = 0
        date_clean = (date8 or "").replace("-", "")
        for o in _load_orders().get("orders", []):
            if o.get("product") == product and \
                    (o.get("order_date") or "").replace("-", "") == date_clean:
                n += 1
        return {"ok": True, "batch": "%02d" % (n + 1)}
    except Exception as e:
        return {"ok": False, "msg": str(e)}


@eel.expose
def order_preview(product, mode, qty, tablet, materials=None, per_bottle=None,
                  recipe_override=None):
    try:
        r = recipe_override if isinstance(recipe_override, dict) else load_recipes().get(product)
        if not r:
            return {"ok": False, "msg": "无此产品"}
        try:
            tablet = float(tablet)
        except Exception:
            tablet = float(r.get("tablet_g", 0.5))
        if mode == "kg":
            total_kg = float(qty)
            wan = total_kg * 1000 / tablet / 10000
        else:
            if mode == "bottle":
                pb_in = _positive_int_or_default(
                    per_bottle, r.get("packaging", {}).get("per_bottle") or 60,
                    "每瓶片数")
                wan = float(qty) * pb_in / 10000
            elif mode == "tablet":
                wan = float(qty) / 10000
            else:
                wan = float(qty)
            total_kg = wan * 10000 * tablet / 1000
        meta, rows, alert, warn = build_report()
        by = {x["short"]: x for x in rows}
        mats = materials if materials else r.get("materials", [])
        items = []
        for m in mats:
            need = round(total_kg * float(m["ratio"]) / 100, 3)
            row = by.get(m["name"])
            if not row:
                # Only infer inventory aliases when the substring match is unambiguous.
                nl = m["name"].lower()
                fuzzy = []
                for short, candidate in by.items():
                    sl = short.lower()
                    if nl in sl or sl in nl:
                        fuzzy.append(candidate)
                if len(fuzzy) == 1:
                    row = fuzzy[0]
            avail = round(row["avail"], 3) if row else 0.0
            unit = row["unit"] if row else "KG"
            matched = row["short"] if row else ""
            # 批次分配（按有效期先进先出）；单批不够时自动拆到多个批次
            alloc = ""
            alloc_list = []
            default_batch = ""
            batch_shortage = 0.0
            if row and row.get("batches"):
                plan = ([ (str(m.get("batch") or "").strip(), need) ]
                        if str(m.get("batch") or "").strip()
                        else allocate({"batches": row["batches"]}, need,
                                      occupied=row.get("occupied", 0),
                                      book_qty=row.get("qty")))
                alloc_list = [{"batch": b, "qty": q} for b, q in plan if q > 0]
                alloc = "；".join("%s 取 %g" % (b, q) for b, q in plan if q > 0)
                if not str(m.get("batch") or "").strip():
                    allocated = sum(entry["qty"] for entry in alloc_list)
                    batch_shortage = round(max(0.0, need - allocated), 3)
                    if batch_shortage:
                        alloc += "%s批次可用不足 %g kg" % ("；" if alloc else "", batch_shortage)
                if alloc_list:
                    default_batch = alloc_list[0]["batch"]
            if not default_batch:
                default_batch = (row.get("batch") if row else "") or ""
            multi = len(alloc_list) > 1
            shortage = round(max(need - avail, batch_shortage), 3)
            items.append({"name": m["name"], "batch": m.get("batch", ""),
                          "default_batch": default_batch,
                          "batch_summary": batch_summary(row) if row else "",
                          "batch_options": [
                              {"batch": b.get("batch", ""), "qty": b.get("qty", 0),
                               "in_date": b.get("in_date", ""),
                               "mfg": b.get("mfg", ""), "exp": b.get("exp", "")}
                              for b in batches_at_book_balance(row)[0]
                          ] if row else [],
                          "need": need, "avail": avail, "unit": unit,
                          "enough": shortage <= 0.0005, "shortage": shortage,
                          "batch_shortage": batch_shortage, "alloc": alloc,
                          "alloc_list": alloc_list, "multi": multi,
                          "matched": matched,
                          "status": (row.get("status") if row else "未建档")})
        pk = r.get("packaging", {})
        per_box = _positive_int_or_default(pk.get("per_box"), 500, "每箱瓶数")
        pb = _positive_int_or_default(
            per_bottle, pk.get("per_bottle") or 60, "每瓶片数")
        bottles = int(round(wan * 10000 / pb)) if pb else 0
        boxes = int(round(bottles / per_box)) if per_box else 0
        return {"ok": True, "wan": round(wan, 4), "total_kg": round(total_kg, 3),
                "items": items, "bottles": bottles, "boxes": boxes}
    except Exception as e:
        return {"ok": False, "msg": "%s: %s" % (type(e).__name__, e)}


@eel.expose
@_serialize_stock_write
def generate_order(cfg):
    try:
        from orders import (
            ORDER_TYPES, add_order, delete_test_order, delete_unstarted_order,
            order_exists,
        )
        order_type = cfg.get("order_type", "正式生产")
        if order_type not in ORDER_TYPES:
            return {"ok": False, "msg": "生产记录分类无效，未生成"}
        date8 = cfg.get("date") or datetime.datetime.now().strftime("%Y%m%d")
        serial = "%s-%02d" % (date8, int(cfg.get("seq") or 1))
        if order_exists(serial, cfg.get("product")):
            return {"ok": False, "msg": "订单台账中已存在 %s（%s），未重复生成。"
                                      "请检查生产序号或打开已有档案。" %
                                      (serial, cfg.get("product", ""))}
    except Exception as e:
        return {"ok": False, "msg": "检查订单台账失败，已取消生成：%s" % e}
    order = {
        "serial": serial, "product": cfg.get("product", ""),
        "wan": cfg.get("wan", 0), "tablet": cfg.get("tablet_g", 0.5),
        "total_kg": cfg.get("total_kg"),
        "customer": "", "due": cfg.get("due", ""),
        "sales": cfg.get("manager", ""), "status": "待生产",
        "order_date": date8, "batch": cfg.get("batch", "01"),
        "order_type": order_type, "materials": [], "packaging": [],
    }
    try:
        added, message = add_order(order)
    except Exception as e:
        return {"ok": False, "msg": "订单台账写入失败，未生成文件：%s" % e}
    if not added:
        return {"ok": False, "msg": "订单台账写入失败，未生成文件：%s" % message}

    try:
        res = generate(cfg)
    except Exception as e:
        res = {"ok": False, "msg": "%s: %s" % (type(e).__name__, e)}
    if not res.get("ok"):
        try:
            if order_type in ("测试单", "实验单"):
                delete_test_order(serial, order["product"])
            else:
                delete_unstarted_order(serial, order["product"])
        except Exception as e:
            res["msg"] += "；临时台账记录回滚失败，请检查订单 %s：%s" % (serial, e)
        return res

    try:
        from orders import set_order_file
        set_order_file(serial, order["product"], res["file"], {
            "wan": res["plan_wan"], "tablet": res["tablet_g"],
            "total_kg": res["total_kg"], "materials": res.get("materials", []),
            "packaging": res.get("packaging", []),
        })
        res["ledger"] = "订单已登记并关联生产文件"
    except Exception as e:
        res["ledger"] = (
            "订单已写入台账且可刷新查看，但文件路径同步失败：%s" % e
        )
    return res


@eel.expose
def get_stock(kw=""):
    try:
        meta, rows, alert, warn = build_report()
        k = (kw or "").lower()
        out = []
        for r in rows:
            if k and k not in r["short"].lower() and k not in r["name"].lower() \
                    and k not in (r.get("spec") or "").lower() and k not in r["cat"].lower():
                continue
            out.append({"short": r["short"], "cat": r["cat"].replace("原材料-", ""),
                        "qty": r["qty"], "unit": r["unit"],
                        "occupied": r.get("occupied", 0), "avail": r.get("avail", r["qty"]),
                        "line": r["line"], "status": r["status"],
                        "batches": batch_summary(r)})
        # 按紧急程度排序：断料 > 不足 > 偏低 > 正常，组内按类别
        order_rank = {"断料": 0, "不足": 1, "偏低": 2, "正常": 3}
        out.sort(key=lambda x: (order_rank.get(x["status"], 9), x["cat"], x["short"]))
        fg = []
        sp = os.path.join(APP, "stock.json")
        data = _read_json(sp) if os.path.exists(sp) else {"items": []}
        for it in data.get("items", []):
            if it["cat"].startswith("产成品"):
                if k and k not in it["name"].lower() and k not in it.get("spec", "").lower():
                    continue
                fg.append({"name": it["name"], "qty": it["qty"], "unit": it["unit"],
                           "spec": it.get("spec", "")})
        return {"ok": True, "raw": out, "fg": fg, "stock_date": meta.get("stock_date", "-"),
                "alerts": len(alert), "warns": len(warn)}
    except Exception as e:
        return {"ok": False, "msg": "%s: %s" % (type(e).__name__, e)}


@eel.expose
def search_stock(kw="", limit=30):
    """物料/包材名称联想检索：输入部分字即可匹配，返回可直接回填一行的完整信息。
    排序：前缀命中 > 包含命中；同类按紧急度。"""
    try:
        meta, rows, alert, warn = build_report()
        k = (kw or "").strip().lower()
        hits = []
        for r in rows:
            short = r["short"] or ""
            sl = short.lower()
            nl = (r.get("name") or "").lower()
            spec = (r.get("spec") or "").lower()
            cat = (r.get("cat") or "").lower()
            if not k:
                score = 50
            elif sl == k or nl == k:
                score = 0
            elif sl.startswith(k) or nl.startswith(k):
                score = 1
            elif k in sl or k in nl:
                score = 2
            elif k in spec or k in cat:
                score = 3
            else:
                continue
            # 默认批号：按有效期先进先出的第一批
            default_batch = ""
            if r.get("batches"):
                plan = allocate({"batches": r["batches"]}, 1e9,
                                occupied=r.get("occupied", 0), book_qty=r.get("qty"))
                positive = [b for b, q in plan if q > 0]
                if positive:
                    default_batch = positive[0]
                elif r["batches"]:
                    default_batch = sorted(r["batches"], key=batch_fifo_key)[0].get("batch", "")
            if not default_batch:
                default_batch = r.get("batch", "") or ""
            rank = {"断料": 0, "不足": 1, "偏低": 2, "正常": 3}
            hits.append({
                "short": short,
                "name": r.get("name", short),
                "cat": r.get("cat", "").replace("原材料-", "").replace("包装材料-", ""),
                "spec": r.get("spec", ""),
                "qty": r["qty"], "unit": r["unit"],
                "avail": r.get("avail", r["qty"]),
                "occupied": r.get("occupied", 0),
                "status": r["status"],
                "default_batch": default_batch,
                "batches": batch_summary(r),
                "batch_options": [
                    {"batch": b.get("batch", ""), "qty": b.get("qty", 0),
                     "in_date": b.get("in_date", ""), "mfg": b.get("mfg", ""),
                     "exp": b.get("exp", "")}
                    for b in batches_at_book_balance(r)[0]
                ],
                "_s": score, "_r": rank.get(r["status"], 9),
            })
        hits.sort(key=lambda x: (x["_s"], x["_r"], len(x["short"]), x["short"]))
        hits = hits[:int(limit)]
        for h in hits:
            h.pop("_s", None)
            h.pop("_r", None)
        return {"ok": True, "items": hits, "kw": kw}
    except Exception as e:
        return {"ok": False, "msg": "%s: %s" % (type(e).__name__, e)}


@eel.expose
@_serialize_stock_write
def import_stock():
    try:
        p = find_latest_stock_file() or latest_inbox()
        if not p:
            return {"ok": False, "msg": "没找到库存表：已搜索 inbox、D:\\Downloads、用户 Downloads，"
                                        "都没有含「库存/余额」的 xls 文件。请先在柠檬云点「导出」。"}
        rows, diffs, report, n, company, sd = compare_stock(p)
        return {"ok": True, "n": n, "sd": sd, "report": report, "source": p,
                "diffs": [{"short": d["short"], "delta": d["delta"],
                           "expected": d["expected"], "diff": d["diff"]} for d in diffs]}
    except Exception as e:
        return {"ok": False, "msg": "%s: %s" % (type(e).__name__, e)}


@eel.expose
@_serialize_stock_write
def import_batches(path=None, force=False):
    """导入柠檬云「批次状态表 / 批次跟踪表」，把批号写进库存（这是批号能自动带出的前提）"""
    try:
        p = path or find_latest_batch_file()
        if not p or not os.path.exists(p):
            return {"ok": False, "updated": False,
                    "msg": "没找到批次表。请到柠檬云进销存 → 库存管理 → 报表 → 批次状态表（或批次跟踪表），"
                           "导出后放到 D:\\Downloads 或系统 inbox 目录，文件名要含「批次」。"}
        st = os.stat(p)
        sig = "B|%s|%d|%d" % (os.path.abspath(p), st.st_size, int(st.st_mtime))
        sig_file = os.path.join(APP, "data", "last_import.json")
        last = {}
        if os.path.exists(sig_file):
            try:
                last = _read_json(sig_file)
            except Exception:
                last = {}
        if (not force) and last.get("batch_sig") == sig:
            return {"ok": True, "updated": False, "reason": "无新批次表",
                    "source": p, "last_time": last.get("batch_time", "")}
        hit, nb, miss = import_batch_xls(p)
        try:
            os.makedirs(os.path.dirname(sig_file), exist_ok=True)
            last["batch_sig"] = sig
            last["batch_time"] = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            last["batch_source"] = p
            save_json_atomic(sig_file, last)
        except Exception:
            pass
        return {"ok": True, "updated": True, "items": hit, "batches": nb,
                "source": p, "miss": miss[:20], "n_miss": len(miss),
                "time": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")}
    except Exception as e:
        return {"ok": False, "updated": False, "msg": "%s: %s" % (type(e).__name__, e)}


def _fail_count():
    try:
        f = os.path.join(APP, "data", "last_import.json")
        d = _read_json(f) if os.path.exists(f) else {}
        return int(d.get("fail_count", 0))
    except Exception:
        return 0


def _bump_fail(n=1, why=""):
    """记录自动更新连续失败次数；连续 3 次以上就要求人工介入"""
    try:
        f = os.path.join(APP, "data", "last_import.json")
        d = _read_json(f) if os.path.exists(f) else {}
        d["fail_count"] = int(d.get("fail_count", 0)) + n
        d["fail_why"] = why
        d["fail_at"] = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        os.makedirs(os.path.dirname(f), exist_ok=True)
        save_json_atomic(f, d)
        return d["fail_count"]
    except Exception:
        return 0


def _clear_fail():
    try:
        f = os.path.join(APP, "data", "last_import.json")
        d = _read_json(f) if os.path.exists(f) else {}
        d["fail_count"] = 0
        os.makedirs(os.path.dirname(f), exist_ok=True)
        save_json_atomic(f, d)
    except Exception:
        pass


@eel.expose
def batch_reconcile_report():
    """只读：列出批次合计与账面不一致的物料。批次可能是按账面估算，仅供人工核对，不改数据。"""
    try:
        data = _read_json(os.path.join(APP, "stock.json"))
        rows = []
        for it in data.get("items", []):
            bs = it.get("batches") or []
            if not bs:
                continue
            total = round(sum(float(b.get("qty", 0)) for b in bs), 3)
            q = float(it.get("qty", 0))
            if abs(total - q) > max(0.01, abs(q) * 0.001):
                rows.append({"short": it.get("short", it["name"]), "book": q,
                             "batch_total": total, "diff": round(total - q, 3)})
        rows.sort(key=lambda r: -abs(r["diff"]))
        return {"ok": True, "n": len(rows), "items": rows}
    except Exception as e:
        return {"ok": False, "msg": "%s: %s" % (type(e).__name__, e)}


def _batch_audit_rows(data):
    """批次核对数据（纯函数）：无批号清单 + 批次合计与账面不符清单。"""
    missing, mismatch = [], []
    for it in data.get("items", []):
        cat = str(it.get("cat") or "")
        if cat.startswith("产成品"):
            continue
        name = it.get("short") or it.get("name") or ""
        qty = float(it.get("qty", 0) or 0)
        unit = it.get("unit", "")
        batches = it.get("batches") or []
        if not batches:
            missing.append({"name": name, "qty": qty, "unit": unit,
                            "cat": cat.replace("原材料-", "")})
            continue
        total = round(sum(float(b.get("qty", 0)) for b in batches), 3)
        if abs(total - qty) > max(0.01, abs(qty) * 0.001):
            mismatch.append({"name": name, "book": qty, "batch_total": total,
                             "diff": round(total - qty, 3), "unit": unit,
                             "n_batches": len(batches)})
    mismatch.sort(key=lambda r: -abs(r["diff"]))
    missing.sort(key=lambda r: -abs(r["qty"]))
    return missing, mismatch


@eel.expose
def export_batch_audit():
    """导出批次核对表（.md）：无批号物料清单 + 批次合计与账面不符清单，供盘点时照单核对。"""
    try:
        data = _read_json(os.path.join(APP, "stock.json"))
        stock_date = data.get("stock_date") or data.get("date") or "-"
        missing, mismatch = _batch_audit_rows(data)
        stamp = datetime.datetime.now().strftime("%Y%m%d_%H%M")
        out_dir = os.path.join(APP, "data", "compare")
        os.makedirs(out_dir, exist_ok=True)
        path = os.path.join(out_dir, "批次核对_%s.md" % stamp)
        lines = ["# 批次核对表（%s 生成，库存日期 %s）" % (
            datetime.datetime.now().strftime("%Y-%m-%d %H:%M"), stock_date), ""]
        lines.append("## 一、还没有批号的物料（%d 项）" % len(missing))
        if missing:
            lines += ["", "| 物料 | 账面数量 | 单位 | 类别 |", "|---|---|---|---|"]
            lines += ["| %s | %g | %s | %s |" % (m["name"], m["qty"], m["unit"], m["cat"])
                      for m in missing]
        else:
            lines.append("全部原材料都有批号。")
        lines += ["", "## 二、批次合计与账面不符（%d 项）" % len(mismatch)]
        if mismatch:
            lines += ["", "| 物料 | 账面数量 | 批次合计 | 差值（批-账） | 批次数 |",
                      "|---|---|---|---|---|"]
            lines += ["| %s | %g | %g | %+g | %d |" % (
                m["name"], m["book"], m["batch_total"], m["diff"], m["n_batches"])
                for m in mismatch]
        else:
            lines.append("已有批号的物料，批次合计与账面一致。")
        lines += ["", "## 三、处理建议", "",
                  "1. 差异先做实物盘点确认；确认后可在库存页「按账面修正批次数」或重新导入柠檬云批次表。",
                  "2. 无批号物料下次导出《批次跟踪表》时补齐，预占和先进先出才会准确。",
                  "3. 本表只是核对清单，不会修改任何数据。"]
        with open(path, "w", encoding="utf-8") as f:
            f.write("\n".join(lines) + "\n")
        return {"ok": True, "path": path, "missing": len(missing),
                "mismatch": len(mismatch)}
    except Exception as e:
        return {"ok": False, "msg": "%s: %s" % (type(e).__name__, e)}


@eel.expose
@_serialize_stock_write
def fix_batch_qty():
    """按柠檬云账面数量等比修正批次数量（批次数据过期时用，让批次跟库存保持一致）。
    只会缩放数量，不会删批号；批次总量为 0 或账面为 0 时跳过。"""
    try:
        p = os.path.join(APP, "stock.json")
        data = _read_json(p)
        fixed = []
        for it in data.get("items", []):
            bs = it.get("batches") or []
            if not bs:
                continue
            total = sum(float(b.get("qty", 0)) for b in bs)
            q = float(it.get("qty", 0))
            if q <= 0:
                continue
            if total <= 0:
                # 批次数量全部未知（如跟踪表只有出库流水）：按最早入库批次承接账面数量
                bs.sort(key=batch_fifo_key)
                for b in bs:
                    b["qty"] = 0.0
                bs[0]["qty"] = round(q, 3)
                fixed.append({"short": it.get("short", it["name"]),
                              "from": 0.0, "to": q, "note": "数量未知，已归到最早入库批次"})
                continue
            if abs(total - q) <= max(0.01, abs(q) * 0.001):
                continue
            k = q / total
            for b in bs:
                b["qty"] = round(float(b.get("qty", 0)) * k, 3)
            # 尾差补到最后一个批次上，保证合计严格等于账面
            diff = round(q - sum(float(b["qty"]) for b in bs), 3)
            if abs(diff) > 0.0005:
                bs[-1]["qty"] = round(float(bs[-1]["qty"]) + diff, 3)
            fixed.append({"short": it.get("short", it["name"]),
                          "from": round(total, 3), "to": q})
        save_json_atomic(p, data)
        return {"ok": True, "n": len(fixed), "items": fixed}
    except Exception as e:
        return {"ok": False, "msg": "%s: %s" % (type(e).__name__, e)}


@eel.expose
@_serialize_stock_write
def auto_import_stock(force=False):
    """自动库存更新：发现新的柠檬云库存导出表就自动导入（以柠檬云为准）。
    靠「文件路径 + 大小 + 修改时间」做指纹，已导入过的不会重复导入。
    force=True 时强制重新导入（手动备用方案走这个）。"""
    try:
        p = None
        for attempt in range(3):          # 防护：一次不成功立刻再试，最多 3 次
            p = find_latest_stock_file() or latest_inbox()
            if p and os.path.exists(p):
                break
            time.sleep(3)
        if not p:
            sig_path = os.path.join(APP, "data", "last_import.json")
            try:
                prev = _read_json(sig_path) if os.path.exists(sig_path) else {}
            except Exception:
                prev = {}
            if prev.get("sig"):
                # 曾成功导入过：源表被移走属于正常，只提示、不计失败
                _clear_fail()
                return {"ok": True, "updated": False, "reason": "无新表",
                        "stock_date": prev.get("sd", ""), "last_time": prev.get("time", ""),
                        "fail_count": 0,
                        "msg": "未发现新的库存表（上次导入的源文件已移走，属正常）。有新导出时放到 D:\\Downloads 或 inbox 即可。"}
            _bump_fail(3, "未找到库存表")
            return {"ok": True, "updated": False, "reason": "未找到库存表",
                    "fail_count": _fail_count(),
                    "msg": "连试 3 次都没找到柠檬云导出的库存表（搜索范围：inbox、D:\\Downloads、用户 Downloads）。"
                           "请手动导出后放到 D:\\Downloads，或点库存页「导入最新库存表（手动）」"}
        st = os.stat(p)
        sig = "%s|%d|%d" % (os.path.abspath(p), st.st_size, int(st.st_mtime))
        sig_file = os.path.join(APP, "data", "last_import.json")
        last = {}
        if os.path.exists(sig_file):
            try:
                last = _read_json(sig_file)
            except Exception:
                last = {}
        if (not force) and last.get("sig") == sig:
            # 库存表没变，但批次表可能是新的 → 顺便检查一次批次表
            _clear_fail()
            br = import_batches(force=force)
            return {"ok": True, "updated": False, "reason": "无新表",
                    "stock_date": last.get("sd", ""), "source": p,
                    "last_time": last.get("time", ""), "batch": br, "fail_count": 0}
        rows, diffs, report, n, company, sd = compare_stock(p)
        try:
            os.makedirs(os.path.dirname(sig_file), exist_ok=True)
            last.update({"sig": sig, "sd": sd, "source": p, "n": n,
                         "time": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")})
            save_json_atomic(sig_file, last)
        except Exception:
            pass
        # 库存表更新后批次会保留，再导入一次批次表（可能是新表）
        _clear_fail()
        br = import_batches(force=force)
        return {"ok": True, "updated": True, "n": n, "sd": sd, "source": p,
                "report": report, "batch": br, "fail_count": 0,
                "time": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                "diffs": [{"short": d["short"], "delta": d["delta"],
                           "expected": d["expected"], "diff": d["diff"]} for d in diffs]}
    except Exception as e:
        # 解析失败也重试：文件可能正在下载/被 Excel 占用
        last_err = None
        for attempt in range(2):
            try:
                time.sleep(4)
                p = find_latest_stock_file() or latest_inbox()
                if not p:
                    continue
                rows, diffs, report, n, company, sd = compare_stock(p)
                _clear_fail()
                br = import_batches(force=force)
                return {"ok": True, "updated": True, "n": n, "sd": sd, "source": p,
                        "report": report, "batch": br, "retry": attempt + 1, "fail_count": 0,
                        "time": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                        "diffs": [{"short": d["short"], "delta": d["delta"],
                                   "expected": d["expected"], "diff": d["diff"]} for d in diffs]}
            except Exception as e2:
                last_err = e2
        fc = _bump_fail(1, "%s: %s" % (type(e).__name__, e))
        return {"ok": False, "msg": "%s: %s" % (type(e).__name__, e), "updated": False,
                "fail_count": fc, "need_manual": fc >= 3}


@eel.expose
def ai_status():
    """AI 助手状态（不含密钥）。"""
    try:
        import ai_assistant
        return {"ok": True, **ai_assistant.status()}
    except Exception as e:
        return {"ok": False, "msg": "%s: %s" % (type(e).__name__, e)}


def _stock_material_names():
    try:
        data = _read_json(os.path.join(APP, "stock.json"))
        return [it.get("short") or it.get("name") or ""
                for it in data.get("items", [])
                if not str(it.get("cat") or "").startswith("产成品")]
    except Exception:
        return []


def _ai_form_from_docx(temp_path):
    from sales_form_import import ai_extract_form
    return ai_extract_form(docx_path=temp_path, stock_names=_stock_material_names())


@eel.expose
def ai_preview_sales_form_upload(filename, encoded_content):
    """AI 兜底解析销售/配方单（新排版、解析器失败时用）：纯提取不落盘。
    返回与普通预览同构的配方信息 + 溯源引用 + 校验结果。"""
    try:
        safe_name = os.path.basename(str(filename or "").replace("\\", "/"))
        if not safe_name.lower().endswith(".docx"):
            return {"ok": False, "msg": "请选择 Word .docx 文件"}
        try:
            content = base64.b64decode(encoded_content, validate=True)
        except (ValueError, TypeError):
            return {"ok": False, "msg": "上传文件数据无效，请重新选择文件"}
        if not content or len(content) > 20 * 1024 * 1024:
            return {"ok": False, "msg": "文件为空或超过 20 MB，请检查后重试"}
        with tempfile.TemporaryDirectory(prefix="ai-preview-") as directory:
            source = os.path.join(directory, safe_name)
            with open(source, "wb") as stream:
                stream.write(content)
            extracted = _ai_form_from_docx(source)
        form = extracted["form"]
        raw = extracted.get("raw") or {}
        quotes = []
        for item in raw.get("materials") or []:
            if isinstance(item, dict) and str(item.get("quote") or "").strip():
                quotes.append({"field": item.get("name") or "原料",
                               "quote": str(item.get("quote")).strip()})
        if isinstance(raw.get("coat"), dict) and str(
                raw["coat"].get("quote") or "").strip():
            quotes.append({"field": raw["coat"].get("name") or "包衣剂",
                           "quote": str(raw["coat"]["quote"]).strip()})
        return {
            "ok": True, "ai": extracted["ai"],
            "product": form["product"], "tablet_g": form["tablet_g"],
            "tablet_range": "", "hardness": "",
            "per_bottle": form["per_bottle"], "coat_type": form["coat_type"],
            "materials": [{"name": m["name"], "ratio": m["ratio"]}
                          for m in form["materials"]],
            "quotes": quotes,
            "order_date": form["source_date"],
            "next_serial": "", "ledger_day_orders": [], "suggest_serial": "",
            "hints": extracted["warnings"] + (
                ["校验未通过：" + "；".join(extracted["problems"])]
                if extracted["problems"] else []),
            "problems": extracted["problems"],
            "unknown_materials": extracted["unknown_materials"],
            "form_json": extracted["form"],
        }
    except Exception as e:
        return {"ok": False, "msg": "%s: %s" % (type(e).__name__, e)}


@eel.expose
def ai_extract_order_text(text, filename="", encoded_content=""):
    """文字订单（可选附 Word 单据合并解析）→ AI 提取结构化信息，纯提取不落盘。"""
    docx_path = None
    directory = None
    try:
        text = str(text or "").strip()
        safe_name = os.path.basename(str(filename or "").replace("\\", "/"))
        if safe_name:
            if not safe_name.lower().endswith(".docx"):
                return {"ok": False, "msg": "附带的单据必须是 Word .docx 文件"}
            try:
                content = base64.b64decode(encoded_content, validate=True)
            except (ValueError, TypeError):
                return {"ok": False, "msg": "上传文件数据无效，请重新选择文件"}
            if not content or len(content) > 20 * 1024 * 1024:
                return {"ok": False, "msg": "文件为空或超过 20 MB，请检查后重试"}
            directory = tempfile.TemporaryDirectory(prefix="ai-order-")
            docx_path = os.path.join(directory.name, safe_name)
            with open(docx_path, "wb") as stream:
                stream.write(content)
        if not text and not docx_path:
            return {"ok": False, "msg": "请粘贴订单文字或选择 Word 单据（至少一项）"}
        if len(text) > 8000:
            text = text[:8000]
        from sales_form_import import ai_extract_form
        extracted = ai_extract_form(docx_path=docx_path, raw_text=text,
                                    stock_names=_stock_material_names())
        form = extracted["form"]
        raw = extracted.get("raw") or {}
        quotes = []
        for item in raw.get("materials") or []:
            if isinstance(item, dict) and str(item.get("quote") or "").strip():
                quotes.append({"field": item.get("name") or "原料",
                               "quote": str(item.get("quote")).strip()})
        if isinstance(raw.get("coat"), dict) and str(
                raw["coat"].get("quote") or "").strip():
            quotes.append({"field": raw["coat"].get("name") or "包衣剂",
                           "quote": str(raw["coat"]["quote"]).strip()})
        return {"ok": True, "ai": extracted["ai"], "form_json": form,
                "product": form["product"], "tablet_g": form["tablet_g"],
                "per_bottle": form["per_bottle"], "coat_type": form["coat_type"],
                "materials": [{"name": m["name"], "ratio": m["ratio"]}
                              for m in form["materials"]],
                "quotes": quotes,
                "order_date": form["source_date"],
                "problems": extracted["problems"],
                "warnings": extracted["warnings"],
                "unknown_materials": extracted["unknown_materials"]}
    except Exception as e:
        return {"ok": False, "msg": "%s: %s" % (type(e).__name__, e)}
    finally:
        if directory is not None:
            directory.cleanup()


@eel.expose
def ai_progress():
    """AI 任务实时进度（前端轮询，判断进行中/卡住/失败）。"""
    try:
        import ai_assistant
        state = dict(ai_assistant.PROGRESS)
        state["configured"] = ai_assistant.is_configured()
        return state
    except Exception as e:
        return {"stage": "", "detail": "", "done": True,
                "error": "%s: %s" % (type(e).__name__, e), "configured": False}


@eel.expose
def lemon_sync_run(mode="auto"):
    """柠檬云库存同步（auto=核对+安全修正；report=仅核对）。"""
    try:
        import lemon_sync
        return lemon_sync.run(APP, mode if mode in ("auto", "report") else "auto")
    except Exception as e:
        return {"ok": False, "msg": "%s: %s" % (type(e).__name__, e)}


@eel.expose
def lemon_status():
    """上次柠檬云同步状态 + 修正历史尾部（供库存页卡片）。"""
    lemon_dir = os.path.join(APP, "data", "lemon")
    state_path = os.path.join(lemon_dir, "last_sync.json")
    state = {}
    if os.path.exists(state_path):
        try:
            with open(state_path, encoding="utf-8-sig") as fp:
                state = json.load(fp)
        except (OSError, ValueError):
            state = {}
    history = []
    history_path = os.path.join(lemon_dir, "修正历史.csv")
    if os.path.exists(history_path):
        try:
            import csv as _csv
            with open(history_path, encoding="utf-8-sig", newline="") as fp:
                rows = list(_csv.reader(fp))
            history = ["%s %s" % (r[0], r[1]) for r in rows[-6:] if len(r) >= 2]
        except (OSError, ValueError):
            history = []
    forecast_rows = []
    try:
        import csv as _csv
        from lemon_sync import forecast
        stock_data = _read_stock()
        result = forecast(stock_data.get("items", []),
                          _load_orders_safe(), horizon_days=14)
        forecast_rows = result["rows"][:8]
    except Exception:
        forecast_rows = []
    return {"ok": True, "state": state, "history": history,
            "forecast": forecast_rows}


def _read_stock():
    with open(os.path.join(APP, "stock.json"), encoding="utf-8") as fp:
        return json.load(fp)


def _load_orders_safe():
    try:
        from orders import load as load_orders
        return load_orders().get("orders", [])
    except Exception:
        return []


@eel.expose
def purchase_forecast_view():
    """采购预测明细（历史消耗速率 + 待生产占用 → 14 天覆盖建议）。"""
    try:
        from lemon_sync import forecast
        result = forecast(_read_stock().get("items", []), _load_orders_safe())
        return {"ok": True, **result}
    except Exception as e:
        return {"ok": False, "msg": "%s: %s" % (type(e).__name__, e)}


@eel.expose
def ai_ask(question):
    """台账问答助手（只读工具 + 多轮检索）。"""
    try:
        from ai_agent import ask
        return ask(question)
    except Exception as e:
        return {"ok": False, "msg": "%s: %s" % (type(e).__name__, e)}


@eel.expose
def ai_daily_briefing():
    """汇总自检/补料/生产跟踪/批次核对，请 AI 写一段今日晨报，存 data\\daily。"""
    try:
        import ai_assistant
        if not ai_assistant.is_configured():
            return {"ok": False, "msg": "AI 未配置（data\\ai_config.json）"}
        stats = {}
        try:
            check = self_check()
            stats["self_check"] = [
                {"name": i["name"], "level": i["level"], "detail": i["detail"][:120]}
                for i in check.get("items", []) if i["level"] != "ok"]
        except Exception:
            stats["self_check"] = []
        try:
            replenish = get_replenish()
            stats["replenish"] = {"need": replenish.get("need", [])[:12],
                                  "low": replenish.get("low", [])[:8]}
        except Exception:
            stats["replenish"] = {}
        try:
            from orders import production_dashboard
            orders_now = production_dashboard(True)
            stats["orders"] = [
                {"serial": o.get("serial"), "product": o.get("product"),
                 "status": o.get("status"), "total_kg": o.get("total_kg"),
                 "due": o.get("due"), "stale": o.get("tracking_stale")}
                for o in orders_now if o.get("status") in ("待生产", "生产中")]
        except Exception:
            stats["orders"] = []
        try:
            missing_rows, mismatch_rows = _batch_audit_rows(
                _read_json(os.path.join(APP, "stock.json")))
            stats["batch_audit"] = {"missing": len(missing_rows),
                                    "mismatch": len(mismatch_rows)}
        except Exception:
            stats["batch_audit"] = {}
        try:
            lemon_path = os.path.join(APP, "data", "lemon", "last_sync.json")
            if os.path.exists(lemon_path):
                with open(lemon_path, encoding="utf-8-sig") as fp:
                    last = json.load(fp)
                stats["lemon_sync"] = {
                    "at": last.get("at"), "diffs": last.get("diffs"),
                    "fixes": (last.get("fixes") or [])[:5],
                    "warnings": (last.get("warnings") or [])[:5]}
        except Exception:
            stats["lemon_sync"] = {}
        try:
            from lemon_sync import forecast
            result = forecast(_read_stock().get("items", []), _load_orders_safe())
            stats["purchase_forecast"] = [
                {"name": r["name"], "qty": r["qty"], "committed": r["committed"],
                 "daily": r["daily"], "cover_days": r["cover_days"],
                 "suggest": r["suggest"], "urgent": r["urgent"]}
                for r in result["rows"][:10]]
        except Exception:
            stats["purchase_forecast"] = []
        system_prompt = (
            "你是食品厂跟单助理。根据提供的经营数据 JSON 写一段今日晨报："
            "1) 今天要盯的生产单（编号、产品、状态、交期）；2) 库存缺口与需采购物料"
            "（结合采购预测的紧急项，写明建议数量）；3) 自检发现的问题与建议处理顺序；"
            "4) 批次数据质量与柠檬云核对差异提醒。"
            "用简体中文，条目化，不超过 300 字，不要编造数据里没有的信息。")
        briefing = ai_assistant.chat(
            [{"role": "system", "content": system_prompt},
             {"role": "user", "content": json.dumps(stats, ensure_ascii=False)}])
        stamp = datetime.datetime.now().strftime("%Y%m%d")
        out_dir = os.path.join(APP, "data", "daily")
        os.makedirs(out_dir, exist_ok=True)
        path = os.path.join(out_dir, "晨报_%s.md" % stamp)
        with open(path, "w", encoding="utf-8") as fp:
            fp.write("# 今日晨报（%s，模型 %s）\n\n%s\n" % (
                datetime.datetime.now().strftime("%Y-%m-%d %H:%M"),
                ai_assistant.load_config().get("model", ""), briefing))
        return {"ok": True, "text": briefing, "path": path}
    except Exception as e:
        return {"ok": False, "msg": "%s: %s" % (type(e).__name__, e)}


@eel.expose
def preview_sales_form_upload(filename, encoded_content):
    """导入前解析预览：只解析不落盘，返回配方、规格与台账批次建议。"""
    try:
        safe_name = os.path.basename(str(filename or "").replace("\\", "/"))
        if not safe_name.lower().endswith(".docx"):
            return {"ok": False, "msg": "请选择 Word .docx 文件"}
        try:
            content = base64.b64decode(encoded_content, validate=True)
        except (ValueError, TypeError):
            return {"ok": False, "msg": "上传文件数据无效，请重新选择文件"}
        if not content or len(content) > 20 * 1024 * 1024:
            return {"ok": False, "msg": "文件为空或超过 20 MB，请检查后重试"}
        from sales_form_import import preview_sales_form
        with tempfile.TemporaryDirectory(prefix="sales-preview-") as directory:
            source = os.path.join(directory, safe_name)
            with open(source, "wb") as stream:
                stream.write(content)
            return preview_sales_form(source)
    except Exception as e:
        return {"ok": False, "msg": "%s: %s" % (type(e).__name__, e)}


@eel.expose
def mark_production_run(serial, run_no, completed):
    """核销/撤销某一生产批次（预占释放 + 状态流转）。"""
    try:
        from orders import mark_production_run as mark
        return mark(str(serial or "").strip(), run_no, bool(completed))
    except Exception as e:
        return {"ok": False, "msg": "%s: %s" % (type(e).__name__, e)}


@eel.expose
def open_doc(kind):
    """打开系统内的常用位置（换电脑后路径自动跟随）：manual / images / inbox / app / data"""
    m = {
        "manual": os.path.join(APP, "使用与维护手册.md"),
        "images": os.path.join(APP, "images"),
        "inbox": os.path.join(APP, "inbox"),
        "app": APP,
        "data": os.path.join(APP, "data"),
        "template": os.path.dirname(TEMPLATE) if TEMPLATE else APP,
    }
    p = m.get(kind, "")
    if not p:
        return {"ok": False, "msg": "未知位置：%s" % kind}
    if kind in ("images", "inbox", "data"):
        try:
            os.makedirs(p, exist_ok=True)
        except Exception:
            pass
    return open_path(p)


@eel.expose
def clear_logs():
    """清空前端报错日志（自检里看到历史旧报错时用）"""
    try:
        jp = os.path.join(APP, "data", "js_errors.log")
        n = 0
        if os.path.exists(jp):
            with open(jp, encoding="utf-8", errors="ignore") as stream:
                n = len([line for line in stream.read().splitlines() if line.strip()])
            with open(jp, "w", encoding="utf-8"):
                pass
        return {"ok": True, "n": n}
    except Exception as e:
        return {"ok": False, "msg": str(e)}


@eel.expose
def self_check():
    """整页自检：连接、库存新鲜度、批次覆盖、模板、配方库、台账、图片，全部列出问题与处理办法"""
    import subprocess
    from generator import TEMPLATE, SHEET
    chk = []

    def add(name, level, detail, fix=""):
        chk.append({"name": name, "level": level, "detail": detail, "fix": fix})

    today = datetime.datetime.now().strftime("%Y-%m-%d")
    # 1 服务
    add("后端服务", "ok", "HTTP 服务正常（127.0.0.1:8765）", "")
    # 2 库存数据
    try:
        data = _read_json(os.path.join(APP, "stock.json"))
        meta = data.get("_meta", {})
        items = data.get("items", [])
        sd = meta.get("stock_date", "")
        imp = meta.get("imported_at", "")
        days = 0
        if sd:
            try:
                days = (datetime.datetime.now() - datetime.datetime.strptime(sd, "%Y-%m-%d")).days
            except Exception:
                days = -1
        if not sd:
            add("库存数据", "err", "库存里没有日期，可能从未导入",
                "点库存页「导入最新库存表（手动）」，或把柠檬云导出的表放进 D:\\Downloads")
        elif days >= 2:
            add("库存新鲜度", "warn", "库存日期 %s，距今 %d 天（导入于 %s）" % (sd, days, imp),
                "重新到柠檬云导出库存余额表，系统 10 分钟内自动识别；或点「强制重新导入」")
        else:
            add("库存新鲜度", "ok", "库存日期 %s（导入于 %s），共 %d 项" % (sd, imp, len(items)), "")
        # 3 批次覆盖
        raw = [i for i in items if not i.get("cat", "").startswith("产成品")]
        withb = [i for i in raw if i.get("batches")]
        pct = int(round(100.0 * len(withb) / len(raw))) if raw else 0
        bsrc = meta.get("batch_source", "")
        if not withb:
            add("批次覆盖", "err", "全部 %d 项原材料都没有批号，指令单上的批号无法自动带出" % len(raw),
                "柠檬云进销存 → 库存管理 → 报表 → 批次状态表（或批次跟踪表）→ 导出，"
                "文件名含「批次」，放到 D:\\Downloads，再点库存页「导入批次表」")
        elif pct < 60:
            add("批次覆盖", "warn", "%d/%d 项原材料有批号（%d%%），来源：%s"
                % (len(withb), len(raw), pct, bsrc or "手工录入"),
                "其余物料在柠檬云未启用批次管理，或批次表里没有它们；可手工在「本次批号」列填")
        else:
            add("批次覆盖", "ok", "%d/%d 项原材料有批号（%d%%），来源：%s"
                % (len(withb), len(raw), pct, bsrc or "手工录入"), "")
        # 批次数量与账面是否对得上（对不上说明批次数据过期）
        mism = []
        for i in withb:
            s = round(sum(float(b.get("qty", 0)) for b in i["batches"]), 3)
            q = float(i.get("qty", 0))
            if abs(s - q) > max(0.5, abs(q) * 0.05):
                mism.append("%s（账面 %g / 批次合计 %g）" % (i.get("short", i["name"]), q, s))
        if mism:
            add("批次数量一致性", "warn",
                "%d 项物料的批次合计与账面不符：%s" % (len(mism), "、".join(mism[:6])),
                "重新导出「批次状态表」并点「导入批次表（批号来源）」，批次会按最新数据覆盖")
        elif withb:
            add("批次数量一致性", "ok", "已有批号的物料，批次合计与账面一致", "")
    except Exception as e:
        add("库存数据", "err", "读取失败：%s" % e, "重新导入库存表")

    # 4 模板
    try:
        if not os.path.exists(TEMPLATE):
            add("指令单模板", "err", "找不到模板文件：%s" % TEMPLATE,
                "恢复程序目录 templates\\生产指令与交接单.xlsx 母版，或检查 data\\config.json 中的模板覆盖路径")
        else:
            import openpyxl
            wb = openpyxl.load_workbook(TEMPLATE, read_only=True)
            try:
                ok_sheet = SHEET in wb.sheetnames
                sz = os.path.getsize(TEMPLATE)
                mt = datetime.datetime.fromtimestamp(os.path.getmtime(TEMPLATE)).strftime("%Y-%m-%d %H:%M")
                if not ok_sheet:
                    add("指令单模板", "err", "模板存在但缺少工作表「%s」（现有：%s）" % (SHEET, "、".join(wb.sheetnames)),
                        "请使用程序目录 templates\\生产指令与交接单.xlsx 中的固定母版")
                else:
                    ws = wb[SHEET]
                    expected_steps = ("包衣液配制", "素片预热", "包衣液雾化", "包衣设备清洗")
                    has_sections = all(
                        str(ws[cell].value or "").startswith("生产指令与交接单")
                        for cell in ("A1", "A32", "A42", "A60"))
                    has_steps = tuple(
                        str(ws["C%d" % row].value or "") for row in range(54, 58)
                    ) == expected_steps
                    if has_sections and has_steps:
                        add("指令单模板", "ok", "%s（%d 字节，更新于 %s），四段母版结构正常"
                            % (os.path.basename(TEMPLATE), sz, mt),
                            "出单与销售单导入共用此母版")
                    else:
                        add("指令单模板", "err", "模板缺少新版混料/压片/包衣/包装区或包衣步骤栏",
                            "请使用程序目录 templates\\生产指令与交接单.xlsx 中的固定母版")
            finally:
                wb.close()
    except Exception as e:
        add("指令单模板", "err", "检查失败：%s" % e, "确认模板文件没被 Excel 占用")

    # 4b 销售单导入
    try:
        import docx  # noqa: F401
        if not os.path.isfile(TEMPLATE):
            add("销售单导入", "err", "缺少统一生产模板 templates\\生产指令与交接单.xlsx",
                "请恢复程序目录中的固定母版")
        else:
            add("销售单导入", "ok", "python-docx 可用，与出单共用固定四段母版", "")
    except ImportError:
        add("销售单导入", "err", "缺少 python-docx 组件", "运行：pip install python-docx")
    except Exception as e:
        add("销售单导入", "err", "检查失败：%s" % e, "确认模板文件没被 Excel 占用")

    # 5 配方库
    try:
        rp = load_recipes()
        bad = []
        for k, v in rp.items():
            s = sum(float(m.get("ratio", 0)) for m in v.get("materials", []))
            if abs(s - 100) > 0.5:
                bad.append("%s(%.2f%%)" % (k, s))
        if bad:
            add("配方库", "warn", "%d 个产品，比例合计不等于 100%%：%s" % (len(rp), "、".join(bad)),
                "在「新增产品」页改比例，重新保存")
        else:
            add("配方库", "ok", "%d 个产品，配比合计均正常" % len(rp), "")
    except Exception as e:
        add("配方库", "err", "读取失败：%s" % e, "检查 recipes.json")

    # 6 台账
    try:
        op = os.path.join(APP, "orders.json")
        od = _read_json(op) if os.path.exists(op) else {}
        lst = od.get("orders", []) if isinstance(od, dict) else od
        pend = [o for o in lst if o.get("status") == "待生产" and not o.get("archived")]
        add("订单台账", "ok", "共 %d 张单，其中待生产 %d 张（待生产会预占库存）" % (len(lst), len(pend)),
            "测试单请及时改状态或删除，避免占用库存")
    except Exception as e:
        add("订单台账", "err", "读取失败：%s" % e, "")

    try:
        from backup_manager import list_backups
        backups = list_backups()
        valid_backups = [backup for backup in backups if backup.get("valid")]
        invalid_backups = [backup for backup in backups if not backup.get("valid")]
        if invalid_backups:
            details = "；".join("%s：%s" % (
                backup["name"], backup.get("validation_error", "校验失败"))
                for backup in invalid_backups[:3])
            add("业务数据备份", "warn",
                "%d 份备份可恢复，另有 %d 份校验失败：%s"
                % (len(valid_backups), len(invalid_backups), details),
                "不要恢复校验失败的文件；保留原件供排查，并重新创建有效备份")
        elif valid_backups:
            add("业务数据备份", "ok", "已有 %d 份校验通过的可恢复备份" % len(valid_backups), "")
        else:
            add("业务数据备份", "warn", "尚未创建业务数据备份",
                "打开「备份恢复」页，点击「立即创建备份」")
    except Exception as e:
        add("业务数据备份", "err", "检查失败：%s" % e,
            "检查 data\\backups 目录权限")

    # 7 图片
    try:
        img = os.path.join(APP, "images")
        n = 0
        if os.path.isdir(img):
            for root, dirs, files in os.walk(img):
                n += len([f for f in files if f.lower().endswith((".png", ".jpg", ".jpeg", ".bmp"))])
        add("包材/成品图片", "ok" if n else "warn",
            ("已有 %d 张图片" % n) if n else "还没有任何产品图片，包材图/成品图不会出现在指令单上",
            "放到 D:\\QMDownload\\OrderSystem\\images\\<产品名>\\ 下，"
            "文件名：压片图 / 物料1…物料6 / 成品图")
    except Exception as e:
        add("图片", "warn", "检查失败：%s" % e, "")

    # 8 最近 JS 错误
    try:
        jp = os.path.join(APP, "data", "js_errors.log")
        if os.path.exists(jp):
            with open(jp, encoding="utf-8", errors="ignore") as stream:
                lines = [line for line in stream.read().splitlines() if line.strip()]
            recent = [l for l in lines[-30:] if today in l or (datetime.datetime.now() -
                      datetime.timedelta(days=1)).strftime("%Y-%m-%d") in l]
            if recent:
                kinds = _classify_frontend_errors(recent)
                detail = "最近 1 天 %d 条：接口连接失败 %d，浏览器确认弹窗被权限拦截 %d，其他脚本异常 %d。最后一条：%s" % (
                    len(recent), kinds["network"], kinds["dialog_permission"],
                    kinds["runtime"], recent[-1][:120])
                add("页面报错", "warn", detail,
                    "连接失败时先看服务是否正在重启；确认类操作已改为页面内对话框。若“其他脚本异常”大于 0，请保存日志供排查。")
            else:
                add("页面报错", "ok", "最近 1 天没有前端报错", "")
        else:
            add("页面报错", "ok", "无报错记录", "")
    except Exception as e:
        add("页面报错", "warn", "读取前端错误日志失败：%s" % e,
            "检查 data\\js_errors.log 文件是否可读")

    # 9 自动更新健康度
    try:
        sf = os.path.join(APP, "data", "last_import.json")
        lf = _read_json(sf) if os.path.exists(sf) else {}
        fail = lf.get("fail_count", 0)
        if fail >= 3:
            add("自动库存更新", "err", "连续 %d 次自动更新失败" % fail,
                "请手动导入：库存页 → 「导入最新库存表（手动）」或「强制重新导入」")
        else:
            add("自动库存更新", "ok",
                "上次导入 %s（库存表）/ %s（批次表），失败计数 %d"
                % (lf.get("time", "无"), lf.get("batch_time", "无"), fail), "")
    except Exception as e:
        add("自动库存更新", "warn", "读取自动更新状态失败：%s" % e,
            "检查 data\\last_import.json 文件")

    errs = [c for c in chk if c["level"] == "err"]
    warns = [c for c in chk if c["level"] == "warn"]
    return {"ok": True, "items": chk, "n_err": len(errs), "n_warn": len(warns),
            "time": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")}


@eel.expose
def get_replenish():
    """需要补料的物料清单（供置顶横幅）"""
    try:
        meta, rows, alert, warn = build_report()
        need = []
        overbooked = False
        for r in alert:
            if r["avail"] < 0:
                overbooked = True
            suggest = round(max(0.0, r["line"] - r["avail"]), 3)
            need.append({"short": r["short"], "avail": r["avail"], "line": r["line"],
                         "unit": r["unit"], "status": r["status"], "suggest": suggest})
        low = [{"short": r["short"], "avail": r["avail"], "line": r["line"],
                "unit": r["unit"]} for r in warn]
        return {"ok": True, "need": need, "low": low, "overbooked": overbooked,
                "stock_date": meta.get("stock_date", "-")}
    except Exception as e:
        return {"ok": False, "msg": "%s: %s" % (type(e).__name__, e)}


@eel.expose
@_serialize_stock_write
def save_product(data):
    try:
        name = (data.get("name") or "").strip()
        if not name:
            return {"ok": False, "msg": "产品名称不能为空"}
        mats = data.get("materials", [])
        if not mats:
            return {"ok": False, "msg": "至少录入一种物料"}
        normalized_mats = []
        seen_names = set()
        for material in mats:
            material_name = (material.get("name") or "").strip()
            if not material_name:
                return {"ok": False, "msg": "物料名称不能为空"}
            if material_name in seen_names:
                return {"ok": False, "msg": "物料「%s」重复，请合并为一行" % material_name}
            seen_names.add(material_name)
            ratio = float(material["ratio"])
            if not math.isfinite(ratio) or ratio <= 0:
                return {"ok": False, "msg": "物料「%s」比例必须是大于 0 的有效数字" % material_name}
            stage = int(material.get("stage", 0))
            if stage < 0 or stage > 3:
                return {"ok": False, "msg": "物料「%s」的工艺阶段无效" % material_name}
            normalized_mats.append(dict(material, name=material_name, ratio=ratio, stage=stage))
        total = sum(material["ratio"] for material in normalized_mats)
        if abs(total - 100) > 0.5:
            return {"ok": False, "msg": "配方比例合计为 %.2f%%，必须在 99.5%% 至 100.5%% 之间；未保存"
                                      % total}
        with open(RECIPE_JSON, "r", encoding="utf-8") as f:
            rec = json.load(f)
        if name in rec:
            return {"ok": False, "msg": "该产品已存在"}
        max_stage = max(m["stage"] for m in normalized_mats)
        stages = [
            {"name": "一次混装原料", "prefix": "先加", "action": "预混", "time": "10-15"},
            {"name": "二次混装原料", "prefix": "再加", "action": "混合", "time": "5"},
            {"name": "三次混装原料", "prefix": "最后加", "action": "混合", "time": "5"},
            {"name": "其他混装原料", "prefix": "再加", "action": "混合", "time": "5"},
        ][:max_stage + 1]
        pkg = data.get("packaging") or {}
        rec[name] = {"category": data.get("category", "压片糖果"),
                     "tablet_g": float(data.get("tablet_g", 0.5)),
                     "stages": stages, "materials": normalized_mats,
                     "packaging": {"per_bottle": int(pkg.get("per_bottle", 60)),
                                   "per_box": int(pkg.get("per_box", 500)),
                                   "coat": pkg.get("coat", "无"),
                                   "shape": pkg.get("shape", "圆形"),
                                   "materials": pkg.get("materials", [])}}
        save_json_atomic(RECIPE_JSON, rec)
        return {"ok": True, "msg": "已保存《%s》，出单页下拉里即可选到（比例合计 %.2f%%）"
                % (name, total), "total": round(total, 2)}
    except Exception as e:
        return {"ok": False, "msg": "%s: %s" % (type(e).__name__, e)}


def _recipe_signature(recipe):
    materials = recipe.get("materials") or []
    composition = []
    for material in materials:
        if not isinstance(material, dict):
            return None
        try:
            ratio = float(material.get("ratio"))
            stage = int(material.get("stage", 0))
        except (TypeError, ValueError):
            return None
        name = str(material.get("name") or "").strip()
        if not name or not math.isfinite(ratio):
            return None
        composition.append((name.casefold(), round(ratio, 6), stage))
    try:
        tablet = round(float(recipe.get("tablet_g")), 6)
    except (TypeError, ValueError):
        return None
    if not math.isfinite(tablet):
        return None
    packaging = recipe.get("packaging") or {}
    coat = str(packaging.get("coat") or "").strip()
    coat_material = str(packaging.get("coat_material") or "").strip().casefold()
    coat_ratio = packaging.get("coat_ratio")
    if coat_ratio not in (None, ""):
        try:
            coat_ratio = round(float(coat_ratio), 6)
        except (TypeError, ValueError):
            return None
        if not math.isfinite(coat_ratio):
            return None
    else:
        coat_ratio = None
    return tablet, tuple(sorted(composition)), coat, coat_material, coat_ratio


def _persist_manual_recipe(base_name, recipe):
    with open(RECIPE_JSON, "r", encoding="utf-8") as stream:
        catalog = json.load(stream)
    if not isinstance(catalog, dict):
        raise ValueError("产品配方库格式无效，未建立订单")
    signature = _recipe_signature(recipe)
    if signature is None:
        raise ValueError("本单配方无法生成产品版本")

    version_prefix = base_name + "（配方版本"
    versions = []
    for name in catalog:
        if name == base_name:
            versions.append((name, 1))
        elif name.startswith(version_prefix) and name.endswith("）"):
            version = name[len(version_prefix):-1]
            if version.isdigit():
                versions.append((name, int(version)))
    for name, _ in sorted(versions, key=lambda item: item[1]):
        if _recipe_signature(catalog.get(name, {})) == signature:
            return name, False

    if not versions:
        product_name = base_name
    else:
        next_version = max(version for _, version in versions) + 1
        product_name = "%s（配方版本%d）" % (base_name, next_version)
        while product_name in catalog:
            next_version += 1
            product_name = "%s（配方版本%d）" % (base_name, next_version)
    if len(product_name) > 160:
        raise ValueError("产品名称过长，无法创建配方版本（最多 160 字）")
    catalog[product_name] = recipe
    save_json_atomic(RECIPE_JSON, catalog)
    return product_name, True


def _remove_manual_recipe(product_name, expected_recipe):
    with open(RECIPE_JSON, "r", encoding="utf-8") as stream:
        catalog = json.load(stream)
    if product_name not in catalog:
        return
    if _recipe_signature(catalog[product_name]) != _recipe_signature(expected_recipe):
        raise ValueError("新产品配方已被其他操作修改，未自动删除")
    del catalog[product_name]
    save_json_atomic(RECIPE_JSON, catalog)


@eel.expose
def get_production_dashboard(include_closed=False):
    try:
        from orders import production_dashboard
        orders = production_dashboard(bool(include_closed))
        hidden = 0
        if not include_closed:
            hidden = len(production_dashboard(True)) - len(orders)
        return {"ok": True, "orders": orders, "hidden_closed": max(0, hidden)}
    except Exception as e:
        return {"ok": False, "msg": "%s: %s" % (type(e).__name__, e)}


@eel.expose
def confirm_tracking_order(serial, product):
    try:
        from orders import confirm_tracking_order as confirm
        return confirm(serial, product)
    except Exception as e:
        return {"ok": False, "msg": "%s: %s" % (type(e).__name__, e)}


def _production_sheet_number(value):
    if isinstance(value, bool) or value in (None, ""):
        return None
    try:
        number = float(value)
    except (TypeError, ValueError):
        match = re.search(r"-?\d+(?:\.\d+)?", str(value).replace(",", ""))
        if not match:
            return None
        number = float(match.group(0))
    return number if math.isfinite(number) else None


def _normalize_production_serial(value):
    """归一化手工填写的生产单号：2026102901 → 20261029-01，全角横线转半角。"""
    if isinstance(value, bool):
        return ""
    if isinstance(value, float) and value.is_integer():
        value = int(value)
    text = re.sub(r"[\s\u3000]+", "", str(value or "").strip())
    for dash in ("－", "—", "–"):
        text = text.replace(dash, "-")
    if re.fullmatch(r"\d{10}", text):
        return "%s-%s" % (text[:8], text[8:])
    return text


def _parse_production_sheet(content, order_type):
    from openpyxl import load_workbook

    if order_type not in ("正式生产", "测试单", "实验单"):
        raise ValueError("请为导入的生产表选择有效分类")
    stream = io.BytesIO(content)
    with zipfile.ZipFile(stream) as archive:
        entries = archive.infolist()
        if ("xl/workbook.xml" not in archive.namelist()
                or len(entries) > 2000
                or sum(entry.file_size for entry in entries) > 100 * 1024 * 1024):
            raise ValueError("文件不是有效的 Excel 工作簿或内容超出安全限制")
    stream.seek(0)
    workbook = load_workbook(stream, read_only=True, data_only=True)
    try:
        if SHEET not in workbook.sheetnames:
            raise ValueError("未找到系统固定生产母版工作表「%s」" % SHEET)
        sheet = workbook[SHEET]
        raw_serial = sheet["E3"].value
        serial = _normalize_production_serial(raw_serial)
        product = str(sheet["D2"].value or "").strip()
        if not re.fullmatch(r"\d{8}-\d{2}", serial):
            shown = str(raw_serial).strip() if raw_serial is not None else ""
            if len(shown) > 40:
                shown = shown[:40] + "…"
            raise ValueError(
                "生产表中的生产单号格式无效（E3 当前值：%s；应为 8 位日期-2 位序号，"
                "如 20261029-01）" % (shown or "空"))
        try:
            datetime.datetime.strptime(serial[:8], "%Y%m%d")
        except ValueError:
            raise ValueError("生产单号中的日期不存在：%s（应为真实日期，如 20261029-01）" % serial)
        if not product or len(product) > 120:
            raise ValueError("生产表中的产品名称为空或过长")

        quantity_text = str(sheet["G2"].value or "").strip()
        wan_match = re.search(r"(\d+(?:\.\d+)?)\s*万片", quantity_text)
        kg_match = re.search(r"(\d+(?:\.\d+)?)\s*(?:kg(?![a-z])|公斤)", quantity_text, re.IGNORECASE)
        wan = float(wan_match.group(1)) if wan_match else None
        total_kg = float(kg_match.group(1)) if kg_match else None
        tablet_text = str(sheet["C38"].value or "")
        tablet_match = re.search(
            r"(\d+(?:\.\d+)?)\s*g\s*/\s*片", tablet_text, re.IGNORECASE)
        tablet = float(tablet_match.group(1)) if tablet_match else None
        if wan is not None and (not math.isfinite(wan) or wan <= 0):
            raise ValueError("生产表中的计划片数无效")
        if total_kg is not None and (not math.isfinite(total_kg) or total_kg <= 0):
            raise ValueError("生产表中的计划重量无效")
        if tablet is not None and (not math.isfinite(tablet) or tablet <= 0 or tablet > 100):
            raise ValueError("生产表中的片重无效")
        if total_kg is None and wan is not None and tablet is not None:
            total_kg = round(wan * 10000 * tablet / 1000, 3)
            if total_kg <= 0:
                raise ValueError("生产表计划重量无效，未导入")
        if wan is None and total_kg is None:
            raise ValueError("生产表计划产量无法识别，未导入")

        material_map = {}
        for row in range(6, 20):
            name = str(sheet["B%d" % row].value or "").strip()
            if not name:
                continue
            kg = _production_sheet_number(sheet["H%d" % row].value)
            if kg is None or kg < 0.001:
                raise ValueError("生产表物料「%s」的用量无效" % name)
            batch = str(sheet["E%d" % row].value or "").strip()
            if name not in material_map:
                material_map[name] = {"name": name, "kg": 0.0, "batch": ""}
            entry = material_map[name]
            entry["kg"] = round(entry["kg"] + kg, 3)
            if batch and batch not in entry["batch"].split("；"):
                entry["batch"] = "；".join(
                    part for part in (entry["batch"], batch) if part)
        materials = list(material_map.values())
        if not materials:
            raise ValueError("生产表中没有可识别的混料物料行，未导入")
        if any(material["kg"] <= 0 for material in materials):
            raise ValueError("生产表物料用量精度不足，未导入")

        due_cell = sheet["J3"].value
        if isinstance(due_cell, (datetime.date, datetime.datetime)):
            due = due_cell.strftime("%Y-%m-%d")
        else:
            due_text = str(due_cell or "").strip()
            due = due_text if re.fullmatch(r"\d{4}-\d{2}-\d{2}", due_text) else ""
        return {
            "serial": serial,
            "product": product,
            "wan": wan,
            "tablet": tablet,
            "total_kg": total_kg,
            "order_no": str(sheet["B3"].value or "").strip(),
            "batch": str(sheet["H3"].value or "01").strip() or "01",
            "due": due,
            "sales": str(sheet["L3"].value or "").strip(),
            "order_date": serial[:8],
            "order_type": order_type,
            "materials": materials,
            "packaging": [],
        }
    finally:
        workbook.close()


@eel.expose
@_serialize_stock_write
def import_production_sheet(filename, encoded_content, order_type):
    """Import a canonical production workbook into the tracking ledger."""
    try:
        from orders import import_production_sheet as save_imported_sheet

        safe_name = os.path.basename(str(filename or "").replace("\\", "/"))
        if not safe_name.lower().endswith(".xlsx"):
            return {"ok": False, "msg": "请选择系统生产表 .xlsx 文件"}
        try:
            content = base64.b64decode(encoded_content, validate=True)
        except (ValueError, TypeError):
            return {"ok": False, "msg": "上传文件数据无效，请重新选择文件"}
        if not content or len(content) > 20 * 1024 * 1024:
            return {"ok": False, "msg": "文件为空或超过 20 MB，请检查后重试"}
        try:
            parsed = _parse_production_sheet(content, order_type)
            result = save_imported_sheet(
                parsed, safe_name, hashlib.sha256(content).hexdigest())
        except (zipfile.BadZipFile, OSError) as exc:
            return {"ok": False, "msg": "无法读取 Excel 生产表：%s" % exc}
        result.setdefault("warnings", [])
        return result
    except Exception as e:
        return {"ok": False, "msg": "%s: %s" % (type(e).__name__, e)}


@eel.expose
@_serialize_stock_write
def create_manual_tracking(cfg):
    try:
        from orders import ORDER_TYPES, add_order, order_exists

        product = str(cfg.get("product") or "").strip()
        if not product:
            return {"ok": False, "msg": "请输入或选择产品名称"}
        order_type = cfg.get("order_type", "正式生产")
        if order_type not in ORDER_TYPES:
            return {"ok": False, "msg": "生产记录分类无效"}
        try:
            total_kg = float(cfg.get("total_kg"))
        except (TypeError, ValueError):
            return {"ok": False, "msg": "计划投料量必须是大于 0 的数字"}
        if not math.isfinite(total_kg) or total_kg <= 0:
            return {"ok": False, "msg": "计划投料量必须是大于 0 的数字"}
        rounded_total_kg = round(total_kg, 3)
        if rounded_total_kg <= 0 or abs(total_kg - rounded_total_kg) > 1e-9:
            return {"ok": False, "msg": "计划投料量最多填写三位小数（kg）"}
        total_kg = rounded_total_kg
        try:
            raw_sequence = cfg.get("seq")
            if isinstance(raw_sequence, bool):
                raise ValueError("生产序号无效")
            date8 = datetime.datetime.strptime(
                str(cfg.get("date") or ""), "%Y-%m-%d").strftime("%Y%m%d")
            sequence = int(raw_sequence)
            if isinstance(raw_sequence, float) and not raw_sequence.is_integer():
                raise ValueError("生产序号无效")
        except (TypeError, ValueError):
            return {"ok": False, "msg": "生产日期或生产序号无效"}
        if not 1 <= sequence <= 99:
            return {"ok": False, "msg": "生产序号必须在 01-99 之间"}
        due = str(cfg.get("due") or "").strip()
        if due:
            datetime.datetime.strptime(due, "%Y-%m-%d")
        order_no = str(cfg.get("order_no") or "").strip()
        if len(order_no) > 100:
            return {"ok": False, "msg": "销售订单号最多 100 字"}

        recipe_mode = cfg.get("recipe_mode", "library")
        if recipe_mode == "library":
            recipe = load_recipes().get(product)
            if not isinstance(recipe, dict):
                return {"ok": False, "msg": "产品配方不存在；请先维护配方，或为本单录入临时配方"}
            try:
                tablet = float(recipe.get("tablet_g", 0.5))
            except (TypeError, ValueError):
                return {"ok": False, "msg": "产品配方片重无效"}
            materials = recipe.get("materials") or []
        elif recipe_mode == "custom":
            try:
                tablet = float(cfg.get("tablet_g"))
            except (TypeError, ValueError):
                return {"ok": False, "msg": "本单配方的片重必须是大于 0 的数字"}
            materials = cfg.get("materials")
            if (not math.isfinite(tablet) or tablet <= 0
                    or tablet > 100):
                return {"ok": False, "msg": "本单配方片重无效"}
            if not isinstance(materials, list) or not materials or len(materials) > 100:
                return {"ok": False, "msg": "请录入 1 至 100 项本单配方原料"}
            normalized_materials = []
            seen_names = set()
            for material in materials:
                if not isinstance(material, dict):
                    return {"ok": False, "msg": "本单配方原料格式无效"}
                material_name = str(material.get("name") or "").strip()
                try:
                    ratio = float(material.get("ratio"))
                    raw_stage = material.get("stage", 0)
                    if isinstance(raw_stage, bool):
                        raise ValueError("工艺阶段无效")
                    stage = int(raw_stage)
                except (TypeError, ValueError):
                    return {"ok": False, "msg": "原料比例或工艺阶段无效"}
                if (not material_name or len(material_name) > 120
                        or not math.isfinite(ratio) or ratio <= 0
                        or stage < 0 or stage > 3
                        or (isinstance(raw_stage, float) and not raw_stage.is_integer())):
                    return {"ok": False, "msg": "原料名称、比例或工艺阶段无效"}
                if material_name.casefold() in seen_names:
                    return {"ok": False, "msg": "本单配方中原料名称重复：%s" % material_name}
                seen_names.add(material_name.casefold())
                normalized_materials.append({
                    "name": material_name, "ratio": ratio, "stage": stage, "batch": "",
                })
            materials = normalized_materials
            packaging = {
                "per_bottle": _positive_int_or_default(
                    cfg.get("per_bottle"), 60, "每瓶片数"),
                "per_box": _positive_int_or_default(
                    cfg.get("per_box"), 500, "每箱瓶数"),
                "coat": str(cfg.get("coat") or "").strip(),
                "coat_material": str(cfg.get("coat_material") or "").strip(),
                "shape": str(cfg.get("shape") or "").strip(),
                "materials": [],
            }
            raw_coat_ratio = cfg.get("coat_ratio")
            if raw_coat_ratio in (None, ""):
                packaging["coat_ratio"] = None
            else:
                coat_ratio = float(raw_coat_ratio)
                if not math.isfinite(coat_ratio) or coat_ratio <= 0 or coat_ratio > 100:
                    return {"ok": False, "msg": "包衣比例须大于 0 且不超过 100%"}
                if not packaging["coat_material"]:
                    return {"ok": False, "msg": "填写包衣比例时须同时填写包衣材料名称"}
                packaging["coat_ratio"] = coat_ratio
            category = str(cfg.get("category") or "压片糖果").strip()
            if not category or len(category) > 80:
                return {"ok": False, "msg": "产品类别无效"}
            recipe = {
                "category": category,
                "tablet_g": tablet,
                "stages": [
                    {"name": "一次混装原料", "prefix": "先加", "action": "预混", "time": "10-15"},
                    {"name": "二次混装原料", "prefix": "再加", "action": "混合", "time": "5"},
                    {"name": "三次混装原料", "prefix": "最后加", "action": "混合", "time": "5"},
                    {"name": "其他混装原料", "prefix": "再加", "action": "混合", "time": "5"},
                ][:max(material["stage"] for material in materials) + 1],
                "materials": materials,
                "packaging": packaging,
            }
        else:
            return {"ok": False, "msg": "配方来源无效"}
        if not isinstance(materials, list):
            return {"ok": False, "msg": "产品配方原料格式无效"}
        ratios = [float(material.get("ratio", 0)) for material in materials
                  if isinstance(material, dict)]
        if (not materials or len(ratios) != len(materials)
                or not math.isfinite(tablet) or tablet <= 0
                or any(not math.isfinite(ratio) or ratio <= 0 for ratio in ratios)
                or abs(sum(ratios) - 100) > 0.5):
            return {"ok": False, "msg": "产品配方比例无效，未建立跟踪单"}

        serial = "%s-%02d" % (date8, sequence)
        preview = order_preview(
            product, "kg", total_kg, tablet, recipe_override=recipe)
        if not preview.get("ok"):
            return {"ok": False, "msg": "计算手动跟踪单用料失败：%s"
                    % preview.get("msg", "未知错误")}
        if len(preview.get("items", [])) != len(materials):
            return {"ok": False, "msg": "配方与用料计算结果不一致，未建立跟踪单"}
        new_recipe_name = None
        recipe_created = False
        if recipe_mode == "custom":
            new_recipe_name, recipe_created = _persist_manual_recipe(product, recipe)
            product = new_recipe_name
        if order_exists(serial, product):
            if recipe_created:
                try:
                    _remove_manual_recipe(new_recipe_name, recipe)
                except Exception as cleanup_error:
                    return {"ok": False, "msg":
                            "订单重复，未写入台账；新产品配方回滚失败：%s" % cleanup_error}
            return {"ok": False, "msg": "订单台账中已存在 %s（%s）" % (serial, product)}
        material_records = []
        for material, item in zip(materials, preview["items"]):
            material_records.append({
                "name": material["name"], "ratio": float(material["ratio"]),
                "kg": item["need"], "batch": item.get("alloc", ""),
                "multi": item.get("multi", False),
                "avail": item.get("avail", 0),
                "shortage": item.get("shortage", 0),
                "batch_shortage": item.get("batch_shortage", 0),
            })
        packaging = recipe.get("packaging") or {}
        coat_material = str(packaging.get("coat_material") or "").strip()
        coat_ratio = packaging.get("coat_ratio")
        if coat_ratio not in (None, ""):
            coat_ratio = float(coat_ratio)
            if (not coat_material or not math.isfinite(coat_ratio)
                    or coat_ratio <= 0 or coat_ratio > 100):
                if recipe_mode == "library":
                    return {"ok": False, "msg": "产品配方中的包衣材料或比例无效"}
            else:
                coat_need = round(total_kg * coat_ratio / 100, 3)
                coat_recipe = {
                    "tablet_g": tablet,
                    "materials": [{"name": coat_material, "ratio": 100}],
                    "packaging": {},
                }
                coat_preview = order_preview(
                    product, "kg", coat_need, tablet,
                    recipe_override=coat_recipe)
                if not coat_preview.get("ok") or len(coat_preview.get("items", [])) != 1:
                    return {"ok": False, "msg": "计算包衣材料用量失败：%s"
                            % coat_preview.get("msg", "物料结果不一致")}
                coat_item = coat_preview["items"][0]
                material_records.append({
                    "name": coat_material, "ratio": coat_ratio, "kg": coat_item["need"],
                    "batch": coat_item.get("alloc", ""),
                    "multi": coat_item.get("multi", False),
                    "avail": coat_item.get("avail", 0),
                    "shortage": coat_item.get("shortage", 0),
                    "batch_shortage": coat_item.get("batch_shortage", 0),
                    "kind": "coating", "basis": "素片重量",
                })
        order = {
            "serial": serial, "product": product,
            "wan": round(total_kg * 1000 / tablet / 10000, 4),
            "tablet": tablet, "total_kg": round(total_kg, 3),
            "customer": "", "due": due, "sales": str(cfg.get("manager") or ""),
            "status": "待生产", "order_date": date8,
            "batch": str(cfg.get("batch") or "01"),
            "order_no": order_no, "order_type": order_type,
            "materials": material_records,
            "packaging": [], "manual_tracking": True,
            "recipe_source": "产品库配方" if recipe_mode == "library" else "新建产品配方",
            "product_base_name": str(cfg.get("product") or "").strip()
                                 if recipe_mode == "custom" and product != cfg.get("product")
                                 else "",
            "tracking_confirmed_at": datetime.datetime.now().isoformat(timespec="seconds"),
            "tracking_confirmation_source": "手动新增跟踪单",
            "tracking_history": [{
                "at": datetime.datetime.now().isoformat(timespec="seconds"),
                "source": "手动新增跟踪单",
            }],
            "note": "手动跟踪单（未生成 Excel）",
        }
        try:
            added, message = add_order(order)
        except Exception as add_error:
            if recipe_created:
                try:
                    _remove_manual_recipe(new_recipe_name, recipe)
                except Exception as cleanup_error:
                    return {"ok": False, "msg":
                            "订单写入失败，且新产品配方回滚失败：%s；原错误：%s"
                            % (cleanup_error, add_error)}
            raise
        if not added:
            if recipe_created:
                try:
                    _remove_manual_recipe(new_recipe_name, recipe)
                except Exception as cleanup_error:
                    return {"ok": False, "msg":
                            "订单未写入，且新产品配方回滚失败：%s" % cleanup_error}
            return {"ok": False, "msg": message}
        return {"ok": True, "msg": message, "order": order,
                "created_product": recipe_created,
                "product_name": product}
    except Exception as e:
        return {"ok": False, "msg": "%s: %s" % (type(e).__name__, e)}


@eel.expose
def update_production_step(serial, product, step, progress=None, qc=None,
                           operator=None, note=None):
    try:
        from orders import update_production_step as update_step
        return update_step(serial, product, step, progress, qc, operator, note)
    except Exception as e:
        return {"ok": False, "msg": "%s: %s" % (type(e).__name__, e)}


@eel.expose
def update_order_status(serial, product, status):
    try:
        from orders import update_order_status as update_status
        return update_status(serial, product, status)
    except Exception as e:
        return {"ok": False, "msg": "%s: %s" % (type(e).__name__, e)}


@eel.expose
def update_order_info(serial, product, new_product=None, total_kg=None, note=""):
    try:
        from orders import update_order_info as update_info
        return update_info(serial, product, new_product, total_kg, note)
    except Exception as e:
        return {"ok": False, "msg": "%s: %s" % (type(e).__name__, e)}


@eel.expose
def archive_or_void_order(serial, product, reason):
    try:
        from orders import archive_or_void_order as archive
        return archive(serial, product, reason)
    except Exception as e:
        return {"ok": False, "msg": "%s: %s" % (type(e).__name__, e)}


@eel.expose
def delete_unstarted_order(serial, product):
    try:
        from orders import delete_unstarted_order as delete
        return delete(serial, product)
    except Exception as e:
        return {"ok": False, "msg": "%s: %s" % (type(e).__name__, e)}


@eel.expose
def delete_test_order(serial, product):
    try:
        from orders import delete_test_order as delete
        return delete(serial, product)
    except Exception as e:
        return {"ok": False, "msg": "%s: %s" % (type(e).__name__, e)}


@eel.expose
def create_backup():
    try:
        from backup_manager import create_backup as create
        return create()
    except Exception as e:
        return {"ok": False, "msg": "%s: %s" % (type(e).__name__, e)}


@eel.expose
def list_backups():
    try:
        from backup_manager import list_backups as list_all
        return {"ok": True, "backups": list_all()}
    except Exception as e:
        return {"ok": False, "msg": "%s: %s" % (type(e).__name__, e)}


@eel.expose
def restore_backup(path):
    try:
        from backup_manager import restore_backup as restore
        return restore(path)
    except Exception as e:
        return {"ok": False, "msg": "%s: %s" % (type(e).__name__, e)}


@eel.expose
def open_inbox():
    try:
        inbox = os.path.join(APP, "inbox")
        os.makedirs(inbox, exist_ok=True)
        os.startfile(inbox)
        return {"ok": True}
    except Exception as e:
        return {"ok": False, "msg": str(e)}


_BLOCKED_OPEN_EXT = {".exe", ".bat", ".cmd", ".com", ".ps1", ".psm1", ".vbs", ".vbe", ".js", ".jse",
                     ".wsf", ".msi", ".lnk", ".scr", ".reg", ".hta", ".jar", ".dll", ".py", ".pyw"}


@eel.expose
def open_path(path):
    try:
        p = (path or "").strip()
        if p and os.path.exists(p):
            if os.path.isfile(p) and os.path.splitext(p)[1].lower() in _BLOCKED_OPEN_EXT:
                return {"ok": False, "msg": "出于安全考虑，不能通过系统打开可执行/脚本文件"}
            os.startfile(p)
            return {"ok": True}
        return {"ok": False, "msg": "路径不存在：%s" % p}
    except Exception as e:
        return {"ok": False, "msg": str(e)}


@eel.expose
def import_sales_form_upload(filename, encoded_content, plan_qty=None, plan_unit="kg",
                             serial="", form_json=""):
    """Import a browser-uploaded DOCX without exposing an arbitrary filesystem path.
    form_json：AI 解析并经人工确认后的结构化配方（JSON 字符串），提供时替代 Word 解析。"""
    try:
        safe_name = os.path.basename(str(filename or "").replace("\\", "/"))
        if not safe_name.lower().endswith(".docx"):
            return {"ok": False, "msg": "请选择 Word .docx 文件"}
        try:
            content = base64.b64decode(encoded_content, validate=True)
        except (ValueError, TypeError):
            return {"ok": False, "msg": "上传文件数据无效，请重新选择文件"}
        if not content or len(content) > 20 * 1024 * 1024:
            return {"ok": False, "msg": "文件为空或超过 20 MB，请检查后重试"}
        options = {}
        if plan_unit not in ("kg", "bottle", "tablet"):
            return {"ok": False, "msg": "计划单位只能是 kg、瓶或片"}
        if plan_qty not in (None, ""):
            options["plan_qty"] = float(plan_qty)
            options["plan_unit"] = plan_unit
        options["serial"] = str(serial or "").strip()
        if str(form_json or "").strip():
            try:
                override = json.loads(form_json)
            except ValueError:
                return {"ok": False, "msg": "AI 解析结果数据无效，请重新解析"}
            from sales_form_import import validate_ai_form
            stock_names = _stock_material_names()
            payload = override.get("form") if isinstance(override, dict) \
                and override.get("form") else override
            problems, _warnings, _unknown = validate_ai_form(payload, stock_names)
            if problems:
                return {"ok": False, "msg": "AI 解析结果未通过校验：%s" %
                        "；".join(problems)}
            options["form_override"] = override
        template = TEMPLATE
        if not os.path.isfile(template):
            return {"ok": False, "msg": "系统模板不存在：%s" % template}
        from sales_form_import import import_sales_form
        _, current_stock_rows, _, _ = build_report()
        with tempfile.TemporaryDirectory(prefix="sales-form-") as directory:
            source = os.path.join(directory, safe_name)
            with open(source, "wb") as stream:
                stream.write(content)
            result = import_sales_form(source, template, options=options,
                                       stock_rows=current_stock_rows)
        result["folder"] = os.path.dirname(result["output"])
        return {"ok": True, **result}
    except Exception as e:
        return {"ok": False, "msg": "%s: %s" % (type(e).__name__, e)}


@eel.expose
def ai_import_text_order(text_json, plan_qty=None, plan_unit="kg", serial="",
                         order_date=""):
    """文字订单（AI 提取并经人工确认）→ 直接按模板生成交接单并登记台账。"""
    try:
        try:
            payload = json.loads(text_json) if str(text_json or "").strip() else None
        except ValueError:
            return {"ok": False, "msg": "AI 提取数据无效，请重新提取"}
        if not payload:
            return {"ok": False, "msg": "请先完成 AI 提取并核对"}
        options = {"form_override": payload}
        if plan_unit not in ("kg", "bottle", "tablet"):
            return {"ok": False, "msg": "计划单位只能是 kg、瓶或片"}
        if plan_qty not in (None, ""):
            options["plan_qty"] = float(plan_qty)
            options["plan_unit"] = plan_unit
        options["serial"] = str(serial or "").strip()
        if str(order_date or "").strip():
            options["date"] = str(order_date)
        template = TEMPLATE
        if not os.path.isfile(template):
            return {"ok": False, "msg": "系统模板不存在：%s" % template}
        from sales_form_import import import_sales_form
        _, current_stock_rows, _, _ = build_report()
        virtual_name = "文字订单_%s.docx" % (str(order_date) or
                                             datetime.datetime.now().strftime("%Y%m%d"))
        result = import_sales_form(virtual_name, template, options=options,
                                   stock_rows=current_stock_rows)
        result["folder"] = os.path.dirname(result["output"])
        name_match = re.match(r"(\d{8})-(\d{2})_", os.path.basename(result["output"]))
        if name_match:
            result["serial_display"] = "%s-%s" % (name_match.group(1), name_match.group(2))
        import ai_assistant
        ai_assistant.set_progress("完成", "交接单已生成", done=True)
        return {"ok": True, **result}
    except Exception as e:
        return {"ok": False, "msg": "%s: %s" % (type(e).__name__, e)}


@eel.expose
def preview_manual_recipe_upload(filename, encoded_content):
    """Parse a DOCX recipe into the manual tracking form without creating a document."""
    try:
        safe_name = os.path.basename(str(filename or "").replace("\\", "/"))
        if not safe_name.lower().endswith(".docx"):
            return {"ok": False, "msg": "请选择 Word .docx 文件"}
        try:
            content = base64.b64decode(encoded_content, validate=True)
        except (ValueError, TypeError):
            return {"ok": False, "msg": "上传文件数据无效，请重新选择文件"}
        if not content or len(content) > 20 * 1024 * 1024:
            return {"ok": False, "msg": "文件为空或超过 20 MB，请检查后重试"}
        from sales_form_import import _stage_for, parse_sales_form
        warnings = []
        with tempfile.TemporaryDirectory(prefix="manual-recipe-") as directory:
            source = os.path.join(directory, safe_name)
            with open(source, "wb") as stream:
                stream.write(content)
            form = parse_sales_form(source)
        materials = []
        for material in form["materials"]:
            try:
                stage = _stage_for(material["name"], form.get("press_process", ""))
            except ValueError:
                stage = 0
                warnings.append("未能从工艺说明识别「%s」的阶段，请核对后再建立产品"
                                % material["name"])
            materials.append({
                "name": material["name"], "ratio": material["ratio"], "stage": stage,
            })
        if not form.get("coat_type"):
            warnings.append("文件中未识别包衣类型；请按已确认资料选择或留空")
        return {
            "ok": True,
            "recipe": {
                "product": form["product"],
                "tablet_g": form["tablet_g"],
                "materials": materials,
                "coat": form.get("coat_type") or "",
                "coat_material": form.get("coat_name") or "",
                "coat_ratio": form.get("coat_ratio"),
                "per_bottle": form.get("per_bottle"),
                "category": "压片糖果",
            },
            "warnings": warnings,
        }
    except Exception as e:
        return {"ok": False, "msg": "%s: %s" % (type(e).__name__, e)}


@eel.expose
def open_report():
    try:
        cdir = os.path.join(APP, "data", "compare")
        if os.path.isdir(cdir):
            files = sorted([f for f in os.listdir(cdir) if f.endswith(".md")])
            if files:
                os.startfile(os.path.join(cdir, files[-1]))
                return {"ok": True}
        return {"ok": False, "msg": "暂无比对报告"}
    except Exception as e:
        return {"ok": False, "msg": str(e)}


if __name__ == "__main__":
    eel.init(os.path.join(APP, "webui"))
    eel.start("index.html", size=(1080, 820), port=8765, block=True)
