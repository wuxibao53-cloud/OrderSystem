# -*- coding: utf-8 -*-
"""台账问答助手：只读工具 + 简化 ReAct 循环。
AI 通过 JSON 协议调用本地只读工具（查台账/库存/批次/快照），最后给出中文回答。
铁律：所有工具只读；工具结果超限自动截断；最多 6 轮；未配置 AI 时直接报错降级。"""
import csv
import datetime
import json
import os

import ai_assistant

APP = os.path.dirname(os.path.abspath(__file__))
SNAPSHOTS = os.path.join(APP, "data", "snapshots")
MAX_ROUNDS = 6
MAX_TOOL_CHARS = 6000


# ---------------- 只读工具 ----------------

def _orders_data():
    from orders import load as load_orders
    return load_orders().get("orders", [])


def tool_query_orders(args):
    status = str(args.get("status") or "").strip()
    keyword = str(args.get("keyword") or "").strip().lower()
    out = []
    for o in _orders_data():
        if status and o.get("status") != status:
            continue
        text = "%s %s" % (o.get("serial"), o.get("product"))
        if keyword and keyword not in text.lower():
            continue
        runs = o.get("production_runs") or []
        out.append({
            "serial": o.get("serial"), "product": o.get("product"),
            "status": o.get("status"), "order_type": o.get("order_type"),
            "total_kg": o.get("total_kg"), "due": o.get("due"),
            "sales": o.get("sales"),
            "runs_done": sum(1 for r in runs if r.get("status") == "已完成"),
            "runs_total": len(runs),
        })
    return out[:25]


def tool_order_detail(args):
    serial = str(args.get("serial") or "").strip()
    matches = [o for o in _orders_data() if o.get("serial") == serial]
    if len(matches) != 1:
        return {"error": "订单 %s 不存在或匹配不唯一" % serial}
    o = matches[0]
    runs = []
    for run in o.get("production_runs") or []:
        runs.append({
            "run_no": run.get("run_no"), "run_date": run.get("run_date"),
            "plan_kg": run.get("plan_kg"), "status": run.get("status", ""),
            "file": run.get("file"), "note": run.get("note"),
            "materials": run.get("materials") or [],
        })
    return {
        "serial": o.get("serial"), "product": o.get("product"),
        "status": o.get("status"), "total_kg": o.get("total_kg"),
        "batch": o.get("batch"), "due": o.get("due"), "sales": o.get("sales"),
        "production_runs": runs,
        "pending_materials": o.get("materials") or [],
        "tracking_history_tail": (o.get("tracking_history") or [])[-8:],
    }


def tool_query_stock(args):
    from stock import build_report
    keyword = str(args.get("keyword") or "").strip().lower()
    _meta, rows, _alert, _warn = build_report()
    out = []
    for r in rows:
        text = "%s %s %s" % (r.get("short"), r.get("name"), r.get("cat"))
        if keyword and keyword not in text.lower():
            continue
        batches = r.get("batches") or []
        batch_total = round(sum(float(b.get("qty") or 0) for b in batches), 3)
        qty = float(r.get("qty") or 0)
        entry = {"short": r.get("short"), "cat": r.get("cat"),
                 "qty": r.get("qty"), "unit": r.get("unit"),
                 "occupied": r.get("occupied"), "avail": r.get("avail"),
                 "batches": [{"batch": b.get("batch"), "qty": b.get("qty"),
                              "in_date": b.get("in_date")} for b in batches[:8]]}
        if batches and abs(batch_total - qty) > max(0.01, abs(qty) * 0.001):
            entry["batch_total"] = batch_total
            entry["data_quality_note"] = (
                "批次合计 %.3f 与账面 %.3f 不一致（差 %+.3f），该物料批号数据待核对"
                % (batch_total, qty, batch_total - qty))
        out.append(entry)
        if len(out) >= 20:
            break
    return out


def tool_replenish(_args):
    from stock import build_report
    _meta, _rows, alert, warn = build_report()
    return {
        "need_purchase": [{"short": r["short"], "avail": r["avail"],
                           "line": r["line"], "unit": r["unit"]}
                          for r in alert],
        "low": [{"short": r["short"], "avail": r["avail"],
                 "line": r["line"], "unit": r["unit"]} for r in warn],
    }


def tool_batch_audit(args):
    from app_gui import _batch_audit_rows
    data = json.load(open(os.path.join(APP, "stock.json"), encoding="utf-8"))
    missing, mismatch = _batch_audit_rows(data)
    kind = str(args.get("kind") or "all")
    result = {}
    if kind in ("missing", "all"):
        result["missing_batches"] = missing[:20]
    if kind in ("mismatch", "all"):
        result["mismatch"] = mismatch[:20]
    return result


def tool_consumption_history(args):
    """从每日库存快照 CSV 取某物料近 N 天数量变化。"""
    short = str(args.get("short") or "").strip()
    days = int(args.get("days") or 30)
    if not short:
        return {"error": "需要 short 参数（物料名）"}
    csv_path = os.path.join(SNAPSHOTS, "stock_history.csv")
    if not os.path.exists(csv_path):
        return {"error": "还没有库存快照数据（每天 18:00 自动积累）"}
    days = max(1, min(days, 180))
    rows = []
    with open(csv_path, encoding="utf-8", newline="") as fp:
        for row in csv.reader(fp):
            if len(row) >= 4 and row[1] == short:
                rows.append({"date": row[0], "qty": row[2], "unit": row[3]})
    return {"short": short, "history": rows[-days:]}


def tool_query_recipe(args):
    """查产品库配方（原料配比/片重/每瓶），用于换算「一批 X kg 产品需要多少某原料」。"""
    keyword = str(args.get("keyword") or "").strip().lower()
    path = os.path.join(APP, "recipes.json")
    if not os.path.exists(path):
        return {"error": "产品库为空"}
    with open(path, encoding="utf-8") as fp:
        data = json.load(fp)
    out = []
    for name, recipe in data.items():
        if str(name).startswith("_"):
            continue
        if keyword and keyword not in name.lower():
            continue
        out.append({
            "product": name,
            "tablet_g": recipe.get("tablet_g"),
            "per_bottle": (recipe.get("packaging") or {}).get("per_bottle"),
            "coat": (recipe.get("packaging") or {}).get("coat"),
            "materials": [{"name": m.get("name"), "ratio": m.get("ratio")}
                          for m in recipe.get("materials") or []],
        })
        if len(out) >= 5:
            break
    return out


TOOLS = {
    "query_orders": tool_query_orders,
    "order_detail": tool_order_detail,
    "query_stock": tool_query_stock,
    "query_recipe": tool_query_recipe,
    "replenish": tool_replenish,
    "batch_audit": tool_batch_audit,
    "consumption_history": tool_consumption_history,
}

TOOL_DESCRIPTIONS = """- query_orders {keyword?, status?}：查生产单列表（编号/产品/状态/累计kg/交期/批次数）
- order_detail {serial}：某生产单详情（各批次记录、物料批号预占、跟踪历史）
- query_stock {keyword?}：查库存（账面/占用/可用/警戒线/批次号与数量）
- query_recipe {keyword}：查产品库配方（各原料配比%、片重、每瓶片数）
- replenish {}：需采购与偏低物料清单
- batch_audit {kind: missing|mismatch|all}：无批号物料 / 批次合计与账面不符清单
- consumption_history {short, days?}：某物料近 N 天每日数量快照（看消耗趋势）"""

SYSTEM_PROMPT = """你是食品厂生产系统的台账问答助手。回答规则：
1) 一次规划一个工具调用：只输出一个 JSON 对象 {"tool": "工具名", "args": {...}}，不要解释；
2) 可用工具：\n%s\n
3) 信息足够时输出 {"answer": "最终中文回答"}：直接给结论和数字，注明编号出处，控制在 200 字以内；数据里没有的就说不确定，禁止编造；
4) 数量推算（如"还够做几批"）用可用量÷单位需求向下取整，并写出算式；
5) 工具返回里带 data_quality_note 或涉及批次与账面不一致时，必须在回答末尾提醒用户先核对批次数据；
6) 「还够做几批 X kg」里的 X 指产品批量，不是原料量：必须先 query_recipe 查该产品配比，按 可用量÷(X×配比%%) 向下取整并写出算式；query_recipe 里没有该产品（或名字只是相近）时，改用 order_detail 查该产品最近批次的实际投料配比，并在回答开头声明「产品库无此配方，按某单实际投料/最接近配方 X 估算」；
7) 你只能读数据，不能建议任何绕过系统的操作。""" % TOOL_DESCRIPTIONS


def _clip(result):
    text = json.dumps(result, ensure_ascii=False)
    if len(text) > MAX_TOOL_CHARS:
        return {"clipped": True, "data": text[:MAX_TOOL_CHARS]}
    return result


def ask(question):
    """问答主入口。返回 {ok, answer, rounds, tools_used}。"""
    import ai_assistant
    if not ai_assistant.is_configured():
        return {"ok": False, "msg": "AI 未配置（data\\ai_config.json）"}
    question = str(question or "").strip()
    if not question:
        return {"ok": False, "msg": "请输入问题"}
    if len(question) > 500:
        question = question[:500]
    messages = [{"role": "system", "content": SYSTEM_PROMPT},
                {"role": "user", "content": question}]
    tools_used = []
    try:
        for round_no in range(1, MAX_ROUNDS + 1):
            ai_assistant.set_progress("检索台账", "第 %d/%d 轮工具调用…" % (round_no, MAX_ROUNDS))
            reply = ai_assistant.chat(messages)
            data = ai_assistant._extract_json(reply)
            if "answer" in data and not data.get("tool"):
                return {"ok": True, "answer": str(data["answer"]),
                        "rounds": round_no, "tools_used": tools_used}
            tool = str(data.get("tool") or "").strip()
            if tool not in TOOLS:
                messages.append({"role": "assistant", "content": reply[:500]})
                messages.append({"role": "user", "content":
                                 "没有 %s 这个工具。可用工具：\n%s\n请重新只输出一个 JSON。"
                                 % (tool, TOOL_DESCRIPTIONS)})
                continue
            args = data.get("args") if isinstance(data.get("args"), dict) else {}
            try:
                result = _clip(TOOLS[tool](args))
            except Exception as exc:
                result = {"error": "%s: %s" % (type(exc).__name__, exc)}
            tools_used.append({"tool": tool, "args": args, "round": round_no})
            messages.append({"role": "assistant", "content": reply[:500]})
            messages.append({"role": "user", "content":
                             "工具 %s 返回：\n%s\n\n继续：信息足够就输出 "
                             '{"answer": "..."}，还需要就再调用下一个工具。'
                             % (tool, json.dumps(result, ensure_ascii=False))})
        return {"ok": True, "answer": "查询轮次已达上限，未能得出结论；请把问题拆小一点再问。",
                "rounds": MAX_ROUNDS, "tools_used": tools_used}
    except ai_assistant.AiError as exc:
        text = str(exc)
        if "JSON" in text or "回复" in text:
            return {"ok": False, "msg": "AI 回复格式异常：%s" % text}
        return {"ok": False, "msg": "AI 接口异常：%s" % text}
    except ValueError as exc:
        return {"ok": False, "msg": "AI 回复格式异常：%s" % exc}
