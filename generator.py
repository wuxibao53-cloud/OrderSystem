# -*- coding: utf-8 -*-
"""生产指令与交接单生成器（适配新模板：混料/压片/包装三段 + 图片 + 包材）"""
import os
import sys
import json
import datetime
import tempfile
from copy import copy

from openpyxl import load_workbook

APP = os.path.dirname(os.path.abspath(__file__))   # 便携：系统放哪个目录都能跑
RECIPE_FILE = os.path.join(APP, "recipes.json")
DEFAULT_DESKTOP = os.path.join(os.path.expanduser("~"), "Desktop")


TPL_NAME = "生产指令与交接单.xlsx"


def _template_path():
    """Use the same bundled master for every production-document workflow."""
    return os.path.join(APP, "templates", TPL_NAME)


TEMPLATE = _template_path()
SHEET = "批生产指令与交接单"
IMG_DIR = os.path.join(APP, "images")
MANAGERS = ["曾帅", "付豪", "房浩绅", "马士龙"]


def load_recipes():
    with open(RECIPE_FILE, "r", encoding="utf-8") as f:
        data = json.load(f)
    data.pop("_说明", None)
    for k in list(data.keys()):
        if isinstance(data[k], dict):
            data[k].pop("_备注", None)
    return data


def put(ws, addr, value):
    cell = ws[addr]
    for rng in ws.merged_cells.ranges:
        if cell.coordinate in rng:
            if (rng.min_row, rng.min_col) != (cell.row, cell.column):
                return False
            break
    cell.value = value
    return True


def put_img(ws, addr, path, w=88, h=66):
    if not (path and os.path.exists(path)):
        return False
    try:
        from openpyxl.drawing.image import Image as XLImage
        img = XLImage(path)
        img.width, img.height = w, h
        ws.add_image(img, addr)
        return True
    except Exception:
        return False


def find_img(product, key):
    """images/<产品>/key.* → 返回第一个匹配的图片文件路径"""
    d = os.path.join(IMG_DIR, product)
    if not os.path.isdir(d):
        return None
    for fn in sorted(os.listdir(d)):
        stem, ext = os.path.splitext(fn)
        if stem == key and ext.lower() in (".png", ".jpg", ".jpeg", ".bmp", ".gif"):
            return os.path.join(d, fn)
    return None


def mix_note(stage, mats):
    names = "、".join(m["name"] for m in mats)
    return "%s%s%s%smin" % (stage.get("prefix", ""), names,
                            stage.get("action", "混合"), stage.get("time", "5"))


def stock_rows():
    """Read the inventory report required for safe batch allocation."""
    from stock import build_report
    meta, rows, alert, warn = build_report()
    stock_rows.meta = meta
    return rows
stock_rows.meta = {}


def lookup_row(rows, name):
    """精确 → 模糊匹配库存行"""
    if not name:
        return None
    for r in rows:
        if r.get("short") == name:
            return r
    nl = name.lower()
    for r in rows:
        sl = (r.get("short") or "").lower()
        if sl and (nl in sl or sl in nl):
            return r
    return None


def default_batch_of(row):
    """库存里该物料按有效期先进先出的批号（默认批号）"""
    if not row or not row.get("batches"):
        return (row.get("batch") if row else "") or ""
    from stock import allocate
    plan = allocate({"batches": row["batches"]}, 1e9,
                    occupied=row.get("occupied", 0), book_qty=row.get("qty"))
    positive = [batch for batch, quantity in plan if quantity > 0]
    return positive[0] if positive else ""


def plan_alloc(row, need_kg, manual_batch=""):
    """返回 [(批号, 用量)]：单批够 → 一行；不够 → 按有效期先进先出拆到多个批次。
    手工填了批号则以其为准。"""
    if need_kg <= 0:
        return []
    if manual_batch:
        return [(manual_batch, round(need_kg, 3))]
    if not row or not row.get("batches"):
        b = (row.get("batch") if row else "") or ""
        return [(b, round(need_kg, 3))]
    from stock import allocate
    plan = allocate({"batches": row["batches"]}, need_kg,
                    occupied=row.get("occupied", 0), book_qty=row.get("qty"))
    plan = [(b, q) for b, q in plan if q > 0]
    if not plan:
        return [] if row.get("batches") else [("", round(need_kg, 3))]
    return plan


def fmt_due(due):
    """交货日期只体现 月.日（如 10.30），与生产日期格式一致"""
    d = (due or "").strip()
    if len(d) == 8 and d.isdigit():
        return "%d.%02d" % (int(d[4:6]), int(d[6:8]))
    return d


def _safe_name_part(value):
    value = str(value or "").strip()
    value = "".join("_" if ord(char) < 32 or char in '<>:"/\\|?*' else char
                    for char in value)
    return value.rstrip(" .")[:100] or "未命名"


def generate(cfg, out_root=None):
    recipes = load_recipes()
    product = cfg["product"]
    if product not in recipes:
        return {"ok": False, "msg": "配方库中没有产品：%s" % product}

    r = recipes[product]
    tablet = float(cfg.get("tablet_g") or r.get("tablet_g", 0.5))
    wan = float(cfg.get("wan") or 0)
    bottles_txt = cfg.get("bottles", "")
    date8 = cfg.get("date") or datetime.datetime.now().strftime("%Y%m%d")
    seq = int(cfg.get("seq") or 1)
    manager = cfg.get("manager", "")
    maker = cfg.get("maker") or "李浩冉"
    batch_no = cfg.get("batch", "01")
    due = (cfg.get("due") or "").strip()
    stock_in = cfg.get("stock") or {}
    pk = dict(r.get("packaging", {}))
    pack_direct = False
    if "pack_materials" in cfg:
        # 用户本次手动录入（可为空数组 = 包装段留空）；数量即总套数
        pk["materials"] = cfg["pack_materials"]
        pack_direct = True
    elif cfg.get("pack_materials") is None and r.get("packaging", {}).get("materials"):
        pk["materials"] = r["packaging"]["materials"]
    if cfg.get("coat"):
        pk["coat"] = cfg["coat"]
    if cfg.get("shape"):
        pk["shape"] = cfg["shape"]
    if len(pk.get("materials", [])) > 6:
        return {"ok": False, "msg": "包材超过模板包装区的 6 行上限；未生成文件"}
    bno_map = cfg.get("batch_override") or {}
    per_bottle = int(cfg["per_bottle"]) if cfg.get("per_bottle") else int(pk.get("per_bottle", 60))
    per_box = int(pk.get("per_box", 500))

    # 物料：优先用界面调整后的（增/删/改比例/改阶段），否则用配方库
    mats = cfg.get("materials") or r["materials"]
    norm_mats = []
    for m in mats:
        st = int(m.get("stage", 0))
        if st < 0 or st > 3:
            return {"ok": False, "msg": "物料「%s」的工艺阶段无效" % m.get("name", "")}
        norm_mats.append({"name": m["name"], "ratio": float(m["ratio"]),
                          "stage": st, "batch": m.get("batch", "")})
    max_stage = max([m["stage"] for m in norm_mats] + [0])
    stages_def = list(r.get("stages", []))
    _st_names = ["一次混装原料", "二次混装原料", "三次混装原料", "其他混装原料"]
    while len(stages_def) <= max_stage:
        idx = len(stages_def)
        stages_def.append({"name": _st_names[idx] if idx < len(_st_names) else "其他混装原料",
                           "prefix": "先加" if idx == 0 else "再加", "action": "混合", "time": "5"})

    if cfg.get("total_kg"):
        # 按总量模式：直接给定总量 kg，反算片数
        total_kg = float(cfg["total_kg"])
        wan = round(total_kg * 1000 / tablet / 10000, 4)
        if total_kg <= 0:
            return {"ok": False, "msg": "总量必须大于 0"}
    else:
        wan = float(cfg.get("wan") or 0)
        if wan <= 0:
            return {"ok": False, "msg": "计划产量必须大于 0"}
        total_kg = wan * 10000 * tablet / 1000
    order_type = cfg.get("order_type", "正式生产")
    if order_type not in ("正式生产", "测试单", "实验单"):
        return {"ok": False, "msg": "生产记录分类无效"}
    serial = "%s-%02d" % (date8, seq)
    n_bottles = int(round(wan * 10000 / per_bottle)) if per_bottle else 0
    n_boxes = int(round(n_bottles / per_box)) if per_box else 0

    scale, limiting = 1.0, None
    for m in norm_mats:
        s = stock_in.get(m["name"])
        if s:
            need = total_kg * m["ratio"] / 100
            if need > 0 and s < need:
                k = s / need
                if k < scale:
                    scale, limiting = k, m["name"]
    actual_kg = total_kg * scale
    actual_wan = wan * scale

    out_root = out_root or DEFAULT_DESKTOP
    month = "%d月" % int(date8[4:6])
    quantity_name = ("%gkg" % total_kg) if cfg.get("total_kg") else ("%g万片" % wan)
    document_name = "%s_%s_%s_%s" % (
        _safe_name_part(serial), _safe_name_part(order_type),
        _safe_name_part(product), _safe_name_part(quantity_name))
    folder = os.path.join(out_root, month, date8, document_name)
    out = os.path.join(folder, document_name + ".xlsx")
    if os.path.exists(out):
        return {"ok": False, "msg": "生产指令单已存在，未覆盖：%s。请检查生产序号后再生成。" % out}

    wb = load_workbook(TEMPLATE)
    ws = wb[SHEET]
    bottle_part = str(bottles_txt or "").strip() or ("%d" % n_bottles if n_bottles else "")
    if bottle_part:
        qty_text = "%gkg（%s瓶）" % (total_kg, bottle_part)
    else:
        qty_text = "%gkg" % total_kg
    due_fmt = fmt_due(due)
    short_date = "%d.%02d" % (int(date8[4:6]), int(date8[6:8]))

    # ---- 四段表头 ----
    for order_cell, date_cell, product_cell, qty_cell, maker_cell, serial_cell, batch_cell, due_cell, manager_cell in (
            ("B3", "B2", "D2", "G2", "L2", "E3", "H3", "J3", "L3"),
            ("B34", "B33", "D33", "G33", "L33", "E34", "H34", "J34", "L34"),
            ("B44", "B43", "D43", "G43", "L43", "E44", "H44", "J44", "L44"),
            ("B62", "B61", "D61", "G61", "L61", "E62", "H62", "J62", "L62")):
        put(ws, order_cell, cfg.get("order_no") or None)
        put(ws, date_cell, short_date)
        put(ws, product_cell, product)
        put(ws, qty_cell, qty_text)
        put(ws, maker_cell, maker)
        put(ws, serial_cell, serial)
        put(ws, batch_cell, batch_no)
        put(ws, due_cell, due_fmt or None)
        put(ws, manager_cell, manager)

    # ---- BOM（一次 row6-13 / 二次 row14-16 / 三次+其他 row17-19；批号E 用量kg H）----
    # 行区上限，绝不越到 row20（领料批号核对区）
    SEG = {0: (6, 13), 1: (14, 16)}
    TAIL = (17, 19)
    srows = stock_rows()          # 批号与批次分配一律以库存为准
    groups = {}
    for m in norm_mats:
        groups.setdefault(m["stage"], []).append(m)
    cursors = {0: SEG[0][0], 1: SEG[1][0]}
    tail_cur = TAIL[0]
    notes = {}
    overflow = []
    for si in sorted(groups):
        if si in SEG:
            lo, hi = SEG[si]
            cursor = cursors[si]
        else:
            lo, hi = TAIL
            cursor = tail_cur
        mats = groups[si]
        stage = stages_def[si] if si < len(stages_def) else {"name": "混装原料"}
        for j, m in enumerate(mats):
            if j == 0 and cursor <= hi:
                put(ws, "A%d" % cursor, stage["name"])
            need_kg = round(actual_kg * m["ratio"] / 100, 3)
            manual = bno_map.get(m["name"], m.get("batch", ""))
            stock_row = lookup_row(srows, m["name"])
            plan = plan_alloc(stock_row, need_kg, manual)
            batch_shortage = 0.0
            if stock_row and stock_row.get("batches") and not manual:
                allocated = sum(qty for _, qty in plan if qty > 0)
                batch_shortage = round(max(0.0, need_kg - allocated), 3)
            if cursor > hi:
                overflow.append(m["name"])
                continue
            if batch_shortage:
                batch_text = "；".join(b for b, q in plan if q > 0)
                if batch_text:
                    batch_text += "；"
                batch_text += "批次可用不足 %g kg" % batch_shortage
                put(ws, "B%d" % cursor, m["name"])
                put(ws, "E%d" % cursor, batch_text)
                put(ws, "H%d" % cursor, need_kg)
                put(ws, "J%d" % cursor, round(need_kg * 1000, 1))
                cursor += 1
            elif len(plan) <= 1 or cursor + len(plan) - 1 > hi:
                # 单行：E 列只写批号编码（多个用「；」分隔）；用量只体现在 H/J 列
                if len(plan) > 1:
                    bno = "；".join(b for b, q in plan)
                else:
                    bno = plan[0][0] if plan else manual
                put(ws, "B%d" % cursor, m["name"])
                put(ws, "E%d" % cursor, bno)
                put(ws, "H%d" % cursor, need_kg)
                put(ws, "J%d" % cursor, round(need_kg * 1000, 1))
                cursor += 1
            else:
                # 多批次领料：每批一行（名称重复写，车间按批领）
                for b, q in plan:
                    put(ws, "B%d" % cursor, m["name"])
                    put(ws, "E%d" % cursor, b)
                    put(ws, "H%d" % cursor, q)
                    put(ws, "J%d" % cursor, round(q * 1000, 1))
                    cursor += 1
        if si in SEG:
            cursors[si] = cursor
        else:
            tail_cur = cursor
        notes[si] = mix_note(stage, mats)

    # 清掉模板预置但本单未用到的混装阶段标签
    if 1 not in groups:
        put(ws, "A14", "")
        put(ws, "G14", "")
    if 2 not in groups and 3 not in groups:
        put(ws, "A17", "")
        put(ws, "G17", "")

    put(ws, "G6", notes.get(0, ""))
    put(ws, "G14", notes.get(1, ""))
    tail_g = "；".join(notes[k] for k in (2, 3) if k in notes)
    if tail_g:
        put(ws, "G17", tail_g)
    put(ws, "E25", notes.get(0, ""))
    put(ws, "E26", notes.get(1, ""))
    tail = "；".join(notes[k] for k in (2, 3) if k in notes)
    if tail:
        put(ws, "E27", tail)

    if overflow:
        return {"ok": False,
                "msg": "物料超过模板行区上限，未生成文件；请调整阶段或减少本阶段物料/批次行：%s"
                       % "、".join(overflow),
                "overflow": overflow}

    # ---- 压片段 ----
    put(ws, "C38", "规格要求：%gg/片\n包衣要求：%s     形状：%s"
        % (tablet, pk.get("coat", "无"), pk.get("shape", "圆形")))
    put(ws, "C39", " 片重范围：%g–%gg/片\n硬度：__________ N　厚度：__________ mm；\n"
                   "首件实测：片重 ______g 硬度 ______N 厚度 ______mm"
        % (tablet, round(tablet + 0.05, 2)))
    put_img(ws, "H35", find_img(product, "压片图"))

    # ---- 包衣工艺 ----
    coating_steps = cfg.get("coating_steps") or {}
    if isinstance(coating_steps, (list, tuple)):
        coating_steps = dict(zip(
            ("包衣液配制", "素片预热", "包衣液雾化", "包衣设备清洗"),
            coating_steps))
    for row, label in zip((54, 55, 56, 57),
                          ("包衣液配制", "素片预热", "包衣液雾化", "包衣设备清洗")):
        value = coating_steps.get(label, "")
        put(ws, "E%d" % row, str(value).strip() or None)
        alignment = copy(ws["E%d" % row].alignment)
        alignment.wrap_text = True
        ws["E%d" % row].alignment = alignment

    pm = pk.get("materials", [])
    # ---- 包装段 ----
    row = 69
    for m in pm[:6]:
        sets = int(m.get("qty", 1)) if pack_direct else n_bottles * int(m.get("qty", 1))
        bno = bno_map.get(m["name"], m.get("batch", ""))
        if not bno:
            bno = default_batch_of(lookup_row(srows, m["name"]))
        put(ws, "D%d" % row, m["name"])
        put(ws, "E%d" % row, bno)
        put(ws, "G%d" % row, "%d 套" % sets)
        row += 1
    put(ws, "D75", "重量：%gg/ 片　每瓶 %d 片　每箱 %d 瓶" % (tablet, per_bottle, per_box))

    # ---- 图片 ----
    put_img(ws, "C64", find_img(product, "物料1"))
    put_img(ws, "E64", find_img(product, "物料2"))
    put_img(ws, "G64", find_img(product, "物料3"))
    put_img(ws, "C66", find_img(product, "物料4"))
    put_img(ws, "E66", find_img(product, "物料5"))
    put_img(ws, "G66", find_img(product, "物料6"))
    put_img(ws, "J64", find_img(product, "成品图"), 100, 75)

    os.makedirs(folder, exist_ok=True)
    fd, tmp_out = tempfile.mkstemp(prefix=".production-", suffix=".xlsx", dir=folder)
    os.close(fd)
    try:
        wb.save(tmp_out)
        try:
            os.rename(tmp_out, out)
        except OSError:
            if os.path.exists(out):
                return {"ok": False, "msg": "生产指令单已存在，未覆盖：%s。请检查生产序号后再生成。" % out}
            raise
    finally:
        if os.path.exists(tmp_out):
            os.remove(tmp_out)

    mats_out = []
    for m in norm_mats:
        need_kg = round(actual_kg * m["ratio"] / 100, 3)
        srow = lookup_row(srows, m["name"])
        manual = bno_map.get(m["name"], m.get("batch", ""))
        plan = plan_alloc(srow, need_kg, manual)
        batch_shortage = 0.0
        if srow and srow.get("batches") and not manual:
            batch_shortage = round(max(0.0, need_kg - sum(q for _, q in plan if q > 0)), 3)
        batch = "；".join(b for b, q in plan if q > 0)
        if not batch and len(plan) == 1:
            batch = plan[0][0]
        if batch_shortage:
            batch += "%s批次可用不足 %g kg" % ("；" if batch else "", batch_shortage)
        mats_out.append({"name": m["name"], "ratio": m["ratio"], "kg": need_kg,
                         "batch": batch, "batch_shortage": batch_shortage,
                         "multi": len(plan) > 1,
                         "avail": round(srow.get("avail", 0), 3) if srow else 0})
    pack_out = [{"name": m["name"],
                 "sets": (int(m.get("qty", 1)) if pack_direct else n_bottles * int(m.get("qty", 1)))}
                for m in pm]

    return {"ok": True, "file": out, "folder": folder, "serial": serial,
            "product": product, "tablet_g": tablet,
            "plan_wan": wan, "total_kg": round(total_kg, 3),
            "actual_kg": round(actual_kg, 3), "actual_wan": round(actual_wan, 3),
            "limited": limiting,
            "n_bottles": n_bottles, "n_boxes": n_boxes,
            "materials": mats_out, "packaging": pack_out,
            "overflow": overflow,
            "stock_date": stock_rows.meta.get("stock_date", "")}


def main():
    if len(sys.argv) < 2:
        print(json.dumps({"ok": False, "msg": "缺少输入文件"}, ensure_ascii=False))
        return
    with open(sys.argv[1], "r", encoding="utf-8") as f:
        cfg = json.load(f)
    out_root = sys.argv[2] if len(sys.argv) > 2 else None
    try:
        res = generate(cfg, out_root)
    except Exception as e:
        res = {"ok": False, "msg": "%s: %s" % (type(e).__name__, e)}
    txt = json.dumps(res, ensure_ascii=False)
    try:
        tmp = os.environ.get("TEMP", r"C:\Windows\Temp")
        with open(os.path.join(tmp, "produce_result.json"), "w", encoding="utf-8") as f:
            f.write(txt)
        if res.get("ok"):
            lines = ["OK=1", "FILE=" + res["file"], "FOLDER=" + res["folder"],
                     "SERIAL=" + res["serial"], "PRODUCT=" + res["product"],
                     "TABLET_G=%s" % res["tablet_g"], "PLAN_WAN=%s" % res["plan_wan"],
                     "TOTAL_KG=%s" % res["total_kg"], "ACTUAL_KG=%s" % res["actual_kg"],
                     "ACTUAL_WAN=%s" % res["actual_wan"],
                     "BOTTLES=%d" % res["n_bottles"], "BOXES=%d" % res["n_boxes"],
                     "LIMITED=" + (res.get("limited") or ""),
                     "MATS=" + "|".join("%s:%s" % (m["name"], m["kg"]) for m in res["materials"]),
                     "PACKS=" + "|".join("%s:%d套" % (m["name"], m["sets"]) for m in res["packaging"])]
        else:
            lines = ["OK=0", "MSG=%s" % res.get("msg", "未知错误")]
        with open(os.path.join(tmp, "produce_result.txt"), "w", encoding="utf-8") as f:
            f.write("\n".join(lines))
    except Exception:
        pass
    print(txt)


if __name__ == "__main__":
    main()
