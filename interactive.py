# -*- coding: utf-8 -*-
"""生产指令单生成器 - 交互命令行版
双击桌面「生产指令单生成器」即可运行，按提示输入即可出单。
"""
import os
import sys
import json
import ctypes

APP = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, APP)
from generator import generate, load_recipes, MANAGERS  # noqa: E402


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
    "grn": "\033[92m", "red": "\033[91m", "cyn": "\033[96m",
}


def ask(prompt, default=None):
    tip = prompt
    if default is not None:
        tip += "（回车用 %s）" % default
    v = input(tip + "：").strip()
    if v == "" and default is not None:
        return default
    return v


def load_stock_items():
    p = os.path.join(APP, "stock.json")
    if not os.path.exists(p):
        return []
    try:
        return json.load(open(p, encoding="utf-8")).get("items", [])
    except Exception:
        return []


def main():
    recipes = load_recipes()
    names = [k for k in recipes if not k.startswith("_")]

    print("=" * 56)
    print(" 生产指令单生成器")
    print("=" * 56)
    print(" 使用方法：每次输入后按回车；括号内是默认值，")
    print("            直接回车就使用默认值。")
    print("=" * 56)
    print()
    print("产品列表：")
    for i, n in enumerate(names, start=1):
        print("   %d  %s" % (i, n))
    print()

    idx = ask("请输入产品序号（1-%d）" % len(names))
    try:
        idx = int(idx)
        product = names[idx - 1]
    except Exception:
        print("输入无效：%s" % idx)
        input("按回车退出")
        return

    r = recipes[product]
    tablet = ask("片重 g/片", r.get("tablet_g", 0.5))
    try:
        tablet = float(tablet)
    except Exception:
        tablet = float(r.get("tablet_g", 0.5))

    mode = ask("产量输入方式：1=按片数（万片） 2=按总量（kg）", "1")
    total_kg = None
    if mode == "2":
        total_kg = ask("总量 kg")
        try:
            total_kg = float(total_kg)
        except Exception:
            print("总量必须是数字")
            input("按回车退出")
            return
        wan = round(total_kg * 1000 / tablet / 10000, 4)
        print("  折算：%g kg ≈ %g 万片" % (total_kg, wan))
    else:
        wan = ask("计划产量（万片）")
        try:
            wan = float(wan)
        except Exception:
            print("计划产量必须是数字")
            input("按回车退出")
            return
        if wan <= 0:
            print("计划产量必须大于 0")
            input("按回车退出")
            return

    bottles = ask("瓶数（回车 = 按每瓶片数自动算）", "")
    due = ask("交货日期 8 位（可空）", "")

    print("销售经理：")
    for i, m in enumerate(MANAGERS, start=1):
        print("   %d  %s" % (i, m))
    mi = ask("请输入销售经理序号（1-%d）" % len(MANAGERS), "1")
    try:
        manager = MANAGERS[int(mi) - 1]
    except Exception:
        manager = MANAGERS[0]

    date8 = ask("生产日期（8位）", __import__("datetime").datetime.now().strftime("%Y%m%d"))
    seq = ask("当天第几单", "01")
    maker = ask("制表人", "李浩冉")
    batch = ask("生产批次", "01")

    print()
    print("%s物料批号确认%s（回车 = 用配方库记录；本次进货批号不同就直接输入）" % (C["b"], C["r"]))
    batch_override = {}
    for m in r["materials"]:
        v = ask("  %s 批号（默认 %s）" % (m["name"], m.get("batch") or "-"), "")
        if v:
            batch_override[m["name"]] = v

    pk0 = r.get("packaging", {})
    print()
    coat = ask("包衣要求（回车 = %s）" % pk0.get("coat", "无"), "")
    shape = ask("形状（回车 = %s）" % pk0.get("shape", "圆形"), "")

    print()
    print("%s包材录入%s（本次订单实际用的瓶 / 盖 / 标签 / 干燥剂…）" % (C["b"], C["r"]))
    print("%s  每输完名称会问数量；直接回车 = 沿用配方库默认；输 k 可查库存包材列表%s"
          % (C["dim"], C["r"]))
    stock_items = load_stock_items()
    pack = []
    while True:
        pn = input("  包材名称：").strip()
        if not pn:
            if not pack and pk0.get("materials"):
                print("  %s未输入 → 沿用配方库默认包材%s" % (C["dim"], C["r"]))
            break
        if pn.lower() == "k":
            packs = [i for i in stock_items if i["cat"].startswith("原材料-包装材料")]
            for i, it in enumerate(packs, 1):
                print("    %d  %-24s 库存 %s %s" % (i, it["name"], it["qty"], it["unit"]))
            continue
        q = input("  数量（套，回车 = 1）：").strip()
        pack.append({"name": pn, "qty": int(q) if q.isdigit() else 1})
        print("  %s已录入：%s × %s%s" % (C["grn"], pack[-1]["name"], pack[-1]["qty"], C["r"]))

    print()
    print("库存核对（选填，直接回车跳过）")
    stock = {}
    for m in r["materials"]:
        v = ask("  %s 可用库存 kg" % m["name"], "")
        if v:
            try:
                stock[m["name"]] = float(v)
            except Exception:
                pass

    cfg = {
        "product": product,
        "tablet_g": float(tablet),
        "wan": float(wan),
        "total_kg": total_kg,
        "bottles": bottles,
        "due": due,
        "date": date8,
        "seq": int(seq),
        "manager": manager,
        "maker": maker,
        "batch": batch,
        "stock": stock,
        "pack_materials": pack,
        "coat": coat,
        "shape": shape,
        "batch_override": batch_override,
    }

    print()
    print("生成中...")
    res = generate(cfg)

    print()
    if not res.get("ok"):
        print("生成失败：%s" % res.get("msg"))
    else:
        print("生成成功")
        print("单号：%s   产品：%s" % (res["serial"], res["product"]))
        print("文件：%s" % res["file"])
        print()
        print("投料量：")
        for m in res["materials"]:
            print("   %-14s %10s kg" % (m["name"], m["kg"]))
        print()
        print("理论总重 %s kg    实际投料 %s kg" % (res["total_kg"], res["actual_kg"]))
        print("包材：%d 瓶（%d 箱）" % (res["n_bottles"], res["n_boxes"]))
        for p in res.get("packaging", []):
            print("   %-24s %d 套" % (p["name"], p["sets"]))
        if res.get("limited"):
            print()
            print("注意：受限物料「%s」，库存最多可做 %s 万片（计划 %s 万片）"
                  % (res["limited"], res["actual_wan"], res["plan_wan"]))
        print()
        print("提示：实际产量请在生成单子的「实际生产数量」行填写。")
        try:
            from orders import add_order
            ok, msg = add_order({
                "serial": res["serial"], "product": res["product"],
                "wan": res["plan_wan"], "tablet": float(tablet),
                "total_kg": total_kg,
                "customer": "", "due": due, "sales": manager,
                "status": "待生产", "order_date": date8, "batch": batch,
                "materials": res.get("materials", []),
                "packaging": res.get("packaging", []),
            })
            print("台账：%s" % msg)
        except Exception as e:
            print("台账记录失败（不影响单据）：%s" % e)
    print()
    input("按回车退出")


if __name__ == "__main__":
    main()
